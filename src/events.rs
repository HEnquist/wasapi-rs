use std::slice;
use windows::{
    core::{implement, Result, GUID, PCWSTR},
    Win32::Media::Audio::{
        AudioSessionDisconnectReason, AudioSessionState, AudioSessionStateActive,
        AudioSessionStateExpired, AudioSessionStateInactive, DisconnectReasonDeviceRemoval,
        DisconnectReasonExclusiveModeOverride, DisconnectReasonFormatChanged,
        DisconnectReasonServerShutdown, DisconnectReasonSessionDisconnected,
        DisconnectReasonSessionLogoff, IAudioSessionEvents, IAudioSessionEvents_Impl,
    },
};

use crate::SessionState;

type OptionBox<T> = Option<Box<T>>;

/// A structure holding the callbacks for notifications
pub struct EventCallbacks {
    simple_volume: OptionBox<dyn Fn(f32, bool, GUID) + Send + Sync>,
    channel_volume: OptionBox<dyn Fn(usize, f32, GUID) + Send + Sync>,
    state: OptionBox<dyn Fn(SessionState) + Send + Sync>,
    disconnected: OptionBox<dyn Fn(DisconnectReason) + Send + Sync>,
    iconpath: OptionBox<dyn Fn(String, GUID) + Send + Sync>,
    displayname: OptionBox<dyn Fn(String, GUID) + Send + Sync>,
    groupingparam: OptionBox<dyn Fn(GUID, GUID) + Send + Sync>,
}

impl Default for EventCallbacks {
    fn default() -> Self {
        Self::new()
    }
}

impl EventCallbacks {
    /// Create a new EventCallbacks with no callbacks set
    pub fn new() -> Self {
        Self {
            simple_volume: None,
            channel_volume: None,
            state: None,
            disconnected: None,
            iconpath: None,
            displayname: None,
            groupingparam: None,
        }
    }

    /// Set a callback for OnSimpleVolumeChanged notifications
    pub fn set_simple_volume_callback(
        &mut self,
        c: impl Fn(f32, bool, GUID) + 'static + Sync + Send,
    ) {
        self.simple_volume = Some(Box::new(c));
    }
    /// Remove a callback for OnSimpleVolumeChanged notifications
    pub fn unset_simple_volume_callback(&mut self) {
        self.simple_volume = None;
    }

    /// Set a callback for OnChannelVolumeChanged notifications
    pub fn set_channel_volume_callback(
        &mut self,
        c: impl Fn(usize, f32, GUID) + 'static + Sync + Send,
    ) {
        self.channel_volume = Some(Box::new(c));
    }
    /// Remove a callback for OnChannelVolumeChanged notifications
    pub fn unset_channel_volume_callback(&mut self) {
        self.channel_volume = None;
    }

    /// Set a callback for OnSessionDisconnected notifications
    pub fn set_disconnected_callback(
        &mut self,
        c: impl Fn(DisconnectReason) + 'static + Sync + Send,
    ) {
        self.disconnected = Some(Box::new(c));
    }
    /// Remove a callback for OnSessionDisconnected notifications
    pub fn unset_disconnected_callback(&mut self) {
        self.disconnected = None;
    }

    /// Set a callback for OnStateChanged notifications
    pub fn set_state_callback(&mut self, c: impl Fn(SessionState) + 'static + Sync + Send) {
        self.state = Some(Box::new(c));
    }
    /// Remove a callback for OnStateChanged notifications
    pub fn unset_state_callback(&mut self) {
        self.state = None;
    }

    /// Set a callback for OnIconPathChanged notifications
    pub fn set_iconpath_callback(&mut self, c: impl Fn(String, GUID) + 'static + Sync + Send) {
        self.iconpath = Some(Box::new(c));
    }
    /// Remove a callback for OnIconPathChanged notifications
    pub fn unset_iconpath_callback(&mut self) {
        self.iconpath = None;
    }

    /// Set a callback for OnDisplayNameChanged notifications
    pub fn set_displayname_callback(&mut self, c: impl Fn(String, GUID) + 'static + Sync + Send) {
        self.displayname = Some(Box::new(c));
    }
    /// Remove a callback for OnDisplayNameChanged notifications
    pub fn unset_displayname_callback(&mut self) {
        self.displayname = None;
    }

    /// Set a callback for OnGroupingParamChanged notifications
    pub fn set_groupingparam_callback(&mut self, c: impl Fn(GUID, GUID) + 'static + Sync + Send) {
        self.groupingparam = Some(Box::new(c));
    }
    /// Remove a callback for OnGroupingParamChanged notifications
    pub fn unset_groupingparam_callback(&mut self) {
        self.groupingparam = None;
    }
}

/// Reason for session disconnect, an enum representing the `DisconnectReasonXxx` values of the
/// [AudioSessionDisconnectReason enum](https://learn.microsoft.com/en-us/windows/win32/api/audiopolicy/nf-audiopolicy-iaudiosessionevents-onsessiondisconnected)
#[derive(Debug)]
pub enum DisconnectReason {
    /// The user removed the audio endpoint device.
    DeviceRemoval,
    /// The Windows audio service has stopped.
    ServerShutdown,
    /// The stream format changed for the device that the audio session is connected to.
    FormatChanged,
    /// The user logged off the Windows Terminal Services (WTS) session that the audio session was running in.
    SessionLogoff,
    /// The WTS session that the audio session was running in was disconnected.
    SessionDisconnected,
    /// The (shared-mode) audio session was disconnected to make the audio endpoint device available for an exclusive-mode connection.
    ExclusiveModeOverride,
    /// An unknown reason was returned.
    Unknown,
}

/// Wrapper for [IAudioSessionEvents](https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nn-audiopolicy-iaudiosessionevents).
#[implement(IAudioSessionEvents)]
pub(crate) struct AudioSessionEvents {
    callbacks: EventCallbacks,
}

impl AudioSessionEvents {
    /// Create a new [AudioSessionEvents] instance, returned as a [IAudioSessionEvent].
    pub fn new(callbacks: EventCallbacks) -> Self {
        Self { callbacks }
    }
}

impl IAudioSessionEvents_Impl for AudioSessionEvents_Impl {
    fn OnStateChanged(&self, newstate: AudioSessionState) -> Result<()> {
        #[allow(non_upper_case_globals)]
        let state_name = match newstate {
            AudioSessionStateActive => "Active",
            AudioSessionStateInactive => "Inactive",
            AudioSessionStateExpired => "Expired",
            _ => "Unknown",
        };
        trace!("state change to: {state_name}");
        #[allow(non_upper_case_globals)]
        let sessionstate = match newstate {
            AudioSessionStateActive => SessionState::Active,
            AudioSessionStateInactive => SessionState::Inactive,
            AudioSessionStateExpired => SessionState::Expired,
            _ => return Ok(()),
        };
        if let Some(callback) = &self.callbacks.state {
            callback(sessionstate);
        }
        Ok(())
    }

    fn OnSessionDisconnected(&self, disconnectreason: AudioSessionDisconnectReason) -> Result<()> {
        trace!("Disconnected");
        #[allow(non_upper_case_globals)]
        let reason = match disconnectreason {
            DisconnectReasonDeviceRemoval => DisconnectReason::DeviceRemoval,
            DisconnectReasonServerShutdown => DisconnectReason::ServerShutdown,
            DisconnectReasonFormatChanged => DisconnectReason::FormatChanged,
            DisconnectReasonSessionLogoff => DisconnectReason::SessionLogoff,
            DisconnectReasonSessionDisconnected => DisconnectReason::SessionDisconnected,
            DisconnectReasonExclusiveModeOverride => DisconnectReason::ExclusiveModeOverride,
            _ => DisconnectReason::Unknown,
        };

        if let Some(callback) = &self.callbacks.disconnected {
            callback(reason);
        }
        Ok(())
    }

    fn OnDisplayNameChanged(
        &self,
        newdisplayname: &PCWSTR,
        eventcontext: *const GUID,
    ) -> Result<()> {
        let name = unsafe { newdisplayname.to_string().unwrap_or_default() };
        trace!("New display name: {name}");
        if let Some(callback) = &self.callbacks.displayname {
            let context = unsafe { *eventcontext };
            callback(name, context);
        }
        Ok(())
    }

    fn OnIconPathChanged(&self, newiconpath: &PCWSTR, eventcontext: *const GUID) -> Result<()> {
        let path = unsafe { newiconpath.to_string().unwrap_or_default() };
        trace!("New icon path: {path}");
        if let Some(callback) = &self.callbacks.iconpath {
            let context = unsafe { *eventcontext };
            callback(path, context);
        }
        Ok(())
    }

    fn OnSimpleVolumeChanged(
        &self,
        newvolume: f32,
        newmute: windows_core::BOOL,
        eventcontext: *const GUID,
    ) -> Result<()> {
        trace!("New volume: {newvolume}, mute: {newmute:?}");
        if let Some(callback) = &self.callbacks.simple_volume {
            let context = unsafe { *eventcontext };
            callback(newvolume, bool::from(newmute), context);
        }
        Ok(())
    }

    fn OnChannelVolumeChanged(
        &self,
        channelcount: u32,
        newchannelvolumearray: *const f32,
        changedchannel: u32,
        eventcontext: *const GUID,
    ) -> Result<()> {
        trace!("New channel volume for channel: {changedchannel}");
        let volslice =
            unsafe { slice::from_raw_parts(newchannelvolumearray, channelcount as usize) };
        if let Some(callback) = &self.callbacks.channel_volume {
            let context = unsafe { *eventcontext };
            if changedchannel == u32::MAX {
                // special meaning by specs: (DWORD)(-1) - "more than one channel have changed"
                // using all channels
                for (idx, newvol) in volslice.iter().enumerate() {
                    callback(idx, *newvol, context);
                }
            }
            if (changedchannel as usize) < volslice.len() {
                let newvol = volslice[changedchannel as usize];
                callback(changedchannel as usize, newvol, context);
            } else {
                warn!(
                        "OnChannelVolumeChanged: received unsupported changedchannel value {} for volume array length of {}",
                        changedchannel,
                        volslice.len()
                    );
                return Ok(());
            }
        }
        Ok(())
    }

    fn OnGroupingParamChanged(
        &self,
        newgroupingparam: *const GUID,
        eventcontext: *const GUID,
    ) -> Result<()> {
        trace!("Grouping changed");
        if let Some(callback) = &self.callbacks.groupingparam {
            let context = unsafe { *eventcontext };
            let grouping = unsafe { *newgroupingparam };
            callback(grouping, context);
        }
        Ok(())
    }
}
