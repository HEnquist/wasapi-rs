use crate::Windows;
use crate::Windows::{
    Win32::Foundation::{BOOL, PWSTR, S_OK},
    Win32::Media::Audio::CoreAudio::{
        AudioSessionDisconnectReason, AudioSessionState, AudioSessionStateActive,
        AudioSessionStateExpired, AudioSessionStateInactive, DisconnectReasonDeviceRemoval,
        DisconnectReasonExclusiveModeOverride, DisconnectReasonFormatChanged,
        DisconnectReasonServerShutdown, DisconnectReasonSessionDisconnected,
        DisconnectReasonSessionLogoff,
    },
};
use std::rc::Weak;
use std::slice;
use widestring::U16CString;
use windows::runtime::GUID;
use windows::runtime::HRESULT;
use windows_macros::implement;

use crate::SessionState;

type OptionBox<T> = Option<Box<T>>;

/// A structure holding the callbacks for notifications
pub struct EventCallbacks {
    simple_volume: OptionBox<dyn Fn(f32, bool, GUID)>,
    channel_volume: OptionBox<dyn Fn(usize, f32, GUID)>,
    state: OptionBox<dyn Fn(SessionState)>,
    disconnected: OptionBox<dyn Fn(DisconnectReason)>,
    iconpath: OptionBox<dyn Fn(String, GUID)>,
    displayname: OptionBox<dyn Fn(String, GUID)>,
    groupingparam: OptionBox<dyn Fn(GUID, GUID)>,
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
    pub fn set_simple_volume_callback(&mut self, c: impl Fn(f32, bool, GUID) + 'static) {
        self.simple_volume = Some(Box::new(c));
    }
    /// Remove a callback for OnSimpleVolumeChanged notifications
    pub fn unset_simple_volume_callback(&mut self) {
        self.simple_volume = None;
    }

    /// Set a callback for OnChannelVolumeChanged notifications
    pub fn set_channel_volume_callback(&mut self, c: impl Fn(usize, f32, GUID) + 'static) {
        self.channel_volume = Some(Box::new(c));
    }
    /// Remove a callback for OnChannelVolumeChanged notifications
    pub fn unset_channel_volume_callback(&mut self) {
        self.channel_volume = None;
    }

    /// Set a callback for OnSessionDisconnected notifications
    pub fn set_disconnected_callback(&mut self, c: impl Fn(DisconnectReason) + 'static) {
        self.disconnected = Some(Box::new(c));
    }
    /// Remove a callback for OnSessionDisconnected notifications
    pub fn unset_disconnected_callback(&mut self) {
        self.disconnected = None;
    }

    /// Set a callback for OnStateChanged notifications
    pub fn set_state_callback(&mut self, c: impl Fn(SessionState) + 'static) {
        self.state = Some(Box::new(c));
    }
    /// Remove a callback for OnStateChanged notifications
    pub fn unset_state_callback(&mut self) {
        self.state = None;
    }

    /// Set a callback for OnIconPathChanged notifications
    pub fn set_iconpath_callback(&mut self, c: impl Fn(String, GUID) + 'static) {
        self.iconpath = Some(Box::new(c));
    }
    /// Remove a callback for OnIconPathChanged notifications
    pub fn unset_iconpath_callback(&mut self) {
        self.iconpath = None;
    }

    /// Set a callback for OnDisplayNameChanged notifications
    pub fn set_displayname_callback(&mut self, c: impl Fn(String, GUID) + 'static) {
        self.displayname = Some(Box::new(c));
    }
    /// Remove a callback for OnDisplayNameChanged notifications
    pub fn unset_displayname_callback(&mut self) {
        self.displayname = None;
    }

    /// Set a callback for OnGroupingParamChanged notifications
    pub fn set_groupingparam_callback(&mut self, c: impl Fn(GUID, GUID) + 'static) {
        self.groupingparam = Some(Box::new(c));
    }
    /// Remove a callback for OnGroupingParamChanged notifications
    pub fn unset_groupingparam_callback(&mut self) {
        self.groupingparam = None;
    }
}

/// Reason for session disconnect
#[derive(Debug)]
pub enum DisconnectReason {
    DeviceRemoval,
    ServerShutdown,
    FormatChanged,
    SessionLogoff,
    SessionDisconnected,
    ExclusiveModeOverride,
    Unknown,
}

/// Wrapper for IAudioSessionEvents
#[implement(Windows::Win32::Media::Audio::CoreAudio::IAudioSessionEvents)]
pub(crate) struct AudioSessionEvents {
    callbacks: Weak<EventCallbacks>,
}

#[allow(non_snake_case)]
impl AudioSessionEvents {
    /// Create a new AudioSessionEvents instance, returned as a IAudioSessionEvent.
    #[allow(clippy::new_ret_no_self)]
    pub fn new(callbacks: Weak<EventCallbacks>) -> Self {
        Self { callbacks }
    }

    fn OnStateChanged(&mut self, newstate: AudioSessionState) -> HRESULT {
        #[allow(non_upper_case_globals)]
        let state_name = match newstate {
            AudioSessionStateActive => "Active",
            AudioSessionStateInactive => "Inactive",
            AudioSessionStateExpired => "Expired",
            _ => "Unknown",
        };
        trace!("state change to: {}", state_name);
        #[allow(non_upper_case_globals)]
        let sessionstate = match newstate {
            AudioSessionStateActive => SessionState::Active,
            AudioSessionStateInactive => SessionState::Inactive,
            AudioSessionStateExpired => SessionState::Expired,
            _ => return S_OK,
        };
        if let Some(callbacks) = &mut self.callbacks.upgrade() {
            if let Some(callback) = &callbacks.state {
                callback(sessionstate);
            }
        }
        S_OK
    }

    fn OnSessionDisconnected(&mut self, disconnectreason: AudioSessionDisconnectReason) -> HRESULT {
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

        if let Some(callbacks) = &mut self.callbacks.upgrade() {
            if let Some(callback) = &callbacks.disconnected {
                callback(reason);
            }
        }
        S_OK
    }

    fn OnDisplayNameChanged(
        &mut self,
        newdisplayname: PWSTR,
        eventcontext: *const GUID,
    ) -> HRESULT {
        let wide_name = unsafe { U16CString::from_ptr_str(newdisplayname.0) };
        let name = wide_name.to_string_lossy();
        trace!("New display name: {}", name);
        if let Some(callbacks) = &mut self.callbacks.upgrade() {
            if let Some(callback) = &callbacks.displayname {
                let context = unsafe { *eventcontext };
                callback(name, context);
            }
        }
        S_OK
    }

    fn OnIconPathChanged(&mut self, newiconpath: PWSTR, eventcontext: *const GUID) -> HRESULT {
        let wide_path = unsafe { U16CString::from_ptr_str(newiconpath.0) };
        let path = wide_path.to_string_lossy();
        trace!("New icon path: {}", path);
        if let Some(callbacks) = &mut self.callbacks.upgrade() {
            if let Some(callback) = &callbacks.iconpath {
                let context = unsafe { *eventcontext };
                callback(path, context);
            }
        }
        S_OK
    }

    fn OnSimpleVolumeChanged(
        &mut self,
        newvolume: f32,
        newmute: BOOL,
        eventcontext: *const GUID,
    ) -> HRESULT {
        trace!("New volume: {}, mute: {:?}", newvolume, newmute);
        if let Some(callbacks) = &mut self.callbacks.upgrade() {
            if let Some(callback) = &callbacks.simple_volume {
                let context = unsafe { *eventcontext };
                callback(newvolume, bool::from(newmute), context);
            }
        }
        S_OK
    }

    fn OnChannelVolumeChanged(
        &mut self,
        channelcount: u32,
        newchannelvolumearray: *const f32,
        changedchannel: u32,
        eventcontext: *const GUID,
    ) -> HRESULT {
        trace!("New channel volume for channel: {}", changedchannel);
        let volslice =
            unsafe { slice::from_raw_parts(newchannelvolumearray, channelcount as usize) };
        let newvol = volslice[changedchannel as usize];
        if let Some(callbacks) = &mut self.callbacks.upgrade() {
            if let Some(callback) = &callbacks.channel_volume {
                let context = unsafe { *eventcontext };
                callback(changedchannel as usize, newvol, context);
            }
        }
        S_OK
    }

    fn OnGroupingParamChanged(
        &mut self,
        newgroupingparam: *const GUID,
        eventcontext: *const GUID,
    ) -> HRESULT {
        trace!("Grouping changed");
        if let Some(callbacks) = &mut self.callbacks.upgrade() {
            if let Some(callback) = &callbacks.groupingparam {
                let context = unsafe { *eventcontext };
                let grouping = unsafe { *newgroupingparam };
                callback(grouping, context);
            }
        }
        S_OK
    }
}
