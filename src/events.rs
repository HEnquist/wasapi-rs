use std::slice;
use std::string::FromUtf16Error;
use windows::{
    core::{implement, Result, GUID, PCWSTR},
    Win32::Foundation::PROPERTYKEY,
    Win32::Media::Audio::{
        AudioSessionDisconnectReason, AudioSessionState, AudioSessionStateActive,
        AudioSessionStateExpired, AudioSessionStateInactive, DisconnectReasonDeviceRemoval,
        DisconnectReasonExclusiveModeOverride, DisconnectReasonFormatChanged,
        DisconnectReasonServerShutdown, DisconnectReasonSessionDisconnected,
        DisconnectReasonSessionLogoff, EDataFlow, ERole, IAudioSessionEvents,
        IAudioSessionEvents_Impl, IMMNotificationClient, IMMNotificationClient_Impl, DEVICE_STATE,
    },
};

use crate::{DeviceState, Direction, Role, SessionState};

type OptionBox<T> = Option<Box<T>>;

/// Read a [PCWSTR] that points to a string owned by the caller.
/// Returns Ok(None) if the pointer is null, which the audio system uses
/// to signal that there is no device.
/// An unreadable string gives an error, and must not be confused
/// with the absence of a device.
fn read_pcwstr(pcwstr: &PCWSTR) -> std::result::Result<Option<String>, FromUtf16Error> {
    if pcwstr.is_null() {
        return Ok(None);
    }
    unsafe { pcwstr.to_string() }.map(Some)
}

/// Read the id of the device that a notification refers to.
/// Returns None, after logging the reason, if no usable id was provided.
fn read_device_id(pcwstr: &PCWSTR, notification: &str) -> Option<String> {
    match read_pcwstr(pcwstr) {
        Ok(Some(id)) => Some(id),
        Ok(None) => {
            warn!("{notification}: received a null device id");
            None
        }
        Err(err) => {
            warn!("{notification}: received an unreadable device id, {err}");
            None
        }
    }
}

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

/// A structure holding the callbacks for device change notifications
pub struct DeviceEventCallbacks {
    device_state: OptionBox<dyn Fn(String, DeviceState) + Send + Sync>,
    device_added: OptionBox<dyn Fn(String) + Send + Sync>,
    device_removed: OptionBox<dyn Fn(String) + Send + Sync>,
    default_device: OptionBox<dyn Fn(Direction, Role, Option<String>) + Send + Sync>,
    property_value: OptionBox<dyn Fn(String, PROPERTYKEY) + Send + Sync>,
}

impl Default for DeviceEventCallbacks {
    fn default() -> Self {
        Self::new()
    }
}

impl DeviceEventCallbacks {
    /// Create a new DeviceEventCallbacks with no callbacks set
    pub fn new() -> Self {
        Self {
            device_state: None,
            device_added: None,
            device_removed: None,
            default_device: None,
            property_value: None,
        }
    }

    /// Set a callback for OnDeviceStateChanged notifications.
    /// The parameters are the device id and the new state.
    pub fn set_device_state_callback(
        &mut self,
        c: impl Fn(String, DeviceState) + 'static + Sync + Send,
    ) {
        self.device_state = Some(Box::new(c));
    }
    /// Remove a callback for OnDeviceStateChanged notifications
    pub fn unset_device_state_callback(&mut self) {
        self.device_state = None;
    }

    /// Set a callback for OnDeviceAdded notifications.
    /// The parameter is the device id.
    pub fn set_device_added_callback(&mut self, c: impl Fn(String) + 'static + Sync + Send) {
        self.device_added = Some(Box::new(c));
    }
    /// Remove a callback for OnDeviceAdded notifications
    pub fn unset_device_added_callback(&mut self) {
        self.device_added = None;
    }

    /// Set a callback for OnDeviceRemoved notifications.
    /// The parameter is the device id.
    pub fn set_device_removed_callback(&mut self, c: impl Fn(String) + 'static + Sync + Send) {
        self.device_removed = Some(Box::new(c));
    }
    /// Remove a callback for OnDeviceRemoved notifications
    pub fn unset_device_removed_callback(&mut self) {
        self.device_removed = None;
    }

    /// Set a callback for OnDefaultDeviceChanged notifications.
    /// The parameters are the direction and role of the new default device,
    /// and its device id. The id is None when there is no longer
    /// a default device for that direction and role.
    pub fn set_default_device_callback(
        &mut self,
        c: impl Fn(Direction, Role, Option<String>) + 'static + Sync + Send,
    ) {
        self.default_device = Some(Box::new(c));
    }
    /// Remove a callback for OnDefaultDeviceChanged notifications
    pub fn unset_default_device_callback(&mut self) {
        self.default_device = None;
    }

    /// Set a callback for OnPropertyValueChanged notifications.
    /// The parameters are the device id and the key of the changed property.
    pub fn set_property_value_callback(
        &mut self,
        c: impl Fn(String, PROPERTYKEY) + 'static + Sync + Send,
    ) {
        self.property_value = Some(Box::new(c));
    }
    /// Remove a callback for OnPropertyValueChanged notifications
    pub fn unset_property_value_callback(&mut self) {
        self.property_value = None;
    }
}

/// Wrapper for [IMMNotificationClient](https://learn.microsoft.com/en-us/windows/win32/api/mmdeviceapi/nn-mmdeviceapi-immnotificationclient).
#[implement(IMMNotificationClient)]
pub(crate) struct NotificationClient {
    callbacks: DeviceEventCallbacks,
}

impl NotificationClient {
    /// Create a new [NotificationClient] instance.
    pub fn new(callbacks: DeviceEventCallbacks) -> Self {
        Self { callbacks }
    }
}

impl IMMNotificationClient_Impl for NotificationClient_Impl {
    fn OnDeviceStateChanged(&self, pwstrdeviceid: &PCWSTR, dwnewstate: DEVICE_STATE) -> Result<()> {
        let Some(id) = read_device_id(pwstrdeviceid, "OnDeviceStateChanged") else {
            return Ok(());
        };
        let state = match DeviceState::try_from(dwnewstate) {
            Ok(state) => state,
            Err(err) => {
                warn!("OnDeviceStateChanged: {err}");
                return Ok(());
            }
        };
        trace!("Device {id} changed state to: {state}");
        if let Some(callback) = &self.callbacks.device_state {
            callback(id, state);
        }
        Ok(())
    }

    fn OnDeviceAdded(&self, pwstrdeviceid: &PCWSTR) -> Result<()> {
        let Some(id) = read_device_id(pwstrdeviceid, "OnDeviceAdded") else {
            return Ok(());
        };
        trace!("Device added: {id}");
        if let Some(callback) = &self.callbacks.device_added {
            callback(id);
        }
        Ok(())
    }

    fn OnDeviceRemoved(&self, pwstrdeviceid: &PCWSTR) -> Result<()> {
        let Some(id) = read_device_id(pwstrdeviceid, "OnDeviceRemoved") else {
            return Ok(());
        };
        trace!("Device removed: {id}");
        if let Some(callback) = &self.callbacks.device_removed {
            callback(id);
        }
        Ok(())
    }

    fn OnDefaultDeviceChanged(
        &self,
        flow: EDataFlow,
        role: ERole,
        pwstrdefaultdeviceid: &PCWSTR,
    ) -> Result<()> {
        // A null id means that there is no longer a default device.
        // An unreadable id must not be reported as no device, so it is skipped.
        let id = match read_pcwstr(pwstrdefaultdeviceid) {
            Ok(id) => id,
            Err(err) => {
                warn!("OnDefaultDeviceChanged: received an unreadable device id, {err}");
                return Ok(());
            }
        };
        let direction = match Direction::try_from(flow) {
            Ok(direction) => direction,
            Err(err) => {
                warn!("OnDefaultDeviceChanged: {err}");
                return Ok(());
            }
        };
        let device_role = match Role::try_from(role) {
            Ok(role) => role,
            Err(err) => {
                warn!("OnDefaultDeviceChanged: {err}");
                return Ok(());
            }
        };
        trace!("New default {direction} device for role {device_role}: {id:?}");
        if let Some(callback) = &self.callbacks.default_device {
            callback(direction, device_role, id);
        }
        Ok(())
    }

    fn OnPropertyValueChanged(&self, pwstrdeviceid: &PCWSTR, key: &PROPERTYKEY) -> Result<()> {
        let Some(id) = read_device_id(pwstrdeviceid, "OnPropertyValueChanged") else {
            return Ok(());
        };
        trace!("Property {key:?} changed for device {id}");
        if let Some(callback) = &self.callbacks.property_value {
            callback(id, *key);
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{Arc, Mutex};
    use windows::core::HSTRING;
    use windows::Win32::Media::Audio::{
        eAll, eCapture, eConsole, eRender, DEVICE_STATE_ACTIVE, DEVICE_STATE_UNPLUGGED,
    };

    const TEST_ID: &str = "{0.0.0.00000000}.{6e6f7420-6120-7265-616c-206465766963}";

    /// Build a client that appends a description of every notification to the returned vector.
    fn logging_client() -> (IMMNotificationClient, Arc<Mutex<Vec<String>>>) {
        let log: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
        let mut callbacks = DeviceEventCallbacks::new();

        let added = log.clone();
        callbacks.set_device_added_callback(move |id| added.lock().unwrap().push(format!("+{id}")));
        let removed = log.clone();
        callbacks
            .set_device_removed_callback(move |id| removed.lock().unwrap().push(format!("-{id}")));
        let state = log.clone();
        callbacks.set_device_state_callback(move |id, newstate| {
            state.lock().unwrap().push(format!("{id} is {newstate}"))
        });
        let default = log.clone();
        callbacks.set_default_device_callback(move |direction, role, id| {
            default
                .lock()
                .unwrap()
                .push(format!("default {direction} {role} {id:?}"))
        });
        let property = log.clone();
        callbacks.set_property_value_callback(move |id, key| {
            property.lock().unwrap().push(format!("{id} {}", key.pid))
        });

        (NotificationClient::new(callbacks).into(), log)
    }

    #[test]
    fn notifications_reach_the_callbacks() {
        let (client, log) = logging_client();
        let id = HSTRING::from(TEST_ID);
        let id = PCWSTR::from_raw(id.as_ptr());
        let key = PROPERTYKEY {
            fmtid: GUID::zeroed(),
            pid: 14,
        };

        unsafe {
            client.OnDeviceAdded(id).unwrap();
            client.OnDeviceRemoved(id).unwrap();
            client
                .OnDeviceStateChanged(id, DEVICE_STATE_UNPLUGGED)
                .unwrap();
            client
                .OnDefaultDeviceChanged(eCapture, eConsole, id)
                .unwrap();
            client.OnPropertyValueChanged(id, key).unwrap();
        }

        assert_eq!(
            *log.lock().unwrap(),
            vec![
                format!("+{TEST_ID}"),
                format!("-{TEST_ID}"),
                format!("{TEST_ID} is Unplugged"),
                format!("default Capture Console Some(\"{TEST_ID}\")"),
                format!("{TEST_ID} 14"),
            ]
        );
    }

    /// The audio system passes a null pointer when the last default device disappears.
    #[test]
    fn a_null_device_id_is_handled() {
        let (client, log) = logging_client();

        unsafe {
            client
                .OnDefaultDeviceChanged(eRender, eConsole, PCWSTR::null())
                .unwrap();
            // The remaining notifications always carry an id, but must not
            // dereference a null pointer if one arrives anyway.
            client.OnDeviceAdded(PCWSTR::null()).unwrap();
            client.OnDeviceRemoved(PCWSTR::null()).unwrap();
        }

        assert_eq!(*log.lock().unwrap(), vec!["default Render Console None"]);
    }

    /// An id that cannot be read is not the same thing as a missing device,
    /// and must not be reported as one.
    #[test]
    fn an_unreadable_device_id_is_skipped() {
        let (client, log) = logging_client();
        // A lone surrogate is not valid UTF-16.
        let invalid = [0xd800u16, 0];
        let id = PCWSTR::from_raw(invalid.as_ptr());

        unsafe {
            client
                .OnDefaultDeviceChanged(eRender, eConsole, id)
                .unwrap();
            client.OnDeviceAdded(id).unwrap();
        }

        assert!(log.lock().unwrap().is_empty());
    }

    /// Values that have no counterpart in the wrapper enums are skipped, not passed on.
    #[test]
    fn unsupported_values_are_skipped() {
        let (client, log) = logging_client();
        let id = HSTRING::from(TEST_ID);
        let id = PCWSTR::from_raw(id.as_ptr());

        unsafe {
            client.OnDefaultDeviceChanged(eAll, eConsole, id).unwrap();
            client.OnDeviceStateChanged(id, DEVICE_STATE(0)).unwrap();
        }

        assert!(log.lock().unwrap().is_empty());
    }

    /// A client with no callbacks set should accept every notification.
    #[test]
    fn notifications_without_callbacks() {
        let client: IMMNotificationClient =
            NotificationClient::new(DeviceEventCallbacks::new()).into();
        let id = HSTRING::from(TEST_ID);
        let id = PCWSTR::from_raw(id.as_ptr());

        unsafe {
            client.OnDeviceAdded(id).unwrap();
            client.OnDeviceRemoved(id).unwrap();
            client
                .OnDeviceStateChanged(id, DEVICE_STATE_ACTIVE)
                .unwrap();
            client
                .OnDefaultDeviceChanged(eRender, eConsole, id)
                .unwrap();
            client
                .OnPropertyValueChanged(
                    id,
                    PROPERTYKEY {
                        fmtid: GUID::zeroed(),
                        pid: 0,
                    },
                )
                .unwrap();
        }
    }
}
