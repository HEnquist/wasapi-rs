use num_integer::Integer;
use std::cmp;
use std::collections::VecDeque;
use std::mem::{size_of, ManuallyDrop};
use std::ops::Deref;
use std::pin::Pin;
use std::sync::{Arc, Condvar, Mutex};
use std::{fmt, ptr, slice};
use widestring::U16CString;
use windows::Win32::Foundation::{E_INVALIDARG, E_NOINTERFACE, PROPERTYKEY};
use windows::Win32::Media::Audio::{
    ActivateAudioInterfaceAsync, AudioClientProperties, EDataFlow, ERole,
    IAcousticEchoCancellationControl, IActivateAudioInterfaceAsyncOperation,
    IActivateAudioInterfaceCompletionHandler, IActivateAudioInterfaceCompletionHandler_Impl,
    IAudioClient2, IAudioEffectsManager, IMMEndpoint, PKEY_AudioEngine_DeviceFormat,
    AUDCLNT_STREAMOPTIONS, AUDIOCLIENT_ACTIVATION_PARAMS, AUDIOCLIENT_ACTIVATION_PARAMS_0,
    AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK, AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS,
    AUDIO_EFFECT, AUDIO_STREAM_CATEGORY, PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE,
    PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE, VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
};
use windows::Win32::Media::KernelStreaming::AUDIO_EFFECT_TYPE_ACOUSTIC_ECHO_CANCELLATION;
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::Com::StructuredStorage::PropVariantClear;
use windows::Win32::System::Variant::VT_BLOB;
use windows::{
    core::{HRESULT, PCSTR},
    Win32::Devices::FunctionDiscovery::{
        PKEY_DeviceInterface_FriendlyName, PKEY_Device_DeviceDesc, PKEY_Device_FriendlyName,
    },
    Win32::Foundation::{HANDLE, WAIT_OBJECT_0},
    Win32::Media::Audio::{
        eCapture, eCommunications, eConsole, eMultimedia, eRender, AudioSessionStateActive,
        AudioSessionStateExpired, AudioSessionStateInactive, IAudioCaptureClient, IAudioClient,
        IAudioClock, IAudioRenderClient, IAudioSessionControl, IAudioSessionEvents, IMMDevice,
        IMMDeviceCollection, IMMDeviceEnumerator, MMDeviceEnumerator,
        AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY, AUDCLNT_BUFFERFLAGS_SILENT,
        AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR, AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM, AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        AUDCLNT_STREAMFLAGS_LOOPBACK, AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY, DEVICE_STATE_ACTIVE,
        DEVICE_STATE_DISABLED, DEVICE_STATE_NOTPRESENT, DEVICE_STATE_UNPLUGGED, WAVEFORMATEX,
        WAVEFORMATEXTENSIBLE,
    },
    Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE,
    Win32::System::Com::StructuredStorage::{
        PropVariantToStringAlloc, PROPVARIANT, PROPVARIANT_0, PROPVARIANT_0_0, PROPVARIANT_0_0_0,
    },
    Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
        COINIT_MULTITHREADED,
    },
    Win32::System::Com::{BLOB, STGM_READ},
    Win32::System::Threading::{CreateEventA, WaitForSingleObject},
};
use windows_core::{implement, IUnknown, Interface, Ref, HSTRING, PCWSTR};

use crate::{make_channelmasks, AudioSessionEvents, EventCallbacks, WasapiError, WaveFormat};

pub(crate) type WasapiRes<T> = Result<T, WasapiError>;

/// Initializes COM for use by the calling thread for the multi-threaded apartment (MTA).
pub fn initialize_mta() -> HRESULT {
    unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) }
}

/// Initializes COM for use by the calling thread for a single-threaded apartment (STA).
pub fn initialize_sta() -> HRESULT {
    unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }
}

/// Close the COM library on the current thread.
pub fn deinitialize() {
    unsafe { CoUninitialize() }
}

/// Audio direction, playback or capture.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Direction {
    Render,
    Capture,
}

impl fmt::Display for Direction {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            Direction::Render => write!(f, "Render"),
            Direction::Capture => write!(f, "Capture"),
        }
    }
}

impl TryFrom<&EDataFlow> for Direction {
    type Error = WasapiError;

    fn try_from(value: &EDataFlow) -> Result<Self, Self::Error> {
        match value {
            EDataFlow(0) => Ok(Self::Render),
            EDataFlow(1) => Ok(Self::Capture),
            // EDataFlow(2) => All/Both,
            x => Err(WasapiError::IllegalDeviceDirection(x.0)),
        }
    }
}
impl TryFrom<EDataFlow> for Direction {
    type Error = WasapiError;

    fn try_from(value: EDataFlow) -> Result<Self, Self::Error> {
        Self::try_from(&value)
    }
}

impl From<&Direction> for EDataFlow {
    fn from(value: &Direction) -> Self {
        match value {
            Direction::Capture => eCapture,
            Direction::Render => eRender,
        }
    }
}
impl From<Direction> for EDataFlow {
    fn from(value: Direction) -> Self {
        Self::from(&value)
    }
}

/// Wrapper for [ERole](https://learn.microsoft.com/en-us/windows/win32/api/mmdeviceapi/ne-mmdeviceapi-erole).
/// Console is the role used by most applications
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Role {
    Console,
    Multimedia,
    Communications,
}

impl fmt::Display for Role {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            Role::Console => write!(f, "Console"),
            Role::Multimedia => write!(f, "Multimedia"),
            Role::Communications => write!(f, "Communications"),
        }
    }
}

impl TryFrom<&ERole> for Role {
    type Error = WasapiError;

    fn try_from(value: &ERole) -> Result<Self, Self::Error> {
        match value {
            ERole(0) => Ok(Self::Console),
            ERole(1) => Ok(Self::Multimedia),
            ERole(2) => Ok(Self::Communications),
            x => Err(WasapiError::IllegalDeviceRole(x.0)),
        }
    }
}
impl TryFrom<ERole> for Role {
    type Error = WasapiError;

    fn try_from(value: ERole) -> Result<Self, Self::Error> {
        Self::try_from(&value)
    }
}

impl From<&Role> for ERole {
    fn from(value: &Role) -> Self {
        match value {
            Role::Communications => eCommunications,
            Role::Multimedia => eMultimedia,
            Role::Console => eConsole,
        }
    }
}
impl From<Role> for ERole {
    fn from(value: Role) -> Self {
        Self::from(&value)
    }
}

/// Helper enum for initializing an [AudioClient].
/// There are four main modes that can be specified,
/// corresponding to the four possible combinations of sharing mode and timing.
/// The enum variants only expose the parameters that can be set in each mode.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum StreamMode {
    /// Shared mode using polling for timing.
    /// The parameters that can be set are the device buffer duration (in units on 100 ns)
    /// and whether automatic format conversion should be enabled.
    /// The audio engine decides the period, and this cannot be changed.
    PollingShared {
        autoconvert: bool,
        buffer_duration_hns: i64,
    },
    /// Exclusive mode using polling for timing.
    /// Both device period and buffer duration are given, in units of 100 ns.
    PollingExclusive {
        buffer_duration_hns: i64,
        period_hns: i64,
    },
    /// Shared mode using event driven timing.
    /// The parameters that can be set are the device buffer duration (in units on 100 ns)
    /// and whether automatic format conversion should be enabled.
    /// The audio engine decides the period, and this cannot be changed.
    EventsShared {
        autoconvert: bool,
        buffer_duration_hns: i64,
    },
    /// Exclusive mode using event driven timing.
    /// The period and buffer duration must be set to the same value.
    /// Only device period is given here, in units of 100 ns.
    EventsExclusive { period_hns: i64 },
}

/// Sharemode for device
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ShareMode {
    Shared,
    Exclusive,
}

/// Timing mode for device
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TimingMode {
    Polling,
    Events,
}

impl fmt::Display for ShareMode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            ShareMode::Shared => write!(f, "Shared"),
            ShareMode::Exclusive => write!(f, "Exclusive"),
        }
    }
}

/// Sample type, float or integer
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SampleType {
    Float,
    Int,
}

impl fmt::Display for SampleType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            SampleType::Float => write!(f, "Float"),
            SampleType::Int => write!(f, "Int"),
        }
    }
}

/// Possible states for an [AudioSessionControl], an enum representing the
/// [AudioSessionStateXxx constants](https://learn.microsoft.com/en-us/windows/win32/api/audiosessiontypes/ne-audiosessiontypes-audiosessionstate)
#[derive(Debug, Eq, PartialEq)]
pub enum SessionState {
    /// The audio session is active. (At least one of the streams in the session is running.)
    Active,
    /// The audio session is inactive. (It contains at least one stream, but none of the streams in the session is currently running.)
    Inactive,
    /// The audio session has expired. (It contains no streams.)
    Expired,
}

impl fmt::Display for SessionState {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            SessionState::Active => write!(f, "Active"),
            SessionState::Inactive => write!(f, "Inactive"),
            SessionState::Expired => write!(f, "Expired"),
        }
    }
}

/// Possible states for an [IMMDevice], an enum representing the
/// [DEVICE_STATE_XXX constants](https://learn.microsoft.com/en-us/windows/win32/coreaudio/device-state-xxx-constants)
#[derive(Debug, Eq, PartialEq)]
pub enum DeviceState {
    /// The audio endpoint device is active. That is, the audio adapter that connects to the
    /// endpoint device is present and enabled. In addition, if the endpoint device plugs int
    /// a jack on the adapter, then the endpoint device is plugged in.
    Active,
    /// The audio endpoint device is disabled. The user has disabled the device in the Windows
    /// multimedia control panel, Mmsys.cpl
    Disabled,
    /// The audio endpoint device is not present because the audio adapter that connects to the
    /// endpoint device has been removed from the system, or the user has disabled the adapter
    /// device in Device Manager.
    NotPresent,
    /// The audio endpoint device is unplugged. The audio adapter that contains the jack for the
    /// endpoint device is present and enabled, but the endpoint device is not plugged into the
    /// jack. Only a device with jack-presence detection can be in this state.
    Unplugged,
}

impl fmt::Display for DeviceState {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            DeviceState::Active => write!(f, "Active"),
            DeviceState::Disabled => write!(f, "Disabled"),
            DeviceState::NotPresent => write!(f, "NotPresent"),
            DeviceState::Unplugged => write!(f, "Unplugged"),
        }
    }
}

/// Get the default playback or capture device for the console role
pub fn get_default_device(direction: &Direction) -> WasapiRes<Device> {
    get_default_device_for_role(direction, &Role::Console)
}

/// Get the default playback or capture device for a specific role
pub fn get_default_device_for_role(direction: &Direction, role: &Role) -> WasapiRes<Device> {
    let dir = direction.into();
    let e_role = role.into();

    let enumerator: IMMDeviceEnumerator =
        unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)? };
    let device = unsafe { enumerator.GetDefaultAudioEndpoint(dir, e_role)? };

    let dev = Device {
        device,
        direction: *direction,
    };
    debug!("default device {:?}", dev.get_friendlyname());
    Ok(dev)
}

/// Calculate a period in units of 100ns that corresponds to the given number of buffer frames at the given sample rate.
/// See the [IAudioClient documentation](https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nf-audioclient-iaudioclient-initialize#remarks).
pub fn calculate_period_100ns(frames: i64, samplerate: i64) -> i64 {
    ((10000.0 * 1000.0 / samplerate as f64 * frames as f64) + 0.5) as i64
}

/// Struct wrapping an [IMMDeviceCollection](https://docs.microsoft.com/en-us/windows/win32/api/mmdeviceapi/nn-mmdeviceapi-immdevicecollection).
pub struct DeviceCollection {
    collection: IMMDeviceCollection,
    direction: Direction,
}

impl DeviceCollection {
    /// Get an [IMMDeviceCollection] of all active playback or capture devices
    pub fn new(direction: &Direction) -> WasapiRes<DeviceCollection> {
        let dir: EDataFlow = direction.into();
        let enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)? };
        let devs = unsafe { enumerator.EnumAudioEndpoints(dir, DEVICE_STATE_ACTIVE)? };
        Ok(DeviceCollection {
            collection: devs,
            direction: *direction,
        })
    }

    /// Get the number of devices in an [IMMDeviceCollection]
    pub fn get_nbr_devices(&self) -> WasapiRes<u32> {
        let count = unsafe { self.collection.GetCount()? };
        Ok(count)
    }

    /// Get a device from an [IMMDeviceCollection] using index
    pub fn get_device_at_index(&self, idx: u32) -> WasapiRes<Device> {
        let device = unsafe { self.collection.Item(idx)? };
        Ok(Device {
            device,
            direction: self.direction,
        })
    }

    /// Get a device from an [IMMDeviceCollection] using name
    pub fn get_device_with_name(&self, name: &str) -> WasapiRes<Device> {
        let count = unsafe { self.collection.GetCount()? };
        trace!("nbr devices {}", count);
        for n in 0..count {
            let device = self.get_device_at_index(n)?;
            let devname = device.get_friendlyname()?;
            if name == devname {
                return Ok(device);
            }
        }
        Err(WasapiError::DeviceNotFound(name.to_owned()))
    }

    /// Get the direction for this [DeviceCollection]
    pub fn get_direction(&self) -> Direction {
        self.direction
    }
}

/// Iterator for [DeviceCollection]
pub struct DeviceCollectionIter<'a> {
    collection: &'a DeviceCollection,
    index: u32,
}

impl Iterator for DeviceCollectionIter<'_> {
    type Item = WasapiRes<Device>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index < self.collection.get_nbr_devices().unwrap() {
            let device = self.collection.get_device_at_index(self.index);
            self.index += 1;
            Some(device)
        } else {
            None
        }
    }
}

/// Implement iterator for [DeviceCollection]
impl<'a> IntoIterator for &'a DeviceCollection {
    type Item = WasapiRes<Device>;
    type IntoIter = DeviceCollectionIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        DeviceCollectionIter {
            collection: self,
            index: 0,
        }
    }
}

/// Struct wrapping an [IMMDevice](https://docs.microsoft.com/en-us/windows/win32/api/mmdeviceapi/nn-mmdeviceapi-immdevice).
pub struct Device {
    device: IMMDevice,
    direction: Direction,
}

impl Device {
    /// Build a [Device] from a supplied [IMMDevice] and [Direction]
    ///
    /// # Safety
    ///
    /// The caller must ensure that the [IMMDevice]'s data flow direction
    /// is the same as the [Direction] supplied to the function.
    ///
    /// Use [Device::from_immdevice], which queries the endpoint, for safe construction.
    pub unsafe fn from_raw(device: IMMDevice, direction: Direction) -> Device {
        Device { device, direction }
    }

    /// Attempts to build a [Device] from a supplied [IMMDevice],
    /// querying the endpoint for its data flow direction.
    pub fn from_immdevice(device: IMMDevice) -> WasapiRes<Device> {
        let endpoint: IMMEndpoint = device.cast()?;
        let direction: Direction = unsafe { endpoint.GetDataFlow()? }.try_into()?;

        Ok(Device { device, direction })
    }

    /// Get an [IAudioClient] from an [IMMDevice]
    pub fn get_iaudioclient(&self) -> WasapiRes<AudioClient> {
        let audio_client = unsafe { self.device.Activate::<IAudioClient>(CLSCTX_ALL, None)? };
        Ok(AudioClient {
            client: audio_client,
            direction: self.direction,
            sharemode: None,
            timingmode: None,
            bytes_per_frame: None,
        })
    }

    /// Read state from an [IMMDevice]
    pub fn get_state(&self) -> WasapiRes<DeviceState> {
        let state = unsafe { self.device.GetState()? };
        trace!("state: {:?}", state);
        let state_enum = match state {
            _ if state == DEVICE_STATE_ACTIVE => DeviceState::Active,
            _ if state == DEVICE_STATE_DISABLED => DeviceState::Disabled,
            _ if state == DEVICE_STATE_NOTPRESENT => DeviceState::NotPresent,
            _ if state == DEVICE_STATE_UNPLUGGED => DeviceState::Unplugged,
            x => return Err(WasapiError::IllegalDeviceState(x.0)),
        };
        Ok(state_enum)
    }

    /// Read the friendly name of the endpoint device (for example, "Speakers (XYZ Audio Adapter)")
    pub fn get_friendlyname(&self) -> WasapiRes<String> {
        self.get_string_property(&PKEY_Device_FriendlyName)
    }

    /// Read the friendly name of the audio adapter to which the endpoint device is attached (for example, "XYZ Audio Adapter")
    pub fn get_interface_friendlyname(&self) -> WasapiRes<String> {
        self.get_string_property(&PKEY_DeviceInterface_FriendlyName)
    }

    /// Read the device description of the endpoint device (for example, "Speakers")
    pub fn get_description(&self) -> WasapiRes<String> {
        self.get_string_property(&PKEY_Device_DeviceDesc)
    }

    /// Read the device format of the endpoint device, which is the format that the user has selected for the stream
    /// that flows between the audio engine and the audio endpoint device when the device operates in shared mode.
    pub fn get_device_format(&self) -> WasapiRes<WaveFormat> {
        let data = self.get_blob_property(&PKEY_AudioEngine_DeviceFormat)?;
        // SAFETY: PKEY_AudioEngine_DeviceFormat is guaranteed to be a WAVEFORMATEX structure based on MSFT docs:
        // https://learn.microsoft.com/en-us/windows/win32/coreaudio/pkey-audioengine-deviceformat
        let waveformatex: &WAVEFORMATEX = unsafe { &*(data.as_ptr() as *const _) };
        WaveFormat::parse(waveformatex)
    }

    /// Read a string property from an [IMMDevice]
    fn get_string_property(&self, key: &PROPERTYKEY) -> WasapiRes<String> {
        self.get_property(key, Self::parse_string_property)
    }

    /// Read a BLOB property from an [IMMDevice]
    fn get_blob_property(&self, key: &PROPERTYKEY) -> WasapiRes<Vec<u8>> {
        self.get_property(key, Self::parse_blob_property)
    }

    /// Read a property from an [IMMDevice] and parse it
    fn get_property<T>(
        &self,
        key: &PROPERTYKEY,
        parse: impl FnOnce(&PROPVARIANT) -> WasapiRes<T>,
    ) -> WasapiRes<T> {
        let store = unsafe { self.device.OpenPropertyStore(STGM_READ)? };
        let mut prop = unsafe { store.GetValue(key)? };
        let ret = parse(&prop);
        unsafe { PropVariantClear(&mut prop) }?;
        ret
    }

    /// Parse a device string property to String
    fn parse_string_property(prop: &PROPVARIANT) -> WasapiRes<String> {
        let propstr = unsafe { PropVariantToStringAlloc(prop)? };
        let wide_name = unsafe { U16CString::from_ptr_str(propstr.0) };
        let name = wide_name.to_string_lossy();
        trace!("name: {}", name);
        Ok(name)
    }

    /// Parse a device blob property to Vec<u8>
    fn parse_blob_property(prop: &PROPVARIANT) -> WasapiRes<Vec<u8>> {
        if prop.vt() != VT_BLOB {
            return Err(windows::core::Error::from(E_INVALIDARG).into());
        }
        let blob = unsafe { prop.Anonymous.Anonymous.Anonymous.blob };
        let blob_slice = unsafe { slice::from_raw_parts(blob.pBlobData, blob.cbSize as usize) };
        let data = blob_slice.to_vec();
        Ok(data)
    }

    /// Get the Id of an [IMMDevice]
    pub fn get_id(&self) -> WasapiRes<String> {
        let idstr = unsafe { self.device.GetId()? };
        let wide_id = unsafe { U16CString::from_ptr_str(idstr.0) };
        let id = wide_id.to_string_lossy();
        trace!("id: {}", id);
        Ok(id)
    }

    /// Get the direction for this Device
    pub fn get_direction(&self) -> Direction {
        self.direction
    }
}

#[implement(IActivateAudioInterfaceCompletionHandler)]
struct Handler(Arc<(Mutex<bool>, Condvar)>);

impl Handler {
    pub fn new(object: Arc<(Mutex<bool>, Condvar)>) -> Handler {
        Handler(object)
    }
}

impl IActivateAudioInterfaceCompletionHandler_Impl for Handler_Impl {
    fn ActivateCompleted(
        &self,
        _activateoperation: Ref<IActivateAudioInterfaceAsyncOperation>,
    ) -> windows::core::Result<()> {
        let (lock, cvar) = &*self.0;
        let mut completed = lock.lock().unwrap();
        *completed = true;
        drop(completed);
        cvar.notify_one();
        Ok(())
    }
}

/// Struct wrapping an [IAudioClient](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iaudioclient).
pub struct AudioClient {
    client: IAudioClient,
    direction: Direction,
    sharemode: Option<ShareMode>,
    timingmode: Option<TimingMode>,
    bytes_per_frame: Option<usize>,
}

impl AudioClient {
    /// Creates a loopback capture [AudioClient] for a specific process.
    ///
    /// `include_tree` is equivalent to [PROCESS_LOOPBACK_MODE](https://learn.microsoft.com/en-us/windows/win32/api/audioclientactivationparams/ne-audioclientactivationparams-process_loopback_mode).
    /// If true, the loopback capture client will capture audio from the target process and all its child processes,
    /// if false only audio from the target process is captured.
    ///
    /// On versions of Windows prior to Windows 10, the thread calling this function
    /// must called in a COM Single-Threaded Apartment (STA).
    ///
    /// Additionally when calling [AudioClient::initialize_client] on the client returned by this method,
    /// the caller must use [Direction::Capture], and [ShareMode::Shared].
    /// Finally calls to [AudioClient::get_device_period] do not work,
    /// however the period passed by the caller to [AudioClient::initialize_client] is irrelevant.
    ///
    /// # Non-functional methods
    /// In process loopback mode, the functionality of the AudioClient is limited.
    /// The following methods either do not work, or return incorrect results:
    /// * `get_mixformat` just returns `Not implemented`.
    /// * `is_supported` just returns `Not implemented` even if the format and mode work.
    /// * `is_supported_exclusive_with_quirks` just returns `Unable to find a supported format`.
    /// * `get_device_period` just returns `Not implemented`.
    /// * `calculate_aligned_period_near` just returns `Not implemented` even for values that would later work.
    /// * `get_buffer_size` returns huge values like 3131961357 but no error.
    /// * `get_current_padding` just returns `Not implemented`.
    /// * `get_available_space_in_frames` just returns `Client has not been initialised` even if it has.
    /// * `get_audiorenderclient` just returns `No such interface supported`.
    /// * `get_audiosessioncontrol` just returns `No such interface supported`.
    /// * `get_audioclock` just returns `No such interface supported`.
    /// * `get_sharemode` always returns `None` when it should return `Shared` after initialisation.
    ///
    /// # Example
    /// ```
    /// use wasapi::{WaveFormat, SampleType, AudioClient, Direction, StreamMode, initialize_mta};
    /// let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2, None);
    /// let buffer_duration_hns = 200_000; // 20ms in hundreds of nanoseconds
    /// let autoconvert = true;
    /// let include_tree = false;
    /// let process_id = std::process::id();
    ///
    /// initialize_mta().ok().unwrap(); // Don't do this on a UI thread
    /// let mut audio_client = AudioClient::new_application_loopback_client(process_id, include_tree).unwrap();
    /// let mode = StreamMode::EventsShared { autoconvert, buffer_duration_hns };
    /// audio_client.initialize_client(
    ///     &desired_format,
    ///     &Direction::Capture,
    ///     &mode
    /// ).unwrap();
    /// ```
    pub fn new_application_loopback_client(process_id: u32, include_tree: bool) -> WasapiRes<Self> {
        unsafe {
            // Create audio client
            let mut audio_client_activation_params = AUDIOCLIENT_ACTIVATION_PARAMS {
                ActivationType: AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK,
                Anonymous: AUDIOCLIENT_ACTIVATION_PARAMS_0 {
                    ProcessLoopbackParams: AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS {
                        TargetProcessId: process_id,
                        ProcessLoopbackMode: if include_tree {
                            PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE
                        } else {
                            PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE
                        },
                    },
                },
            };
            let pinned_params = Pin::new(&mut audio_client_activation_params);

            let raw_prop = PROPVARIANT {
                Anonymous: PROPVARIANT_0 {
                    Anonymous: ManuallyDrop::new(PROPVARIANT_0_0 {
                        vt: VT_BLOB,
                        wReserved1: 0,
                        wReserved2: 0,
                        wReserved3: 0,
                        Anonymous: PROPVARIANT_0_0_0 {
                            blob: BLOB {
                                cbSize: size_of::<AUDIOCLIENT_ACTIVATION_PARAMS>() as u32,
                                pBlobData: pinned_params.get_mut() as *const _ as *mut _,
                            },
                        },
                    }),
                },
            };

            let activation_prop = ManuallyDrop::new(raw_prop);
            let pinned_prop = Pin::new(activation_prop.deref());
            let activation_params = Some(pinned_prop.get_ref() as *const _);

            // Create completion handler
            let setup = Arc::new((Mutex::new(false), Condvar::new()));
            let callback: IActivateAudioInterfaceCompletionHandler =
                Handler::new(setup.clone()).into();

            // Activate audio interface
            let operation = ActivateAudioInterfaceAsync(
                VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
                &IAudioClient::IID,
                activation_params,
                &callback,
            )?;

            // Wait for completion
            let (lock, cvar) = &*setup;
            let mut completed = lock.lock().unwrap();
            while !*completed {
                completed = cvar.wait(completed).unwrap();
            }
            drop(completed);

            // Get audio client and result
            let mut audio_client: Option<IUnknown> = Default::default();
            let mut result: HRESULT = Default::default();
            operation.GetActivateResult(&mut result, &mut audio_client)?;

            // Ensure successful activation
            result.ok()?;
            // always safe to unwrap if result above is checked first
            let audio_client: IAudioClient = audio_client.unwrap().cast()?;

            Ok(AudioClient {
                client: audio_client,
                direction: Direction::Render,
                sharemode: Some(ShareMode::Shared),
                timingmode: None,
                bytes_per_frame: None,
            })
        }
    }

    /// Get MixFormat of the device. This is the format the device uses in shared mode and should always be accepted.
    pub fn get_mixformat(&self) -> WasapiRes<WaveFormat> {
        let temp_fmt_ptr = unsafe { self.client.GetMixFormat()? };
        let temp_fmt = unsafe { *temp_fmt_ptr };
        let mix_format =
            if temp_fmt.cbSize == 22 && temp_fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE {
                unsafe {
                    WaveFormat {
                        wave_fmt: (temp_fmt_ptr as *const _ as *const WAVEFORMATEXTENSIBLE).read(),
                    }
                }
            } else {
                WaveFormat::from_waveformatex(temp_fmt)?
            };
        Ok(mix_format)
    }

    /// Check if a format is supported.
    /// If it's directly supported, this returns Ok(None). If not, but a similar format is, then the nearest matching supported format is returned as Ok(Some(WaveFormat)).
    ///
    /// NOTE: For exclusive mode, this function may not always give the right result for 1- and 2-channel formats.
    /// From the [Microsoft documentation](https://docs.microsoft.com/en-us/windows/win32/coreaudio/device-formats#specifying-the-device-format):
    /// > For exclusive-mode formats, the method queries the device driver.
    /// > Some device drivers will report that they support a 1-channel or 2-channel PCM format if the format is specified by a stand-alone WAVEFORMATEX structure,
    /// > but will reject the same format if it is specified by a WAVEFORMATEXTENSIBLE structure.
    /// > To obtain reliable results from these drivers, exclusive-mode applications should call IsFormatSupported twice for each 1-channel or 2-channel PCM format.
    /// > One call should use a stand-alone WAVEFORMATEX structure to specify the format, and the other call should use a WAVEFORMATEXTENSIBLE structure to specify the same format.
    ///
    /// If the first call fails, use [WaveFormat::to_waveformatex] to get a copy of the WaveFormat in the simpler WAVEFORMATEX representation.
    /// Then call this function again with the new WafeFormat structure.
    /// If the driver then reports that the format is supported, use the original WaveFormat structure when calling [AudioClient::initialize_client].
    ///
    /// See also the helper function [is_supported_exclusive_with_quirks](AudioClient::is_supported_exclusive_with_quirks).
    pub fn is_supported(
        &self,
        wave_fmt: &WaveFormat,
        sharemode: &ShareMode,
    ) -> WasapiRes<Option<WaveFormat>> {
        let supported = match sharemode {
            ShareMode::Exclusive => {
                unsafe {
                    self.client
                        .IsFormatSupported(
                            AUDCLNT_SHAREMODE_EXCLUSIVE,
                            wave_fmt.as_waveformatex_ref(),
                            None,
                        )
                        .ok()?
                };
                None
            }
            ShareMode::Shared => {
                let mut supported_format: *mut WAVEFORMATEX = std::ptr::null_mut();
                unsafe {
                    self.client
                        .IsFormatSupported(
                            AUDCLNT_SHAREMODE_SHARED,
                            wave_fmt.as_waveformatex_ref(),
                            Some(&mut supported_format),
                        )
                        .ok()?
                };
                // Check if we got a pointer to a WAVEFORMATEX structure.
                if supported_format.is_null() {
                    // The pointer is still null, thus the format is supported as is.
                    debug!("The requested format is supported");
                    None
                } else {
                    // Read the structure
                    let temp_fmt: WAVEFORMATEX = unsafe { supported_format.read() };
                    debug!("The requested format is not supported but a simular one is");
                    let new_fmt = if temp_fmt.cbSize == 22
                        && temp_fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE
                    {
                        debug!("got the nearest matching format as a WAVEFORMATEXTENSIBLE");
                        let temp_fmt_ext: WAVEFORMATEXTENSIBLE = unsafe {
                            (supported_format as *const _ as *const WAVEFORMATEXTENSIBLE).read()
                        };
                        WaveFormat {
                            wave_fmt: temp_fmt_ext,
                        }
                    } else {
                        debug!("got the nearest matching format as a WAVEFORMATEX, converting..");
                        WaveFormat::from_waveformatex(temp_fmt)?
                    };
                    Some(new_fmt)
                }
            }
        };
        Ok(supported)
    }

    /// A helper function for checking if a format is supported.
    /// It calls `is_supported` several times with different options
    /// in order to find a format that the device accepts.
    ///
    /// The alternatives it tries are:
    /// - The format as given.
    /// - If one or two channels, try with the format as WAVEFORMATEX.
    /// - Try with different channel masks:
    ///   - If channels <= 8: Recommended mask(s) from ksmedia.h.
    ///   - If channels <= 18: Simple mask.
    ///   - Zero mask.
    ///
    /// If an accepted format is found, this is returned.
    /// An error means no accepted format was found.
    pub fn is_supported_exclusive_with_quirks(
        &self,
        wave_fmt: &WaveFormat,
    ) -> WasapiRes<WaveFormat> {
        let mut wave_fmt = wave_fmt.clone();
        let supported_direct = self.is_supported(&wave_fmt, &ShareMode::Exclusive);
        if supported_direct.is_ok() {
            debug!("The requested format is supported as provided");
            return Ok(wave_fmt);
        }
        if wave_fmt.get_nchannels() <= 2 {
            debug!("Repeating query with format as WAVEFORMATEX");
            let wave_formatex = wave_fmt.to_waveformatex().unwrap();
            if self
                .is_supported(&wave_formatex, &ShareMode::Exclusive)
                .is_ok()
            {
                debug!("The requested format is supported as WAVEFORMATEX");
                return Ok(wave_formatex);
            }
        }
        let masks = make_channelmasks(wave_fmt.get_nchannels() as usize);
        for mask in masks {
            debug!("Repeating query with channel mask: {:#010b}", mask);
            wave_fmt.wave_fmt.dwChannelMask = mask;
            if self.is_supported(&wave_fmt, &ShareMode::Exclusive).is_ok() {
                debug!(
                    "The requested format is supported with a modified mask: {:#010b}",
                    mask
                );
                return Ok(wave_fmt);
            }
        }
        Err(WasapiError::UnsupportedFormat)
    }

    /// Get default and minimum periods in 100-nanosecond units
    pub fn get_device_period(&self) -> WasapiRes<(i64, i64)> {
        let mut def_time = 0;
        let mut min_time = 0;
        unsafe {
            self.client
                .GetDevicePeriod(Some(&mut def_time), Some(&mut min_time))?
        };
        trace!("default period {}, min period {}", def_time, min_time);
        Ok((def_time, min_time))
    }

    #[deprecated(
        since = "0.17.0",
        note = "please use the new function name `get_device_period` instead"
    )]
    pub fn get_periods(&self) -> WasapiRes<(i64, i64)> {
        self.get_device_period()
    }

    /// Helper function for calculating a period size in 100-nanosecond units that is near a desired value,
    /// and always larger than the minimum value supported by the device.
    /// The returned value leads to a device buffer size that is aligned both to the frame size of the format,
    /// and the optional align_bytes value.
    /// This parameter is used for devices that require the buffer size to be a multiple of a certain number of bytes.
    /// Give None, Some(0) or Some(1) if the device has no special requirements for the alignment.
    /// For example, all devices following the Intel High Definition Audio specification require buffer sizes in multiples of 128 bytes.
    ///
    /// See also the `playnoise_exclusive` example.
    pub fn calculate_aligned_period_near(
        &self,
        desired_period: i64,
        align_bytes: Option<u32>,
        wave_fmt: &WaveFormat,
    ) -> WasapiRes<i64> {
        let (_default_period, min_period) = self.get_device_period()?;
        let adjusted_desired_period = cmp::max(desired_period, min_period);
        let frame_bytes = wave_fmt.get_blockalign();
        let period_alignment_bytes = match align_bytes {
            Some(0) => frame_bytes,
            Some(bytes) => frame_bytes.lcm(&bytes),
            None => frame_bytes,
        };
        let period_alignment_frames = period_alignment_bytes as i64 / frame_bytes as i64;
        let desired_period_frames =
            (adjusted_desired_period as f64 * wave_fmt.get_samplespersec() as f64 / 10000000.0)
                .round() as i64;
        let min_period_frames =
            (min_period as f64 * wave_fmt.get_samplespersec() as f64 / 10000000.0).ceil() as i64;
        let mut nbr_segments = desired_period_frames / period_alignment_frames;
        if nbr_segments * period_alignment_frames < min_period_frames {
            // Add one segment if the value got rounded down below the minimum
            nbr_segments += 1;
        }
        let aligned_period = calculate_period_100ns(
            period_alignment_frames * nbr_segments,
            wave_fmt.get_samplespersec() as i64,
        );
        Ok(aligned_period)
    }

    /// Initialize an [AudioClient] for the given direction, sharemode, timing mode and format.
    /// This method wraps [IAudioClient::Initialize()](https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nf-audioclient-iaudioclient-initialize).
    ///
    /// ### Sharing mode
    /// In WASAPI, sharing mode determines how multiple audio applications interact with the same audio endpoint.
    /// There are two primary sharing modes: Shared and Exclusive.
    /// #### Shared Mode ([ShareMode::Shared])
    /// - Multiple applications can simultaneously access the audio device.
    /// - The system's audio engine mixes the audio streams from all applications.
    /// - The application has no control over the sample rate and format used by the device.
    /// - The audio engine can perform automatic sample rate and format conversion,
    ///   meaning that almost any format can be accepted.
    ///
    /// #### Exclusive Mode ([ShareMode::Exclusive])
    /// - Only one application can access the audio device at a time.
    /// - This mode provides lower latency but requires the device to support the exact audio format requested.
    /// - The application can control the sample rate and format used by the device.
    ///
    /// ### Timing mode
    /// Event-driven mode and polling mode are two different ways of handling audio buffer updates.
    ///
    /// #### Event-Driven Mode ([TimingMode::Events])
    ///   - In this mode, the application registers an event handle using [AudioClient::set_get_eventhandle()].
    ///   - The system signals this event whenever a new buffer of audio data is ready to be processed (either for rendering or capture).
    ///   - The application's audio processing thread waits on this event ([Handle::wait_for_event()]).
    ///   - When the event is signaled, the thread wakes up to processes the available data, and then goes back to waiting.
    ///   - This mode is generally more efficient because the application only wakes up when there's work to do.
    ///   - It's suitable for real-time audio applications where low latency is important.
    ///   - This mode is not supported by all devices in exclusive mode (but all devices are supported in shared mode).
    ///   - In exclusive mode, devices using the standard Windows USB audio driver can have issues
    ///     with stuttering sound on playback.
    ///
    /// #### Polling Mode ([TimingMode::Polling])
    ///   - In this mode, the application periodically calls [AudioClient::get_current_padding()] (for capture)
    ///     or [AudioClient::get_available_space_in_frames()] (for playback)
    ///     to check how much data is available or required.
    ///   - The thread processes the data, and then goes to sleep, for example by calling [std::thread::sleep()].
    ///   - This mode is less efficient and is more prone to glitches when running at low latency.
    ///   - In exclusive mode, it supports more devices, and does not have the stuttering issue with USB audio devices.
    pub fn initialize_client(
        &mut self,
        wavefmt: &WaveFormat,
        direction: &Direction,
        stream_mode: &StreamMode,
    ) -> WasapiRes<()> {
        let sharemode = match stream_mode {
            StreamMode::PollingShared { .. } | StreamMode::EventsShared { .. } => ShareMode::Shared,
            StreamMode::PollingExclusive { .. } | StreamMode::EventsExclusive { .. } => {
                ShareMode::Exclusive
            }
        };
        let timing = match stream_mode {
            StreamMode::PollingShared { .. } | StreamMode::PollingExclusive { .. } => {
                TimingMode::Polling
            }
            StreamMode::EventsShared { .. } | StreamMode::EventsExclusive { .. } => {
                TimingMode::Events
            }
        };
        let mut streamflags = match (&self.direction, direction, sharemode) {
            (Direction::Render, Direction::Capture, ShareMode::Shared) => {
                AUDCLNT_STREAMFLAGS_LOOPBACK
            }
            (Direction::Render, Direction::Capture, ShareMode::Exclusive) => {
                return Err(WasapiError::LoopbackWithExclusiveMode);
            }
            (Direction::Capture, Direction::Render, _) => {
                return Err(WasapiError::RenderToCaptureDevice);
            }
            _ => 0,
        };
        match stream_mode {
            StreamMode::PollingShared { autoconvert, .. }
            | StreamMode::EventsShared { autoconvert, .. } => {
                if *autoconvert {
                    streamflags |= AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM
                        | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY;
                }
            }
            _ => {}
        }
        if timing == TimingMode::Events {
            streamflags |= AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
        }
        let mode = match sharemode {
            ShareMode::Exclusive => AUDCLNT_SHAREMODE_EXCLUSIVE,
            ShareMode::Shared => AUDCLNT_SHAREMODE_SHARED,
        };
        let (period, buffer_duration) = match stream_mode {
            StreamMode::PollingShared {
                buffer_duration_hns,
                ..
            } => (0, *buffer_duration_hns),
            StreamMode::EventsShared {
                buffer_duration_hns,
                ..
            } => (0, *buffer_duration_hns),
            StreamMode::PollingExclusive {
                period_hns,
                buffer_duration_hns,
            } => (*period_hns, *buffer_duration_hns),
            StreamMode::EventsExclusive { period_hns, .. } => (*period_hns, *period_hns),
        };
        unsafe {
            self.client.Initialize(
                mode,
                streamflags,
                buffer_duration,
                period,
                wavefmt.as_waveformatex_ref(),
                None,
            )?;
        }
        self.direction = *direction;
        self.sharemode = Some(sharemode);
        self.timingmode = Some(timing);
        self.bytes_per_frame = Some(wavefmt.get_blockalign() as usize);
        Ok(())
    }

    /// Create and return an event handle for an [AudioClient].
    /// This is required when using an [AudioClient] initialized for event driven mode, [TimingMode::Events].
    pub fn set_get_eventhandle(&self) -> WasapiRes<Handle> {
        let h_event = unsafe { CreateEventA(None, false, false, PCSTR::null())? };
        unsafe { self.client.SetEventHandle(h_event)? };
        Ok(Handle { handle: h_event })
    }

    /// Get buffer size in frames,
    /// see [IAudioClient::GetBufferSize](https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nf-audioclient-iaudioclient-getbuffersize).
    pub fn get_buffer_size(&self) -> WasapiRes<u32> {
        let buffer_frame_count = unsafe { self.client.GetBufferSize()? };
        trace!("buffer_frame_count {}", buffer_frame_count);
        Ok(buffer_frame_count)
    }

    #[deprecated(
        since = "0.17.0",
        note = "please use the new function name `get_buffer_size` instead"
    )]
    pub fn get_bufferframecount(&self) -> WasapiRes<u32> {
        self.get_buffer_size()
    }

    /// Get current padding in frames.
    /// This represents the number of frames currently in the buffer, for both capture and render devices.
    /// The exact meaning depends on how the AudioClient was initialized, see
    /// [IAudioClient::GetCurrentPadding](https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nf-audioclient-iaudioclient-getcurrentpadding).
    pub fn get_current_padding(&self) -> WasapiRes<u32> {
        let padding_count = unsafe { self.client.GetCurrentPadding()? };
        trace!("padding_count {}", padding_count);
        Ok(padding_count)
    }

    /// Get buffer size minus padding in frames.
    /// Use this to find out how much free space is available in the buffer.
    pub fn get_available_space_in_frames(&self) -> WasapiRes<u32> {
        let frames = match (self.sharemode, self.timingmode) {
            (Some(ShareMode::Exclusive), Some(TimingMode::Events)) => {
                let buffer_frame_count = unsafe { self.client.GetBufferSize()? };
                trace!("buffer_frame_count {}", buffer_frame_count);
                buffer_frame_count
            }
            (Some(_), Some(_)) => {
                let padding_count = unsafe { self.client.GetCurrentPadding()? };
                let buffer_frame_count = unsafe { self.client.GetBufferSize()? };

                buffer_frame_count - padding_count
            }
            _ => return Err(WasapiError::ClientNotInit),
        };
        Ok(frames)
    }

    /// Start the stream on an [IAudioClient]
    pub fn start_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Start()? };
        Ok(())
    }

    /// Stop the stream on an [IAudioClient]
    pub fn stop_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Stop()? };
        Ok(())
    }

    /// Reset the stream on an [IAudioClient]
    pub fn reset_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Reset()? };
        Ok(())
    }

    /// Get a rendering (playback) client
    pub fn get_audiorenderclient(&self) -> WasapiRes<AudioRenderClient> {
        let client = unsafe { self.client.GetService::<IAudioRenderClient>()? };
        Ok(AudioRenderClient {
            client,
            bytes_per_frame: self.bytes_per_frame.unwrap_or_default(),
        })
    }

    /// Get a capture client
    pub fn get_audiocaptureclient(&self) -> WasapiRes<AudioCaptureClient> {
        let client = unsafe { self.client.GetService::<IAudioCaptureClient>()? };
        Ok(AudioCaptureClient {
            client,
            sharemode: self.sharemode,
            bytes_per_frame: self.bytes_per_frame.unwrap_or_default(),
        })
    }

    /// Get the [AudioSessionControl]
    pub fn get_audiosessioncontrol(&self) -> WasapiRes<AudioSessionControl> {
        let control = unsafe { self.client.GetService::<IAudioSessionControl>()? };
        Ok(AudioSessionControl { control })
    }

    /// Get the [AudioClock]
    pub fn get_audioclock(&self) -> WasapiRes<AudioClock> {
        let clock = unsafe { self.client.GetService::<IAudioClock>()? };
        Ok(AudioClock { clock })
    }

    /// Get the direction for this [AudioClient]
    pub fn get_direction(&self) -> Direction {
        self.direction
    }

    /// Get the sharemode for this [AudioClient].
    /// The sharemode is decided when the client is initialized.
    pub fn get_sharemode(&self) -> Option<ShareMode> {
        self.sharemode
    }

    /// Get the timing mode for this [AudioClient].
    /// The mode is decided when the client is initialized.
    pub fn get_timing_mode(&self) -> Option<TimingMode> {
        self.timingmode
    }

    /// Get the Acoustic Echo Cancellation Control.
    /// If it succeeds, the capture endpoint supports control of the loopback reference endpoint for AEC.
    pub fn get_aec_control(&self) -> WasapiRes<AcousticEchoCancellationControl> {
        let control = unsafe {
            self.client
                .GetService::<IAcousticEchoCancellationControl>()?
        };
        Ok(AcousticEchoCancellationControl { control })
    }

    /// Get the Audio Effects Manager.
    pub fn get_audio_effects_manager(&self) -> WasapiRes<AudioEffectsManager> {
        let manager = unsafe { self.client.GetService::<IAudioEffectsManager>()? };
        Ok(AudioEffectsManager { manager })
    }

    /// Set the category of an audio stream.
    ///
    /// This function is a subset of the `set_client_properties` method, as it only sets the audio stream category, and
    /// hence it is recommended to use `set_client_properties` instead.
    pub fn set_audio_stream_category(&self, category: AUDIO_STREAM_CATEGORY) -> WasapiRes<()> {
        let audio_client_2 = self.client.cast::<IAudioClient2>()?;

        let audio_client_property = AudioClientProperties {
            cbSize: size_of::<AudioClientProperties>() as u32,
            eCategory: category,
            ..Default::default()
        };

        unsafe { audio_client_2.SetClientProperties(&audio_client_property as *const _)? };
        Ok(())
    }

    /// Set properties of the client's audio stream.
    pub fn set_client_properties(
        &self,
        is_offload: bool,
        category: AUDIO_STREAM_CATEGORY,
        options: AUDCLNT_STREAMOPTIONS,
    ) -> WasapiRes<()> {
        let audio_client_2 = self.client.cast::<IAudioClient2>()?;

        let audio_client_property = AudioClientProperties {
            cbSize: size_of::<AudioClientProperties>() as u32,
            bIsOffload: is_offload.into(),
            eCategory: category,
            Options: options,
        };

        unsafe { audio_client_2.SetClientProperties(&audio_client_property as *const _)? };
        Ok(())
    }

    /// Check if the Acoustic Echo Cancellation (AEC) is supported.
    pub fn is_aec_supported(&self) -> WasapiRes<bool> {
        if !self.is_aec_effect_present()? {
            return Ok(false);
        }

        match unsafe { self.client.GetService::<IAcousticEchoCancellationControl>() } {
            Ok(_) => Ok(true),
            Err(err) if err == E_NOINTERFACE.into() => Ok(false),
            Err(err) => Err(err.into()),
        }
    }

    /// Check if the Acoustic Echo Cancellation (AEC) effect is currently present.
    fn is_aec_effect_present(&self) -> WasapiRes<bool> {
        // IAudioEffectsManager requires Windows 11 (build 22000 or higher).
        let audio_effects_manager = match self.get_audio_effects_manager() {
            Ok(manager) => manager,
            Err(WasapiError::Windows(win_err)) if win_err == E_NOINTERFACE.into() => {
                // Audio effects manager is not supported, so clearly not present.
                return Ok(false);
            }
            Err(err) => return Err(err),
        };

        if let Some(audio_effects) = audio_effects_manager.get_audio_effects()? {
            // Check if the AEC effect is present in the list of audio effects.
            let is_present = audio_effects
                .iter()
                .any(|effect| effect.id == AUDIO_EFFECT_TYPE_ACOUSTIC_ECHO_CANCELLATION);
            return Ok(is_present);
        }

        Ok(false)
    }
}

/// Struct wrapping an [IAudioSessionControl](https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nn-audiopolicy-iaudiosessioncontrol).
pub struct AudioSessionControl {
    control: IAudioSessionControl,
}

impl AudioSessionControl {
    /// Get the current session state
    pub fn get_state(&self) -> WasapiRes<SessionState> {
        let state = unsafe { self.control.GetState()? };
        #[allow(non_upper_case_globals)]
        let sessionstate = match state {
            _ if state == AudioSessionStateActive => SessionState::Active,
            _ if state == AudioSessionStateInactive => SessionState::Inactive,
            _ if state == AudioSessionStateExpired => SessionState::Expired,
            x => return Err(WasapiError::IllegalSessionState(x.0)),
        };
        Ok(sessionstate)
    }

    /// Register to receive notifications.
    /// Returns a [EventRegistration] struct.
    /// The notifications are unregistered when this struct is dropped.
    /// Make sure to store the [EventRegistration] in a variable that remains
    /// in scope for as long as the event notifications are needed.
    ///
    /// The function takes ownership of the provided [EventCallbacks].
    pub fn register_session_notification(
        &self,
        callbacks: EventCallbacks,
    ) -> WasapiRes<EventRegistration> {
        let events: IAudioSessionEvents = AudioSessionEvents::new(callbacks).into();

        match unsafe { self.control.RegisterAudioSessionNotification(&events) } {
            Ok(()) => Ok(EventRegistration {
                events,
                control: self.control.clone(),
            }),
            Err(err) => Err(WasapiError::RegisterNotifications(err)),
        }
    }
}

/// Struct for keeping track of the registered notifications.
pub struct EventRegistration {
    events: IAudioSessionEvents,
    control: IAudioSessionControl,
}

impl Drop for EventRegistration {
    fn drop(&mut self) {
        let _ = unsafe {
            self.control
                .UnregisterAudioSessionNotification(&self.events)
        };
    }
}

/// Struct wrapping an [IAudioClock](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iaudioclock).
pub struct AudioClock {
    clock: IAudioClock,
}

impl AudioClock {
    /// Get the frequency for this [AudioClock].
    /// Note that the unit for the value is undefined.
    pub fn get_frequency(&self) -> WasapiRes<u64> {
        let freq = unsafe { self.clock.GetFrequency()? };
        Ok(freq)
    }

    /// Get the current device position. Returns the position, as well as the value of the
    /// performance counter at the time the position values was taken.
    /// The unit for the position value is undefined, but the frequency and position values are
    /// in the same unit. Dividing the position with the frequency gets the position in seconds.
    pub fn get_position(&self) -> WasapiRes<(u64, u64)> {
        let mut pos = 0;
        let mut timer = 0;
        unsafe { self.clock.GetPosition(&mut pos, Some(&mut timer))? };
        Ok((pos, timer))
    }
}

/// Struct wrapping an [IAudioRenderClient](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iaudiorenderclient).
pub struct AudioRenderClient {
    client: IAudioRenderClient,
    bytes_per_frame: usize,
}

impl AudioRenderClient {
    /// Write raw bytes data to a device from a slice.
    /// The number of frames to write should first be checked with the
    /// [AudioClient::get_available_space_in_frames()] method.
    /// The buffer_flags argument can be used to mark a buffer as silent.
    pub fn write_to_device(
        &self,
        nbr_frames: usize,
        data: &[u8],
        buffer_flags: Option<BufferFlags>,
    ) -> WasapiRes<()> {
        if nbr_frames == 0 {
            return Ok(());
        }
        let nbr_bytes = nbr_frames * self.bytes_per_frame;
        if nbr_bytes != data.len() {
            return Err(WasapiError::DataLengthMismatch {
                received: data.len(),
                expected: nbr_bytes,
            });
        }
        let bufferptr = unsafe { self.client.GetBuffer(nbr_frames as u32)? };
        let bufferslice = unsafe { slice::from_raw_parts_mut(bufferptr, nbr_bytes) };
        bufferslice.copy_from_slice(data);
        let flags = match buffer_flags {
            Some(bflags) => bflags.to_u32(),
            None => 0,
        };
        unsafe { self.client.ReleaseBuffer(nbr_frames as u32, flags)? };
        trace!("wrote {} frames", nbr_frames);
        Ok(())
    }

    /// Write raw bytes data to a device from a deque.
    /// The number of frames to write should first be checked with the
    /// [AudioClient::get_available_space_in_frames()] method.
    /// The buffer_flags argument can be used to mark a buffer as silent.
    pub fn write_to_device_from_deque(
        &self,
        nbr_frames: usize,
        data: &mut VecDeque<u8>,
        buffer_flags: Option<BufferFlags>,
    ) -> WasapiRes<()> {
        if nbr_frames == 0 {
            return Ok(());
        }
        let nbr_bytes = nbr_frames * self.bytes_per_frame;
        if nbr_bytes > data.len() {
            return Err(WasapiError::DataLengthTooShort {
                received: data.len(),
                expected: nbr_bytes,
            });
        }
        let bufferptr = unsafe { self.client.GetBuffer(nbr_frames as u32)? };
        let bufferslice = unsafe { slice::from_raw_parts_mut(bufferptr, nbr_bytes) };
        for element in bufferslice.iter_mut() {
            *element = data.pop_front().unwrap();
        }
        let flags = match buffer_flags {
            Some(bflags) => bflags.to_u32(),
            None => 0,
        };
        unsafe { self.client.ReleaseBuffer(nbr_frames as u32, flags)? };
        trace!("wrote {} frames", nbr_frames);
        Ok(())
    }
}

/// Struct representing information on data read from an audio client buffer.
#[derive(Debug)]
pub struct BufferInfo {
    /// Decoded audio client flags.
    pub flags: BufferFlags,
    /// The index of the first frame that was read from the buffer.
    pub index: u64,
    /// The timestamp in 100-nanosecond units of the first frame that was read from the buffer.
    pub timestamp: u64,
}

impl BufferInfo {
    /// Creates a new [BufferInfo] struct from the `u32` flags value, and `u64` index and timestamp.
    pub fn new(flags: u32, index: u64, timestamp: u64) -> Self {
        Self {
            flags: BufferFlags::new(flags),
            index,
            timestamp,
        }
    }

    pub fn none() -> Self {
        Self {
            flags: BufferFlags::none(),
            index: 0,
            timestamp: 0,
        }
    }
}

/// Struct representing the [ _AUDCLNT_BUFFERFLAGS enum values](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/ne-audioclient-_audclnt_bufferflags).
#[derive(Debug)]
pub struct BufferFlags {
    /// AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY
    pub data_discontinuity: bool,
    /// AUDCLNT_BUFFERFLAGS_SILENT
    pub silent: bool,
    /// AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR
    pub timestamp_error: bool,
}

impl BufferFlags {
    /// Create a new [BufferFlags] struct from a `u32` value.
    pub fn new(flags: u32) -> Self {
        BufferFlags {
            data_discontinuity: flags & AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY.0 as u32 > 0,
            silent: flags & AUDCLNT_BUFFERFLAGS_SILENT.0 as u32 > 0,
            timestamp_error: flags & AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR.0 as u32 > 0,
        }
    }

    pub fn none() -> Self {
        BufferFlags {
            data_discontinuity: false,
            silent: false,
            timestamp_error: false,
        }
    }

    /// Convert a [BufferFlags] struct to a `u32` value.
    pub fn to_u32(&self) -> u32 {
        let mut value = 0;
        if self.data_discontinuity {
            value += AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY.0 as u32;
        }
        if self.silent {
            value += AUDCLNT_BUFFERFLAGS_SILENT.0 as u32;
        }
        if self.timestamp_error {
            value += AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR.0 as u32;
        }
        value
    }
}

/// Struct wrapping an [IAudioCaptureClient](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iaudiocaptureclient).
pub struct AudioCaptureClient {
    client: IAudioCaptureClient,
    sharemode: Option<ShareMode>,
    bytes_per_frame: usize,
}

impl AudioCaptureClient {
    /// Get number of frames in next packet when in shared mode.
    /// In exclusive mode it returns `None`, instead use [AudioClient::get_buffer_size()] or [AudioClient::get_current_padding()].
    /// See [IAudioCaptureClient::GetNextPacketSize](https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nf-audioclient-iaudiocaptureclient-getnextpacketsize).
    pub fn get_next_packet_size(&self) -> WasapiRes<Option<u32>> {
        if let Some(ShareMode::Exclusive) = self.sharemode {
            return Ok(None);
        }
        let nbr_frames = unsafe { self.client.GetNextPacketSize()? };
        Ok(Some(nbr_frames))
    }

    #[deprecated(
        since = "0.17.0",
        note = "please use the new function name `get_next_packet_size` instead"
    )]
    pub fn get_next_nbr_frames(&self) -> WasapiRes<Option<u32>> {
        self.get_next_packet_size()
    }

    /// Read raw bytes from a device into a slice. Returns the number of frames
    /// that was read, and the `BufferInfo` describing the buffer that the data was read from.
    /// The slice must be large enough to hold all data.
    /// If it is longer that needed, the unused elements will not be modified.
    pub fn read_from_device(&self, data: &mut [u8]) -> WasapiRes<(u32, BufferInfo)> {
        let data_len_in_frames = data.len() / self.bytes_per_frame;
        if data_len_in_frames == 0 {
            return Ok((0, BufferInfo::none()));
        }
        let mut buffer_ptr = ptr::null_mut();
        let mut nbr_frames_returned = 0;
        let mut index: u64 = 0;
        let mut timestamp: u64 = 0;
        let mut flags = 0;
        unsafe {
            self.client.GetBuffer(
                &mut buffer_ptr,
                &mut nbr_frames_returned,
                &mut flags,
                Some(&mut index),
                Some(&mut timestamp),
            )?
        };
        let buffer_info = BufferInfo::new(flags, index, timestamp);
        if nbr_frames_returned == 0 {
            unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
            return Ok((0, buffer_info));
        }
        if data_len_in_frames < nbr_frames_returned as usize {
            unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
            return Err(WasapiError::DataLengthTooShort {
                received: data_len_in_frames,
                expected: nbr_frames_returned as usize,
            });
        }
        let len_in_bytes = nbr_frames_returned as usize * self.bytes_per_frame;
        let bufferslice = unsafe { slice::from_raw_parts(buffer_ptr, len_in_bytes) };
        data[..len_in_bytes].copy_from_slice(bufferslice);
        if nbr_frames_returned > 0 {
            unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
        }
        trace!("read {} frames", nbr_frames_returned);
        Ok((nbr_frames_returned, buffer_info))
    }

    /// Read raw bytes data from a device into a deque.
    /// Returns the [BufferInfo] describing the buffer that the data was read from.
    pub fn read_from_device_to_deque(&self, data: &mut VecDeque<u8>) -> WasapiRes<BufferInfo> {
        let mut buffer_ptr = ptr::null_mut();
        let mut nbr_frames_returned = 0;
        let mut index: u64 = 0;
        let mut timestamp: u64 = 0;
        let mut flags = 0;
        unsafe {
            self.client.GetBuffer(
                &mut buffer_ptr,
                &mut nbr_frames_returned,
                &mut flags,
                Some(&mut index),
                Some(&mut timestamp),
            )?
        };
        let buffer_info = BufferInfo::new(flags, index, timestamp);
        if nbr_frames_returned == 0 {
            // There is no need to release a buffer of 0 bytes
            return Ok(buffer_info);
        }
        let len_in_bytes = nbr_frames_returned as usize * self.bytes_per_frame;
        let bufferslice = unsafe { slice::from_raw_parts(buffer_ptr, len_in_bytes) };
        for element in bufferslice.iter() {
            data.push_back(*element);
        }
        if nbr_frames_returned > 0 {
            unsafe { self.client.ReleaseBuffer(nbr_frames_returned).unwrap() };
        }
        trace!("read {} frames", nbr_frames_returned);
        Ok(buffer_info)
    }

    /// Get the sharemode for this [AudioCaptureClient].
    /// The sharemode is decided when the client is initialized.
    pub fn get_sharemode(&self) -> Option<ShareMode> {
        self.sharemode
    }
}

/// Struct wrapping a [HANDLE] to an [Event Object](https://docs.microsoft.com/en-us/windows/win32/sync/event-objects).
pub struct Handle {
    handle: HANDLE,
}

impl Handle {
    /// Wait for an event on a handle, with a timeout given in ms
    pub fn wait_for_event(&self, timeout_ms: u32) -> WasapiRes<()> {
        let retval = unsafe { WaitForSingleObject(self.handle, timeout_ms) };
        if retval.0 != WAIT_OBJECT_0.0 {
            return Err(WasapiError::EventTimeout);
        }
        Ok(())
    }
}

// Struct wrapping an [IAudioEffectsManager](https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iaudioeffectsmanager).
pub struct AudioEffectsManager {
    manager: IAudioEffectsManager,
}

impl AudioEffectsManager {
    /// Gets the current list of audio effects for the associated audio stream.
    pub fn get_audio_effects(&self) -> WasapiRes<Option<Vec<AUDIO_EFFECT>>> {
        let mut audio_effects: *mut AUDIO_EFFECT = std::ptr::null_mut();
        let mut num_effects: u32 = 0;

        unsafe {
            self.manager
                .GetAudioEffects(&mut audio_effects, &mut num_effects)?;
        }

        if num_effects > 0 {
            let effects_slice =
                unsafe { slice::from_raw_parts(audio_effects, num_effects as usize) };
            let effects_vec = effects_slice.to_vec();
            // Free the memory allocated for the audio effects.
            unsafe { CoTaskMemFree(Some(audio_effects as *mut _)) };
            Ok(Some(effects_vec))
        } else {
            Ok(None)
        }
    }
}

/// Struct wrapping an [AcousticEchoCancellationControl](https://learn.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iacousticechocancellationcontrol).
pub struct AcousticEchoCancellationControl {
    control: IAcousticEchoCancellationControl,
}

impl AcousticEchoCancellationControl {
    /// Sets the audio render endpoint to be used as the reference stream for acoustic echo cancellation (AEC).
    ///
    /// # Parameters
    /// - `endpoint_id`: An optional string containing the device ID of the audio render endpoint to use as the loopback reference.
    ///   If set to `None`, Windows will automatically select the reference device.
    ///   You can obtain the device ID by calling [Device::get_id()].
    ///
    /// # Errors
    /// Returns an error if setting the echo cancellation render endpoint fails.
    pub fn set_echo_cancellation_render_endpoint(
        &self,
        endpoint_id: Option<String>,
    ) -> WasapiRes<()> {
        let endpoint_id = if let Some(endpoint_id) = endpoint_id {
            PCWSTR::from_raw(HSTRING::from(endpoint_id).as_ptr())
        } else {
            PCWSTR::null()
        };
        unsafe {
            self.control
                .SetEchoCancellationRenderEndpoint(endpoint_id)?
        };
        Ok(())
    }
}
