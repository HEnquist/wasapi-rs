use num_integer::Integer;
use std::cmp;
use std::collections::VecDeque;
use std::mem::{size_of, ManuallyDrop};
use std::ops::Deref;
use std::pin::Pin;
use std::rc::Rc;
use std::sync::{Arc, Condvar, Mutex};
use std::{fmt, ptr, slice};
use widestring::U16CString;
use windows::Win32::Foundation::PROPERTYKEY;
use windows::Win32::Media::Audio::{
    ActivateAudioInterfaceAsync, EDataFlow, ERole, IActivateAudioInterfaceAsyncOperation,
    IActivateAudioInterfaceCompletionHandler, IActivateAudioInterfaceCompletionHandler_Impl,
    IMMEndpoint, AUDIOCLIENT_ACTIVATION_PARAMS, AUDIOCLIENT_ACTIVATION_PARAMS_0,
    AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK, AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS,
    PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE,
    PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE, VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
};
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
use windows_core::{implement, IUnknown, Interface, Ref};

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

/// Sharemode for device
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ShareMode {
    Shared,
    Exclusive,
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

    /// Read the FriendlyName of an [IMMDevice]
    fn get_string_property(&self, key: &PROPERTYKEY) -> WasapiRes<String> {
        let store = unsafe { self.device.OpenPropertyStore(STGM_READ)? };
        let prop = unsafe { store.GetValue(key)? };
        let propstr = unsafe { PropVariantToStringAlloc(&prop)? };
        let wide_name = unsafe { U16CString::from_ptr_str(propstr.0) };
        let name = wide_name.to_string_lossy();
        trace!("name: {}", name);
        Ok(name)
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
    /// Finally calls to [AudioClient::get_periods] do not work,
    /// however the period passed by the caller to [AudioClient::initialize_client] is irrelevant.
    ///
    /// # Non-functional methods
    /// In process loopback mode, the functionality of the AudioClient is limited.
    /// The following methods either do not work, or return incorrect results:
    /// * `get_mixformat` just returns `Not implemented`.
    /// * `is_supported` just returns `Not implemented` even if the format and mode work.
    /// * `is_supported_exclusive_with_quirks` just returns `Unable to find a supported format`.
    /// * `get_periods` just returns `Not implemented`.
    /// * `calculate_aligned_period_near` just returns `Not implemented` even for values that would later work.
    /// * `get_bufferframecount` returns huge values like 3131961357 but no error.
    /// * `get_current_padding` just returns `Not implemented`.
    /// * `get_available_space_in_frames` just returns `Client has not been initialised` even if it has.
    /// * `get_audiorenderclient` just returns `No such interface supported`.
    /// * `get_audiosessioncontrol` just returns `No such interface supported`.
    /// * `get_audioclock` just returns `No such interface supported`.
    /// * `get_sharemode` always returns `None` when it should return `Shared` after initialisation.
    ///
    /// # Example
    /// ```
    /// use wasapi::{WaveFormat, SampleType, AudioClient, Direction, ShareMode, initialize_mta};
    /// let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2, None);
    /// let hnsbufferduration = 200_000; // 20ms in hundreds of nanoseconds
    /// let autoconvert = true;
    /// let include_tree = false;
    /// let process_id = std::process::id();
    ///
    /// initialize_mta().ok().unwrap(); // Don't do this on a UI thread
    /// let mut audio_client = AudioClient::new_application_loopback_client(process_id, include_tree).unwrap();
    /// audio_client.initialize_client(&desired_format, hnsbufferduration, &Direction::Capture, &ShareMode::Shared, autoconvert).unwrap();
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
    pub fn get_periods(&self) -> WasapiRes<(i64, i64)> {
        let mut def_time = 0;
        let mut min_time = 0;
        unsafe {
            self.client
                .GetDevicePeriod(Some(&mut def_time), Some(&mut min_time))?
        };
        trace!("default period {}, min period {}", def_time, min_time);
        Ok((def_time, min_time))
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
        let (_default_period, min_period) = self.get_periods()?;
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

    /// Initialize an [IAudioClient] for the given direction, sharemode and format.
    /// Setting `convert` to true enables automatic samplerate and format conversion, meaning that almost any format will be accepted.
    pub fn initialize_client(
        &mut self,
        wavefmt: &WaveFormat,
        period: i64,
        direction: &Direction,
        sharemode: &ShareMode,
        convert: bool,
    ) -> WasapiRes<()> {
        if sharemode == &ShareMode::Exclusive && convert {
            return Err(WasapiError::AutomaticFormatConversionInExclusiveMode);
        }
        let mut streamflags = match (&self.direction, direction, sharemode) {
            (Direction::Render, Direction::Capture, ShareMode::Shared) => {
                AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_LOOPBACK
            }
            (Direction::Render, Direction::Capture, ShareMode::Exclusive) => {
                return Err(WasapiError::LoopbackWithExclusiveMode);
            }
            (Direction::Capture, Direction::Render, _) => {
                return Err(WasapiError::RenderToCaptureDevice);
            }
            _ => AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        };
        if convert {
            streamflags |=
                AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY;
        }
        let mode = match sharemode {
            ShareMode::Exclusive => AUDCLNT_SHAREMODE_EXCLUSIVE,
            ShareMode::Shared => AUDCLNT_SHAREMODE_SHARED,
        };
        let device_period = match sharemode {
            ShareMode::Exclusive => period,
            ShareMode::Shared => 0,
        };
        self.sharemode = Some(*sharemode);
        unsafe {
            self.client.Initialize(
                mode,
                streamflags,
                period,
                device_period,
                wavefmt.as_waveformatex_ref(),
                None,
            )?;
        }
        self.bytes_per_frame = Some(wavefmt.get_blockalign() as usize);
        Ok(())
    }

    /// Create and return an event handle for an [IAudioClient]
    pub fn set_get_eventhandle(&self) -> WasapiRes<Handle> {
        let h_event = unsafe { CreateEventA(None, false, false, PCSTR::null())? };
        unsafe { self.client.SetEventHandle(h_event)? };
        Ok(Handle { handle: h_event })
    }

    /// Get buffer size in frames
    pub fn get_bufferframecount(&self) -> WasapiRes<u32> {
        let buffer_frame_count = unsafe { self.client.GetBufferSize()? };
        trace!("buffer_frame_count {}", buffer_frame_count);
        Ok(buffer_frame_count)
    }

    /// Get current padding in frames.
    /// This represents the number of frames currently in the buffer, for both capture and render devices.
    pub fn get_current_padding(&self) -> WasapiRes<u32> {
        let padding_count = unsafe { self.client.GetCurrentPadding()? };
        trace!("padding_count {}", padding_count);
        Ok(padding_count)
    }

    /// Get buffer size minus padding in frames.
    /// Use this to find out how much free space is available in the buffer.
    pub fn get_available_space_in_frames(&self) -> WasapiRes<u32> {
        let frames = match self.sharemode {
            Some(ShareMode::Exclusive) => {
                let buffer_frame_count = unsafe { self.client.GetBufferSize()? };
                trace!("buffer_frame_count {}", buffer_frame_count);
                buffer_frame_count
            }
            Some(ShareMode::Shared) => {
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
        Ok(AudioSessionControl {
            control: Rc::new(control),
        })
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
}

/// Struct wrapping an [IAudioSessionControl](https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nn-audiopolicy-iaudiosessioncontrol).
pub struct AudioSessionControl {
    control: Rc<IAudioSessionControl>,
}

impl AudioSessionControl {
    /// Get the current state
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
    /// Returns a [SessionEvents] struct.
    /// The notifications are unregistered when this struct is dropped.
    pub fn register_session_notification(
        &self,
        callbacks: std::sync::Weak<EventCallbacks>,
    ) -> WasapiRes<SessionEvents> {
        let events: IAudioSessionEvents = AudioSessionEvents::new(callbacks).into();

        match unsafe { self.control.RegisterAudioSessionNotification(&events) } {
            Ok(()) => Ok(SessionEvents {
                events,
                control: self.control.downgrade().unwrap(),
            }),
            Err(err) => Err(WasapiError::RegisterNotifications(err)),
        }
    }
}

/// Struct for keeping track of the registered notifications.
pub struct SessionEvents {
    events: IAudioSessionEvents,
    control: windows_core::Weak<IAudioSessionControl>,
}

impl Drop for SessionEvents {
    fn drop(&mut self) {
        if let Some(control) = self.control.upgrade() {
            let _ = unsafe { control.UnregisterAudioSessionNotification(&self.events) };
        }
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
    /// `get_available_space_in_frames()` method on the [AudioClient].
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
    /// `get_available_space_in_frames()` method on the [AudioClient].
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
    /// In exclusive mode it returns None, instead use [AudioClient::get_bufferframecount()].
    pub fn get_next_nbr_frames(&self) -> WasapiRes<Option<u32>> {
        if let Some(ShareMode::Exclusive) = self.sharemode {
            return Ok(None);
        }
        let nbr_frames = unsafe { self.client.GetNextPacketSize()? };
        Ok(Some(nbr_frames))
    }

    /// Read raw bytes from a device into a slice. Returns the number of frames
    /// that was read, and the BufferFlags describing the buffer that the data was read from.
    /// The slice must be large enough to hold all data.
    /// If it is longer that needed, the unused elements will not be modified.
    pub fn read_from_device(&self, data: &mut [u8]) -> WasapiRes<(u32, BufferFlags)> {
        let data_len_in_frames = data.len() / self.bytes_per_frame;
        if data_len_in_frames == 0 {
            return Ok((0, BufferFlags::none()));
        }
        let mut buffer_ptr = ptr::null_mut();
        let mut nbr_frames_returned = 0;
        let mut flags = 0;
        unsafe {
            self.client.GetBuffer(
                &mut buffer_ptr,
                &mut nbr_frames_returned,
                &mut flags,
                None,
                None,
            )?
        };
        let bufferflags = BufferFlags::new(flags);
        if nbr_frames_returned == 0 {
            unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
            return Ok((0, bufferflags));
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
        Ok((nbr_frames_returned, bufferflags))
    }

    /// Read raw bytes data from a device into a deque.
    /// Returns the [BufferFlags] describing the buffer that the data was read from.
    pub fn read_from_device_to_deque(&self, data: &mut VecDeque<u8>) -> WasapiRes<BufferFlags> {
        let mut buffer_ptr = ptr::null_mut();
        let mut nbr_frames_returned = 0;
        let mut flags = 0;
        unsafe {
            self.client.GetBuffer(
                &mut buffer_ptr,
                &mut nbr_frames_returned,
                &mut flags,
                None,
                None,
            )?
        };
        let bufferflags = BufferFlags::new(flags);
        if nbr_frames_returned == 0 {
            // There is no need to release a buffer of 0 bytes
            return Ok(bufferflags);
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
        Ok(bufferflags)
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
