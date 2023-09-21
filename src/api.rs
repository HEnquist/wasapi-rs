use num_integer::Integer;
use std::cmp;
use std::collections::VecDeque;
use std::rc::Weak;
use std::{error, fmt, ptr, slice};
use widestring::U16CString;
use windows::Win32::UI::Shell::PropertiesSystem::PROPERTYKEY;
use windows::{
    core::PCSTR,
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
        WAVEFORMATEX, WAVEFORMATEXTENSIBLE,
    },
    Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE,
    Win32::System::Com::STGM_READ,
    Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
        COINIT_MULTITHREADED,
    },
    Win32::System::Threading::{CreateEventA, WaitForSingleObject},
    Win32::UI::Shell::PropertiesSystem::PropVariantToStringAlloc,
};

use crate::{make_channelmasks, AudioSessionEvents, EventCallbacks, WaveFormat};

pub(crate) type WasapiRes<T> = Result<T, Box<dyn error::Error>>;

/// Error returned by the Wasapi crate.
#[derive(Debug)]
pub struct WasapiError {
    desc: String,
}

impl fmt::Display for WasapiError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.desc)
    }
}

impl error::Error for WasapiError {
    fn description(&self) -> &str {
        &self.desc
    }
}

impl WasapiError {
    pub fn new(desc: &str) -> Self {
        WasapiError {
            desc: desc.to_owned(),
        }
    }
}

/// Initializes COM for use by the calling thread for the multi-threaded apartment (MTA).
pub fn initialize_mta() -> Result<(), windows::core::Error> {
    unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) }
}

/// Initializes COM for use by the calling thread for a single-threaded apartment (STA).
pub fn initialize_sta() -> Result<(), windows::core::Error> {
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

/// Audio direction, playback or capture.
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

/// States of an AudioSession
#[derive(Debug, Eq, PartialEq)]
pub enum SessionState {
    Active,
    Inactive,
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
/// Get the default playback or capture device for Console purposes, which is usually what's wanted
/// https://learn.microsoft.com/en-us/windows/win32/api/mmdeviceapi/ne-mmdeviceapi-erole
pub fn get_default_device(direction: &Direction) -> WasapiRes<Device> {
    get_default_device_for_role(direction, &Role::Console)
}

/// Get the default playback or capture device for a specific role
pub fn get_default_device_for_role(direction: &Direction, role: &Role) -> WasapiRes<Device> {
    let dir = match direction {
        Direction::Capture => eCapture,
        Direction::Render => eRender,
    };

    let e_role = match role {
        Role::Console => eConsole,
        Role::Multimedia => eMultimedia,
        Role::Communications => eCommunications,
    };

    let enumerator: IMMDeviceEnumerator =
        unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)? };
    let device = unsafe { enumerator.GetDefaultAudioEndpoint(dir, e_role)? };

    let dev = Device {
        device,
        direction: direction.clone(),
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
    /// Get an IMMDeviceCollection of all active playback or capture devices
    pub fn new(direction: &Direction) -> WasapiRes<DeviceCollection> {
        let dir = match direction {
            Direction::Capture => eCapture,
            Direction::Render => eRender,
        };
        let enumerator: IMMDeviceEnumerator =
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)? };
        let devs = unsafe { enumerator.EnumAudioEndpoints(dir, DEVICE_STATE_ACTIVE)? };
        Ok(DeviceCollection {
            collection: devs,
            direction: direction.clone(),
        })
    }

    /// Get the number of devices in an IMMDeviceCollection
    pub fn get_nbr_devices(&self) -> WasapiRes<u32> {
        let count = unsafe { self.collection.GetCount()? };
        Ok(count)
    }

    /// Get a device from an IMMDeviceCollection using index
    pub fn get_device_at_index(&self, idx: u32) -> WasapiRes<Device> {
        let device = unsafe { self.collection.Item(idx)? };
        Ok(Device {
            device,
            direction: self.direction.clone(),
        })
    }

    /// Get a device from an IMMDeviceCollection using name
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
        Err(WasapiError::new(format!("Unable to find device {}", name).as_str()).into())
    }

    /// Get the direction for this DeviceCollection
    pub fn get_direction(&self) -> Direction {
        self.direction
    }
}

/// Iterator for DeviceCollection
pub struct DeviceCollectionIter {
    collection: DeviceCollection,
    index: u32,
}

impl Iterator for DeviceCollectionIter {
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

/// Implement iterator for DeviceCollection
impl IntoIterator for DeviceCollection {
    type Item = WasapiRes<Device>;
    type IntoIter = DeviceCollectionIter;

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
    /// Get an IAudioClient from an IMMDevice
    pub fn get_iaudioclient(&self) -> WasapiRes<AudioClient> {
        let audio_client = unsafe { self.device.Activate::<IAudioClient>(CLSCTX_ALL, None)? };
        Ok(AudioClient {
            client: audio_client,
            direction: self.direction.clone(),
            sharemode: None,
        })
    }

    /// Read state from an IMMDevice
    pub fn get_state(&self) -> WasapiRes<u32> {
        let state: u32 = unsafe { self.device.GetState()? };
        trace!("state: {:?}", state);
        Ok(state)
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

    /// Read the FriendlyName of an IMMDevice
    fn get_string_property(&self, key: &PROPERTYKEY) -> WasapiRes<String> {
        let store = unsafe { self.device.OpenPropertyStore(STGM_READ)? };
        let prop = unsafe { store.GetValue(key)? };
        let propstr = unsafe { PropVariantToStringAlloc(&prop)? };
        let wide_name = unsafe { U16CString::from_ptr_str(propstr.0) };
        let name = wide_name.to_string_lossy();
        trace!("name: {}", name);
        Ok(name)
    }

    /// Get the Id of an IMMDevice
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

/// Struct wrapping an [IAudioClient](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iaudioclient).
pub struct AudioClient {
    client: IAudioClient,
    direction: Direction,
    sharemode: Option<ShareMode>,
}

impl AudioClient {
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
        Err(WasapiError::new("Unable to find a supported format").into())
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

    /// Initialize an IAudioClient for the given direction, sharemode and format.
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
            return Err(
                WasapiError::new("Cant use automatic format conversion in exclusive mode").into(),
            );
        }
        let mut streamflags = match (&self.direction, direction, sharemode) {
            (Direction::Render, Direction::Capture, ShareMode::Shared) => {
                AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_LOOPBACK
            }
            (Direction::Render, Direction::Capture, ShareMode::Exclusive) => {
                return Err(WasapiError::new("Cant use Loopback with exclusive mode").into());
            }
            (Direction::Capture, Direction::Render, _) => {
                return Err(WasapiError::new("Cant render to a capture device").into());
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
        self.sharemode = Some(sharemode.clone());
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
        Ok(())
    }

    /// Create and return an event handle for an IAudioClient
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
            _ => return Err(WasapiError::new("Client has not been initialized").into()),
        };
        Ok(frames)
    }

    /// Start the stream on an IAudioClient
    pub fn start_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Start()? };
        Ok(())
    }

    /// Stop the stream on an IAudioClient
    pub fn stop_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Stop()? };
        Ok(())
    }

    /// Reset the stream on an IAudioClient
    pub fn reset_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Reset()? };
        Ok(())
    }

    /// Get a rendering (playback) client
    pub fn get_audiorenderclient(&self) -> WasapiRes<AudioRenderClient> {
        let client = unsafe { self.client.GetService::<IAudioRenderClient>()? };
        Ok(AudioRenderClient { client })
    }

    /// Get a capture client
    pub fn get_audiocaptureclient(&self) -> WasapiRes<AudioCaptureClient> {
        let client = unsafe { self.client.GetService::<IAudioCaptureClient>()? };
        Ok(AudioCaptureClient {
            client,
            sharemode: self.sharemode.clone(),
        })
    }

    /// Get the AudioSessionControl
    pub fn get_audiosessioncontrol(&self) -> WasapiRes<AudioSessionControl> {
        let control = unsafe { self.client.GetService::<IAudioSessionControl>()? };
        Ok(AudioSessionControl { control })
    }

    /// Get the AudioClock
    pub fn get_audioclock(&self) -> WasapiRes<AudioClock> {
        let clock = unsafe { self.client.GetService::<IAudioClock>()? };
        Ok(AudioClock { clock })
    }

    /// Get the direction for this AudioClient
    pub fn get_direction(&self) -> Direction {
        self.direction
    }

    /// Get the sharemode for this AudioClient.
    /// The sharemode is decided when the client is initialized.
    pub fn get_sharemode(&self) -> Option<ShareMode> {
        self.sharemode
    }
}

/// Struct wrapping an [IAudioSessionControl](https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nn-audiopolicy-iaudiosessioncontrol).
pub struct AudioSessionControl {
    control: IAudioSessionControl,
}

impl AudioSessionControl {
    /// Get the current state
    pub fn get_state(&self) -> WasapiRes<SessionState> {
        let state = unsafe { self.control.GetState()? };
        #[allow(non_upper_case_globals)]
        let sessionstate = match state {
            AudioSessionStateActive => SessionState::Active,
            AudioSessionStateInactive => SessionState::Inactive,
            AudioSessionStateExpired => SessionState::Expired,
            _ => {
                return Err(WasapiError::new("Got an illegal state").into());
            }
        };
        Ok(sessionstate)
    }

    /// Register to receive notifications
    pub fn register_session_notification(&self, callbacks: Weak<EventCallbacks>) -> WasapiRes<()> {
        let events: IAudioSessionEvents = AudioSessionEvents::new(callbacks).into();

        match unsafe { self.control.RegisterAudioSessionNotification(&events) } {
            Ok(()) => Ok(()),
            Err(err) => {
                Err(WasapiError::new(&format!("Failed to register notifications, {}", err)).into())
            }
        }
    }
}

/// Struct wrapping an [IAudioClock](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-iaudioclock).
pub struct AudioClock {
    clock: IAudioClock,
}

impl AudioClock {
    /// Get the frequency for this AudioClock.
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
}

impl AudioRenderClient {
    /// Write raw bytes data to a device from a slice.
    /// The number of frames to write should first be checked with the
    /// `get_available_space_in_frames()` method on the `AudioClient`.
    /// The buffer_flags argument can be used to mark a buffer as silent.
    pub fn write_to_device(
        &self,
        nbr_frames: usize,
        byte_per_frame: usize,
        data: &[u8],
        buffer_flags: Option<BufferFlags>,
    ) -> WasapiRes<()> {
        let nbr_bytes = nbr_frames * byte_per_frame;
        if nbr_bytes != data.len() {
            return Err(WasapiError::new(
                format!(
                    "Wrong length of data, got {}, expected {}",
                    data.len(),
                    nbr_bytes
                )
                .as_str(),
            )
            .into());
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
    /// `get_available_space_in_frames()` method on the `AudioClient`.
    /// The buffer_flags argument can be used to mark a buffer as silent.
    pub fn write_to_device_from_deque(
        &self,
        nbr_frames: usize,
        byte_per_frame: usize,
        data: &mut VecDeque<u8>,
        buffer_flags: Option<BufferFlags>,
    ) -> WasapiRes<()> {
        let nbr_bytes = nbr_frames * byte_per_frame;
        if nbr_bytes > data.len() {
            return Err(WasapiError::new(
                format!("To little data, got {}, need {}", data.len(), nbr_bytes).as_str(),
            )
            .into());
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

/// Struct representing the [ _AUDCLNT_BUFFERFLAGS enums](https://docs.microsoft.com/en-us/windows/win32/api/audioclient/ne-audioclient-_audclnt_bufferflags).
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
    /// Create a new BufferFlags struct from a u32 value.
    pub fn new(flags: u32) -> Self {
        BufferFlags {
            data_discontinuity: flags & AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY.0 as u32 > 0,
            silent: flags & AUDCLNT_BUFFERFLAGS_SILENT.0 as u32 > 0,
            timestamp_error: flags & AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR.0 as u32 > 0,
        }
    }

    /// Convert a BufferFlags struct to a u32 value.
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
}

impl AudioCaptureClient {
    /// Get number of frames in next packet when in shared mode.
    /// In exclusive mode it returns None, instead use `get_bufferframecount()` on the AudioClient.
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
    pub fn read_from_device(
        &self,
        bytes_per_frame: usize,
        data: &mut [u8],
    ) -> WasapiRes<(u32, BufferFlags)> {
        let data_len_in_frames = data.len() / bytes_per_frame;
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
            return Err(WasapiError::new(
                format!(
                    "Wrong length of data, got {} frames, expected at least {} frames",
                    data_len_in_frames, nbr_frames_returned
                )
                .as_str(),
            )
            .into());
        }
        let len_in_bytes = nbr_frames_returned as usize * bytes_per_frame;
        let bufferslice = unsafe { slice::from_raw_parts(buffer_ptr, len_in_bytes) };
        data[..len_in_bytes].copy_from_slice(bufferslice);
        unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
        trace!("read {} frames", nbr_frames_returned);
        Ok((nbr_frames_returned, bufferflags))
    }

    /// Read raw bytes data from a device into a deque.
    /// Returns the BufferFlags describing the buffer that the data was read from.
    pub fn read_from_device_to_deque(
        &self,
        bytes_per_frame: usize,
        data: &mut VecDeque<u8>,
    ) -> WasapiRes<BufferFlags> {
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
        let len_in_bytes = nbr_frames_returned as usize * bytes_per_frame;
        let bufferslice = unsafe { slice::from_raw_parts(buffer_ptr, len_in_bytes) };
        for element in bufferslice.iter() {
            data.push_back(*element);
        }
        unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
        trace!("read {} frames", nbr_frames_returned);
        Ok(bufferflags)
    }

    /// Get the sharemode for this AudioCaptureClient.
    /// The sharemode is decided when the client is initialized.
    pub fn get_sharemode(&self) -> Option<ShareMode> {
        self.sharemode
    }
}

/// Struct wrapping a HANDLE to an [Event Object](https://docs.microsoft.com/en-us/windows/win32/sync/event-objects).
pub struct Handle {
    handle: HANDLE,
}

impl Handle {
    /// Wait for an event on a handle, with a timeout given in ms
    pub fn wait_for_event(&self, timeout_ms: u32) -> WasapiRes<()> {
        let retval = unsafe { WaitForSingleObject(self.handle, timeout_ms) };
        if retval.0 != WAIT_OBJECT_0.0 {
            return Err(WasapiError::new("Wait timed out").into());
        }
        Ok(())
    }
}
