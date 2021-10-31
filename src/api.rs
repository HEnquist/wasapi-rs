use crate::Windows;
use crate::Windows::{
    Win32::Foundation::{HANDLE, PSTR},
    Win32::Media::Audio::CoreAudio::{
        eCapture, eConsole, eRender, AudioSessionStateActive, AudioSessionStateExpired,
        AudioSessionStateInactive, IAudioCaptureClient, IAudioClient, IAudioRenderClient,
        IAudioSessionControl, IAudioSessionEvents, IMMDevice, IMMDeviceCollection,
        IMMDeviceEnumerator, MMDeviceEnumerator, AUDCLNT_SHAREMODE_EXCLUSIVE,
        AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
        AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_LOOPBACK,
        AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY, DEVICE_STATE_ACTIVE, WAVE_FORMAT_EXTENSIBLE,
    },
    Win32::Media::Multimedia::{WAVEFORMATEX, WAVEFORMATEXTENSIBLE},
    Win32::Storage::StructuredStorage::STGM_READ,
    Win32::System::Com::CLSCTX_ALL,
    Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED,
    },
    Win32::System::PropertiesSystem::PropVariantToStringAlloc,
    Win32::System::Threading::{CreateEventA, WaitForSingleObject},
};
use std::collections::VecDeque;
use std::error;
use std::fmt;
use std::mem;
use std::ptr;
use std::rc::Weak;
use std::slice;
use widestring::U16CString;
use windows::runtime::Interface;
use windows::Win32::System::Threading::WAIT_OBJECT_0;
use Windows::Win32::System::SystemServices::{
    DEVPKEY_Device_DeviceDesc, DEVPKEY_Device_FriendlyName,
};

use crate::WaveFormat;
use crate::{AudioSessionEvents, EventCallbacks};

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
pub fn initialize_mta() -> Result<(), windows::runtime::Error> {
    unsafe { CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED) }
}

/// Initializes COM for use by the calling thread for a single-threaded apartment (STA).
pub fn initialize_sta() -> Result<(), windows::runtime::Error> {
    unsafe { CoInitializeEx(std::ptr::null_mut(), COINIT_APARTMENTTHREADED) }
}

/// Audio direction, playback or capture.
#[derive(Clone)]
pub enum Direction {
    Render,
    Capture,
}

/// Sharemode for device
#[derive(Clone)]
pub enum ShareMode {
    Shared,
    Exclusive,
}

/// Sample type, float or integer
#[derive(Clone)]
pub enum SampleType {
    Float,
    Int,
}

/// Get the default playback or capture device
pub fn get_default_device(direction: &Direction) -> WasapiRes<Device> {
    let dir = match direction {
        Direction::Capture => eCapture,
        Direction::Render => eRender,
    };

    let enumerator: IMMDeviceEnumerator =
        unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)? };
    let device = unsafe { enumerator.GetDefaultAudioEndpoint(dir, eConsole)? };
    debug!("default device {:?}", device);

    //match device {
    //    Some(dev) => Ok(Device {
    //        device: dev,
    //        direction: direction.clone(),
    //    }),
    //    None => Err(WasapiError::new("Failed to get default device").into()),
    //}
    Ok(Device {
        device,
        direction: direction.clone(),
    })
}

/// Struct wrapping an IMMDeviceCollection.
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
}

/// Struct wrapping an IMMDevice.
pub struct Device {
    device: IMMDevice,
    direction: Direction,
}

impl Device {
    /// Get an IAudioClient from an IMMDevice
    pub fn get_iaudioclient(&self) -> WasapiRes<AudioClient> {
        let mut audio_client: mem::MaybeUninit<IAudioClient> = mem::MaybeUninit::zeroed();
        unsafe {
            self.device.Activate(
                &IAudioClient::IID,
                CLSCTX_ALL.0,
                ptr::null_mut(),
                audio_client.as_mut_ptr() as *mut _,
            )?;
            Ok(AudioClient {
                client: audio_client.assume_init(),
                direction: self.direction.clone(),
                sharemode: None,
            })
        }
    }

    /// Read state from an IMMDevice
    pub fn get_state(&self) -> WasapiRes<u32> {
        let state: u32 = unsafe { self.device.GetState()? };
        trace!("state: {:?}", state);
        Ok(state)
    }

    /// Read the FriendlyName of an IMMDevice
    pub fn get_friendlyname(&self) -> WasapiRes<String> {
        let store = unsafe { self.device.OpenPropertyStore(STGM_READ as u32)? };
        let prop = unsafe { store.GetValue(&DEVPKEY_Device_FriendlyName)? };
        let propstr = unsafe { PropVariantToStringAlloc(&prop)? };
        let wide_name = unsafe { U16CString::from_ptr_str(propstr.0) };
        let name = wide_name.to_string_lossy();
        trace!("name: {}", name);
        Ok(name)
    }

    /// Read the Description of an IMMDevice
    pub fn get_description(&self) -> WasapiRes<String> {
        let store = unsafe { self.device.OpenPropertyStore(STGM_READ as u32)? };
        let prop = unsafe { store.GetValue(&DEVPKEY_Device_DeviceDesc)? };
        let propstr = unsafe { PropVariantToStringAlloc(&prop)? };
        let wide_desc = unsafe { U16CString::from_ptr_str(propstr.0) };
        let desc = wide_desc.to_string_lossy();
        trace!("description: {}", desc);
        Ok(desc)
    }

    /// Get the Id of an IMMDevice
    pub fn get_id(&self) -> WasapiRes<String> {
        let idstr = unsafe { self.device.GetId()? };
        let wide_id = unsafe { U16CString::from_ptr_str(idstr.0) };
        let id = wide_id.to_string_lossy();
        trace!("id: {}", id);
        Ok(id)
    }
}

/// Struct wrapping an IAudioClient.
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
    /// If it's directly supported, this returns Ok(None). If not, but a similar format is, then the supported format is returned as Ok(Some(WaveFormat)).
    pub fn is_supported(
        &self,
        wave_fmt: &WaveFormat,
        sharemode: &ShareMode,
    ) -> WasapiRes<Option<WaveFormat>> {
        let supported = match sharemode {
            ShareMode::Exclusive => {
                unsafe {
                    self.client.IsFormatSupported(
                        AUDCLNT_SHAREMODE_EXCLUSIVE,
                        wave_fmt.as_waveformatex_ptr(),
                    )?
                };
                None
            }
            ShareMode::Shared => {
                let supported_format = unsafe {
                    self.client
                        .IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, wave_fmt.as_waveformatex_ptr())
                }?;
                // Check if we got a pointer to a WAVEFORMATEX structure.
                if supported_format.is_null() {
                    // The pointer is still null, thus the format is supported as is.
                    debug!("requested format is directly supported");
                    None
                } else {
                    // Read the structure
                    let temp_fmt: WAVEFORMATEX = unsafe { supported_format.read() };
                    debug!("requested format is not directly supported");
                    let new_fmt = if temp_fmt.cbSize == 22
                        && temp_fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE
                    {
                        debug!("got the supported format as a WAVEFORMATEXTENSIBLE");
                        let temp_fmt_ext: WAVEFORMATEXTENSIBLE = unsafe {
                            (supported_format as *const _ as *const WAVEFORMATEXTENSIBLE).read()
                        };
                        WaveFormat {
                            wave_fmt: temp_fmt_ext,
                        }
                    } else {
                        debug!("got the supported format as a WAVEFORMATEX, converting..");
                        WaveFormat::from_waveformatex(temp_fmt)?
                    };
                    Some(new_fmt)
                }
            }
        };
        Ok(supported)
    }

    /// Get default and minimum periods in 100-nanosecond units
    pub fn get_periods(&self) -> WasapiRes<(i64, i64)> {
        let mut def_time = 0;
        let mut min_time = 0;
        unsafe { self.client.GetDevicePeriod(&mut def_time, &mut min_time)? };
        trace!("default period {}, min period {}", def_time, min_time);
        Ok((def_time, min_time))
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
        self.sharemode = Some(sharemode.clone());
        unsafe {
            self.client.Initialize(
                mode,
                streamflags,
                period,
                period,
                wavefmt.as_waveformatex_ptr(),
                std::ptr::null(),
            )?;
        }
        Ok(())
    }

    /// Create and return an event handle for an IAudioClient
    pub fn set_get_eventhandle(&self) -> WasapiRes<Handle> {
        let h_event = unsafe { CreateEventA(std::ptr::null_mut(), false, false, PSTR::default()) };
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
        let mut renderclient_ptr = ptr::null_mut();
        unsafe {
            self.client
                .GetService(&IAudioRenderClient::IID, &mut renderclient_ptr)?
        };
        if renderclient_ptr.is_null() {
            return Err(WasapiError::new("Failed getting IAudioCaptureClient").into());
        }
        let client = unsafe { mem::transmute(renderclient_ptr) };
        Ok(AudioRenderClient { client })
    }

    /// Get a capture client
    pub fn get_audiocaptureclient(&self) -> WasapiRes<AudioCaptureClient> {
        let mut renderclient_ptr = ptr::null_mut();
        unsafe {
            self.client
                .GetService(&IAudioCaptureClient::IID, &mut renderclient_ptr)?
        };
        if renderclient_ptr.is_null() {
            return Err(WasapiError::new("Failed getting IAudioCaptureClient").into());
        }
        let client = unsafe { mem::transmute(renderclient_ptr) };
        Ok(AudioCaptureClient {
            client,
            sharemode: self.sharemode.clone(),
        })
    }

    /// Get the AudioSessionControl
    pub fn get_audiosessioncontrol(&self) -> WasapiRes<AudioSessionControl> {
        let mut sessioncontrol_ptr = ptr::null_mut();
        unsafe {
            self.client
                .GetService(&IAudioSessionControl::IID, &mut sessioncontrol_ptr)?
        };
        if sessioncontrol_ptr.is_null() {
            return Err(WasapiError::new("Failed getting IAudioSessionControl").into());
        }
        let control = unsafe { mem::transmute(sessioncontrol_ptr) };
        Ok(AudioSessionControl { control })
    }
}

/// Struct wrapping an IAudioSessionControl.
pub struct AudioSessionControl {
    control: IAudioSessionControl,
}

/// States of an AudioSession
#[derive(Debug)]
pub enum SessionState {
    Active,
    Inactive,
    Expired,
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

        match unsafe { self.control.RegisterAudioSessionNotification(events) } {
            Ok(()) => Ok(()),
            Err(err) => {
                Err(WasapiError::new(&format!("Failed to register notifications, {}", err)).into())
            }
        }
    }
}

/// Struct wrapping an IAudioRenderClient.
pub struct AudioRenderClient {
    client: IAudioRenderClient,
}

impl AudioRenderClient {
    /// Write raw bytes data to a device from a slice.
    /// The number of frames to write should first be checked with the
    /// `get_available_space_in_frames()` method on the `AudioClient`.
    pub fn write_to_device(
        &self,
        nbr_frames: usize,
        byte_per_frame: usize,
        data: &[u8],
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
        unsafe { self.client.ReleaseBuffer(nbr_frames as u32, 0)? };
        trace!("wrote {} frames", nbr_frames);
        Ok(())
    }

    /// Write raw bytes data to a device from a deque.
    /// The number of frames to write should first be checked with the
    /// `get_available_space_in_frames()` method on the `AudioClient`.
    pub fn write_to_device_from_deque(
        &self,
        nbr_frames: usize,
        byte_per_frame: usize,
        data: &mut VecDeque<u8>,
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
        unsafe { self.client.ReleaseBuffer(nbr_frames as u32, 0)? };
        trace!("wrote {} frames", nbr_frames);
        Ok(())
    }
}

/// Struct wrapping an IAudioCaptureClient.
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

    /// Read raw bytes from a device into a slice, returns the number of frames read.
    /// The slice must be large enough to hold all data.
    /// If it is longer that needed, the unused elements will not be modified.
    pub fn read_from_device(&self, bytes_per_frame: usize, data: &mut [u8]) -> WasapiRes<u32> {
        let data_len_in_frames = data.len() / bytes_per_frame;
        let mut buffer = mem::MaybeUninit::uninit();
        let mut nbr_frames_returned = 0;
        unsafe {
            self.client.GetBuffer(
                buffer.as_mut_ptr(),
                &mut nbr_frames_returned,
                &mut 0,
                ptr::null_mut(),
                ptr::null_mut(),
            )?
        };
        if nbr_frames_returned == 0 {
            unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
            return Ok(0);
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
        let bufferptr = unsafe { buffer.assume_init() };
        let bufferslice = unsafe { slice::from_raw_parts(bufferptr, len_in_bytes) };
        data[..len_in_bytes].copy_from_slice(bufferslice);
        unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
        trace!("read {} frames", nbr_frames_returned);
        Ok(nbr_frames_returned)
    }

    /// Read raw bytes data from a device into a deque.
    pub fn read_from_device_to_deque(
        &self,
        bytes_per_frame: usize,
        data: &mut VecDeque<u8>,
    ) -> WasapiRes<()> {
        let mut buffer = mem::MaybeUninit::uninit();
        let mut nbr_frames_returned = 0;
        unsafe {
            self.client.GetBuffer(
                buffer.as_mut_ptr(),
                &mut nbr_frames_returned,
                &mut 0,
                ptr::null_mut(),
                ptr::null_mut(),
            )?
        };
        let len_in_bytes = nbr_frames_returned as usize * bytes_per_frame;
        let bufferptr = unsafe { buffer.assume_init() };
        let bufferslice = unsafe { slice::from_raw_parts(bufferptr, len_in_bytes) };
        for element in bufferslice.iter() {
            data.push_back(*element);
        }
        unsafe { self.client.ReleaseBuffer(nbr_frames_returned)? };
        trace!("read {} frames", nbr_frames_returned);
        Ok(())
    }
}

/// Struct wrapping a HANDLE (event handle).
pub struct Handle {
    handle: HANDLE,
}

impl Handle {
    /// Wait for an event on a handle, with a timeout given in ms
    pub fn wait_for_event(&self, timeout_ms: u32) -> WasapiRes<()> {
        let retval = unsafe { WaitForSingleObject(self.handle, timeout_ms) };
        if retval != WAIT_OBJECT_0 {
            return Err(WasapiError::new("Wait timed out").into());
        }
        Ok(())
    }
}
