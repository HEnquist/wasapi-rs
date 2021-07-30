use crate::{
    PKEY_Device_DeviceDesc, PKEY_Device_FriendlyName,
    Windows::Win32::Foundation::{BOOL, E_NOINTERFACE, HANDLE, PSTR, PWSTR, S_OK},
    Windows::Win32::Media::Audio::CoreAudio::{
        eCapture, eConsole, eRender, AudioSessionDisconnectReason, AudioSessionState,
        AudioSessionStateActive, AudioSessionStateExpired, AudioSessionStateInactive,
        DisconnectReasonDeviceRemoval, DisconnectReasonExclusiveModeOverride,
        DisconnectReasonFormatChanged, DisconnectReasonServerShutdown,
        DisconnectReasonSessionDisconnected, DisconnectReasonSessionLogoff, IAudioCaptureClient,
        IAudioClient, IAudioRenderClient, IAudioSessionControl, IAudioSessionEvents,
        IAudioSessionEvents_abi, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator,
        MMDeviceEnumerator, AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM, AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        AUDCLNT_STREAMFLAGS_LOOPBACK, AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY, DEVICE_STATE_ACTIVE,
        WAVE_FORMAT_EXTENSIBLE,
    },
    Windows::Win32::Media::Multimedia::{
        KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, KSDATAFORMAT_SUBTYPE_PCM, WAVEFORMATEX,
        WAVEFORMATEXTENSIBLE, WAVEFORMATEXTENSIBLE_0, WAVE_FORMAT_IEEE_FLOAT, WAVE_FORMAT_PCM,
    },
    Windows::Win32::Storage::StructuredStorage::STGM_READ,
    Windows::Win32::System::Com::CLSCTX_ALL,
    Windows::Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED,
        COINIT_MULTITHREADED,
    },
    Windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc,
    Windows::Win32::System::Threading::{CreateEventA, WaitForSingleObject, WAIT_OBJECT_0},
};
use std::collections::VecDeque;
use std::error;
use std::fmt;
use std::mem;
use std::ptr;
use std::slice;
use widestring::U16CString;
use windows::Guid;
use windows::IUnknown;
use windows::Interface;
use windows::HRESULT;

type WasapiRes<T> = Result<T, Box<dyn error::Error>>;

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
pub fn initialize_mta() -> Result<(), windows::Error> {
    unsafe { CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED) }
}

/// Initializes COM for use by the calling thread for a single-threaded apartment (STA).
pub fn initialize_sta() -> Result<(), windows::Error> {
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
        unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER)? };
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
            unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER)? };
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
        let prop = unsafe { store.GetValue(&PKEY_Device_FriendlyName)? };
        let propstr = unsafe { PropVariantToStringAlloc(&prop)? };
        let wide_name = unsafe { U16CString::from_ptr_str(propstr.0) };
        let name = wide_name.to_string_lossy();
        trace!("name: {}", name);
        Ok(name)
    }

    /// Read the Description of an IMMDevice
    pub fn get_description(&self) -> WasapiRes<String> {
        let store = unsafe { self.device.OpenPropertyStore(STGM_READ as u32)? };
        let prop = unsafe { store.GetValue(&PKEY_Device_DeviceDesc)? };
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
        let mut mix_format: mem::MaybeUninit<*mut WAVEFORMATEX> = mem::MaybeUninit::zeroed();
        unsafe { self.client.GetMixFormat(mix_format.as_mut_ptr())? };
        let temp_fmt = unsafe { mix_format.assume_init().read() };
        let mix_format = if temp_fmt.cbSize == 22
            && temp_fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE
        {
            unsafe {
                WaveFormat {
                    wave_fmt: (mix_format.assume_init() as *const _ as *const WAVEFORMATEXTENSIBLE)
                        .read(),
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
                        ptr::null_mut(),
                    )?
                };
                None
            }
            ShareMode::Shared => {
                let mut supported_format: mem::MaybeUninit<*mut WAVEFORMATEX> =
                    mem::MaybeUninit::zeroed();
                unsafe {
                    self.client.IsFormatSupported(
                        AUDCLNT_SHAREMODE_SHARED,
                        wave_fmt.as_waveformatex_ptr(),
                        supported_format.as_mut_ptr(),
                    )
                }?;

                let temp_fmt = unsafe { supported_format.assume_init().read() };
                // Check if anything was written to the waveformatex structure
                if temp_fmt.cbSize == 0 && temp_fmt.wFormatTag == 0 {
                    // Nothing was written, thus the format is supported as is
                    debug!("requested format is directly supported");
                    None
                } else {
                    debug!("requested format is not directly supported");
                    let new_fmt = if temp_fmt.cbSize == 22
                        && temp_fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE
                    {
                        debug!("got a WAVEFORMATEXTENSIBLE");
                        unsafe {
                            WaveFormat {
                                wave_fmt: (supported_format.assume_init() as *const _
                                    as *const WAVEFORMATEXTENSIBLE)
                                    .read(),
                            }
                        }
                    } else {
                        debug!("got a WAVEFORMATEX, converting..");
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

    /// Create an return an event handle for an IAudioClient
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
        let renderclient: Option<IAudioRenderClient> = unsafe { self.client.GetService().ok() };
        match renderclient {
            Some(client) => Ok(AudioRenderClient { client }),
            None => Err(WasapiError::new("Failed getting IAudioRenderClient").into()),
        }
    }

    /// Get a capture client
    pub fn get_audiocaptureclient(&self) -> WasapiRes<AudioCaptureClient> {
        let renderclient: Option<IAudioCaptureClient> = unsafe { self.client.GetService().ok() };
        match renderclient {
            Some(client) => Ok(AudioCaptureClient {
                client,
                sharemode: self.sharemode.clone(),
            }),
            None => Err(WasapiError::new("Failed getting IAudioCaptureClient").into()),
        }
    }

    /// Get the AudioSessionControl
    pub fn get_audiosessioncontrol(&self) -> WasapiRes<AudioSessionControl> {
        let sessioncontrol: Option<IAudioSessionControl> = unsafe { self.client.GetService().ok() };
        match sessioncontrol {
            Some(control) => Ok(AudioSessionControl { control }),
            None => Err(WasapiError::new("Failed getting IAudioSessionControl").into()),
        }
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
    pub fn register_session_notification(&self, callbacks: &mut EventCallbacks) -> WasapiRes<()> {
        let events = AudioSessionEvents::new(callbacks);

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
        let mut buffer = mem::MaybeUninit::uninit();
        unsafe {
            self.client
                .GetBuffer(nbr_frames as u32, buffer.as_mut_ptr())?
        };
        let bufferptr = unsafe { buffer.assume_init() };
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
        let mut buffer = mem::MaybeUninit::uninit();
        unsafe {
            self.client
                .GetBuffer(nbr_frames as u32, buffer.as_mut_ptr())?
        };
        let bufferptr = unsafe { buffer.assume_init() };
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

/// Struct wrapping a WAVEFORMATEXTENSIBLE format descriptor.
#[derive(Clone)]
pub struct WaveFormat {
    wave_fmt: WAVEFORMATEXTENSIBLE,
}

impl fmt::Debug for WaveFormat {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("WaveFormat")
            .field("nAvgBytesPerSec", &{ self.wave_fmt.Format.nAvgBytesPerSec })
            .field("cbSize", &{ self.wave_fmt.Format.cbSize })
            .field("nBlockAlign", &{ self.wave_fmt.Format.nBlockAlign })
            .field("wBitsPerSample", &{ self.wave_fmt.Format.wBitsPerSample })
            .field("nSamplesPerSec", &{ self.wave_fmt.Format.nSamplesPerSec })
            .field("wFormatTag", &{ self.wave_fmt.Format.wFormatTag })
            .field("wValidBitsPerSample", &unsafe {
                self.wave_fmt.Samples.wValidBitsPerSample
            })
            .field("SubFormat", &{ self.wave_fmt.SubFormat })
            .field("nChannel", &{ self.wave_fmt.Format.nChannels })
            .field("dwChannelMask", &{ self.wave_fmt.dwChannelMask })
            .finish()
    }
}

impl WaveFormat {
    /// Build a WAVEFORMATEXTENSIBLE struct for the given parameters
    pub fn new(
        storebits: usize,
        validbits: usize,
        sample_type: &SampleType,
        samplerate: usize,
        channels: usize,
    ) -> Self {
        let blockalign = channels * storebits / 8;
        let byterate = samplerate * blockalign;

        let wave_format = WAVEFORMATEX {
            cbSize: 22,
            nAvgBytesPerSec: byterate as u32,
            nBlockAlign: blockalign as u16,
            nChannels: channels as u16,
            nSamplesPerSec: samplerate as u32,
            wBitsPerSample: storebits as u16,
            wFormatTag: WAVE_FORMAT_EXTENSIBLE as u16,
        };
        let sample = WAVEFORMATEXTENSIBLE_0 {
            wValidBitsPerSample: validbits as u16,
        };
        let subformat = match sample_type {
            SampleType::Float => KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
            SampleType::Int => KSDATAFORMAT_SUBTYPE_PCM,
        };
        let mut mask = 0;
        for n in 0..channels {
            mask += 1 << n;
        }
        let wave_fmt = WAVEFORMATEXTENSIBLE {
            Format: wave_format,
            Samples: sample,
            SubFormat: subformat,
            dwChannelMask: mask,
        };
        WaveFormat { wave_fmt }
    }

    /// Create from a WAVEFORMATEX structure
    pub fn from_waveformatex(wavefmt: WAVEFORMATEX) -> WasapiRes<Self> {
        let validbits = wavefmt.wBitsPerSample as usize;
        let blockalign = wavefmt.nBlockAlign as usize;
        let samplerate = wavefmt.nSamplesPerSec as usize;
        let formattag = wavefmt.wFormatTag;
        let channels = wavefmt.nChannels as usize;
        let sample_type = match formattag as u32 {
            WAVE_FORMAT_PCM => SampleType::Int,
            WAVE_FORMAT_IEEE_FLOAT => SampleType::Float,
            _ => {
                return Err(WasapiError::new("Unsupported format").into());
            }
        };
        let storebits = 8 * blockalign / channels;
        Ok(WaveFormat::new(
            storebits,
            validbits,
            &sample_type,
            samplerate,
            channels,
        ))
    }

    /// get a pointer of type WAVEFORMATEX, used internally
    pub fn as_waveformatex_ptr(&self) -> *const WAVEFORMATEX {
        &self.wave_fmt as *const _ as *const WAVEFORMATEX
    }

    /// Read nBlockAlign.
    pub fn get_blockalign(&self) -> u32 {
        self.wave_fmt.Format.nBlockAlign as u32
    }

    /// Read nAvgBytesPerSec.
    pub fn get_avgbytespersec(&self) -> u32 {
        self.wave_fmt.Format.nAvgBytesPerSec
    }

    /// Read wBitsPerSample.
    pub fn get_bitspersample(&self) -> u16 {
        self.wave_fmt.Format.wBitsPerSample
    }

    /// Read wValidBitsPerSample.
    pub fn get_validbitspersample(&self) -> u16 {
        unsafe { self.wave_fmt.Samples.wValidBitsPerSample }
    }

    /// Read nSamplesPerSec.
    pub fn get_samplespersec(&self) -> u32 {
        self.wave_fmt.Format.nSamplesPerSec
    }

    /// Read nChannels.
    pub fn get_nchannels(&self) -> u16 {
        self.wave_fmt.Format.nChannels
    }

    /// Read dwChannelMask.
    pub fn get_dwchannelmask(&self) -> u32 {
        self.wave_fmt.dwChannelMask
    }

    /// Read SubFormat.
    pub fn get_subformat(&self) -> WasapiRes<SampleType> {
        let subfmt = match self.wave_fmt.SubFormat {
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT => SampleType::Float,
            KSDATAFORMAT_SUBTYPE_PCM => SampleType::Int,
            _ => {
                return Err(WasapiError::new(
                    format!("Unknown subformat {:?}", { self.wave_fmt.SubFormat }).as_str(),
                )
                .into());
            }
        };
        Ok(subfmt)
    }
}

type OptionBox<T> = Option<Box<T>>;

/// A structure holding the callbacks for notifications
pub struct EventCallbacks {
    simple_volume: OptionBox<dyn FnMut(f32, bool, Guid)>,
    channel_volume: OptionBox<dyn FnMut(usize, f32, Guid)>,
    state: OptionBox<dyn FnMut(SessionState)>,
    disconnected: OptionBox<dyn FnMut(DisconnectReason)>,
    iconpath: OptionBox<dyn FnMut(String, Guid)>,
    displayname: OptionBox<dyn FnMut(String, Guid)>,
    groupingparam: OptionBox<dyn FnMut(Guid, Guid)>,
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
    pub fn set_simple_volume_callback(&mut self, c: impl FnMut(f32, bool, Guid) + 'static) {
        self.simple_volume = Some(Box::new(c));
    }
    /// Remove a callback for OnSimpleVolumeChanged notifications
    pub fn unset_simple_volume_callback(&mut self) {
        self.simple_volume = None;
    }

    /// Set a callback for OnChannelVolumeChanged notifications
    pub fn set_channel_volume_callback(&mut self, c: impl FnMut(usize, f32, Guid) + 'static) {
        self.channel_volume = Some(Box::new(c));
    }
    /// Remove a callback for OnChannelVolumeChanged notifications
    pub fn unset_channel_volume_callback(&mut self) {
        self.channel_volume = None;
    }

    /// Set a callback for OnSessionDisconnected notifications
    pub fn set_disconnected_callback(&mut self, c: impl FnMut(DisconnectReason) + 'static) {
        self.disconnected = Some(Box::new(c));
    }
    /// Remove a callback for OnSessionDisconnected notifications
    pub fn unset_disconnected_callback(&mut self) {
        self.disconnected = None;
    }

    /// Set a callback for OnStateChanged notifications
    pub fn set_state_callback(&mut self, c: impl FnMut(SessionState) + 'static) {
        self.state = Some(Box::new(c));
    }
    /// Remove a callback for OnStateChanged notifications
    pub fn unset_state_callback(&mut self) {
        self.state = None;
    }

    /// Set a callback for OnIconPathChanged notifications
    pub fn set_iconpath_callback(&mut self, c: impl FnMut(String, Guid) + 'static) {
        self.iconpath = Some(Box::new(c));
    }
    /// Remove a callback for OnIconPathChanged notifications
    pub fn unset_iconpath_callback(&mut self) {
        self.iconpath = None;
    }

    /// Set a callback for OnDisplayNameChanged notifications
    pub fn set_displayname_callback(&mut self, c: impl FnMut(String, Guid) + 'static) {
        self.displayname = Some(Box::new(c));
    }
    /// Remove a callback for OnDisplayNameChanged notifications
    pub fn unset_displayname_callback(&mut self) {
        self.displayname = None;
    }

    /// Set a callback for OnGroupingParamChanged notifications
    pub fn set_groupingparam_callback(&mut self, c: impl FnMut(Guid, Guid) + 'static) {
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
struct AudioSessionEvents<'a> {
    _abi: Box<IAudioSessionEvents_abi>,
    ref_cnt: u32,
    callbacks: &'a mut EventCallbacks,
}

#[allow(dead_code)]
impl<'a> AudioSessionEvents<'a> {
    /// Create a new AudioSessionEvents instance, returned as a IAudioSessionEvent.
    #[allow(clippy::new_ret_no_self)]
    pub fn new(callbacks: &'a mut EventCallbacks) -> IAudioSessionEvents {
        let target = Box::new(Self {
            _abi: Box::new(IAudioSessionEvents_abi(
                Self::_query_interface,
                Self::_add_ref,
                Self::_release,
                Self::_on_display_name_changed,
                Self::_on_icon_path_changed,
                Self::_on_simple_volume_changed,
                Self::_on_channel_volume_changed,
                Self::_on_grouping_param_changed,
                Self::_on_state_changed,
                Self::_on_session_disconnected,
            )),
            ref_cnt: 1,
            callbacks,
        });

        unsafe {
            let ptr = Box::into_raw(target);
            mem::transmute(ptr)
        }
    }

    fn query_interface(
        &mut self,
        iid: &::windows::Guid,
        interface: *mut ::windows::RawPtr,
    ) -> HRESULT {
        if iid == &IAudioSessionEvents::IID || iid == &IUnknown::IID {
            unsafe {
                *interface = self as *mut Self as *mut _;
            }
            self.add_ref();
            S_OK
        } else {
            E_NOINTERFACE
        }
    }

    fn add_ref(&mut self) -> u32 {
        self.ref_cnt += 1;
        self.ref_cnt
    }

    fn release(&mut self) -> u32 {
        self.ref_cnt -= 1;
        let res = self.ref_cnt;

        if res == 0 {
            unsafe {
                Box::from_raw(self as *mut Self);
            }
        }

        res
    }

    fn on_state_changed(&mut self, newstate: AudioSessionState) -> HRESULT {
        trace!("state change: {:?}", newstate);
        #[allow(non_upper_case_globals)]
        let sessionstate = match newstate {
            AudioSessionStateActive => SessionState::Active,
            AudioSessionStateInactive => SessionState::Inactive,
            AudioSessionStateExpired => SessionState::Expired,
            _ => return S_OK,
        };
        if let Some(callback) = &mut self.callbacks.state {
            callback(sessionstate);
        }
        S_OK
    }

    fn on_session_disconnected(
        &mut self,
        disconnectreason: AudioSessionDisconnectReason,
    ) -> HRESULT {
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

        if let Some(callback) = &mut self.callbacks.disconnected {
            callback(reason);
        }
        S_OK
    }

    fn on_display_name_changed(
        &mut self,
        newdisplayname: PWSTR,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        let wide_name = unsafe { U16CString::from_ptr_str(newdisplayname.0) };
        let name = wide_name.to_string_lossy();
        trace!("New display name: {}", name);
        if let Some(callback) = &mut self.callbacks.displayname {
            let context = unsafe { *eventcontext };
            callback(name, context);
        }
        S_OK
    }

    fn on_icon_path_changed(
        &mut self,
        newiconpath: PWSTR,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        let wide_path = unsafe { U16CString::from_ptr_str(newiconpath.0) };
        let path = wide_path.to_string_lossy();
        trace!("New icon path: {}", path);
        if let Some(callback) = &mut self.callbacks.iconpath {
            let context = unsafe { *eventcontext };
            callback(path, context);
        }
        S_OK
    }

    fn on_simple_volume_changed(
        &mut self,
        newvolume: f32,
        newmute: BOOL,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        trace!("New volume: {}, mute: {:?}", newvolume, newmute);
        if let Some(callback) = &mut self.callbacks.simple_volume {
            let context = unsafe { *eventcontext };
            callback(newvolume, bool::from(newmute), context);
        }
        S_OK
    }

    fn on_channel_volume_changed(
        &mut self,
        channelcount: u32,
        newchannelvolumearray: *mut f32,
        changedchannel: u32,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        trace!("New channel volume for channel: {}", changedchannel);
        let volslice =
            unsafe { slice::from_raw_parts(newchannelvolumearray, channelcount as usize) };
        let newvol = volslice[changedchannel as usize];
        if let Some(callback) = &mut self.callbacks.channel_volume {
            let context = unsafe { *eventcontext };
            callback(changedchannel as usize, newvol, context);
        }
        S_OK
    }

    fn on_grouping_param_changed(
        &mut self,
        newgroupingparam: *const ::windows::Guid,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        trace!("Grouping changed");
        if let Some(callback) = &mut self.callbacks.groupingparam {
            let context = unsafe { *eventcontext };
            let grouping = unsafe { *newgroupingparam };
            callback(grouping, context);
        }
        S_OK
    }

    unsafe extern "system" fn _query_interface(
        this: ::windows::RawPtr,
        iid: &::windows::Guid,
        interface: *mut ::windows::RawPtr,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).query_interface(iid, interface)
    }

    unsafe extern "system" fn _add_ref(this: ::windows::RawPtr) -> u32 {
        (*(this as *mut Self)).add_ref()
    }

    unsafe extern "system" fn _release(this: ::windows::RawPtr) -> u32 {
        (*(this as *mut Self)).release()
    }

    unsafe extern "system" fn _on_display_name_changed(
        this: ::windows::RawPtr,
        newdisplayname: PWSTR,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).on_display_name_changed(newdisplayname, eventcontext)
    }

    unsafe extern "system" fn _on_icon_path_changed(
        this: ::windows::RawPtr,
        newiconpath: PWSTR,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).on_icon_path_changed(newiconpath, eventcontext)
    }

    unsafe extern "system" fn _on_simple_volume_changed(
        this: ::windows::RawPtr,
        newvolume: f32,
        newmute: BOOL,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).on_simple_volume_changed(newvolume, newmute, eventcontext)
    }

    unsafe extern "system" fn _on_channel_volume_changed(
        this: ::windows::RawPtr,
        channelcount: u32,
        newchannelvolumearray: *mut f32,
        changedchannel: u32,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).on_channel_volume_changed(
            channelcount,
            newchannelvolumearray,
            changedchannel,
            eventcontext,
        )
    }

    unsafe extern "system" fn _on_grouping_param_changed(
        this: ::windows::RawPtr,
        newgroupingparam: *const ::windows::Guid,
        eventcontext: *const ::windows::Guid,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).on_grouping_param_changed(newgroupingparam, eventcontext)
    }

    unsafe extern "system" fn _on_state_changed(
        this: ::windows::RawPtr,
        newstate: AudioSessionState,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).on_state_changed(newstate)
    }

    unsafe extern "system" fn _on_session_disconnected(
        this: ::windows::RawPtr,
        disconnectreason: AudioSessionDisconnectReason,
    ) -> ::windows::HRESULT {
        (*(this as *mut Self)).on_session_disconnected(disconnectreason)
    }
}
