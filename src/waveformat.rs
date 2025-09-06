use std::fmt;
use windows::{
    core::GUID,
    Win32::Media::Audio::{
        WAVEFORMATEX, WAVEFORMATEXTENSIBLE, WAVEFORMATEXTENSIBLE_0, WAVE_FORMAT_PCM,
    },
    Win32::Media::KernelStreaming::{
        KSDATAFORMAT_SUBTYPE_PCM, SPEAKER_BACK_CENTER, SPEAKER_BACK_LEFT, SPEAKER_BACK_RIGHT,
        SPEAKER_FRONT_CENTER, SPEAKER_FRONT_LEFT, SPEAKER_FRONT_LEFT_OF_CENTER,
        SPEAKER_FRONT_RIGHT, SPEAKER_FRONT_RIGHT_OF_CENTER, SPEAKER_LOW_FREQUENCY,
        SPEAKER_SIDE_LEFT, SPEAKER_SIDE_RIGHT, WAVE_FORMAT_EXTENSIBLE,
    },
    Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT},
};

use crate::{SampleType, WasapiError, WasapiRes};

// Definitions from ksmedia.h of the windows sdk.
// Covers 1, 2, 4, 6 and 8 channels.
const KSAUDIO_SPEAKER_MONO: u32 = SPEAKER_FRONT_CENTER;
const KSAUDIO_SPEAKER_STEREO: u32 = SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT;
const KSAUDIO_SPEAKER_QUAD: u32 =
    SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT | SPEAKER_BACK_LEFT | SPEAKER_BACK_RIGHT;
const KSAUDIO_SPEAKER_SURROUND: u32 =
    SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT | SPEAKER_FRONT_CENTER | SPEAKER_BACK_CENTER;

// Marked as obsolete in ksmedia.h, kept for compatibility
const KSAUDIO_SPEAKER_5POINT1: u32 = SPEAKER_FRONT_LEFT
    | SPEAKER_FRONT_RIGHT
    | SPEAKER_FRONT_CENTER
    | SPEAKER_LOW_FREQUENCY
    | SPEAKER_BACK_LEFT
    | SPEAKER_BACK_RIGHT;
const KSAUDIO_SPEAKER_7POINT1: u32 = SPEAKER_FRONT_LEFT
    | SPEAKER_FRONT_RIGHT
    | SPEAKER_FRONT_CENTER
    | SPEAKER_LOW_FREQUENCY
    | SPEAKER_BACK_LEFT
    | SPEAKER_BACK_RIGHT
    | SPEAKER_FRONT_LEFT_OF_CENTER
    | SPEAKER_FRONT_RIGHT_OF_CENTER;

// Recommended 6 and 8 channel layouts from ksmedia.h
const KSAUDIO_SPEAKER_5POINT1_SURROUND: u32 = SPEAKER_FRONT_LEFT
    | SPEAKER_FRONT_RIGHT
    | SPEAKER_FRONT_CENTER
    | SPEAKER_LOW_FREQUENCY
    | SPEAKER_SIDE_LEFT
    | SPEAKER_SIDE_RIGHT;
const KSAUDIO_SPEAKER_7POINT1_SURROUND: u32 = SPEAKER_FRONT_LEFT
    | SPEAKER_FRONT_RIGHT
    | SPEAKER_FRONT_CENTER
    | SPEAKER_LOW_FREQUENCY
    | SPEAKER_BACK_LEFT
    | SPEAKER_BACK_RIGHT
    | SPEAKER_SIDE_LEFT
    | SPEAKER_SIDE_RIGHT;

// Custom layouts for 3, 5 and 7 channels, not part of ksmedia.h
const CUSTOM_SPEAKER_2POINT1: u32 = KSAUDIO_SPEAKER_STEREO | SPEAKER_LOW_FREQUENCY;
const CUSTOM_SPEAKER_4POINT1: u32 = KSAUDIO_SPEAKER_QUAD | SPEAKER_LOW_FREQUENCY;
const CUSTOM_SPEAKER_4POINT1_SURROUND: u32 = KSAUDIO_SPEAKER_SURROUND | SPEAKER_LOW_FREQUENCY;
const CUSTOM_SPEAKER_6POINT1: u32 = KSAUDIO_SPEAKER_5POINT1 | SPEAKER_BACK_CENTER;
const CUSTOM_SPEAKER_6POINT1_SURROUND: u32 = KSAUDIO_SPEAKER_5POINT1_SURROUND | SPEAKER_BACK_CENTER;

/// Struct wrapping a [WAVEFORMATEXTENSIBLE](https://docs.microsoft.com/en-us/windows/win32/api/mmreg/ns-mmreg-waveformatextensible) format descriptor.
#[derive(Clone)]
pub struct WaveFormat {
    pub wave_fmt: WAVEFORMATEXTENSIBLE,
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
    /// Parse a [WAVEFORMATEX](https://docs.microsoft.com/en-us/previous-versions/dd757713(v=vs.85)) structure and
    /// return a [WaveFormat] instance. If the underlying structure is a WAVEFORMATEXTENSIBLE, as specified by
    /// wFormatTag, then use as-is. If not, assume it is only a WAVEFORMATEX structure.
    pub fn parse(waveformatex: &WAVEFORMATEX) -> WasapiRes<Self> {
        if waveformatex.wFormatTag == WAVE_FORMAT_EXTENSIBLE as u16 {
            const ATLEAST_SIZE: usize =
                size_of::<WAVEFORMATEXTENSIBLE>() - size_of::<WAVEFORMATEX>();
            if waveformatex.cbSize < ATLEAST_SIZE as u16 {
                return Err(WasapiError::UnsupportedFormat);
            }
            // SAFETY: Both wFormatTag and size check passed, so the pointed-to memory is a full WAVEFORMATEXTENSIBLE.
            let waveformatextensible: WAVEFORMATEXTENSIBLE = unsafe {
                std::ptr::read(waveformatex as *const WAVEFORMATEX as *const WAVEFORMATEXTENSIBLE)
            };
            return Ok(waveformatextensible.into());
        }
        Self::from_waveformatex(*waveformatex)
    }

    /// Build a [WAVEFORMATEXTENSIBLE](https://docs.microsoft.com/en-us/windows/win32/api/mmreg/ns-mmreg-waveformatextensible) struct for the given parameters.
    /// `channel_mask` is optional. If a mask is provided, it will be used. If not, a default mask will be created.
    /// This can be used to work around quirks for some device drivers.
    /// If the default is not accepted, try again using a zero mask, `Some(0)`.
    pub fn new(
        storebits: usize,
        validbits: usize,
        sample_type: &SampleType,
        samplerate: usize,
        channels: usize,
        channel_mask: Option<u32>,
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
        // Only max 18 mask channel positions are defined,
        // https://docs.microsoft.com/en-us/windows/win32/api/mmreg/ns-mmreg-waveformatextensible
        let mask = if let Some(given_mask) = channel_mask {
            given_mask
        } else {
            match channels {
                ch if ch <= 18 => {
                    // setting bit for each channel
                    (1 << ch) - 1
                }
                _ => 0,
            }
        };
        let wave_fmt = WAVEFORMATEXTENSIBLE {
            Format: wave_format,
            Samples: sample,
            SubFormat: subformat,
            dwChannelMask: mask,
        };
        WaveFormat { wave_fmt }
    }

    /// Create from a [WAVEFORMATEX](https://docs.microsoft.com/en-us/previous-versions/dd757713(v=vs.85)) structure
    pub fn from_waveformatex(wavefmt: WAVEFORMATEX) -> WasapiRes<Self> {
        let validbits = wavefmt.wBitsPerSample as usize;
        let blockalign = wavefmt.nBlockAlign as usize;
        let samplerate = wavefmt.nSamplesPerSec as usize;
        let formattag = wavefmt.wFormatTag;
        let channels = wavefmt.nChannels as usize;
        let sample_type = match formattag as u32 {
            WAVE_FORMAT_PCM => SampleType::Int,
            WAVE_FORMAT_IEEE_FLOAT => SampleType::Float,
            _ => return Err(WasapiError::UnsupportedFormat),
        };
        let storebits = 8 * blockalign / channels;
        Ok(WaveFormat::new(
            storebits,
            validbits,
            &sample_type,
            samplerate,
            channels,
            None,
        ))
    }

    /// Return a copy in the simpler [WAVEFORMATEX](https://docs.microsoft.com/en-us/previous-versions/dd757713(v=vs.85)) format.
    pub fn to_waveformatex(&self) -> WasapiRes<Self> {
        let blockalign = self.wave_fmt.Format.nBlockAlign;
        let samplerate = self.wave_fmt.Format.nSamplesPerSec;
        let channels = self.wave_fmt.Format.nChannels;
        let byterate = self.wave_fmt.Format.nAvgBytesPerSec;
        let storebits = self.wave_fmt.Format.wBitsPerSample;
        let sample_type = match self.wave_fmt.SubFormat {
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT => WAVE_FORMAT_IEEE_FLOAT,
            KSDATAFORMAT_SUBTYPE_PCM => WAVE_FORMAT_PCM,
            _ => return Err(WasapiError::UnsupportedFormat),
        };
        let wave_format = WAVEFORMATEX {
            cbSize: 0,
            nAvgBytesPerSec: byterate,
            nBlockAlign: blockalign,
            nChannels: channels,
            nSamplesPerSec: samplerate,
            wBitsPerSample: storebits,
            wFormatTag: sample_type as u16,
        };
        let sample = WAVEFORMATEXTENSIBLE_0 {
            wValidBitsPerSample: 0,
        };
        let subformat = GUID::zeroed();
        let mask = 0;
        let wave_fmt = WAVEFORMATEXTENSIBLE {
            Format: wave_format,
            Samples: sample,
            SubFormat: subformat,
            dwChannelMask: mask,
        };
        Ok(WaveFormat { wave_fmt })
    }

    /// get a reference of type &WAVEFORMATEX, used internally
    pub fn as_waveformatex_ref(&self) -> &WAVEFORMATEX {
        &self.wave_fmt.Format
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
            _ => return Err(WasapiError::UnsupportedSubformat(self.wave_fmt.SubFormat)),
        };
        Ok(subfmt)
    }
}

impl From<WAVEFORMATEXTENSIBLE> for WaveFormat {
    fn from(wave_fmt: WAVEFORMATEXTENSIBLE) -> Self {
        WaveFormat { wave_fmt }
    }
}

/// Return a vector with suggested channel masks for the given number of channels.
/// Used to find a format that a device accepts in exclusive mode.
/// The values are sorted according to how likely they are to be accepted, with the most likely first.
pub fn make_channelmasks(channels: usize) -> Vec<u32> {
    match channels {
        1 => vec![KSAUDIO_SPEAKER_MONO, make_simple_channelmask(channels), 0],
        2 => vec![KSAUDIO_SPEAKER_STEREO, 0],
        3 => vec![CUSTOM_SPEAKER_2POINT1, make_simple_channelmask(channels), 0],
        4 => vec![
            KSAUDIO_SPEAKER_QUAD,
            KSAUDIO_SPEAKER_SURROUND,
            make_simple_channelmask(channels),
            0,
        ],
        5 => vec![
            CUSTOM_SPEAKER_4POINT1,
            CUSTOM_SPEAKER_4POINT1_SURROUND,
            make_simple_channelmask(channels),
            0,
        ],
        6 => vec![
            KSAUDIO_SPEAKER_5POINT1_SURROUND,
            KSAUDIO_SPEAKER_5POINT1,
            make_simple_channelmask(channels),
            0,
        ],
        7 => vec![
            CUSTOM_SPEAKER_6POINT1_SURROUND,
            CUSTOM_SPEAKER_6POINT1,
            make_simple_channelmask(channels),
            0,
        ],
        8 => vec![
            KSAUDIO_SPEAKER_7POINT1_SURROUND,
            KSAUDIO_SPEAKER_7POINT1,
            make_simple_channelmask(channels),
            0,
        ],
        9..=18 => vec![make_simple_channelmask(channels), 0],
        _ => vec![0],
    }
}

/// Make a simple channel mask by adding the correct number of bits.
/// Above the 18 channel positions [that are defined](https://docs.microsoft.com/en-us/windows/win32/api/mmreg/ns-mmreg-waveformatextensible)
/// it returns a zero.
pub fn make_simple_channelmask(channels: usize) -> u32 {
    match channels {
        1..=18 => {
            // setting bit for each channel
            (1 << channels) - 1
        }
        _ => 0,
    }
}
