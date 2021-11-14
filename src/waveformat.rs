use crate::Windows::{
    Win32::Media::Audio::{
        WAVEFORMATEX, WAVEFORMATEXTENSIBLE, WAVEFORMATEXTENSIBLE_0, WAVE_FORMAT_PCM,
    },
    Win32::Media::KernelStreaming::{KSDATAFORMAT_SUBTYPE_PCM, WAVE_FORMAT_EXTENSIBLE},
    Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT},
};
use std::error;
use std::fmt;

use crate::{SampleType, WasapiError};

pub(crate) type WasapiRes<T> = Result<T, Box<dyn error::Error>>;

/// Struct wrapping a WAVEFORMATEXTENSIBLE format descriptor.
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
