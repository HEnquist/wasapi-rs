//! Windows API bindings used by the crate.
//!
//! Almost everything comes from `generated.rs`, which is written by the generator in
//! `tools/bindings` and must not be edited by hand.
//! This file only adds what the default metadata gets wrong.

#[allow(
    non_snake_case,
    non_camel_case_types,
    non_upper_case_globals,
    dead_code,
    clippy::all
)]
pub(crate) mod generated;
pub use generated::*;

use windows_core::GUID;

/// Define constants with the type that the crate uses them as.
///
/// The default metadata gives most constants as `i32`,
/// whatever type the functions and fields that take them use.
/// This converts each one once, instead of at every place it is used.
/// A constant defined here shadows the generated one of the same name,
/// since a named item takes precedence over a glob import.
/// The `HRESULT` form wraps the value instead of casting it.
macro_rules! typed_consts {
    (HRESULT, $doc:expr; $($name:ident),* $(,)?) => {
        $(
            #[doc = $doc]
            pub const $name: windows_core::HRESULT =
                windows_core::HRESULT($crate::bindings::generated::$name);
        )*
    };
    ($ty:ty, $doc:expr; $($name:ident),* $(,)?) => {
        $(
            #[doc = $doc]
            pub const $name: $ty = $crate::bindings::generated::$name as $ty;
        )*
    };
}
pub(crate) use typed_consts;

typed_consts!(
    u32,
    "Retyped as the `u32` that the Windows functions and fields take.";
    AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY,
    AUDCLNT_BUFFERFLAGS_SILENT,
    AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR,
    AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
    AUDCLNT_STREAMFLAGS_LOOPBACK,
    AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY,
    CLSCTX_ALL,
    COINIT_APARTMENTTHREADED,
    COINIT_MULTITHREADED,
    DEVICE_STATE_ACTIVE,
    DEVICE_STATE_DISABLED,
    DEVICE_STATE_NOTPRESENT,
    DEVICE_STATE_UNPLUGGED,
    ENDPOINT_HARDWARE_SUPPORT_METER,
    ENDPOINT_HARDWARE_SUPPORT_MUTE,
    ENDPOINT_HARDWARE_SUPPORT_VOLUME,
    ERROR_INSUFFICIENT_BUFFER,
    ERROR_MORE_DATA,
    FILE_SHARE_READ,
    FILE_SHARE_WRITE,
    GENERIC_WRITE,
    IOCTL_KS_PROPERTY,
    KSPROPERTY_TYPE_GET,
    OPEN_EXISTING,
    STGM_READ,
    WAIT_OBJECT_0,
);

typed_consts!(
    u16,
    "Retyped as the `u16` of a `wFormatTag` or `VARTYPE`.";
    VT_BLOB,
    WAVE_FORMAT_EXTENSIBLE,
    WAVE_FORMAT_IEEE_FLOAT,
    WAVE_FORMAT_PCM,
);

// These are declared with DEFINE_GUIDSTRUCT in the Windows headers.
// The default metadata has no GUID value for them, and the generator
// emits an empty struct instead of a constant, so they are defined here.
/// The type of the acoustic echo cancellation effect, see [AudioEffect](crate::AudioEffect).
pub const AUDIO_EFFECT_TYPE_ACOUSTIC_ECHO_CANCELLATION: GUID =
    GUID::from_u128(0x6f64adbe_8211_11e2_8c70_2c27d7f001fa);
/// The integer PCM subformat, see [SampleType::Int](crate::SampleType::Int).
pub const KSDATAFORMAT_SUBTYPE_PCM: GUID = GUID::from_u128(0x00000001_0000_0010_8000_00aa00389b71);
/// The IEEE float subformat, see [SampleType::Float](crate::SampleType::Float).
pub const KSDATAFORMAT_SUBTYPE_IEEE_FLOAT: GUID =
    GUID::from_u128(0x00000003_0000_0010_8000_00aa00389b71);
pub const KSDATAFORMAT_TYPE_AUDIO: GUID = GUID::from_u128(0x73647561_0000_0010_8000_00aa00389b71);
#[allow(non_upper_case_globals)]
pub const KSPROPSETID_Pin: GUID = GUID::from_u128(0x8c134960_51ad_11cf_878a_94f801c10000);

// The metadata does not mark the output of GetChannelsPeakValues as an array,
// so the generated wrapper hands Windows room for a single value,
// and any channel count above 1 makes Windows write past it.
// The generated method can not be removed, so clippy.toml forbids calling it,
// and this one takes a slice instead.
impl IAudioMeterInformation {
    /// Fill `peaks` with the peak value of each channel, one per element.
    #[allow(non_snake_case)]
    pub unsafe fn GetChannelsPeakValuesSlice(&self, peaks: &mut [f32]) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).GetChannelsPeakValues)(
                windows_core::Interface::as_raw(self),
                peaks.len() as u32,
                peaks.as_mut_ptr(),
            )
            .ok()
        }
    }
}
