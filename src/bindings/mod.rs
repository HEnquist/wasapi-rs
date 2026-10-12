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
mod generated;
pub use generated::*;

use windows_core::GUID;

// These are declared with DEFINE_GUIDSTRUCT in the Windows headers.
// The default metadata has no GUID value for them, and the generator
// emits an empty struct instead of a constant, so they are defined here.
/// The type of the acoustic echo cancellation effect, see [AudioEffect](crate::AudioEffect).
pub const AUDIO_EFFECT_TYPE_ACOUSTIC_ECHO_CANCELLATION: GUID =
    GUID::from_u128(0x6f64adbe_8211_11e2_8c70_2c27d7f001fa);
pub const KSDATAFORMAT_SUBTYPE_PCM: GUID = GUID::from_u128(0x00000001_0000_0010_8000_00aa00389b71);
pub const KSDATAFORMAT_SUBTYPE_IEEE_FLOAT: GUID =
    GUID::from_u128(0x00000003_0000_0010_8000_00aa00389b71);
pub const KSDATAFORMAT_TYPE_AUDIO: GUID = GUID::from_u128(0x73647561_0000_0010_8000_00aa00389b71);
#[allow(non_upper_case_globals)]
pub const KSPROPSETID_Pin: GUID = GUID::from_u128(0x8c134960_51ad_11cf_878a_94f801c10000);
