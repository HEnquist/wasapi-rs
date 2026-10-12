#![doc = include_str!("../README.md")]

mod api;
mod bindings;
mod capabilities;
mod dataranges;
mod errors;
mod events;
mod waveformat;
pub use api::*;
pub use capabilities::*;
pub use dataranges::*;
pub use errors::*;
pub use events::*;
pub use waveformat::*;
// These appear in the public API, re-exported so that callers need no windows-core
// dependency of their own, and can not end up with a mismatched version of it.
pub use windows_core::{Error as WindowsError, GUID, HRESULT};
// The subformats a DataRange or an UnsupportedSubformat error can be compared with.
pub use bindings::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, KSDATAFORMAT_SUBTYPE_PCM};

#[macro_use]
extern crate log;

extern crate num_integer;
