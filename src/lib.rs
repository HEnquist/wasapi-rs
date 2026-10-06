#![doc = include_str!("../README.md")]

mod api;
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
pub use windows::core::GUID;
// Re-exported so that `AudioClient::from_iaudioclient` and the raw accessors can be
// used without depending on a matching version of the windows crate. HANDLE is left
// out on purpose: it would collide with this crate's own Handle in the generated
// documentation, since the two names differ only in case and the crate only ever
// builds for Windows, where filenames do not.
pub use windows::Win32::Media::Audio::{IAudioClient, IMMDevice};

#[macro_use]
extern crate log;

extern crate num_integer;
