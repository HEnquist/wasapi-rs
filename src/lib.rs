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
// Re-exported so that [AudioClient::from_iaudioclient] can be used without
// depending on a matching version of the windows crate.
pub use windows::Win32::Media::Audio::IAudioClient;

#[macro_use]
extern crate log;

extern crate num_integer;
