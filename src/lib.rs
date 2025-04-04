#![doc = include_str!("../README.md")]

mod api;
mod errors;
mod events;
mod waveformat;
pub use api::*;
pub use errors::*;
pub use events::*;
pub use waveformat::*;
pub use windows::core::GUID;

#[macro_use]
extern crate log;

extern crate num_integer;
