//! # Wasapi bindings for Rust
//!
//! The aim of this crate is to provide easy and safe access to the Wasapi API for audio playback and capture.
//!
//! The presented API is all safe Rust, but structs and functions closely follow the original Windows API.
//!
//! For details on how to use Wasapi, please see [the Windows documentation](https://docs.microsoft.com/en-us/windows/win32/coreaudio/core-audio-interfaces).
//!
//! Bindings are generated automatically using the [windows](https://crates.io/crates/windows) crate.
//!
//! ## Supported functionality
//!
//! These things have been implemented so far:
//!
//! - Audio playback and capture
//! - Shared and exclusive modes
//! - Event-driven buffering
//! - Loopback capture
//! - Notifications for volume change, device disconnect etc
//!
//! ## Included examples
//!
//! | Example               | Description                                                                                            |
//! | --------------------- | ------------------------------------------------------------------------------------------------------ |
//! | `playsine`            | Plays a sine wave in shared mode on the default output device.                                         |
//! | `playsine_events`     | Similar to `playsine` but also listens to notifications.                                               |
//! | `playnoise_exclusive` | Plays white noise in exclusive mode on the default output device. Shows how to handle HRESULT errors.  |
//! | `loopback`            | Shows how to simultaneously capture and render sound, with separate threads for capture and render.    |
//! | `record`              | Records audio from the default device, and saves the raw samples to a file.                            |
//! | `devices`             | Lists all available audio devices and displays the default devices.                                    |
//! | `record_application`  | Records audio from the a single application, and saves the raw samples to a file.                      |

mod api;
mod events;
mod waveformat;
pub use api::*;
pub use events::*;
pub use waveformat::*;
pub use windows::core::GUID;

#[macro_use]
extern crate log;

extern crate num_integer;
