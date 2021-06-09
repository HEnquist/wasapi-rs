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
//! - Loopback capture
//! - Notifications for volume change, device disconnect etc
//!
//! ## Examples
//!
//! - The `playsine` example plays a sine wave in shared mode on the default output device.
//!
//! - The `playsine_events` example is similar to `playsine` but also listens to notifications.
//!
//! - The `loopback` example shows how to simultaneously capture and render sound, with separate threads for capture and render.

::windows::include_bindings!();
use Windows::Win32::System::PropertiesSystem::PROPERTYKEY;
mod api;
pub use api::*;
pub use windows::Guid;

#[macro_use]
extern crate log;

#[allow(non_upper_case_globals)]
const PKEY_Device_FriendlyName: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows::Guid::from_values(
        0xA45C254E,
        0xDF1C,
        0x4EFD,
        [0x80, 0x20, 0x67, 0xD1, 0x46, 0xA8, 0x50, 0xE0],
    ),
    pid: 14,
};

#[allow(non_upper_case_globals)]
const PKEY_Device_DeviceDesc: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows::Guid::from_values(
        0xA45C254E,
        0xDF1C,
        0x4EFD,
        [0x80, 0x20, 0x67, 0xD1, 0x46, 0xA8, 0x50, 0xE0],
    ),
    pid: 2,
};
