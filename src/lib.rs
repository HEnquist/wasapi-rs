//! # Wasapi bindings for Rust
//!
//! This create is meant as a thin wrapper for the Wasapi API.
//!
//! Most things map directly to something in the Windows API.
//!
//! Both shared and exclusive modes are supported. 
//!
//! Bindings are generated automatically using the `windows` crate.
//! 
//! The `loopback` example shows how to simultaneously capture and render sound. 

::windows::include_bindings!();
use Windows::Win32::System::PropertiesSystem::PROPERTYKEY;
pub mod wasapi;

#[allow(non_upper_case_globals)]
pub const PKEY_Device_FriendlyName: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows::Guid::from_values(
        0xA45C254E,
        0xDF1C,
        0x4EFD,
        [0x80, 0x20, 0x67, 0xD1, 0x46, 0xA8, 0x50, 0xE0],
    ),
    pid: 14,
};

#[allow(non_upper_case_globals)]
pub const PKEY_Device_DeviceDesc: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows::Guid::from_values(
        0xA45C254E,
        0xDF1C,
        0x4EFD,
        [0x80, 0x20, 0x67, 0xD1, 0x46, 0xA8, 0x50, 0xE0],
    ),
    pid: 2,
};


