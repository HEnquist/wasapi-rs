// Listen to device change notifications for one minute.
//
// Plug or unplug a device, or change the default device in the
// Windows sound settings, to see the notifications arrive.

use std::thread;
use std::time::Duration;
use wasapi::*;

fn main() {
    initialize_mta().unwrap();

    let enumerator = DeviceEnumerator::new().unwrap();

    let mut callbacks = DeviceEventCallbacks::new();

    callbacks.set_device_added_callback(|id| println!("Device added: {id}"));
    callbacks.set_device_removed_callback(|id| println!("Device removed: {id}"));
    callbacks.set_device_state_callback(|id, state| println!("Device {id} is now {state}"));
    callbacks.set_default_device_callback(|direction, role, id| match id {
        Some(id) => println!("New default {direction} device for role {role}: {id}"),
        None => println!("There is no longer a default {direction} device for role {role}"),
    });
    callbacks.set_property_value_callback(|id, key| {
        println!("Property {:?} of device {id} changed", key.fmtid)
    });

    // The notifications are unregistered when this value is dropped,
    // so it must be kept in scope for as long as they are needed.
    let _registered_events = enumerator
        .register_notification_callback(callbacks)
        .unwrap();

    println!("Listening for device changes for 60 seconds...");
    thread::sleep(Duration::from_secs(60));
}
