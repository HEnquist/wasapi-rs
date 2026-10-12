// The raw COM pointer accessors.
//
// These exist so that a caller can reach what the crate does not wrap, with whatever
// Windows bindings it uses itself. The tests are about the crate's side of that: a
// pointer handed out can be wrapped again, the wrapper holds a reference of its own,
// the interface is looked up rather than assumed, and the event handle works in a
// wait on several handles at once. No binding crate is needed for any of it.

use crate::support;
use std::ffi::c_void;
use wasapi::*;

windows_core::link!("kernel32.dll" "system" fn CreateEventW(lpeventattributes : *const c_void, bmanualreset : i32, binitialstate : i32, lpname : *const u16) -> *mut c_void);
windows_core::link!("kernel32.dll" "system" fn SetEvent(hevent : *mut c_void) -> i32);
windows_core::link!("kernel32.dll" "system" fn CloseHandle(hobject : *mut c_void) -> i32);
windows_core::link!("kernel32.dll" "system" fn WaitForMultipleObjects(ncount : u32, lphandles : *const *mut c_void, bwaitall : i32, dwmilliseconds : u32) -> u32);

/// A device wrapped from its own raw pointer is the same endpoint, in the same direction.
#[test]
fn device_round_trips_through_raw() {
    let Some(fx) = support::fixture() else { return };

    for device in [&fx.cable.render, &fx.cable.capture] {
        let again = unsafe { Device::from_raw(device.as_raw()) }.unwrap();
        assert_eq!(again.get_id().unwrap(), device.get_id().unwrap());
        assert_eq!(again.get_direction(), device.get_direction());
    }
}

/// A client wrapped from a raw pointer keeps working after the original is dropped.
///
/// That is what taking a reference of its own means: the pointer is only borrowed,
/// so the original [AudioClient] may release its reference as soon as it likes.
#[test]
fn audio_client_holds_its_own_reference() {
    let Some(fx) = support::fixture() else { return };
    let original = fx.cable.render.get_iaudioclient().unwrap();
    let format = original.get_mixformat().unwrap();

    let wrapped = unsafe { AudioClient::from_raw(original.as_raw(), Direction::Render) }.unwrap();
    drop(original);

    assert_eq!(
        wrapped.get_mixformat().unwrap().get_samplespersec(),
        format.get_samplespersec()
    );
    assert_eq!(wrapped.get_direction(), Direction::Render);
    assert_eq!(wrapped.get_sharemode(), None);
}

/// A pointer to an object without the wanted interface gives `E_NOINTERFACE`.
///
/// The interface is looked up with QueryInterface, so handing over the wrong kind of
/// object is reported rather than called through the wrong vtable.
#[test]
fn wrong_interface_is_refused() {
    let Some(fx) = support::fixture() else { return };

    let result = unsafe { AudioClient::from_raw(fx.cable.render.as_raw(), Direction::Render) };
    let Err(error) = result else {
        panic!("an IMMDevice was accepted as an IAudioClient");
    };
    assert!(
        support::is_hresult(&error, E_NOINTERFACE),
        "wrapping a device as a client gave {error}, expected E_NOINTERFACE"
    );
}

/// The stream's event handle works in a wait on several handles at once.
///
/// This is what the raw handle is for: a render loop that has to wake on its own
/// shutdown event as well as on the device. The second handle is never signalled, so
/// the wait has to come back with the first one.
#[test]
fn event_handle_in_a_multiple_object_wait() {
    let Some(fx) = support::fixture() else { return };
    let Some(opened) = support::open_or_skip(
        &fx.cable.render,
        &Direction::Render,
        ShareMode::Shared,
        TimingMode::Events,
    ) else {
        return;
    };
    let event = opened.client.set_get_eventhandle().unwrap();

    // Stands in for a shutdown event. Manual reset, starts unsignalled.
    let quit = unsafe { CreateEventW(std::ptr::null(), 1, 0, std::ptr::null()) };
    assert!(!quit.is_null(), "could not create the second event");

    let handles = [event.as_handle(), quit];
    opened.client.start_stream().unwrap();
    let result = unsafe { WaitForMultipleObjects(2, handles.as_ptr(), 0, 2000) };
    opened.client.stop_stream().unwrap();

    // WAIT_OBJECT_0 is the first handle, so the device woke the wait.
    assert_eq!(
        result, 0,
        "the wait returned {result:#x}, expected the audio event at index 0"
    );

    // And the other way round: signalling the second handle on a stopped stream, where
    // the device never fires, has to come back with index 1 rather than time out.
    assert_ne!(unsafe { SetEvent(quit) }, 0, "could not signal the event");
    let result = unsafe { WaitForMultipleObjects(2, handles.as_ptr(), 0, 2000) };
    assert_eq!(
        result, 1,
        "the wait returned {result:#x}, expected the signalled handle at index 1"
    );

    // The crate's Handle closes its own event on drop, this one is the test's to close.
    unsafe { CloseHandle(quit) };
}
