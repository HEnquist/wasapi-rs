// The raw COM accessors.
//
// These exist so that a caller can reach what the crate does not wrap, so the tests
// are about whether that actually works, not about the interfaces themselves. Each
// one does the thing the accessor was added for: activate a service with no wrapper
// here, and wait on the stream's event alongside another handle.

use crate::support;
use wasapi::*;
use windows::Win32::Foundation::CloseHandle;
use windows::Win32::Media::Audio::Endpoints::IAudioEndpointVolume;
use windows::Win32::Media::Audio::IAudioClient3;
use windows::Win32::System::Com::CLSCTX_ALL;
use windows::Win32::System::Threading::{CreateEventW, SetEvent, WaitForMultipleObjects};
use windows::core::Interface;

/// An endpoint service the crate does not wrap can be activated through the device.
///
/// `IAudioEndpointVolume` is the case the accessor was added for: nothing in the
/// crate reaches it, while `HardwareSupport::volume` reports that the endpoint has
/// it. The volume is only read here, never set, since the tests run against whatever
/// machine happens to host them.
#[test]
fn endpoint_volume_through_as_immdevice() {
    let Some(fx) = support::fixture() else { return };

    // The accessor is safe, Activate is not.
    let volume: IAudioEndpointVolume =
        unsafe { fx.cable.render.as_immdevice().Activate(CLSCTX_ALL, None) }
            .expect("could not activate IAudioEndpointVolume");

    let level = unsafe { volume.GetMasterVolumeLevelScalar() }.unwrap();
    let muted = unsafe { volume.GetMute() }.unwrap().as_bool();
    println!("master volume {level:.3}, muted {muted}");
    assert!(
        (0.0..=1.0).contains(&level),
        "scalar volume {level} is outside 0.0 to 1.0"
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
    let quit = unsafe { CreateEventW(None, true, false, None) }.unwrap();

    let handles = [event.as_handle(), quit];
    opened.client.start_stream().unwrap();
    let result = unsafe { WaitForMultipleObjects(&handles, false, 2000) };
    opened.client.stop_stream().unwrap();

    // WAIT_OBJECT_0 is the first handle, so the device woke the wait.
    assert_eq!(
        result.0, 0,
        "the wait returned {:#x}, expected the audio event at index 0",
        result.0
    );

    // And the other way round: signalling the second handle on a stopped stream, where
    // the device never fires, has to come back with index 1 rather than time out.
    unsafe { SetEvent(quit) }.unwrap();
    let result = unsafe { WaitForMultipleObjects(&handles, false, 2000) };
    assert_eq!(
        result.0, 1,
        "the wait returned {:#x}, expected the signalled handle at index 1",
        result.0
    );

    // The crate's Handle closes its own event on drop, this one is the test's to close.
    unsafe { CloseHandle(quit) }.unwrap();
}

/// A newer interface version the crate does not wrap can be reached from the client.
///
/// Nothing here touches `IAudioClient3`, and its engine period getter is the kind of
/// thing the accessor exists for. The client is deliberately left uninitialised: the
/// call does not need it, and casting is all that is under test.
#[test]
fn engine_period_through_as_iaudioclient() {
    let Some(fx) = support::fixture() else { return };
    let client = fx.cable.render.get_iaudioclient().unwrap();
    let format = client.get_mixformat().unwrap();

    let client3: IAudioClient3 = client
        .as_iaudioclient()
        .cast()
        .expect("could not cast to IAudioClient3");

    let mut default = 0;
    let mut fundamental = 0;
    let mut min = 0;
    let mut max = 0;
    unsafe {
        client3.GetSharedModeEnginePeriod(
            format.as_waveformatex_ref(),
            &mut default,
            &mut fundamental,
            &mut min,
            &mut max,
        )
    }
    .expect("GetSharedModeEnginePeriod failed");

    println!(
        "engine period in frames: default {default}, fundamental {fundamental}, min {min}, max {max}"
    );
    assert!(min > 0, "the minimum period is {min} frames");
    assert!(
        min <= default && default <= max,
        "the default is outside the range"
    );
}
