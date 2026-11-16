// Clients activated through `ActivateAudioInterfaceAsync` rather than an `IMMDevice`.
//
// The tests at the end cover `AudioClient::new_default_device_client`, which activates
// the default device with automatic stream routing. They run on whatever the default
// devices are, which is not necessarily the cable, but they still take the fixture so
// that they are serialised with the rest of the suite.
//
// The rest covers `AudioClient::from_iaudioclient`, the escape hatch for finding a device
// by some other means than the `DeviceCollection`, for instance through the WinRT
// `MediaDevice` and `DeviceInformation` classes.
//
// There is no WinRT here. What those classes hand over is a device interface path for
// `ActivateAudioInterfaceAsync`, and that path is built from the same endpoint id an
// `IMMDevice` reports, so `interface_path` below produces what `DeviceInformation.Id`
// would have given and the test takes the same route without pulling the WinRT
// projections into the suite. The completion handler is a copy of the private one in
// the library, which is the point: it is what a caller has to write.

use crate::support;
use std::sync::{Arc, Condvar, Mutex};
use wasapi::*;
use windows::Win32::Foundation::E_INVALIDARG;
use windows::Win32::Media::Audio::{
    ActivateAudioInterfaceAsync, DEVINTERFACE_AUDIO_CAPTURE, DEVINTERFACE_AUDIO_RENDER,
    IActivateAudioInterfaceAsyncOperation, IActivateAudioInterfaceCompletionHandler,
    IActivateAudioInterfaceCompletionHandler_Impl,
};
use windows::core::{HRESULT, Interface, PCWSTR, Ref, implement};

/// Signals the waiting thread when the activation has finished.
#[implement(IActivateAudioInterfaceCompletionHandler)]
struct Handler(Arc<(Mutex<bool>, Condvar)>);

impl IActivateAudioInterfaceCompletionHandler_Impl for Handler_Impl {
    fn ActivateCompleted(
        &self,
        _operation: Ref<IActivateAudioInterfaceAsyncOperation>,
    ) -> windows::core::Result<()> {
        let (lock, cvar) = &*self.0;
        *lock.lock().unwrap() = true;
        cvar.notify_one();
        Ok(())
    }
}

/// The device interface path for an endpoint id, in the form WinRT reports.
///
/// `ActivateAudioInterfaceAsync` does not take the bare endpoint id that
/// `Device::get_id` returns, it takes the software device path that wraps it, and
/// passing the id alone fails with `ERROR_FILE_NOT_FOUND`. The path is the MMDEVAPI
/// software enumerator, the endpoint id, and the interface class of the direction.
fn interface_path(id: &str, direction: Direction) -> String {
    let class = match direction {
        Direction::Render => DEVINTERFACE_AUDIO_RENDER,
        Direction::Capture => DEVINTERFACE_AUDIO_CAPTURE,
    };
    // The Debug impl of GUID is the canonical hyphenated form, not a derived one,
    // and the doubled braces are the literal pair the path needs around it.
    format!("\\\\?\\SWD#MMDEVAPI#{id}#{{{class:?}}}")
}

/// Activate an [IAudioClient] on the endpoint with the given id, the long way round.
///
/// Blocks until the activation completes. No timeout: the callback is made by the
/// audio service, and if that never answers the test harness hanging is a clearer
/// signal than a spurious failure would be.
fn activate(id: &str, direction: Direction) -> windows::core::Result<IAudioClient> {
    let path = interface_path(id, direction);
    println!("activating {path}");
    let path: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();
    let setup = Arc::new((Mutex::new(false), Condvar::new()));
    let callback: IActivateAudioInterfaceCompletionHandler = Handler(setup.clone()).into();

    let operation = unsafe {
        ActivateAudioInterfaceAsync(PCWSTR(path.as_ptr()), &IAudioClient::IID, None, &callback)?
    };

    let (lock, cvar) = &*setup;
    let mut completed = lock.lock().unwrap();
    while !*completed {
        completed = cvar.wait(completed).unwrap();
    }
    drop(completed);

    let mut client = None;
    let mut result = HRESULT::default();
    unsafe { operation.GetActivateResult(&mut result, &mut client)? };
    result.ok()?;
    // Safe to unwrap once the result above is checked.
    client.unwrap().cast()
}

/// A wrapped client streams like one from `Device::get_iaudioclient`.
///
/// Renders a short run of silence on the cable and asserts the rate, which is what
/// `streams::render_mechanics` does to a client obtained the normal way. Only shared
/// mode is covered here: what is under test is the constructor, and the sharing and
/// timing modes are already covered in full by `streams`.
#[test]
fn activated_client_renders() {
    let Some(fixture) = support::fixture() else {
        return;
    };
    let id = fixture.cable.render.get_id().unwrap();
    let interface = activate(&id, Direction::Render).expect("ActivateAudioInterfaceAsync failed");

    // The endpoint is a render endpoint, so that is what the direction describes,
    // and the stream direction passed to initialize_client is Render as well.
    let mut client = AudioClient::from_iaudioclient(interface, Direction::Render);
    let format = client.get_mixformat().unwrap();
    let (def_period, _min_period) = client.get_device_period().unwrap();
    let mode = support::mode_for(
        &client,
        ShareMode::Shared,
        TimingMode::Events,
        def_period,
        &format,
    )
    .unwrap();
    client
        .initialize_client(&format, Direction::Render, mode, None)
        .unwrap();

    let run = support::render_for(&client, &format, support::RUN_TIME).unwrap();
    support::assert_rate(&run, &format, "activated render");
}

/// The accessors report what the constructor was told, before and after initialising.
#[test]
fn activated_client_reports_its_state() {
    let Some(fixture) = support::fixture() else {
        return;
    };
    let id = fixture.cable.capture.get_id().unwrap();
    let interface = activate(&id, Direction::Capture).expect("ActivateAudioInterfaceAsync failed");

    let mut client = AudioClient::from_iaudioclient(interface, Direction::Capture);
    assert_eq!(client.get_direction(), Direction::Capture);
    assert_eq!(client.get_sharemode(), None);
    assert_eq!(client.get_timing_mode(), None);

    let format = client.get_mixformat().unwrap();
    let (def_period, _min_period) = client.get_device_period().unwrap();
    let mode = support::mode_for(
        &client,
        ShareMode::Shared,
        TimingMode::Polling,
        def_period,
        &format,
    )
    .unwrap();
    client
        .initialize_client(&format, Direction::Capture, mode, None)
        .unwrap();

    assert_eq!(client.get_direction(), Direction::Capture);
    assert_eq!(client.get_sharemode(), Some(ShareMode::Shared));
    assert_eq!(client.get_timing_mode(), Some(TimingMode::Polling));
}

/// The default device of a direction, or `None` after a skip note when there is none.
fn default_device(fixture: &support::Fixture, direction: Direction) -> Option<Device> {
    match fixture.enumerator.get_default_device(direction) {
        Ok(device) => {
            println!("default {}", support::describe(&device));
            Some(device)
        }
        Err(e) => {
            support::skip(&format!("no default {direction:?} device: {e}"));
            None
        }
    }
}

/// An auto-routed client activates, reports the direction it was given, and streams
/// in shared mode in both directions.
#[test]
fn default_device_client_streams() {
    let Some(fixture) = support::fixture() else {
        return;
    };
    for direction in [Direction::Render, Direction::Capture] {
        if default_device(&fixture, direction).is_none() {
            continue;
        }
        let mut client = AudioClient::new_default_device_client(direction)
            .expect("new_default_device_client failed");
        assert_eq!(client.get_direction(), direction);
        assert_eq!(client.get_sharemode(), None);
        assert_eq!(client.get_timing_mode(), None);

        let format = client.get_mixformat().unwrap();
        let (def_period, _min_period) = client.get_device_period().unwrap();
        let mode = support::mode_for(
            &client,
            ShareMode::Shared,
            TimingMode::Events,
            def_period,
            &format,
        )
        .unwrap();
        client
            .initialize_client(&format, direction, mode, None)
            .unwrap();

        let what = format!("auto-routed {direction:?}");
        let run = match direction {
            Direction::Render => support::render_for(&client, &format, support::RUN_TIME),
            Direction::Capture => support::capture_for(&client, &format, support::RUN_TIME),
        }
        .unwrap();
        support::assert_rate(&run, &format, &what);
    }
}

/// Exclusive mode fails as the documentation of `new_default_device_client` says.
///
/// The format queries agree with those on the device itself, and then `Initialize`
/// fails with `E_INVALIDARG`. This is Windows behaviour, not the wrapper's, and the
/// test is here to notice if the documented claim goes stale.
#[test]
fn default_device_client_rejects_exclusive() {
    let Some(fixture) = support::fixture() else {
        return;
    };
    for direction in [Direction::Render, Direction::Capture] {
        let Some(device) = default_device(&fixture, direction) else {
            continue;
        };
        let device_format = device.get_device_format().unwrap();
        let on_device = device
            .get_iaudioclient()
            .unwrap()
            .is_supported_exclusive_with_quirks(&device_format);
        let mut client = AudioClient::new_default_device_client(direction).unwrap();
        let auto_routed = client.is_supported_exclusive_with_quirks(&device_format);
        assert_eq!(
            auto_routed.is_ok(),
            on_device.is_ok(),
            "{direction:?}: the device gave {on_device:?}, the auto-routed client {auto_routed:?}"
        );
        let Ok(format) = auto_routed else {
            println!("{direction:?}: the device takes no exclusive format, nothing more to check");
            continue;
        };

        for timing in [TimingMode::Events, TimingMode::Polling] {
            let (def_period, _min_period) = client.get_device_period().unwrap();
            let mode =
                support::mode_for(&client, ShareMode::Exclusive, timing, def_period, &format)
                    .unwrap();
            let error = client
                .initialize_client(&format, direction, mode, None)
                .expect_err("exclusive mode initialised on an auto-routed client");
            assert!(
                support::is_hresult(&error, E_INVALIDARG),
                "{direction:?} {timing:?}: gave {error}, expected E_INVALIDARG"
            );
            // A failed Initialize leaves the client invalidated, so take a fresh one.
            client = AudioClient::new_default_device_client(direction).unwrap();
        }
    }
}
