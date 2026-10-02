// Error paths that the crate itself produces, and the hand-written COM plumbing.
//
// Call it wrong, check what comes back. What is worth testing here is the crate's own
// guards, its buffer length accounting, and the activation of a process loopback
// client, which is a substantial piece of hand-written code. Tests that would only
// confirm Windows returns the result code it documents are deliberately absent.

use crate::support::{self, RUN_TIME};
use std::collections::VecDeque;
use wasapi::*;
use windows::Win32::Foundation::E_NOINTERFACE;
use windows::Win32::Media::Audio::{AUDCLNT_E_ALREADY_INITIALIZED, AUDCLNT_E_NOT_STOPPED};

/// Open a shared mode client on a device, at its mix format.
fn shared(device: &Device) -> support::Opened {
    support::open(
        device,
        &device.get_direction(),
        ShareMode::Shared,
        TimingMode::Polling,
    )
    .expect("a shared mode client at the mix format did not initialise")
}

/// A client can only be initialised once.
///
/// Note that it is not usable afterwards either: a failed `Initialize` leaves the
/// client invalidated and the next call reports `AUDCLNT_E_DEVICE_INVALIDATED`. That
/// matches what Windows documents, which is to release the client and make a new one
/// if `Initialize` fails, so there is nothing further to check.
#[test]
fn initialize_twice_is_refused() {
    let Some(fx) = support::fixture() else { return };
    let mut opened = shared(&fx.cable.render);
    let (def_period, _min) = opened.client.get_device_period().unwrap();

    let again = opened
        .client
        .initialize_client(
            &opened.format,
            &Direction::Render,
            &StreamMode::PollingShared {
                autoconvert: false,
                buffer_duration_hns: def_period,
            },
        )
        .expect_err("a second initialize_client succeeded");
    assert!(
        support::is_hresult(&again, AUDCLNT_E_ALREADY_INITIALIZED),
        "a second initialize_client gave {again}, expected AUDCLNT_E_ALREADY_INITIALIZED"
    );
}

/// The two direction guards in `initialize_client` are the crate's own, and are reached
/// before any call into Windows.
#[test]
fn direction_mismatches_are_refused() {
    let Some(fx) = support::fixture() else { return };

    // Capturing from a render device is loopback, which exclusive mode cannot do.
    let mut client = fx.cable.render.get_iaudioclient().unwrap();
    let format = client.get_mixformat().unwrap();
    assert!(matches!(
        client.initialize_client(
            &format,
            &Direction::Capture,
            &StreamMode::EventsExclusive {
                period_hns: 100_000
            },
        ),
        Err(WasapiError::LoopbackWithExclusiveMode)
    ));

    // Rendering to a capture device is not a thing in either mode.
    let mut client = fx.cable.capture.get_iaudioclient().unwrap();
    let format = client.get_mixformat().unwrap();
    for mode in [
        StreamMode::PollingShared {
            autoconvert: false,
            buffer_duration_hns: 100_000,
        },
        StreamMode::EventsExclusive {
            period_hns: 100_000,
        },
    ] {
        assert!(matches!(
            client.initialize_client(&format, &Direction::Render, &mode),
            Err(WasapiError::RenderToCaptureDevice)
        ));
    }
}

/// The length checks on a write are the crate's own, and report both numbers.
///
/// This is where the `bytes_per_frame` the client recorded at initialise time gets
/// used, and the git history says this is the part of the crate that has actually had
/// bugs. Asking for more than the buffer holds is in here too, since it shares a
/// client and checks the stream survives the refusal.
#[test]
fn write_errors() {
    let Some(fx) = support::fixture() else { return };
    let opened = shared(&fx.cable.render);
    let blockalign = opened.format.get_blockalign() as usize;
    let render = opened.client.get_audiorenderclient().unwrap();

    // Too few bytes for the frames asked for, and too many.
    for (frames, bytes) in [(4usize, 3 * blockalign), (4, 5 * blockalign)] {
        let error = render
            .write_to_device(frames, &vec![0u8; bytes], None)
            .expect_err("a wrong length write succeeded");
        assert!(
            matches!(
                error,
                WasapiError::DataLengthMismatch { received, expected }
                    if received == bytes && expected == frames * blockalign
            ),
            "a {bytes} byte write of {frames} frames gave {error}"
        );
    }

    // The deque variant reports that it is short rather than mismatched, and must not
    // consume anything when it refuses.
    let mut short: VecDeque<u8> = VecDeque::from(vec![0u8; 2 * blockalign]);
    let error = render
        .write_to_device_from_deque(4, &mut short, None)
        .expect_err("a short deque write succeeded");
    assert!(
        matches!(error, WasapiError::DataLengthTooShort { .. }),
        "a short deque write gave {error}"
    );
    assert_eq!(short.len(), 2 * blockalign, "the deque was consumed anyway");

    // Zero frames is a no-op, not an error.
    assert!(render.write_to_device(0, &[], None).is_ok());

    // More frames than the buffer holds is refused, and the client survives it.
    let buffer_size = opened.client.get_buffer_size().unwrap() as usize;
    let too_many = buffer_size + 1;
    assert!(
        render
            .write_to_device(too_many, &vec![0u8; too_many * blockalign], None)
            .is_err(),
        "a write larger than the buffer succeeded"
    );
    let space = opened.client.get_available_space_in_frames().unwrap() as usize;
    assert_eq!(space, buffer_size);
    render
        .write_to_device(space, &vec![0u8; space * blockalign], None)
        .expect("a correct write after an oversized one failed");
}

/// Both deque variants move the same bytes the slice variants do.
///
/// The deque paths index into two ring buffer halves by hand, so they are worth a run
/// of their own rather than being assumed to match the slice versions.
#[test]
fn deque_variants_move_whole_frames() {
    let Some(fx) = support::fixture() else { return };

    // Render: fill the buffer from a deque and check it consumed exactly the frames.
    let opened = shared(&fx.cable.render);
    let blockalign = opened.format.get_blockalign() as usize;
    let render = opened.client.get_audiorenderclient().unwrap();
    let frames = opened.client.get_available_space_in_frames().unwrap() as usize;
    let mut queue: VecDeque<u8> = VecDeque::from(vec![0u8; (frames + 32) * blockalign]);
    render
        .write_to_device_from_deque(frames, &mut queue, None)
        .expect("a deque write failed");
    assert_eq!(
        queue.len(),
        32 * blockalign,
        "a {frames} frame deque write consumed the wrong number of bytes"
    );
    drop(opened);

    // Capture: read into a deque and check it grew by whole frames.
    let opened = shared(&fx.cable.capture);
    let blockalign = opened.format.get_blockalign() as usize;
    let capture = opened.client.get_audiocaptureclient().unwrap();
    let mut queue: VecDeque<u8> = VecDeque::new();
    opened.client.start_stream().unwrap();
    for _ in 0..40 {
        std::thread::sleep(std::time::Duration::from_millis(5));
        capture
            .read_from_device_to_deque(&mut queue)
            .expect("a deque read failed");
    }
    opened.client.stop_stream().unwrap();
    assert!(!queue.is_empty(), "a deque read collected nothing");
    assert_eq!(
        queue.len() % blockalign,
        0,
        "{} bytes is not a whole number of {blockalign} byte frames",
        queue.len()
    );
}

/// Resetting a running stream is refused, and a reset empties the buffer.
#[test]
fn reset_requires_a_stopped_stream() {
    let Some(fx) = support::fixture() else { return };
    let opened = shared(&fx.cable.render);
    let blockalign = opened.format.get_blockalign() as usize;
    let render = opened.client.get_audiorenderclient().unwrap();
    let frames = opened.client.get_available_space_in_frames().unwrap() as usize;
    render
        .write_to_device(frames, &vec![0u8; frames * blockalign], None)
        .unwrap();

    opened.client.start_stream().unwrap();
    let running = opened
        .client
        .reset_stream()
        .expect_err("reset_stream succeeded on a running stream");
    assert!(
        support::is_hresult(&running, AUDCLNT_E_NOT_STOPPED),
        "reset_stream on a running stream gave {running}, expected AUDCLNT_E_NOT_STOPPED"
    );

    opened.client.stop_stream().unwrap();
    opened
        .client
        .reset_stream()
        .expect("reset_stream after stop failed");
    assert_eq!(
        opened.client.get_available_space_in_frames().unwrap(),
        opened.client.get_buffer_size().unwrap(),
        "the buffer was not empty after a reset"
    );
}

/// A device without echo cancellation says so, rather than erroring.
///
/// `is_aec_supported` is not a plain delegation: it looks for the effect first and then
/// asks for the interface. Note it needs an initialised client, in which case a fresh
/// one reports `AUDCLNT_E_NOT_INITIALIZED` rather than false.
#[test]
fn aec_reports_cleanly_on_a_device_without_it() {
    let Some(fx) = support::fixture() else { return };
    let opened = shared(&fx.cable.capture);
    let supported = opened
        .client
        .is_aec_supported()
        .expect("is_aec_supported returned an error instead of false");
    println!("the cable reports AEC support: {supported}");

    if !supported {
        // AcousticEchoCancellationControl has no Debug, so unwrap the error by hand.
        match opened.client.get_aec_control() {
            Ok(_) => panic!("get_aec_control succeeded on a device without AEC"),
            Err(error) => assert!(
                support::is_hresult(&error, E_NOINTERFACE),
                "get_aec_control gave {error}, expected E_NOINTERFACE"
            ),
        }
    }
}

/// Loopback capture of a render endpoint, as far as it can be checked here.
///
/// A render device initialised with `Direction::Capture` and a shared mode is what
/// makes `initialize_client` set `AUDCLNT_STREAMFLAGS_LOOPBACK`, which is the
/// hand-written part. Only the mechanics are asserted: on a CI runner a loopback of a
/// render endpoint delivers nothing but zeros, so there is nothing to say about content.
#[test]
fn loopback_capture_mechanics() {
    let Some(fx) = support::fixture() else { return };
    let mut client = fx.cable.render.get_iaudioclient().unwrap();
    let format = client.get_mixformat().unwrap();
    let (def_period, _min) = client.get_device_period().unwrap();
    client
        .initialize_client(
            &format,
            &Direction::Capture,
            &StreamMode::PollingShared {
                autoconvert: false,
                buffer_duration_hns: def_period,
            },
        )
        .expect("a loopback client on a render endpoint did not initialise");

    // The client now reports the direction of the stream, not of the device.
    assert_eq!(client.get_direction(), Direction::Capture);
    client
        .get_audiocaptureclient()
        .expect("a loopback client has no capture client");

    let run = support::capture_for(&client, &format, RUN_TIME).expect("the capture loop failed");
    println!(
        "loopback delivered {} frames in {} packets at {:.0} Hz",
        run.frames,
        run.packets.len(),
        run.rate()
    );
    assert!(run.frames > 0, "a loopback capture delivered no frames");
}

/// Process loopback capture activates, and is as limited as its documentation says.
///
/// `new_application_loopback_client` is substantial hand-written code: it activates the
/// interface asynchronously and waits on a condvar for a completion handler. The loop
/// over the non-functional methods turns the prose list in its doc comment into an
/// executable one, which already caught one claim that had gone stale.
#[test]
fn process_loopback_reports_its_limitations() {
    let Some(_fx) = support::fixture() else {
        return;
    };
    let mut client = match AudioClient::new_application_loopback_client(std::process::id(), false) {
        Ok(client) => client,
        Err(e) => {
            support::skip(&format!("process loopback capture is not available: {e}"));
            return;
        }
    };

    // Every method the documentation lists as non-functional, except get_buffer_size,
    // which it says returns a huge number rather than an error.
    let format = WaveFormat::new(32, 32, &SampleType::Float, 48000, 2, None);
    let broken: [(&str, bool); 7] = [
        ("get_mixformat", client.get_mixformat().is_err()),
        (
            "is_supported",
            client.is_supported(&format, &ShareMode::Shared).is_err(),
        ),
        (
            "is_supported_exclusive_with_quirks",
            client.is_supported_exclusive_with_quirks(&format).is_err(),
        ),
        ("get_device_period", client.get_device_period().is_err()),
        (
            "calculate_aligned_period_near",
            client
                .calculate_aligned_period_near(100_000, Some(128), &format)
                .is_err(),
        ),
        ("get_current_padding", client.get_current_padding().is_err()),
        (
            "get_available_space_in_frames",
            client.get_available_space_in_frames().is_err(),
        ),
    ];
    for (name, failed) in broken {
        assert!(
            failed,
            "{name} worked, but the documentation says it does not"
        );
    }

    // The period given here is irrelevant in this mode, hence the zero.
    client
        .initialize_client(
            &format,
            &Direction::Capture,
            &StreamMode::EventsShared {
                autoconvert: true,
                buffer_duration_hns: 0,
            },
        )
        .expect("a process loopback client did not initialise");

    // The sharing and timing modes are recorded on the Rust side, so unlike the rest
    // they do get reported correctly even here.
    assert_eq!(client.get_sharemode(), Some(ShareMode::Shared));
    assert_eq!(client.get_timing_mode(), Some(TimingMode::Events));
    assert!(client.get_audiocaptureclient().is_ok());
}
