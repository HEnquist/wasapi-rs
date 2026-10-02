// Playback and capture mechanics.
//
// A stream that asks for data at the format's rate has already proved the whole
// chain: client setup, the timing mode, the buffer accounting and the render or
// capture client. So that is what gets asserted. The data written is silence and the
// data read is thrown away, because nothing here looks at content.
//
// The tolerance on a rate is deliberately wide. A runner is a loaded virtual machine,
// and the bugs this catches are off by a factor, not by a few percent: a wrong block
// align, a format the device silently reinterpreted, a period that was not what was
// asked for.
//
// The first two tests each run all four sharing and timing combinations once and
// assert everything there is to say about those runs, rather than streaming the same
// audio again for every assertion.

use crate::support::{self, RUN_TIME};
use std::time::{Duration, Instant};
use wasapi::*;

/// How far the measured frame rate may be from the format's sample rate.
const RATE_TOLERANCE: f64 = 0.20;

/// Every combination of sharing and timing mode.
const MODES: [(ShareMode, TimingMode); 4] = [
    (ShareMode::Shared, TimingMode::Events),
    (ShareMode::Shared, TimingMode::Polling),
    (ShareMode::Exclusive, TimingMode::Events),
    (ShareMode::Exclusive, TimingMode::Polling),
];

/// Assert that a run moved frames at the format's sample rate.
fn assert_rate(run: &support::Run, format: &WaveFormat, what: &str) {
    assert!(
        !run.event_timeout,
        "{what}: timed out waiting for the event handle after {} frames",
        run.frames
    );
    let nominal = f64::from(format.get_samplespersec());
    let measured = run.rate();
    println!(
        "{what}: {} frames in {:.0} ms, {measured:.0} Hz against {nominal:.0} Hz, \
         buffer {} frames",
        run.frames,
        run.elapsed.as_secs_f64() * 1000.0,
        run.buffer_size
    );
    assert!(run.frames > 0, "{what}: no frames at all");
    let error = (measured - nominal).abs() / nominal;
    assert!(
        error < RATE_TOLERANCE,
        "{what}: {measured:.0} Hz is {:.0} % off the format's {nominal:.0} Hz",
        error * 100.0
    );
}

/// Playback asks for data at the rate the format declares.
///
/// This is the test the whole suite exists for.
#[test]
fn render_mechanics() {
    let Some(fx) = support::fixture() else { return };
    for (share, timing) in MODES {
        let Some(opened) =
            support::open_or_skip(&fx.cable.render, &Direction::Render, share, timing)
        else {
            continue;
        };
        let run = support::render_for(&opened.client, &opened.format, RUN_TIME)
            .expect("the render loop failed");
        assert_rate(&run, &opened.format, &format!("render {share} {timing:?}"));
        assert_eq!(opened.client.get_sharemode(), Some(share));
        assert_eq!(opened.client.get_timing_mode(), Some(timing));
    }
}

/// Free space is the buffer minus what is already in it, except in exclusive event
/// driven mode where it is always the whole buffer.
///
/// That branch in `get_available_space_in_frames` is the crate's own, and this is what
/// makes it necessary: in exclusive event driven mode the device owns the buffer, and
/// `get_current_padding` reports it as entirely full even before anything has been
/// written. Subtracting that padding would give zero free space forever, so a client
/// could never write at all. The other three modes report padding the obvious way.
///
/// Checked on a stream that has not been started, because padding and free space cannot
/// be read in one call: on a running stream the device drains frames between the two,
/// and a sum of the two readings means nothing.
#[test]
fn available_space_tracks_padding() {
    let Some(fx) = support::fixture() else { return };
    for (share, timing) in MODES {
        let Some(opened) =
            support::open_or_skip(&fx.cable.render, &Direction::Render, share, timing)
        else {
            continue;
        };
        let what = format!("render {share} {timing:?}");
        let buffer_size = opened.client.get_buffer_size().unwrap();
        let exclusive_events = share == ShareMode::Exclusive && timing == TimingMode::Events;
        let padding = || opened.client.get_current_padding().unwrap();
        let free = || opened.client.get_available_space_in_frames().unwrap();

        // Empty buffer. Exclusive event driven mode already calls it full.
        assert_eq!(
            padding(),
            if exclusive_events { buffer_size } else { 0 },
            "{what}: wrong padding for an empty buffer"
        );
        assert_eq!(
            free(),
            buffer_size,
            "{what}: an empty buffer did not report itself as entirely free"
        );

        // Fill it to the brim, without starting, so nothing drains underneath.
        let blockalign = opened.format.get_blockalign() as usize;
        let render = opened.client.get_audiorenderclient().unwrap();
        let frames = free() as usize;
        render
            .write_to_device(frames, &vec![0u8; frames * blockalign], None)
            .unwrap();
        assert_eq!(
            padding(),
            buffer_size,
            "{what}: a full buffer did not report itself as full"
        );
        assert_eq!(
            free(),
            if exclusive_events { buffer_size } else { 0 },
            "{what}: wrong free space for a full buffer"
        );
    }
}

/// Capture delivers data at the format's rate, in packets that account for it.
///
/// Nothing is playing into the cable. A WASAPI capture stream delivers silence at the
/// device rate regardless, so the two sides never have to be correlated.
///
/// `BufferInfo` has no other coverage, and the position is the part of it that can be
/// checked against something else: it counts frames since the stream started, so
/// consecutive packets should move it on by exactly the frames they carried. Two
/// things stop that being asserted for every packet. A reader that gets descheduled
/// loses packets, which leaves a gap. And `data_discontinuity`, which is meant to mark
/// exactly that, is not set in practice by any driver seen so far, so it cannot be
/// used to tell a real gap from a bug. So what is asserted is the part that holds
/// whatever the scheduler does, that the position never goes backwards and never moves
/// by less than the frames delivered, plus that most steps are exact, which still
/// fails loudly if the position were wrong in general.
#[test]
fn capture_mechanics() {
    let Some(fx) = support::fixture() else { return };
    for (share, timing) in MODES {
        let Some(opened) =
            support::open_or_skip(&fx.cable.capture, &Direction::Capture, share, timing)
        else {
            continue;
        };
        let what = format!("capture {share} {timing:?}");
        let run = support::capture_for(&opened.client, &opened.format, RUN_TIME)
            .expect("the capture loop failed");

        assert_rate(&run, &opened.format, &what);
        assert_eq!(opened.client.get_sharemode(), Some(share));
        assert_eq!(opened.client.get_timing_mode(), Some(timing));
        assert!(
            run.packets.len() > 1,
            "{what}: only {} packets",
            run.packets.len()
        );

        let (mut exact, mut gaps, mut flagged) = (0usize, 0usize, 0usize);
        for pair in run.packets.windows(2) {
            let (frames, before) = &pair[0];
            let (_, after) = &pair[1];
            let step = after.index.checked_sub(before.index).unwrap_or_else(|| {
                panic!(
                    "{what}: position went backwards, {} then {}",
                    before.index, after.index
                )
            });
            assert!(
                step >= u64::from(*frames),
                "{what}: a packet of {frames} frames only moved the position by {step}, \
                 so two packets overlap"
            );
            if step == u64::from(*frames) {
                exact += 1;
            } else {
                gaps += 1;
            }
            if after.flags.data_discontinuity {
                flagged += 1;
            }
            if !after.flags.timestamp_error && !before.flags.timestamp_error {
                assert!(
                    after.timestamp >= before.timestamp,
                    "{what}: timestamp went backwards, {} then {}",
                    before.timestamp,
                    after.timestamp
                );
            }
        }
        let steps = exact + gaps;
        println!(
            "{what}: {} packets, {exact} of {steps} steps exact, {gaps} gaps, \
             {flagged} flagged as a discontinuity",
            run.packets.len()
        );
        assert!(
            exact * 2 > steps,
            "{what}: only {exact} of {steps} steps accounted for their frames exactly"
        );
    }
}

/// A client reports nothing about itself until it is initialised.
///
/// `get_available_space_in_frames` is the crate's own guard rather than a Windows
/// result code, and the two mode getters read back what `initialize_client` recorded.
/// What they report once initialised is asserted by the two tests above, in all four
/// modes.
#[test]
fn an_uninitialised_client_reports_nothing() {
    let Some(fx) = support::fixture() else { return };
    let client = fx.cable.render.get_iaudioclient().unwrap();
    assert_eq!(client.get_sharemode(), None);
    assert_eq!(client.get_timing_mode(), None);
    assert!(matches!(
        client.get_available_space_in_frames(),
        Err(WasapiError::ClientNotInit)
    ));
}

/// The announced packet size is exactly what the next read returns.
///
/// A sharp invariant that only a real device can check, and the only place
/// `get_next_packet_size` is exercised. In exclusive mode it is documented to return
/// `None`, with the padding and the buffer size driving the loop instead.
///
/// Polling only, since this drives the stream by hand and an event driven client would
/// need its event handle set before starting.
#[test]
fn packet_size_matches_the_read() {
    let Some(fx) = support::fixture() else { return };
    for share in [ShareMode::Shared, ShareMode::Exclusive] {
        let Some(opened) = support::open_or_skip(
            &fx.cable.capture,
            &Direction::Capture,
            share,
            TimingMode::Polling,
        ) else {
            continue;
        };
        let what = format!("capture {share} polling");
        let capture = opened.client.get_audiocaptureclient().unwrap();
        let blockalign = opened.format.get_blockalign() as usize;
        let mut buffer = vec![0u8; opened.client.get_buffer_size().unwrap() as usize * blockalign];

        opened.client.start_stream().unwrap();
        let mut checked = 0;
        for _ in 0..200 {
            std::thread::sleep(Duration::from_millis(5));
            match capture.get_next_packet_size().unwrap() {
                // Exclusive mode has no packet size to announce.
                None => {
                    assert_eq!(
                        share,
                        ShareMode::Exclusive,
                        "{what}: no packet size was announced in shared mode"
                    );
                    break;
                }
                Some(0) => continue,
                Some(announced) => {
                    let (read, _info) = capture.read_from_device(&mut buffer).unwrap();
                    assert_eq!(
                        read, announced,
                        "{what}: {announced} frames were announced but {read} were read"
                    );
                    checked += 1;
                    if checked == 5 {
                        break;
                    }
                }
            }
        }
        opened.client.stop_stream().unwrap();
        if share == ShareMode::Shared {
            assert!(checked > 0, "{what}: no packet was ever announced");
        }
        println!("{what}: {checked} announced packet sizes matched the read");
    }
}

/// The period `calculate_aligned_period_near` works out is one the device accepts.
///
/// Real arithmetic on the crate's side, and the value has to survive being handed
/// straight back to an exclusive stream.
#[test]
fn device_period_is_sane_and_accepted() {
    let Some(fx) = support::fixture() else { return };
    let client = fx.cable.render.get_iaudioclient().unwrap();
    let (def_period, min_period) = client.get_device_period().unwrap();
    assert!(min_period > 0, "the minimum period is {min_period}");
    assert!(
        def_period >= min_period,
        "the default period {def_period} is below the minimum {min_period}"
    );

    let format = WaveFormat::new(16, 16, &SampleType::Int, 48000, 2, None);
    let aligned = client
        .calculate_aligned_period_near(min_period, Some(128), &format)
        .unwrap();
    assert!(
        aligned >= min_period,
        "the aligned period {aligned} is below the minimum {min_period}"
    );
    println!("periods: min {min_period}, default {def_period}, aligned {aligned}");
    drop(client);

    // And an exclusive stream really opens with a period from the same helper.
    match support::open(
        &fx.cable.render,
        &Direction::Render,
        ShareMode::Exclusive,
        TimingMode::Events,
    ) {
        Ok(opened) => assert!(opened.client.get_buffer_size().unwrap() > 0),
        Err(e) => support::skip(&format!("no exclusive mode stream: {e}")),
    }
}

/// The session goes active while a stream runs, and inactive again after it stops.
///
/// Polled, because both transitions are asynchronous.
#[test]
fn session_state_follows_the_stream() {
    let Some(fx) = support::fixture() else { return };
    let Some(opened) = support::open_or_skip(
        &fx.cable.render,
        &Direction::Render,
        ShareMode::Shared,
        TimingMode::Polling,
    ) else {
        return;
    };
    let session = opened.client.get_audiosessioncontrol().unwrap();
    let settles_on = |wanted: SessionState| {
        let deadline = Instant::now() + Duration::from_secs(2);
        while Instant::now() < deadline {
            if session.get_state().unwrap() == wanted {
                return true;
            }
            std::thread::sleep(Duration::from_millis(20));
        }
        false
    };

    assert_eq!(session.get_state().unwrap(), SessionState::Inactive);
    let blockalign = opened.format.get_blockalign() as usize;
    let render = opened.client.get_audiorenderclient().unwrap();
    let frames = opened.client.get_available_space_in_frames().unwrap() as usize;
    render
        .write_to_device(frames, &vec![0u8; frames * blockalign], None)
        .unwrap();

    opened.client.start_stream().unwrap();
    assert!(
        settles_on(SessionState::Active),
        "the session never went active"
    );
    opened.client.stop_stream().unwrap();
    assert!(
        settles_on(SessionState::Inactive),
        "the session never went inactive again"
    );
}
