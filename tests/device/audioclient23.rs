// The methods that reach IAudioClient2 and IAudioClient3 by casting.
//
// These tests are about the wrapping, not about WASAPI. No assertion says anything
// about what a period, a buffer limit or an offload answer should be, since those are
// the engine's and the driver's to pick, and a test that pinned them would fail on the
// next machine. What is the crate's own, and what is checked here:
//
// - The casts. Both interfaces come from the Windows audio engine rather than from the
//   driver, so they are there on any Windows that has them at all, IAudioClient2 since
//   Windows 8 and IAudioClient3 since Windows 10. A cast that asked for the wrong thing
//   fails for every endpoint, so reaching the call at all is the result.
// - The out parameters, which are easy to pass in the wrong order and which the
//   compiler cannot tell apart: four `u32` in a row for the engine period, two `i64`
//   for the buffer limits.
// - `WaveFormat::parse`, which `get_current_shared_mode_engine_period` is the only
//   caller of in the whole crate. It reads a `WAVEFORMATEX` through a raw pointer that
//   WASAPI allocated, checks `cbSize` and frees it again.
// - The `sharemode`, `timingmode` and `bytes_per_frame` fields, which
//   `initialize_shared_audio_stream` sets by hand instead of going through
//   `initialize_client`. `get_available_space_in_frames` branches on the first two and
//   the render client sizes its slice with the third.
//
// All of the query methods work on a client that was never initialised, which is why
// most of this needs no stream.

use crate::support;
use wasapi::*;

/// A fresh client on the cable's render endpoint, with the engine's own format.
///
/// The mix format is what the shared mode engine runs at, so it is the format the
/// periodicity queries are meant to be asked about.
fn render_client(fx: &support::Fixture) -> (AudioClient, WaveFormat) {
    let client = fx
        .cable
        .render
        .get_iaudioclient()
        .expect("get_iaudioclient failed");
    let format = client.get_mixformat().expect("get_mixformat failed");
    (client, format)
}

/// `is_offload_capable` answers for every category.
///
/// The categories go through a fourteen arm `From<StreamCategory>` match in the crate,
/// and nothing else exercises it. A category that mapped to something WASAPI does not
/// accept comes back as an error. What the endpoint answers is the endpoint's business,
/// so the answer is printed and not asserted.
#[test]
fn client2_offload_capable_answers_for_every_category() {
    let Some(fx) = support::fixture() else { return };
    let (client, _format) = render_client(&fx);

    // Deprecated by Windows, but still a category an endpoint can be asked about.
    #[allow(deprecated)]
    let categories = [
        StreamCategory::Other,
        StreamCategory::ForegroundOnlyMedia,
        StreamCategory::Communications,
        StreamCategory::Alerts,
        StreamCategory::SoundEffects,
        StreamCategory::GameEffects,
        StreamCategory::GameMedia,
        StreamCategory::GameChat,
        StreamCategory::Speech,
        StreamCategory::Movie,
        StreamCategory::Media,
        StreamCategory::FarFieldSpeech,
        StreamCategory::UniformSpeech,
        StreamCategory::VoiceTyping,
    ];

    for category in categories {
        let capable = client
            .is_offload_capable(category)
            .unwrap_or_else(|e| panic!("is_offload_capable failed for {category:?}: {e}"));
        println!("{category:?}: {capable}");
    }
}

/// `get_buffer_size_limits` reaches IAudioClient2 in both timing modes.
///
/// Only endpoints that support hardware offloaded streams have limits to report, so on
/// a virtual cable this fails, and in CI it always will. The failure is the endpoint
/// declining to answer, which means the call was made, so the only thing worth
/// asserting is that the cast was not what failed. The limits are printed when an
/// endpoint does answer, and not asserted: how long a hardware buffer may be is the
/// driver's to say.
#[test]
fn client2_buffer_size_limits_is_callable() {
    let Some(fx) = support::fixture() else { return };
    let (client, format) = render_client(&fx);

    for timing in [TimingMode::Events, TimingMode::Polling] {
        match client.get_buffer_size_limits(&format, timing) {
            Ok((min, max)) => println!("{timing:?}: {min} to {max} hundred nanoseconds"),
            Err(e) => {
                assert!(
                    !support::is_hresult(&e, E_NOINTERFACE),
                    "{timing:?}: the cast to IAudioClient2 failed"
                );
                println!("{timing:?}: the endpoint declined with {e}");
            }
        }
    }
}

/// The IAudioClient3 period queries work, and the format that comes back parses.
///
/// This is the test that covers `WaveFormat::parse`, which nothing else in the crate
/// calls. WASAPI hands back a `WAVEFORMATEX` it allocated itself, so a rate, channel
/// count and sample size that are all there on the other side mean the pointer was read
/// as a format rather than as something else, and that the extensible case was handled.
/// Which values they are is the engine's choice and is only printed.
///
/// The ordering checks on the four periods are about argument order, not about the
/// engine. They are four `u32` out parameters in a row, so passing them in the wrong
/// order compiles and gives a plausible looking struct; min and max arriving swapped, or
/// the fundamental landing in `default`, is what these catch.
///
/// Neither call needs an initialised client, so this one does not initialise.
#[test]
fn client3_engine_period_queries() {
    let Some(fx) = support::fixture() else { return };
    let (client, mixformat) = render_client(&fx);

    let (format, current) = client
        .get_current_shared_mode_engine_period()
        .expect("get_current_shared_mode_engine_period failed");
    println!(
        "current: {current} frames at {} Hz, {} channels, {} bits",
        format.get_samplespersec(),
        format.get_nchannels(),
        format.get_bitspersample()
    );
    // A zero here is an out parameter that was never written, not a short period.
    assert!(current > 0, "the period came back as zero frames");
    assert!(format.get_samplespersec() > 0, "a parsed format at zero Hz");
    assert!(
        format.get_nchannels() > 0,
        "a parsed format with no channels"
    );
    assert!(
        format.get_bitspersample() > 0,
        "a parsed format with no sample bits"
    );

    let period = client
        .get_shared_mode_engine_period(&mixformat)
        .expect("get_shared_mode_engine_period failed");
    println!("{period:?}");
    assert!(
        period.fundamental > 0,
        "the fundamental came back as zero frames"
    );
    assert!(
        period.min <= period.default && period.default <= period.max,
        "the default of {} frames is outside the {} to {} range, so the four out \
         parameters are in the wrong order",
        period.default,
        period.min,
        period.max
    );
}

/// A shared mode stream initialised with the engine's default period leaves the client
/// describing itself correctly.
///
/// The three fields are the whole point. `initialize_shared_audio_stream` sets them by
/// hand, so they can drift from what `initialize_client` does.
/// `get_available_space_in_frames` reads the first two and returns `ClientNotInit` when
/// either is unset, so getting a number out of it is the check. The write is what proves
/// the third: the slice is sized from the format's block align to fill the free space,
/// so a `bytes_per_frame` that disagrees either writes the wrong number of frames and
/// leaves space free, or is refused.
///
/// Nothing is started. Whether a shared mode stream then runs at the format's rate is
/// WASAPI's side of it, and the render tests in `streams` already ask that of an
/// ordinary client.
#[test]
fn client3_shared_audio_stream_sets_up_the_client() {
    let Some(fx) = support::fixture() else { return };

    for timing in [TimingMode::Events, TimingMode::Polling] {
        let (mut client, format) = render_client(&fx);
        let period = client
            .get_shared_mode_engine_period(&format)
            .expect("get_shared_mode_engine_period failed");

        client
            .initialize_shared_audio_stream(&format, timing, period.default)
            .unwrap_or_else(|e| {
                panic!(
                    "{timing:?}: the engine's own default of {} frames was refused with {e}",
                    period.default
                )
            });

        assert_eq!(
            client.get_sharemode(),
            Some(ShareMode::Shared),
            "{timing:?}: wrong sharemode after initialising"
        );
        assert_eq!(
            client.get_timing_mode(),
            Some(timing),
            "{timing:?}: wrong timing mode after initialising"
        );

        // An answer rather than `ClientNotInit` is what is being checked. How much of a
        // fresh buffer is free is WASAPI's business.
        let available = client
            .get_available_space_in_frames()
            .expect("get_available_space_in_frames failed");
        println!(
            "{timing:?}: {available} frames free, period {} frames",
            period.default
        );

        let blockalign = format.get_blockalign() as usize;
        let render = client
            .get_audiorenderclient()
            .expect("get_audiorenderclient failed");
        let silence = vec![0u8; available as usize * blockalign];
        render
            .write_to_device(&silence, None)
            .unwrap_or_else(|e| panic!("{timing:?}: writing {available} frames of silence: {e}"));
        // The stream is not started, so nothing drains what was just written.
        let left = client
            .get_available_space_in_frames()
            .expect("get_available_space_in_frames failed");
        assert_eq!(
            left, 0,
            "{timing:?}: writing {available} frames left {left} free"
        );
    }
}
