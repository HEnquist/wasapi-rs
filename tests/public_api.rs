// Integration tests that need no audio device.
//
// These run in the normal CI job, on a runner that has no audio endpoint at all.
// That is coverage no other test can give: the behaviour of the enumerator when
// there is nothing to enumerate is asserted here, guarded by WASAPI_EXPECT_NO_DEVICES
// so the same test is harmless on a machine that does have devices.
//
// Being outside the crate, these also use the public surface the way a downstream
// crate does. `WasapiRes` is crate private, so everything here spells out
// `Result<_, WasapiError>`.
//
// What is tested is the crate's own arithmetic: channel masks, periods, and the two
// bit mask wrappers. Nothing here calls into Windows except the enumerator test.

use wasapi::*;

/// Whether the runner is expected to have no audio devices at all.
fn expect_no_devices() -> bool {
    std::env::var_os("WASAPI_EXPECT_NO_DEVICES").is_some()
}

/// The enumerator works on a machine with no audio device.
///
/// Creating it and listing an empty collection has to succeed, while asking for a
/// default device has to fail rather than panic. Only asserted when the
/// environment says there really are no devices, which is the CI job that runs on
/// a bare Windows runner.
#[test]
fn enumerator_works_without_any_device() {
    initialize_mta().ok().unwrap();
    let enumerator = DeviceEnumerator::new().expect("DeviceEnumerator::new failed");

    for direction in [Direction::Render, Direction::Capture] {
        let collection = enumerator
            .get_device_collection(&direction)
            .expect("get_device_collection failed");
        let count = collection
            .get_nbr_devices()
            .expect("get_nbr_devices failed");
        let default = enumerator.get_default_device(&direction);
        match &default {
            Ok(device) => println!(
                "{direction}: {count} devices, default: {}",
                device.get_friendlyname().unwrap_or_default()
            ),
            Err(e) => println!("{direction}: {count} devices, no default device: {e}"),
        }

        if expect_no_devices() {
            assert_eq!(count, 0, "{direction} devices on a runner with none");
            assert!(
                default.is_err(),
                "a default {direction} device appeared on a runner with none"
            );
            // An empty collection still iterates, it just yields nothing.
            assert_eq!((&collection).into_iter().count(), 0);
            assert!(collection.get_device_at_index(0).is_err());
        }
    }
}

/// A channel mask for an n channel layout has exactly n positions in it.
///
/// The doctest on `make_channelmasks` covers three channels. This covers the rest,
/// including the two counts where the list collapses.
#[test]
fn channel_masks_are_consistent() {
    for channels in 1..=18 {
        let masks = make_channelmasks(channels);
        assert_eq!(
            masks.last(),
            Some(&0),
            "the {channels} channel list does not end with the zero mask: {masks:x?}"
        );
        for mask in &masks[..masks.len() - 1] {
            assert_ne!(
                *mask, 0,
                "a zero mask before the end for {channels} channels"
            );
            assert_eq!(
                mask.count_ones() as usize,
                channels,
                "mask {mask:#x} for {channels} channels has {} positions",
                mask.count_ones()
            );
        }
        assert_eq!(make_simple_channelmask(channels), (1 << channels) - 1);
    }

    // Above the 18 defined positions there is nothing left to assign, so the only
    // option is the zero mask.
    for channels in 19..=32 {
        assert_eq!(make_channelmasks(channels), vec![0]);
        assert_eq!(make_simple_channelmask(channels), 0);
    }
}

/// A period converted from frames and back gives the frames again.
#[test]
fn period_arithmetic_round_trips() {
    for rate in [44100, 48000, 88200, 96000, 192000] {
        let mut previous = 0;
        for frames in [1, 16, 128, 441, 480, 1024, 4096, 44100] {
            let period = calculate_period_100ns(frames, rate);
            assert!(period > previous, "{period} is not above {previous}");
            previous = period;

            // 100 ns units, so a period of p covers p * rate / 10_000_000 frames.
            let back = (period as f64 * rate as f64 / 10_000_000.0).round() as i64;
            assert!(
                (back - frames).abs() <= 1,
                "{frames} frames at {rate} Hz became {period} units and back to {back}"
            );
        }
    }
}

/// The three buffer flags survive a round trip through the raw bit mask.
#[test]
fn buffer_flags_round_trip() {
    let none = BufferFlags::none();
    assert!(!none.data_discontinuity && !none.silent && !none.timestamp_error);
    assert_eq!(none.to_u32(), 0);

    // AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY, _SILENT and _TIMESTAMP_ERROR are
    // bits 0, 1 and 2, and each one is read out on its own.
    let discontinuity = BufferFlags::new(0x1);
    assert!(discontinuity.data_discontinuity && !discontinuity.silent);
    let silent = BufferFlags::new(0x2);
    assert!(silent.silent && !silent.data_discontinuity && !silent.timestamp_error);
    let timestamp = BufferFlags::new(0x4);
    assert!(timestamp.timestamp_error && !timestamp.silent);

    // Every combination of the three survives the round trip.
    for mask in 0..8u32 {
        let flags = BufferFlags::new(mask);
        assert_eq!(
            flags.to_u32(),
            mask,
            "{mask:#x} came back as {:#x}",
            flags.to_u32()
        );
    }

    // A bit the crate does not model is dropped rather than carried through.
    assert_eq!(BufferFlags::new(0x8).to_u32(), 0);
}

/// The hardware support bits are read out independently of each other.
#[test]
fn hardware_support_bits() {
    let none = HardwareSupport::new(0);
    assert!(!none.volume && !none.mute && !none.meter);

    // ENDPOINT_HARDWARE_SUPPORT_VOLUME, _MUTE and _METER are bits 0, 1 and 2.
    let volume = HardwareSupport::new(0x1);
    assert!(volume.volume && !volume.mute && !volume.meter);
    let mute = HardwareSupport::new(0x2);
    assert!(!mute.volume && mute.mute && !mute.meter);
    let meter = HardwareSupport::new(0x4);
    assert!(!meter.volume && !meter.mute && meter.meter);

    let all = HardwareSupport::new(0x7);
    assert!(all.volume && all.mute && all.meter);
}

/// A stream option accumulates into the properties rather than replacing what is there.
///
/// `set_properties` itself is a one-line call into Windows, but `set_option` ORs its
/// bits together, which is the crate's own arithmetic. The builder exposes no getter,
/// so the Debug output is the only way to see it from outside.
#[test]
fn client_properties_accumulate_options() {
    let raw = AudioClientProperties::new().set_option(StreamOption::Raw);
    let both = AudioClientProperties::new()
        .set_option(StreamOption::Raw)
        .set_option(StreamOption::MatchFormat);
    let twice = AudioClientProperties::new()
        .set_option(StreamOption::Raw)
        .set_option(StreamOption::Raw);

    // A second, different option has to add to the first rather than replace it.
    assert_ne!(
        format!("{raw:?}"),
        format!("{both:?}"),
        "a second option did not change the properties"
    );
    // The same option twice is still just that option, since the bits are ORed.
    assert_eq!(
        format!("{raw:?}"),
        format!("{twice:?}"),
        "setting the same option twice changed the result"
    );
}
