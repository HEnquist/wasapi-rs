// Format support queries, the exclusive mode capability probe and the data ranges
// a driver declares.
//
// Nothing here plays or records anything. Every test asks the library two questions
// and checks that the answers agree, which is what makes these the sharpest tests in
// the suite: no timing, no tolerances, and a failure is either a real bug or a
// documented driver quirk.

use crate::support;
use wasapi::*;
use windows::Win32::Media::Audio::{
    AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED, AUDCLNT_E_DEVICE_IN_USE, AUDCLNT_E_UNSUPPORTED_FORMAT,
};

/// The formats to ask about, as `(name, storebits, validbits, sample type)`.
const FORMATS: [(&str, usize, usize, SampleType); 5] = [
    ("S16", 16, 16, SampleType::Int),
    ("S24 packed", 24, 24, SampleType::Int),
    ("S24 in 32", 32, 24, SampleType::Int),
    ("S32", 32, 32, SampleType::Int),
    ("F32", 32, 32, SampleType::Float),
];

/// Build one of [FORMATS] at a rate and channel count.
fn format_of(spec: &(&str, usize, usize, SampleType), rate: usize, channels: usize) -> WaveFormat {
    WaveFormat::new(spec.1, spec.2, &spec.3, rate, channels, None)
}

/// A format's bit depths and sample type, enough to tell any two of [FORMATS] apart.
fn describe_bits(format: &WaveFormat) -> (u16, u16, SampleType) {
    (
        format.get_validbitspersample(),
        format.get_bitspersample(),
        format
            .get_subformat()
            .expect("a format with no sample type"),
    )
}

/// Whether `is_supported` approved the format as it was given.
///
/// `Ok(None)` is "supported as is". In shared mode `Ok(Some(_))` means the device
/// offered a near match instead, so the format itself was not accepted.
fn approved(result: &Result<Option<WaveFormat>, WasapiError>) -> bool {
    matches!(result, Ok(None))
}

/// S16 at 48 kHz in stereo, the format VB-Cable is known to have in exclusive mode.
fn s16_48k() -> WaveFormat {
    WaveFormat::new(16, 16, &SampleType::Int, 48000, 2, None)
}

/// What happened when a format was handed to `initialize_client`.
enum Attempt {
    Accepted,
    Refused,
    /// The device was busy or wanted a different buffer alignment, which says
    /// nothing about whether it supports the format.
    Inconclusive(String),
}

/// Try to initialise a fresh client with a format.
///
/// A second `Initialize` on the same client is refused, so every attempt needs a
/// client of its own.
fn try_initialize(device: &Device, share: ShareMode, format: &WaveFormat) -> Attempt {
    let mut client = device.get_iaudioclient().expect("get_iaudioclient failed");
    let (def_period, _min) = client
        .get_device_period()
        .expect("get_device_period failed");
    // Polling, and without autoconvert in shared mode, so the engine has to take the
    // format as it is. That is the same question is_supported was asked.
    let mode = match support::mode_for(&client, share, TimingMode::Polling, def_period, format) {
        Ok(mode) => mode,
        Err(e) => return Attempt::Inconclusive(format!("no aligned period: {e}")),
    };
    match client.initialize_client(format, &device.get_direction(), &mode) {
        Ok(()) => Attempt::Accepted,
        Err(e)
            if support::is_hresult(&e, AUDCLNT_E_DEVICE_IN_USE)
                || support::is_hresult(&e, AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED) =>
        {
            Attempt::Inconclusive(format!("{e}"))
        }
        Err(_) => Attempt::Refused,
    }
}

/// What `is_supported` says has to be what `initialize_client` does.
///
/// The strongest test in the suite. Every mismatch is collected and reported
/// together rather than failing on the first one, so one run says everything about
/// what the device accepts.
///
/// The rates are the device's own mix rate, so shared mode has something it can
/// accept on any machine, plus 96 kHz as one the engine is not running. One and two
/// channels, because that is where drivers are documented to disagree between a
/// `WAVEFORMATEX` and a `WAVEFORMATEXTENSIBLE`.
#[test]
fn is_supported_agrees_with_initialize_client() {
    let Some(fx) = support::fixture() else { return };
    let mut mismatches: Vec<String> = Vec::new();
    let (mut conclusive, mut inconclusive) = (0, 0);

    for (device, side) in [(&fx.cable.render, "render"), (&fx.cable.capture, "capture")] {
        let client = device.get_iaudioclient().expect("get_iaudioclient failed");
        let mix_rate = client.get_mixformat().unwrap().get_samplespersec() as usize;
        for share in [ShareMode::Shared, ShareMode::Exclusive] {
            for rate in [mix_rate, 96000] {
                for channels in [1, 2] {
                    for spec in &FORMATS {
                        let format = format_of(spec, rate, channels);
                        let query = client.is_supported(&format, &share);
                        let says_yes = approved(&query);
                        let what = format!("{side} {share} {} {rate} Hz {channels} ch", spec.0);
                        match try_initialize(device, share, &format) {
                            Attempt::Inconclusive(why) => {
                                inconclusive += 1;
                                println!("  inconclusive, {what}: {why}");
                            }
                            Attempt::Accepted if !says_yes => mismatches.push(format!(
                                "{what}: is_supported said no ({query:?}) but \
                                 initialize_client accepted it"
                            )),
                            Attempt::Refused if says_yes => mismatches.push(format!(
                                "{what}: is_supported said yes but initialize_client refused it"
                            )),
                            Attempt::Accepted => {
                                conclusive += 1;
                                println!("  both yes, {what}");
                            }
                            Attempt::Refused => conclusive += 1,
                        }
                    }
                }
            }
        }
    }

    println!(
        "{conclusive} agreed, {inconclusive} inconclusive, {} disagreed",
        mismatches.len()
    );
    assert!(
        mismatches.is_empty(),
        "is_supported and initialize_client disagree:\n  {}",
        mismatches.join("\n  ")
    );
    assert!(conclusive > 0, "every attempt was inconclusive");
}

/// The mix format is the one format shared mode always takes, and it describes the
/// same stream as the format stored for the endpoint.
///
/// The mix format is what the engine runs, so it is always float. The device format
/// comes from the property store and may be an integer format. They have to agree on
/// the rate and the channel count. This is also the only coverage
/// `WaveFormat::parse_from_blob_bytes` gets against a real registry blob.
#[test]
fn mixformat_and_deviceformat() {
    let Some(fx) = support::fixture() else { return };
    for device in [&fx.cable.render, &fx.cable.capture] {
        let stored = device.get_device_format().unwrap();
        let client = device.get_iaudioclient().unwrap();
        let mix = client.get_mixformat().unwrap();

        assert!(
            approved(&client.is_supported(&mix, &ShareMode::Shared)),
            "the mix format was not accepted as is"
        );
        drop(client);
        assert!(
            matches!(
                try_initialize(device, ShareMode::Shared, &mix),
                Attempt::Accepted
            ),
            "the mix format did not initialise"
        );

        assert_eq!(
            stored.get_samplespersec(),
            mix.get_samplespersec(),
            "the stored and mix formats disagree on the rate"
        );
        assert_eq!(
            stored.get_nchannels(),
            mix.get_nchannels(),
            "the stored and mix formats disagree on the channel count"
        );
        assert_eq!(mix.get_subformat().unwrap(), SampleType::Float);
        assert_eq!(mix.get_bitspersample(), 32);
        assert!(stored.get_blockalign() > 0);
    }
}

/// VB-Cable has the two integer formats in exclusive mode, and not the other three.
///
/// The refusals are only asserted on their result code, not on the fact of being
/// refused, so a future driver that gains a format does not fail this. The sharp
/// version of that invariant is `is_supported_agrees_with_initialize_client`.
#[test]
fn exclusive_supports_s16_and_s24() {
    let Some(fx) = support::fixture() else { return };
    let client = fx.cable.render.get_iaudioclient().unwrap();

    for (storebits, validbits) in [(16, 16), (24, 24)] {
        let format = WaveFormat::new(storebits, validbits, &SampleType::Int, 48000, 2, None);
        // Some drivers only accept one and two channel formats when asked as a plain
        // WAVEFORMATEX, which is what the quirks helper exists for.
        assert!(
            approved(&client.is_supported(&format, &ShareMode::Exclusive))
                || client.is_supported_exclusive_with_quirks(&format).is_ok(),
            "{validbits} in {storebits} bits at 48 kHz is not supported in exclusive mode"
        );
    }

    for spec in FORMATS.iter().filter(|s| s.0 == "S32" || s.0 == "F32") {
        let format = format_of(spec, 48000, 2);
        match client.is_supported(&format, &ShareMode::Exclusive) {
            Ok(None) => println!("note: this device does accept {} in exclusive mode", spec.0),
            Ok(Some(_)) => panic!("exclusive mode offered a near match, which it never does"),
            Err(e) => assert!(
                support::is_hresult(&e, AUDCLNT_E_UNSUPPORTED_FORMAT),
                "{} was refused with {e}, expected AUDCLNT_E_UNSUPPORTED_FORMAT",
                spec.0
            ),
        }
    }
}

/// The autoconvert flag is what makes shared mode take a rate the engine is not
/// running.
#[test]
fn autoconvert_accepts_an_unsupported_rate() {
    let Some(fx) = support::fixture() else { return };
    let mut client = fx.cable.render.get_iaudioclient().unwrap();
    let mix = client.get_mixformat().unwrap();
    let odd_rate = if mix.get_samplespersec() == 44100 {
        48000
    } else {
        44100
    };
    let format = WaveFormat::new(
        32,
        32,
        &SampleType::Float,
        odd_rate,
        mix.get_nchannels() as usize,
        None,
    );
    if approved(&client.is_supported(&format, &ShareMode::Shared)) {
        support::skip(&format!("the engine already runs at {odd_rate} Hz"));
        return;
    }

    // Refused without the flag, accepted with it.
    assert!(
        matches!(
            try_initialize(&fx.cable.render, ShareMode::Shared, &format),
            Attempt::Refused
        ),
        "{odd_rate} Hz was accepted in shared mode without autoconvert"
    );
    let (def_period, _min) = client.get_device_period().unwrap();
    client
        .initialize_client(
            &format,
            &Direction::Render,
            &StreamMode::PollingShared {
                autoconvert: true,
                buffer_duration_hns: def_period,
            },
        )
        .expect("autoconvert did not accept a rate the engine is not running");
}

/// The quirks helper returns a format the device really takes, mask and all.
#[test]
fn quirks_helper_returns_a_usable_format() {
    let Some(fx) = support::fixture() else { return };
    let client = fx.cable.render.get_iaudioclient().unwrap();
    let Ok(found) = client.is_supported_exclusive_with_quirks(&s16_48k()) else {
        support::skip("no S16 48 kHz stereo in exclusive mode");
        return;
    };
    drop(client);

    // The same stream, possibly with a different channel mask.
    assert_eq!(found.get_samplespersec(), 48000);
    assert_eq!(found.get_nchannels(), 2);
    assert_eq!(found.get_bitspersample(), 16);
    assert_eq!(found.get_validbitspersample(), 16);
    assert!(
        make_channelmasks(2).contains(&found.get_dwchannelmask()),
        "the accepted mask {:#x} is not one of {:x?}",
        found.get_dwchannelmask(),
        make_channelmasks(2)
    );
    assert!(
        matches!(
            try_initialize(&fx.cable.render, ShareMode::Exclusive, &found),
            Attempt::Accepted | Attempt::Inconclusive(_)
        ),
        "the format the quirks helper returned did not initialise"
    );
}

/// What the probe reports is exactly what the device accepts.
///
/// The unit tests in src/capabilities.rs run the pruning logic against a fake
/// device. This runs the public entry point against a real one, which nothing
/// covers today.
#[test]
fn probe_matches_is_supported() {
    let Some(fx) = support::fixture() else { return };
    let mut probe = CapabilityProbe::new(&fx.cable.render).expect("CapabilityProbe::new failed");
    let found = probe.supported_formats(48000, 2);
    // Keyed on the sample type as well as the bit depths, since S32 and F32 are both
    // 32 in 32 and would otherwise be indistinguishable here.
    let reported: Vec<(u16, u16, SampleType)> = found.iter().map(describe_bits).collect();
    println!(
        "the probe found {} stereo formats at 48 kHz: {reported:?}",
        found.len()
    );
    assert!(!found.is_empty(), "the probe found no format at all");

    let client = fx.cable.render.get_iaudioclient().unwrap();
    // Everything it reported has to be accepted.
    for format in &found {
        assert!(
            client.is_supported_exclusive_with_quirks(format).is_ok(),
            "the probe reported {}/{} bits, which the device refuses",
            format.get_validbitspersample(),
            format.get_bitspersample()
        );
    }
    // And nothing it left out may be accepted.
    for spec in &FORMATS {
        let format = format_of(spec, 48000, 2);
        if reported.contains(&describe_bits(&format)) {
            continue;
        }
        assert!(
            client.is_supported_exclusive_with_quirks(&format).is_err(),
            "the device accepts {} but the probe did not report it",
            spec.0
        );
    }
}

/// A full scan finds the format the cable is known to have, and stays inside what
/// the driver declared.
///
/// A driver may declare more than it accepts, but never less, so every format the
/// scan finds has to fall inside some declared range. That is the check
/// examples/dataranges.rs performs by hand.
#[test]
fn probe_full_scan_agrees_with_declared_ranges() {
    let Some(fx) = support::fixture() else { return };
    for device in [&fx.cable.render, &fx.cable.capture] {
        let mut probe = CapabilityProbe::new(device).unwrap();
        let ranges = probe.data_ranges().to_vec();

        let start = std::time::Instant::now();
        let found = probe.supported_formats_all_rates();
        println!(
            "the full scan found {} formats in {:.1} s",
            found.len(),
            start.elapsed().as_secs_f64()
        );
        assert!(
            found.iter().any(|f| f.get_samplespersec() == 48000
                && f.get_nchannels() == 2
                && f.get_bitspersample() == 16),
            "the full scan did not find S16 48 kHz stereo"
        );

        if ranges.is_empty() {
            support::skip("the driver declares no data ranges");
            continue;
        }
        for format in &found {
            assert!(
                ranges.iter().any(|range| range.covers(format)),
                "{} ch, {} Hz, {}/{} bits is accepted but outside every declared range",
                format.get_nchannels(),
                format.get_samplespersec(),
                format.get_validbitspersample(),
                format.get_bitspersample()
            );
        }
    }
}

/// Narrowing the scan with the declared ranges must not change what it finds.
///
/// This is the claim the `CapabilityProbe` documentation makes: the ranges only make
/// the scan cheaper, never different.
#[test]
fn probe_with_and_without_ranges_agree() {
    let Some(fx) = support::fixture() else { return };
    let mut probe = CapabilityProbe::new(&fx.cable.render).unwrap();
    if probe.data_ranges().is_empty() {
        support::skip("the driver declares no data ranges, so there is nothing to compare");
        return;
    }

    let describe = |formats: Vec<WaveFormat>| {
        let mut out: Vec<(u16, u16, u16)> = formats
            .iter()
            .map(|f| {
                (
                    f.get_nchannels(),
                    f.get_validbitspersample(),
                    f.get_bitspersample(),
                )
            })
            .collect();
        out.sort_unstable();
        out
    };

    let bounded = describe(probe.supported_formats_at_rate(48000));
    probe.set_data_ranges(Vec::new());
    let staged = describe(probe.supported_formats_at_rate(48000));
    assert_eq!(
        bounded, staged,
        "the bounded scan and the staged scan disagree at 48 kHz"
    );
}

/// The driver declares usable data ranges.
///
/// `Device::get_data_ranges` walks the device topology to the kernel streaming
/// filter and sends it an IOCTL. Only the parsing of the reply is unit tested, so
/// this is the only coverage the walk and the IOCTL get.
#[test]
fn dataranges_are_declared_for_the_cable() {
    let Some(fx) = support::fixture() else { return };
    for device in [&fx.cable.render, &fx.cable.capture] {
        let ranges = device
            .get_data_ranges()
            .expect("get_data_ranges failed for a VB-Cable endpoint");
        assert!(
            !ranges.is_empty(),
            "the driver declared no ranges for {}",
            support::describe(device)
        );

        let mut covers_48k = false;
        let mut has_integer = false;
        for range in &ranges {
            assert!(
                range.min_bits_per_sample <= range.max_bits_per_sample,
                "{range:?} has its bit depths the wrong way round"
            );
            assert!(
                range.min_samplerate <= range.max_samplerate,
                "{range:?} has its rates the wrong way round"
            );
            assert!(
                range.max_channels >= 2,
                "{range:?} allows fewer than 2 channels"
            );
            covers_48k |= range.min_samplerate <= 48000 && range.max_samplerate >= 48000;
            has_integer |= range.sample_type() == Some(SampleType::Int);
        }
        assert!(covers_48k, "no declared range covers 48 kHz");
        assert!(has_integer, "no declared range is an integer format");
    }
}

/// A 16 channel format is accepted, and comes back with a mask the crate offered.
///
/// This is the channel mask retry above 8 channels, where `make_channelmasks` has no
/// named layout left and returns only the simple mask and zero. Nothing else in the
/// suite goes past two channels.
///
/// It deliberately does not compare the two render endpoints. The cable has a 16
/// channel endpoint alongside the plain one, but both declare `max_channels: 16` and
/// both accept up to 16 channels, so there is nothing to tell apart: an earlier version
/// of this test claimed the 16 channel endpoint declared more, which is not true.
#[test]
fn a_sixteen_channel_format_negotiates_a_mask() {
    let Some(fx) = support::fixture() else { return };
    let client = fx.cable.render.get_iaudioclient().unwrap();
    let wanted = WaveFormat::new(16, 16, &SampleType::Int, 48000, 16, None);
    match client.is_supported_exclusive_with_quirks(&wanted) {
        Ok(found) => {
            assert_eq!(found.get_nchannels(), 16);
            assert_eq!(found.get_samplespersec(), 48000);
            let masks = make_channelmasks(16);
            assert!(
                masks.contains(&found.get_dwchannelmask()),
                "the accepted mask {:#x} is not one of the {masks:x?} the crate offers",
                found.get_dwchannelmask()
            );
            println!(
                "16 channels accepted with mask {:#x}",
                found.get_dwchannelmask()
            );
        }
        Err(e) => support::skip(&format!("no 16 channel S16 format in exclusive mode: {e}")),
    }
}
