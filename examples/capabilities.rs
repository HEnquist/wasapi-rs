use std::collections::BTreeMap;
use std::time::Instant;
use wasapi::*;

use simplelog::*;

/// Make a short label for a sample format, such as "S16", "F32" or "S24_in_32".
fn format_name(wave_fmt: &WaveFormat) -> String {
    let sample_type = match wave_fmt.get_subformat() {
        Ok(SampleType::Float) => "F",
        _ => "S",
    };
    let storebits = wave_fmt.get_bitspersample();
    let validbits = wave_fmt.get_validbitspersample();
    if storebits == validbits {
        format!("{sample_type}{storebits}")
    } else {
        format!("{sample_type}{validbits}_in_{storebits}")
    }
}

// Scan the default output device for the formats it supports in exclusive mode.
fn main() {
    let _ = SimpleLogger::init(
        LevelFilter::Info,
        ConfigBuilder::new()
            .set_time_format_rfc3339()
            .set_time_offset_to_local()
            .unwrap()
            .build(),
    );

    initialize_mta().unwrap();

    let enumerator = DeviceEnumerator::new().unwrap();
    let device = enumerator.get_default_device(&Direction::Render).unwrap();
    println!(
        "Scanning device {:?}, this takes a while..",
        device.get_friendlyname().unwrap()
    );

    let mut probe = CapabilityProbe::new(device.get_iaudioclient().unwrap());
    let start = Instant::now();
    let formats = probe.supported_formats_all_rates(DEFAULT_MAX_CHANNELS);
    println!("The scan took {:.1} s.", start.elapsed().as_secs_f32());

    // Group the formats by channel count and sample rate.
    let mut grouped: BTreeMap<u16, BTreeMap<u32, Vec<String>>> = BTreeMap::new();
    for wave_fmt in &formats {
        grouped
            .entry(wave_fmt.get_nchannels())
            .or_default()
            .entry(wave_fmt.get_samplespersec())
            .or_default()
            .push(format_name(wave_fmt));
    }

    if grouped.is_empty() {
        println!("The device supports nothing in exclusive mode.");
        return;
    }
    for (channels, rates) in grouped {
        println!("{channels} channels:");
        for (samplerate, names) in rates {
            println!("  {samplerate} Hz: {}", names.join(", "));
        }
    }
}
