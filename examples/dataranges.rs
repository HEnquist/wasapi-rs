// Compare the capabilities that a driver declares with what the device really accepts.
//
// For every active output and input device this prints the declared data ranges,
// then runs a full scan both with and without them, and compares the results.
// It answers two questions, whether the declared ranges can be trusted,
// and how much they help.
//
// This is a check of the data ranges themselves.
// To simply list what a device supports, use the capabilities example instead.
//
// Give a substring of a device name as an argument to only check the matching devices.

use std::collections::BTreeSet;
use std::time::Instant;

use wasapi::*;

use simplelog::*;

/// A format as rate, channels, stored bits, valid bits and sample type.
type Described = (u32, u16, u16, u16, String);

/// Describe a format as rate, channels, stored bits, valid bits and sample type.
fn describe(wave_fmt: &WaveFormat) -> Described {
    (
        wave_fmt.get_samplespersec(),
        wave_fmt.get_nchannels(),
        wave_fmt.get_bitspersample(),
        wave_fmt.get_validbitspersample(),
        match wave_fmt.get_subformat() {
            Ok(sample_type) => sample_type.to_string(),
            Err(_) => "unknown".to_string(),
        },
    )
}

fn scan(probe: &mut CapabilityProbe) -> (BTreeSet<Described>, u128) {
    let start = Instant::now();
    let formats = probe.supported_formats_all_rates();
    let elapsed = start.elapsed().as_millis();
    (formats.iter().map(describe).collect(), elapsed)
}

fn main() {
    let wanted = std::env::args().nth(1).unwrap_or_default().to_lowercase();
    let _ = SimpleLogger::init(
        LevelFilter::Warn,
        ConfigBuilder::new()
            .set_time_format_rfc3339()
            .set_time_offset_to_local()
            .unwrap()
            .build(),
    );
    initialize_mta().ok().unwrap();

    let enumerator = DeviceEnumerator::new().unwrap();
    let collections = [
        enumerator
            .get_device_collection(&Direction::Render)
            .unwrap(),
        enumerator
            .get_device_collection(&Direction::Capture)
            .unwrap(),
    ];

    for device in collections.iter().flatten() {
        let device = device.unwrap();
        let name = device.get_friendlyname().unwrap_or_default();
        if !wanted.is_empty() && !name.to_lowercase().contains(&wanted) {
            continue;
        }
        println!("\n{} device {name:?}", device.get_direction());

        let ranges = match device.get_data_ranges() {
            Ok(ranges) if ranges.is_empty() => {
                println!("  the driver declares nothing");
                Vec::new()
            }
            Ok(ranges) => {
                for range in &ranges {
                    println!(
                        "  declared: {} ch, {}-{} bits, {}-{} Hz, {}",
                        range.max_channels,
                        range.min_bits_per_sample,
                        range.max_bits_per_sample,
                        range.min_samplerate,
                        range.max_samplerate,
                        match range.sample_type() {
                            Some(sample_type) => sample_type.to_string(),
                            None => "any".to_string(),
                        }
                    );
                }
                ranges
            }
            Err(err) => {
                println!("  could not read the data ranges: {err}");
                Vec::new()
            }
        };

        // Clear the ranges to get the staged scan, for comparison.
        let mut plain = CapabilityProbe::new(&device).unwrap();
        plain.set_data_ranges(Vec::new());
        let (staged, staged_time) = scan(&mut plain);
        println!(
            "  the staged scan found {} formats in {staged_time} ms",
            staged.len()
        );

        if ranges.is_empty() {
            continue;
        }
        let mut bounded = CapabilityProbe::new(&device).unwrap();
        bounded.set_data_ranges(ranges);
        let (found, found_time) = scan(&mut bounded);
        println!(
            "  the bounded scan found {} formats in {found_time} ms",
            found.len()
        );

        for missed in staged.difference(&found) {
            println!("  MISSED by the declared ranges: {missed:?}");
        }
        for extra in found.difference(&staged) {
            println!("  found only with the declared ranges: {extra:?}");
        }
        if staged == found {
            println!("  the two scans agree");
        }
    }
}
