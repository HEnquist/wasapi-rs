// Play a sine wave in shared mode on a device found through WinRT.
//
// The devices are enumerated with the WinRT `DeviceInformation` and `MediaDevice` classes
// from the `windows` crate, and the chosen one is opened by its WinRT id with
// `AudioClient::new_device_interface_client`. Nothing from the `DeviceCollection` of
// this library is used.
//
// Pass part of a device name to play on that device, for example
// `cargo run --example playsine_winrt -- Speakers`.
// Without an argument, the sine plays on the current default output device.

use std::f64::consts::PI;
use wasapi::*;
use windows::Devices::Enumeration::DeviceInformation;
use windows::Media::Devices::{AudioDeviceRole, MediaDevice};

#[macro_use]
extern crate log;
use simplelog::*;

fn main() {
    let _ = SimpleLogger::init(
        LevelFilter::Debug,
        ConfigBuilder::new()
            .set_time_format_rfc3339()
            .set_time_offset_to_local()
            .unwrap()
            .build(),
    );

    // There is no initialize_mta here. The first use of a WinRT class sets up the
    // multithreaded apartment of the process, and a thread that has not initialised
    // COM itself then runs in the implicit MTA, which is enough for the activation.
    // A thread that has to be in a single-threaded apartment, such as a UI thread,
    // must initialise COM itself as usual.
    let selector = MediaDevice::GetAudioRenderSelector().unwrap();
    let devices = DeviceInformation::FindAllAsyncAqsFilter(&selector)
        .unwrap()
        .join()
        .unwrap();
    // DeviceInformation.IsDefault is always false for audio endpoints,
    // so the default is looked up separately.
    let default_id = MediaDevice::GetDefaultAudioRenderId(AudioDeviceRole::Default).unwrap();

    info!("Output devices:");
    for device in &devices {
        let marker = if device.Id().unwrap() == default_id {
            " (default)"
        } else {
            ""
        };
        info!("  {}{marker}", device.Name().unwrap());
    }

    let id = match std::env::args().nth(1) {
        Some(wanted) => {
            let found = devices
                .into_iter()
                .find(|device| device.Name().unwrap().to_string().contains(&wanted));
            match found {
                Some(device) => device.Id().unwrap(),
                None => {
                    error!("No output device with \"{wanted}\" in its name");
                    return;
                }
            }
        }
        None => default_id,
    };
    let id = id.to_string();
    info!("Opening {id}");

    let mut audio_client =
        AudioClient::new_device_interface_client(&id, Direction::Render).unwrap();

    let channels = 2;
    let samplerate = 44100;
    let format = WaveFormat::new(32, 32, SampleType::Float, samplerate, channels, None);
    let blockalign = format.get_blockalign() as usize;
    let (def_time, _min_time) = audio_client.get_device_period().unwrap();
    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: def_time,
    };
    audio_client
        .initialize_client(&format, Direction::Render, mode, None)
        .unwrap();
    let h_event = audio_client.set_get_eventhandle().unwrap();
    let render_client = audio_client.get_audiorenderclient().unwrap();

    let delta_phase = 2.0 * PI * 1000.0 / samplerate as f64;
    let mut phase = 0.0f64;
    let mut write_frames = |nbr_frames: usize| -> Result<(), WasapiError> {
        let mut data = vec![0u8; nbr_frames * blockalign];
        for frame in data.chunks_exact_mut(blockalign) {
            let sample = (0.1 * phase.sin()) as f32;
            phase = (phase + delta_phase) % (2.0 * PI);
            for value in frame.chunks_exact_mut(4) {
                value.copy_from_slice(&sample.to_le_bytes());
            }
        }
        render_client.write_to_device(&data, None)
    };

    write_frames(audio_client.get_available_space_in_frames().unwrap() as usize).unwrap();
    audio_client.start_stream().unwrap();
    info!("Playing for 3 seconds");
    let start = std::time::Instant::now();
    while start.elapsed() < std::time::Duration::from_secs(3) {
        if let Err(err) = h_event.wait_for_event(1000) {
            error!("Waiting for the event failed, {err}");
            break;
        }
        let result = audio_client
            .get_available_space_in_frames()
            .and_then(|frames| write_frames(frames as usize));
        if let Err(err) = result {
            error!("Writing to the device failed, {err}");
            break;
        }
    }
    audio_client.stop_stream().ok();
}
