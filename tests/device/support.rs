// Shared helpers for the device backed tests.
//
// The suite runs on VB-Cable, see .github/workflows/device-tests.yml for how the
// runner gets one. Anywhere the cable is missing, every test prints a note and
// returns, so running the suite on a machine without it is harmless. Setting
// WASAPI_REQUIRE_DEVICE turns those skips into panics, which is what keeps the
// CI job from passing with no coverage at all.
//
// Keep this module small. The tests assert on what the library returns, so there
// is no signal generation and no analysis here: a render run writes zeros and a
// capture run throws the samples away. What gets asserted is the frame count,
// which is enough to prove the stream ran at the format's rate.

use simplelog::{Config, LevelFilter, SimpleLogger};
use std::cell::Cell;
use std::sync::{Mutex, MutexGuard, OnceLock};
use std::time::{Duration, Instant};
use wasapi::*;

/// The interface friendly name of every VB-Cable endpoint.
///
/// This is the name of the adapter rather than of the endpoint, and it is the same
/// for the vendor installer and for the driver only install used in CI, where the
/// endpoint friendly name differs. It also tells the cable apart from the other
/// VB-Audio products and from the unrelated "Virtual Audio Cable".
pub const CABLE_INTERFACE_NAME: &str = "VB-Audio Virtual Cable";

/// How long a test stream runs. Long enough that the prefilled buffer is a small
/// part of the frames counted, short enough that the whole suite stays quick.
pub const RUN_TIME: Duration = Duration::from_millis(500);

/// Sleep between iterations of a polled loop, and the cap on a single wait.
const POLL_SLEEP: Duration = Duration::from_millis(5);
const EVENT_TIMEOUT_MS: u32 = 2000;

/// Guards against a runaway loop if a device stops reporting progress.
const MAX_ITERATIONS: u32 = 10_000;

/// Cap on the read buffer, in frames, generous at a second of 192 kHz audio.
///
/// A process loopback client reports a nonsense buffer size, documented as values
/// like 3131961357, so sizing the buffer from it would ask for tens of gigabytes.
/// A packet larger than this cap surfaces as a `DataLengthTooShort` from the library
/// rather than as an allocation failure.
const MAX_READ_FRAMES: u32 = 192_000;

static DEVICE_LOCK: OnceLock<Mutex<()>> = OnceLock::new();

thread_local! {
    static COM_READY: Cell<bool> = const { Cell::new(false) };
}

/// The endpoints of the cable.
pub struct Cable {
    /// The two channel render endpoint.
    pub render: Device,
    /// "CABLE Output".
    pub capture: Device,
}

/// What a device test runs against. Holds the device lock for its lifetime.
pub struct Fixture {
    pub enumerator: DeviceEnumerator,
    pub cable: Cable,
    // Last on purpose. Fields drop in declaration order, so this releases the lock
    // only after the devices above are gone, which is what frees the endpoint. Any
    // earlier and the next test could take the lock while this one still holds a
    // device open.
    _lock: MutexGuard<'static, ()>,
}

/// The cable, or `None` when the test should skip.
///
/// Panics instead of returning `None` when `WASAPI_REQUIRE_DEVICE` is set.
pub fn fixture() -> Option<Fixture> {
    // Serialise first. An exclusive mode client owns the endpoint, so two tests
    // touching it at once would fail with AUDCLNT_E_DEVICE_IN_USE and look like a
    // library bug. CI passes --test-threads=1 as well, but this keeps a plain
    // `cargo test --features device-tests` from flaking on a dev machine.
    // Poisoning is ignored on purpose: one panicking test must not fail the rest.
    let lock = DEVICE_LOCK
        .get_or_init(|| Mutex::new(()))
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    init_log();
    init_com();
    let enumerator = DeviceEnumerator::new().expect("failed to create a DeviceEnumerator");
    match find_cable(&enumerator) {
        Some(cable) => Some(Fixture {
            enumerator,
            cable,
            _lock: lock,
        }),
        None => {
            skip("no VB-Cable render and capture endpoint found");
            None
        }
    }
}

/// Initialise COM as MTA for the calling thread, at most once.
///
/// Never calls [deinitialize]. Under `--test-threads=1` every test shares one
/// thread, so an unbalanced `CoUninitialize` would tear COM down under the next
/// test, and uninitialising while interface pointers are still alive is undefined
/// behaviour. Nothing in the suite calls `deinitialize`, so the process keeps COM
/// initialised until it exits, which is what the runtime does anyway.
pub fn init_com() {
    COM_READY.with(|ready| {
        if !ready.get() {
            initialize_mta()
                .ok()
                .expect("failed to initialize COM as MTA");
            ready.set(true);
        }
    });
}

/// Turn on the library's own logging at the level in `WASAPI_TEST_LOG`, once.
///
/// Off unless the variable is set. The device CI job passes it through from a
/// workflow input, so a run that failed can be repeated with debug or trace logging
/// without editing anything.
pub fn init_log() {
    static LOGGER: OnceLock<()> = OnceLock::new();
    LOGGER.get_or_init(|| {
        let level = match std::env::var("WASAPI_TEST_LOG")
            .unwrap_or_default()
            .to_ascii_lowercase()
            .as_str()
        {
            "error" => LevelFilter::Error,
            "warn" => LevelFilter::Warn,
            "info" => LevelFilter::Info,
            "debug" => LevelFilter::Debug,
            "trace" => LevelFilter::Trace,
            _ => LevelFilter::Off,
        };
        if level != LevelFilter::Off {
            // Ignore a second init: another test may have got there first.
            let _ = SimpleLogger::init(level, Config::default());
        }
    });
}

/// Report that a test is being skipped, or fail when a device was required.
pub fn skip(reason: &str) {
    let name = std::thread::current().name().unwrap_or("test").to_owned();
    let message = format!("SKIPPED {name}: {reason}");
    assert!(!require_device(), "{message}");
    println!("{message}");
}

/// Whether a missing device should fail the test rather than skip it.
pub fn require_device() -> bool {
    std::env::var_os("WASAPI_REQUIRE_DEVICE").is_some()
}

/// Every endpoint of the given direction that belongs to the cable.
pub fn cable_endpoints(enumerator: &DeviceEnumerator, direction: &Direction) -> Vec<Device> {
    let Ok(collection) = enumerator.get_device_collection(direction) else {
        return Vec::new();
    };
    (&collection)
        .into_iter()
        .flatten()
        .filter(|device| {
            device
                .get_interface_friendlyname()
                .is_ok_and(|name| name == CABLE_INTERFACE_NAME)
        })
        .collect()
}

fn find_cable(enumerator: &DeviceEnumerator) -> Option<Cable> {
    // The cable has two render endpoints and the tests want the plain one, so the 16
    // channel endpoint has to be told apart and discarded. The device format is no
    // help: it reports whatever the control panel last stored, which is two channels
    // for both. The description is "CABLE In 16ch" from the vendor installer and
    // "CABLE In 16 Ch" on the runner, so strip the spaces.
    let is_multichannel = |device: &Device| {
        device
            .get_description()
            .is_ok_and(|desc| desc.to_lowercase().replace(' ', "").contains("16ch"))
    };
    let (multichannel, stereo): (Vec<Device>, Vec<Device>) =
        cable_endpoints(enumerator, &Direction::Render)
            .into_iter()
            .partition(is_multichannel);
    drop(multichannel);
    Some(Cable {
        render: stereo.into_iter().next()?,
        capture: cable_endpoints(enumerator, &Direction::Capture)
            .into_iter()
            .next()?,
    })
}

/// Whether an error is exactly the given Windows result code.
///
/// Lets a test name the `AUDCLNT_E_*` constant it expects instead of matching on a
/// message. False for the crate's own error variants, which carry no result code.
pub fn is_hresult(error: &WasapiError, expected: windows::core::HRESULT) -> bool {
    matches!(error, WasapiError::Windows(e) if e.code() == expected)
}

/// One line describing a device, for the inventory and for failure messages.
pub fn describe(device: &Device) -> String {
    let name = |result: Result<String, WasapiError>| result.unwrap_or_else(|e| format!("<{e}>"));
    format!(
        "{:?} \"{}\" desc \"{}\" iface \"{}\" id {}",
        device.get_direction(),
        name(device.get_friendlyname()),
        name(device.get_description()),
        name(device.get_interface_friendlyname()),
        name(device.get_id()),
    )
}

/// What a short stream reported about itself.
pub struct Run {
    /// Frames written or read after the stream was started.
    pub frames: u64,
    /// Measured from just before `start_stream`, so it matches `frames`.
    pub elapsed: Duration,
    pub buffer_size: u32,
    /// `(frames, BufferInfo)` of every packet read. Empty for a render run.
    pub packets: Vec<(u32, BufferInfo)>,
    /// Set when a wait for the event handle timed out, so the test can report it
    /// rather than the helper panicking somewhere out of sight.
    pub event_timeout: bool,
}

impl Run {
    fn new(buffer_size: u32) -> Self {
        Run {
            frames: 0,
            elapsed: Duration::ZERO,
            buffer_size,
            packets: Vec::new(),
            event_timeout: false,
        }
    }

    /// Frames per second, to compare against the format's sample rate.
    pub fn rate(&self) -> f64 {
        self.frames as f64 / self.elapsed.as_secs_f64()
    }
}

/// An initialised client, and the format it was initialised with.
pub struct Opened {
    pub client: AudioClient,
    pub format: WaveFormat,
}

/// Open and initialise a client on a device.
///
/// Shared mode uses the mix format, which WASAPI guarantees. Exclusive mode asks for
/// S16 at 48 kHz in stereo, which is what VB-Cable offers, through
/// `is_supported_exclusive_with_quirks` so the channel mask is one the device really
/// accepted.
///
/// The exclusive period uses 128 byte alignment, as the exclusive mode examples do,
/// since some devices insist on it. The realign dance those examples perform on
/// `AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED` is deliberately not repeated: no device in CI
/// exercises it, so it would be test code that is never itself tested. A device that
/// wants it fails here with that error, which says so clearly enough.
///
/// Note that an event driven client still needs `set_get_eventhandle` before
/// `start_stream`. [render_for] and [capture_for] do that themselves, but a test that
/// drives the stream by hand has to, or `start_stream` gives
/// `AUDCLNT_E_EVENTHANDLE_NOT_SET`.
pub fn open(
    device: &Device,
    direction: &Direction,
    share: ShareMode,
    timing: TimingMode,
) -> Result<Opened, WasapiError> {
    let mut client = device.get_iaudioclient()?;
    let (def_period, _min_period) = client.get_device_period()?;
    let format = match share {
        ShareMode::Shared => client.get_mixformat()?,
        ShareMode::Exclusive => client.is_supported_exclusive_with_quirks(&WaveFormat::new(
            16,
            16,
            &SampleType::Int,
            48000,
            2,
            None,
        ))?,
    };
    let mode = mode_for(&client, share, timing, def_period, &format)?;
    client.initialize_client(&format, direction, &mode)?;
    Ok(Opened { client, format })
}

/// The [StreamMode] for a sharing and timing mode, with a period the device accepts.
pub fn mode_for(
    client: &AudioClient,
    share: ShareMode,
    timing: TimingMode,
    def_period: i64,
    format: &WaveFormat,
) -> Result<StreamMode, WasapiError> {
    Ok(match (share, timing) {
        (ShareMode::Shared, TimingMode::Events) => StreamMode::EventsShared {
            autoconvert: false,
            buffer_duration_hns: def_period,
        },
        (ShareMode::Shared, TimingMode::Polling) => StreamMode::PollingShared {
            autoconvert: false,
            buffer_duration_hns: def_period,
        },
        (ShareMode::Exclusive, timing) => {
            let period = client.calculate_aligned_period_near(def_period, Some(128), format)?;
            match timing {
                TimingMode::Events => StreamMode::EventsExclusive { period_hns: period },
                TimingMode::Polling => StreamMode::PollingExclusive {
                    period_hns: period,
                    buffer_duration_hns: 16 * period,
                },
            }
        }
    })
}

/// As [open], but prints a skip note and returns `None` when the device cannot do it.
pub fn open_or_skip(
    device: &Device,
    direction: &Direction,
    share: ShareMode,
    timing: TimingMode,
) -> Option<Opened> {
    match open(device, direction, share, timing) {
        Ok(opened) => Some(opened),
        Err(e) => {
            skip(&format!("cannot open {share} {timing:?}: {e}"));
            None
        }
    }
}

/// Wait for the next period, by event or by sleeping. Returns false on timeout.
fn wait(event: Option<&Handle>) -> bool {
    match event {
        Some(handle) => handle.wait_for_event(EVENT_TIMEOUT_MS).is_ok(),
        None => {
            std::thread::sleep(POLL_SLEEP);
            true
        }
    }
}

/// An event handle when the client was initialised for event driven timing.
fn event_handle(client: &AudioClient) -> Result<Option<Handle>, WasapiError> {
    match client.get_timing_mode() {
        Some(TimingMode::Events) => Ok(Some(client.set_get_eventhandle()?)),
        _ => Ok(None),
    }
}

/// Prefill with silence, start, write silence until `dur` has passed, stop.
///
/// The prefill is not counted, and the clock starts at `start_stream`, so
/// [Run::rate] is the rate at which the device asked for data.
pub fn render_for(
    client: &AudioClient,
    fmt: &WaveFormat,
    dur: Duration,
) -> Result<Run, WasapiError> {
    let blockalign = fmt.get_blockalign() as usize;
    let buffer_size = client.get_buffer_size()?;
    let render_client = client.get_audiorenderclient()?;
    let event = event_handle(client)?;
    let silence = vec![0u8; buffer_size as usize * blockalign];

    // Fill the buffer before starting the stream, so playback starts with data.
    // https://learn.microsoft.com/en-us/windows/win32/coreaudio/rendering-a-stream
    let prefill = client.get_available_space_in_frames()? as usize;
    render_client.write_to_device(prefill, &silence[..prefill * blockalign], None)?;

    let mut run = Run::new(buffer_size);
    let start = Instant::now();
    client.start_stream()?;
    for _ in 0..MAX_ITERATIONS {
        if start.elapsed() >= dur {
            break;
        }
        if !wait(event.as_ref()) {
            run.event_timeout = true;
            break;
        }
        let frames = client.get_available_space_in_frames()? as usize;
        render_client.write_to_device(frames, &silence[..frames * blockalign], None)?;
        run.frames += frames as u64;
    }
    run.elapsed = start.elapsed();
    client.stop_stream()?;
    Ok(run)
}

/// Start, read until `dur` has passed, stop. The samples are thrown away.
pub fn capture_for(
    client: &AudioClient,
    fmt: &WaveFormat,
    dur: Duration,
) -> Result<Run, WasapiError> {
    let blockalign = fmt.get_blockalign() as usize;
    let buffer_size = client.get_buffer_size()?;
    let capture_client = client.get_audiocaptureclient()?;
    let event = event_handle(client)?;
    // A packet never exceeds the buffer, and the slack keeps a surprise from turning
    // into a DataLengthTooShort that would look like a library bug. Capped, because
    // a process loopback client reports a nonsense buffer size.
    let read_frames = 2 * buffer_size.min(MAX_READ_FRAMES) as usize;
    let mut buffer = vec![0u8; read_frames * blockalign];

    let mut run = Run::new(buffer_size);
    let start = Instant::now();
    client.start_stream()?;
    for _ in 0..MAX_ITERATIONS {
        if start.elapsed() >= dur {
            break;
        }
        if !wait(event.as_ref()) {
            run.event_timeout = true;
            break;
        }
        // Drain every packet that is ready, not just the first one.
        for _ in 0..MAX_ITERATIONS {
            let (frames, info) = capture_client.read_from_device(&mut buffer)?;
            if frames == 0 {
                break;
            }
            run.frames += u64::from(frames);
            run.packets.push((frames, info));
        }
    }
    run.elapsed = start.elapsed();
    client.stop_stream()?;
    Ok(run)
}

/// How far the measured frame rate may be from the format's sample rate.
const RATE_TOLERANCE: f64 = 0.20;

/// Assert that a run moved frames at the format's sample rate.
pub fn assert_rate(run: &Run, format: &WaveFormat, what: &str) {
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
