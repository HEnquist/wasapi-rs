use std::{collections::VecDeque, sync::{Arc, Condvar, Mutex}};

use wasapi::*;
use sysinfo::{ProcessRefreshKind, RefreshKind, System};

#[macro_use]
extern crate log;
use simplelog::*;

// Handler for completion of audio client activation
struct Handler { condvar: Arc<(Mutex<bool>, Condvar)> }

impl Handler {
    fn new(condvar: Arc<(Mutex<bool>, Condvar)>) -> Self {
        Self { condvar }
    }
}

impl ActivateAudioInterfaceCompletionHandlerImpl for Handler {
    fn activate_completed(&self) {
        debug!("Audio client initialized");
        let (lock, cvar) = &*self.condvar;
        let mut started = lock.lock().unwrap();
        *started = true;
        cvar.notify_one();
    }
}

fn main() {
    let _ = SimpleLogger::init(
        LevelFilter::Trace,
        ConfigBuilder::new()
            .set_time_format_rfc3339()
            .set_time_offset_to_local()
            .unwrap()
            .build(),
    );

    initialize_mta().unwrap();

    // Get process ID, in this case Spotify
    let refreshes = RefreshKind::new().with_processes(ProcessRefreshKind::everything());
    let system = System::new_with_specifics(refreshes);
    let process_ids = system.processes_by_name("Spotify.exe");
    let mut process_id = 0;
    for process in process_ids {
        if system.process(process.parent().unwrap()).unwrap().name() == "Spotify.exe" {
            // Note: When capturing audio windows allows you to capture an app's entire process tree, however you must ensure you get the parent process ID
            process_id = process.parent().unwrap().as_u32();
        }
    }

    // Create completion handler
    let condvar = Arc::new((Mutex::new(false), Condvar::new()));
    let handler = ActivateAudioInterfaceCompletionHandler::new(Box::new(Handler::new(condvar.clone())));

    // Activate audio client
    let operation = AudioClient::create_application_loopback_client(process_id, true, handler).unwrap();

    // Wait for activation to complete
    let (lock, cvar) = &*condvar;
    let mut started = lock.lock().unwrap();
    while !*started {
        started = cvar.wait(started).unwrap();
    }

    // Initialize capture client client
    let mut audio_client = operation.get_audio_client().unwrap();
    let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2, None);

    let blockalign = desired_format.get_blockalign();
    debug!("Desired capture format: {:?}", desired_format);

    // Audio client must be initialised in shared and capture mode, the period is irrelevant
    audio_client.initialize_client(
        &desired_format,
        0,
        &Direction::Capture,
        &ShareMode::Shared,
        false
    ).unwrap();
    debug!("initialized capture");

    let h_event = audio_client.set_get_eventhandle().unwrap();
    let render_client = audio_client.get_audiocaptureclient().unwrap();

    audio_client.start_stream().unwrap();
    let mut available_frames = match render_client.get_next_nbr_frames().unwrap() {
        Some(n) => n,
        None => 0,
    };
    let mut bytes_to_capture = available_frames * blockalign;
    let mut sample_queue = VecDeque::with_capacity(bytes_to_capture as usize);

    loop {
        while available_frames > 0 {
            render_client.read_from_device_to_deque(blockalign as usize, &mut sample_queue).unwrap();
    
            // Do something with the samples
            let mut samples = Vec::with_capacity(bytes_to_capture as usize);
            while sample_queue.len() > 0 {
                samples.push(sample_queue.pop_front().unwrap());
            }
    
            if h_event.wait_for_event(1000000).is_err() {
                error!("error, stopping capture");
                audio_client.stop_stream().unwrap();
                break;
            }
        }

        available_frames = match render_client.get_next_nbr_frames().unwrap() {
            Some(n) => n,
            None => 0,
        };
        bytes_to_capture = available_frames * blockalign;
        sample_queue.reserve_exact(bytes_to_capture as usize);
    }
}