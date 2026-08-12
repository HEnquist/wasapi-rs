use wasapi::*;

fn main() {
    initialize_mta().unwrap();

    let enumerator = DeviceEnumerator::new().unwrap();

    for direction in [Direction::Capture, Direction::Render] {
        println!("The following {direction} devices are being used by:");
        for device in &enumerator.get_device_collection(&direction).unwrap() {
            let dev = device.unwrap();
            let manager = dev.get_iaudiosessionmanager().unwrap();
            let sessions = manager.get_audiosessionenumerator().unwrap();

            println!("Device: {:?}", dev.get_friendlyname().unwrap());

            for i in 0..sessions.get_count().unwrap() {
                let control = sessions.get_session(i).unwrap();
                let state = control.get_state().unwrap();
                if state != SessionState::Active {
                    continue;
                }
                let process_id = control.get_process_id().unwrap();
                let identifier = control.get_session_identifier().unwrap();
                let meter = control.get_audiometerinformation().unwrap();
                let peak = meter.get_peak_value().unwrap();

                println!(" - In use by process: {process_id}, peak: {peak:.3}");
                println!("   session: {identifier}");
            }
        }
    }
}
