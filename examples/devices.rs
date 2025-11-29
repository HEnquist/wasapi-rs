use wasapi::*;

fn main() {
    initialize_mta().unwrap();

    let enumerator = DeviceEnumerator::new().unwrap();

    println!("Found the following output devices:");
    for device in &enumerator
        .get_device_collection(&Direction::Render)
        .unwrap()
    {
        let dev = device.unwrap();
        let state = &dev.get_state().unwrap();
        println!(
            "Device: {:?}. State: {:?}",
            &dev.get_friendlyname().unwrap(),
            state
        );
    }

    println!("Default output devices:");
    [Role::Console, Role::Multimedia, Role::Communications]
        .iter()
        .for_each(|role| {
            println!(
                "{:?}: {:?}",
                role,
                enumerator
                    .get_default_device_for_role(&Direction::Render, role)
                    .unwrap()
                    .get_friendlyname()
                    .unwrap()
            );
        });
}
