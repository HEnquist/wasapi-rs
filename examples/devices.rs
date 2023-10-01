use wasapi::*;

fn main() {
    initialize_mta().unwrap();

    println!("Found the following output devices:");
    for device in &DeviceCollection::new(&Direction::Render).unwrap() {
        let state = &device.unwrap().get_state_enum().unwrap();
        println!("Device: {:?}. State: {:?}", device.unwrap().get_friendlyname().unwrap(), state);
    }

    println!("Default output devices:");
    [Role::Console, Role::Multimedia, Role::Communications]
        .iter()
        .for_each(|role| {
            println!(
                "{:?}: {:?}",
                role,
                get_default_device_for_role(&Direction::Render, role)
                    .unwrap()
                    .get_friendlyname()
                    .unwrap()
            );
        });
}
