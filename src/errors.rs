#[derive(Debug, thiserror::Error)]
pub enum WasapiError {
    #[error("Unable to find device with name: {0}")]
    DeviceNotFound(String),
    #[error("Got an illegal device state: {0}")]
    IllegalDeviceState(u32),
    #[error("Got an illegal device role: {0}")]
    IllegalDeviceRole(i32),
    #[error("Got an illegal device direction: {0}")]
    IllegalDeviceDirection(i32),
    #[error("Got an illegal session state: {0}")]
    IllegalSessionState(i32),
    #[error("Could not find a compatible format")]
    UnsupportedFormat,
    #[error("Got an unknown Subformat: {0:?}")]
    UnsupportedSubformat(windows_core::GUID),
    #[error("Client has not been initialized")]
    ClientNotInit,
    #[error("Couldn't register session notifications: {0}")]
    RegisterNotifications(windows_core::Error),
    #[error("Wrong length of data, got {received}, expected exactly {expected}")]
    DataLengthMismatch { received: usize, expected: usize },
    #[error("Wrong length of data, got {received}, expected at least {expected}")]
    DataLengthTooShort { received: usize, expected: usize },
    #[error("Handle wait timed out")]
    EventTimeout,
    #[error("Cant use automatic format conversion in exclusive mode")]
    AutomaticFormatConversionInExclusiveMode,
    #[error("Cant use Loopback with exclusive mode")]
    LoopbackWithExclusiveMode,
    #[error("Cant render to a capture device")]
    RenderToCaptureDevice,
    #[error("Windows returned an error: {0}")]
    Windows(#[from] windows_core::Error),
}
