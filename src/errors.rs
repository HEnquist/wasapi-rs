use crate::bindings::typed_consts;

pub use crate::bindings::{E_INVALIDARG, E_NOINTERFACE, E_POINTER};

// The audio client result codes as HRESULT, for matching against the code of a
// WasapiError::Windows. The metadata gives them as `i32`.
typed_consts!(
    HRESULT,
    concat!(
        "An audio client result code, see the ",
        "[audio client error codes]",
        "(https://learn.microsoft.com/en-us/windows/win32/coreaudio/audclnt-xxx-error-codes)."
    );
    AUDCLNT_E_ALREADY_INITIALIZED,
    AUDCLNT_E_BUFDURATION_PERIOD_NOT_EQUAL,
    AUDCLNT_E_BUFFER_ERROR,
    AUDCLNT_E_BUFFER_OPERATION_PENDING,
    AUDCLNT_E_BUFFER_SIZE_ERROR,
    AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED,
    AUDCLNT_E_BUFFER_TOO_LARGE,
    AUDCLNT_E_CPUUSAGE_EXCEEDED,
    AUDCLNT_E_DEVICE_INVALIDATED,
    AUDCLNT_E_DEVICE_IN_USE,
    AUDCLNT_E_EFFECT_NOT_AVAILABLE,
    AUDCLNT_E_EFFECT_STATE_READ_ONLY,
    AUDCLNT_E_ENDPOINT_CREATE_FAILED,
    AUDCLNT_E_ENDPOINT_OFFLOAD_NOT_CAPABLE,
    AUDCLNT_E_ENGINE_FORMAT_LOCKED,
    AUDCLNT_E_ENGINE_PERIODICITY_LOCKED,
    AUDCLNT_E_EVENTHANDLE_NOT_EXPECTED,
    AUDCLNT_E_EVENTHANDLE_NOT_SET,
    AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED,
    AUDCLNT_E_EXCLUSIVE_MODE_ONLY,
    AUDCLNT_E_HEADTRACKING_ENABLED,
    AUDCLNT_E_HEADTRACKING_UNSUPPORTED,
    AUDCLNT_E_INCORRECT_BUFFER_SIZE,
    AUDCLNT_E_INVALID_DEVICE_PERIOD,
    AUDCLNT_E_INVALID_SIZE,
    AUDCLNT_E_INVALID_STREAM_FLAG,
    AUDCLNT_E_NONOFFLOAD_MODE_ONLY,
    AUDCLNT_E_NOT_INITIALIZED,
    AUDCLNT_E_NOT_STOPPED,
    AUDCLNT_E_OFFLOAD_MODE_ONLY,
    AUDCLNT_E_OUT_OF_OFFLOAD_RESOURCES,
    AUDCLNT_E_OUT_OF_ORDER,
    AUDCLNT_E_RAW_MODE_UNSUPPORTED,
    AUDCLNT_E_RESOURCES_INVALIDATED,
    AUDCLNT_E_SERVICE_NOT_RUNNING,
    AUDCLNT_E_THREAD_NOT_REGISTERED,
    AUDCLNT_E_UNSUPPORTED_FORMAT,
    AUDCLNT_E_WRONG_ENDPOINT_TYPE,
    AUDCLNT_S_BUFFER_EMPTY,
    AUDCLNT_S_POSITION_STALLED,
    AUDCLNT_S_THREAD_ALREADY_REGISTERED,
);

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
    #[error("Couldn't register notifications: {0}")]
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
    #[error("String conversion error: {0}")]
    FromUtf16Error(#[from] std::string::FromUtf16Error),
}
