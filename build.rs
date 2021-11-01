fn main() {
    windows_macros::build!(
        Windows::Win32::Media::Audio::CoreAudio::{
            AudioSessionState, IAudioSessionEvents, AudioSessionDisconnectReason,
            IAudioClient, IAudioSessionControl, IAudioRenderClient, IAudioCaptureClient,
            IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, IMMDeviceCollection,
            AUDCLNT_SHAREMODE,
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
            AUDCLNT_STREAMFLAGS_LOOPBACK, AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY,
            DEVICE_STATE_ACTIVE,
            WAVE_FORMAT_EXTENSIBLE,
        },
        Windows::Win32::Media::Audio::CoreAudio::{ERole, EDataFlow},
        Windows::Win32::Devices::FunctionDiscovery::IFunctionInstance,
        Windows::Win32::Media::Multimedia::{
            WAVEFORMATEX,
            WAVEFORMATEXTENSIBLE,
            WAVE_FORMAT_PCM,
            WAVE_FORMAT_IEEE_FLOAT,
            KSDATAFORMAT_SUBTYPE_PCM,
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
        },
        Windows::Win32::System::PropertiesSystem::IPropertyStore,
        Windows::Win32::System::Com::CLSCTX,
        Windows::Win32::System::Threading::{
            CreateEventA,
            WaitForSingleObject,
            WAIT_OBJECT_0,
        },
        Windows::Win32::Foundation::{BOOL, E_NOINTERFACE, HANDLE, PSTR, PWSTR, S_OK},
        Windows::Win32::System::PropertiesSystem::PROPERTYKEY,
        Windows::Win32::System::SystemServices::{DEVPKEY_Device_DeviceDesc, DEVPKEY_Device_FriendlyName},
        Windows::Win32::System::Com::StructuredStorage::STGM_READ,
        Windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc,
        Windows::Win32::System::Com::CoCreateInstance,
        Windows::Win32::System::Com::CoInitializeEx,
    );
}
