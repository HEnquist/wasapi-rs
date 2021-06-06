fn main() {
    windows::build!(
        Windows::Win32::Media::Audio::CoreAudio::{
            AudioSessionState, AudioSessionStateActive,AudioSessionStateInactive,AudioSessionStateExpired, IAudioSessionEvents, AudioSessionDisconnectReason,
            DisconnectReasonDeviceRemoval, DisconnectReasonServerShutdown, DisconnectReasonFormatChanged, DisconnectReasonSessionLogoff, DisconnectReasonSessionDisconnected, DisconnectReasonExclusiveModeOverride,
            eConsole, eRender, eCapture, IAudioClient, IAudioSessionControl, IAudioRenderClient, IAudioCaptureClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, IMMDeviceCollection,
            AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_LOOPBACK, DEVICE_STATE_ACTIVE, WAVE_FORMAT_EXTENSIBLE,
            AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM, AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY,
        },
        Windows::Win32::Devices::FunctionDiscovery::IFunctionInstance,
        Windows::Win32::Media::Multimedia::{
            WAVEFORMATEX,
            WAVEFORMATEXTENSIBLE,
            WAVE_FORMAT_PCM,
            WAVE_FORMAT_IEEE_FLOAT,
            KSDATAFORMAT_SUBTYPE_PCM,
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
        },
        Windows::Win32::Media::Audio::DirectMusic::IPropertyStore,
        Windows::Win32::System::Com::CLSCTX_ALL,
        Windows::Win32::System::Threading::{
            CreateEventA,
            WAIT_OBJECT_0,
            WaitForSingleObject,
        },
        Windows::Win32::System::SystemServices::{
            HANDLE, S_OK, S_FALSE, E_NOINTERFACE, BOOL,
        },
        Windows::Win32::System::PropertiesSystem::PROPERTYKEY,
        Windows::Win32::System::SystemServices::PWSTR,
        Windows::Win32::Storage::StructuredStorage::{STGM_READ, PROPVARIANT},
        Windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc,
    );
}
