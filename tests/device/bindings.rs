#[inline]
pub unsafe fn ActivateAudioInterfaceAsync<P0, P3>(
    deviceinterfacepath: P0,
    riid: *const IID,
    activationparams: Option<*const PROPVARIANT>,
    completionhandler: P3,
) -> windows_core::Result<IActivateAudioInterfaceAsyncOperation>
where
    P0: windows_core::Param<windows_core::PCWSTR>,
    P3: windows_core::Param<IActivateAudioInterfaceCompletionHandler>,
{
    windows_core::link!("mmdevapi.dll" "system" fn ActivateAudioInterfaceAsync(deviceinterfacepath : windows_core::PCWSTR, riid : *const IID, activationparams : *const PROPVARIANT, completionhandler : *mut core::ffi::c_void, activationoperation : *mut *mut core::ffi::c_void) -> windows_core::HRESULT);
    unsafe {
        let mut result__ = core::mem::zeroed();
        ActivateAudioInterfaceAsync(
            deviceinterfacepath.param().abi(),
            riid,
            activationparams.unwrap_or(core::mem::zeroed()) as _,
            completionhandler.param().abi(),
            &mut result__,
        )
        .and_then(|| windows_core::imp::from_abi(result__))
    }
}
#[inline]
pub unsafe fn CloseHandle(hobject: HANDLE) -> windows_core::BOOL {
    windows_core::link!("kernel32.dll" "system" fn CloseHandle(hobject : HANDLE) -> windows_core::BOOL);
    unsafe { CloseHandle(hobject) }
}
#[inline]
pub unsafe fn CreateEventW<P3>(
    lpeventattributes: Option<LPSECURITY_ATTRIBUTES>,
    bmanualreset: bool,
    binitialstate: bool,
    lpname: P3,
) -> HANDLE
where
    P3: windows_core::Param<windows_core::PCWSTR>,
{
    windows_core::link!("kernel32.dll" "system" fn CreateEventW(lpeventattributes : LPSECURITY_ATTRIBUTES, bmanualreset : windows_core::BOOL, binitialstate : windows_core::BOOL, lpname : windows_core::PCWSTR) -> HANDLE);
    unsafe {
        CreateEventW(
            lpeventattributes.unwrap_or(core::mem::zeroed()) as _,
            bmanualreset.into(),
            binitialstate.into(),
            lpname.param().abi(),
        )
    }
}
#[inline]
pub unsafe fn SetEvent(hevent: HANDLE) -> windows_core::BOOL {
    windows_core::link!("kernel32.dll" "system" fn SetEvent(hevent : HANDLE) -> windows_core::BOOL);
    unsafe { SetEvent(hevent) }
}
#[inline]
pub unsafe fn WaitForMultipleObjects(
    ncount: DWORD,
    lphandles: *const HANDLE,
    bwaitall: bool,
    dwmilliseconds: DWORD,
) -> DWORD {
    windows_core::link!("kernel32.dll" "system" fn WaitForMultipleObjects(ncount : DWORD, lphandles : *const HANDLE, bwaitall : windows_core::BOOL, dwmilliseconds : DWORD) -> DWORD);
    unsafe { WaitForMultipleObjects(ncount, lphandles, bwaitall.into(), dwmilliseconds) }
}
pub type AUDCLNT_SHAREMODE = i32;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct BLOB {
    pub cbSize: ULONG,
    pub pBlobData: *mut BYTE,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct BSTRBLOB {
    pub cbSize: ULONG,
    pub pData: *mut BYTE,
}
pub type BYTE = u8;
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CABOOL {
    pub cElems: ULONG,
    pub pElems: *mut VARIANT_BOOL,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CABSTR {
    pub cElems: ULONG,
    pub pElems: *mut windows_core::BSTR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CABSTRBLOB {
    pub cElems: ULONG,
    pub pElems: *mut BSTRBLOB,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAC {
    pub cElems: ULONG,
    pub pElems: *mut CHAR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CACLIPDATA {
    pub cElems: ULONG,
    pub pElems: *mut CLIPDATA,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CACLSID {
    pub cElems: ULONG,
    pub pElems: *mut CLSID,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CACY {
    pub cElems: ULONG,
    pub pElems: *mut CY,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CADATE {
    pub cElems: ULONG,
    pub pElems: *mut DATE,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CADBL {
    pub cElems: ULONG,
    pub pElems: *mut DOUBLE,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAFILETIME {
    pub cElems: ULONG,
    pub pElems: *mut FILETIME,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAFLT {
    pub cElems: ULONG,
    pub pElems: *mut FLOAT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAH {
    pub cElems: ULONG,
    pub pElems: *mut LARGE_INTEGER,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAI {
    pub cElems: ULONG,
    pub pElems: *mut SHORT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAL {
    pub cElems: ULONG,
    pub pElems: *mut LONG,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CALPSTR {
    pub cElems: ULONG,
    pub pElems: *mut windows_core::PSTR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CALPWSTR {
    pub cElems: ULONG,
    pub pElems: *mut windows_core::PWSTR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAPROPVARIANT {
    pub cElems: ULONG,
    pub pElems: *mut PROPVARIANT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CASCODE {
    pub cElems: ULONG,
    pub pElems: *mut SCODE,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUB {
    pub cElems: ULONG,
    pub pElems: *mut UCHAR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUH {
    pub cElems: ULONG,
    pub pElems: *mut ULARGE_INTEGER,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUI {
    pub cElems: ULONG,
    pub pElems: *mut USHORT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUL {
    pub cElems: ULONG,
    pub pElems: *mut ULONG,
}
pub type CHAR = i8;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CLIPDATA {
    pub cbSize: ULONG,
    pub ulClipFmt: LONG,
    pub pClipData: *mut BYTE,
}
pub type CLSCTX = i32;
pub const CLSCTX_ALL: CLSCTX = 23;
pub type CLSID = windows_core::GUID;
#[repr(C)]
#[derive(Clone, Copy)]
pub union CY {
    pub Anonymous: CY_0,
    pub int64: LONGLONG,
}
impl Default for CY {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CY_0 {
    pub Lo: ULONG,
    pub Hi: LONG,
}
pub type DATE = f64;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct DECIMAL {
    pub wReserved: USHORT,
    pub Anonymous: DECIMAL_0,
    pub Hi32: ULONG,
    pub Anonymous2: DECIMAL_1,
}
impl Default for DECIMAL {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union DECIMAL_0 {
    pub Anonymous: DECIMAL_0_0,
    pub signscale: USHORT,
}
impl Default for DECIMAL_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DECIMAL_0_0 {
    pub scale: BYTE,
    pub sign: BYTE,
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union DECIMAL_1 {
    pub Anonymous: DECIMAL_1_0,
    pub Lo64: ULONGLONG,
}
impl Default for DECIMAL_1 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DECIMAL_1_0 {
    pub Lo32: ULONG,
    pub Mid32: ULONG,
}
pub const DEVINTERFACE_AUDIO_CAPTURE: windows_core::GUID =
    windows_core::GUID::from_u128(0x2eef81be_33fa_4800_9670_1cd474972c3f);
pub const DEVINTERFACE_AUDIO_RENDER: windows_core::GUID =
    windows_core::GUID::from_u128(0xe6327cad_dcec_4949_ae8a_991e976a79d2);
pub type DOUBLE = f64;
pub type DWORD = u32;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct FILETIME {
    pub dwLowDateTime: DWORD,
    pub dwHighDateTime: DWORD,
}
pub type FLOAT = f32;
pub type HANDLE = *mut core::ffi::c_void;
windows_core::imp::define_interface!(
    IActivateAudioInterfaceAsyncOperation,
    IActivateAudioInterfaceAsyncOperation_Vtbl,
    0x72a22d78_cde4_431d_b8cc_843a71199b6d
);
windows_core::imp::interface_hierarchy!(
    IActivateAudioInterfaceAsyncOperation,
    windows_core::IUnknown
);
impl IActivateAudioInterfaceAsyncOperation {
    pub unsafe fn GetActivateResult(
        &self,
        activateresult: *mut windows_core::HRESULT,
        activatedinterface: *mut Option<windows_core::IUnknown>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetActivateResult)(
                windows_core::Interface::as_raw(self),
                activateresult as _,
                activatedinterface as _,
            )
        }
    }
}
#[repr(C)]
pub struct IActivateAudioInterfaceAsyncOperation_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetActivateResult: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::HRESULT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IActivateAudioInterfaceCompletionHandler,
    IActivateAudioInterfaceCompletionHandler_Vtbl,
    0x41d949ab_9862_444a_80f6_c261334da5eb
);
windows_core::imp::interface_hierarchy!(
    IActivateAudioInterfaceCompletionHandler,
    windows_core::IUnknown
);
impl IActivateAudioInterfaceCompletionHandler {
    pub unsafe fn ActivateCompleted<P0>(&self, activateoperation: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IActivateAudioInterfaceAsyncOperation>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).ActivateCompleted)(
                windows_core::Interface::as_raw(self),
                activateoperation.param().abi(),
            )
        }
    }
}
#[repr(C)]
pub struct IActivateAudioInterfaceCompletionHandler_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub ActivateCompleted: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
pub trait IActivateAudioInterfaceCompletionHandler_Impl: windows_core::IUnknownImpl {
    fn ActivateCompleted(
        &self,
        activateoperation: windows_core::Ref<IActivateAudioInterfaceAsyncOperation>,
    ) -> windows_core::HRESULT;
}
impl IActivateAudioInterfaceCompletionHandler_Vtbl {
    pub const fn new<
        Identity: IActivateAudioInterfaceCompletionHandler_Impl,
        const OFFSET: isize,
    >() -> Self {
        unsafe extern "system" fn ActivateCompleted<
            Identity: IActivateAudioInterfaceCompletionHandler_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            activateoperation: *mut core::ffi::c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IActivateAudioInterfaceCompletionHandler_Impl::ActivateCompleted(
                    this,
                    core::mem::transmute_copy(&activateoperation),
                )
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            ActivateCompleted: ActivateCompleted::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IActivateAudioInterfaceCompletionHandler as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IActivateAudioInterfaceCompletionHandler {}
windows_core::imp::define_interface!(
    IAudioClient,
    IAudioClient_Vtbl,
    0x1cb9ad4c_dbfa_4c32_b178_c2f568a703b2
);
windows_core::imp::interface_hierarchy!(IAudioClient, windows_core::IUnknown);
impl IAudioClient {
    pub unsafe fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: DWORD,
        hnsbufferduration: REFERENCE_TIME,
        hnsperiodicity: REFERENCE_TIME,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: Option<LPCGUID>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Initialize)(
                windows_core::Interface::as_raw(self),
                sharemode,
                streamflags,
                hnsbufferduration,
                hnsperiodicity,
                pformat,
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn GetBufferSize(&self, pnumbufferframes: *mut UINT32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetBufferSize)(
                windows_core::Interface::as_raw(self),
                pnumbufferframes as _,
            )
        }
    }
    pub unsafe fn GetStreamLatency(
        &self,
        phnslatency: *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetStreamLatency)(
                windows_core::Interface::as_raw(self),
                phnslatency as _,
            )
        }
    }
    pub unsafe fn GetCurrentPadding(
        &self,
        pnumpaddingframes: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCurrentPadding)(
                windows_core::Interface::as_raw(self),
                pnumpaddingframes as _,
            )
        }
    }
    pub unsafe fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: Option<*mut *mut WAVEFORMATEX>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).IsFormatSupported)(
                windows_core::Interface::as_raw(self),
                sharemode,
                pformat,
                ppclosestmatch.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn GetMixFormat(
        &self,
        ppdeviceformat: *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetMixFormat)(
                windows_core::Interface::as_raw(self),
                ppdeviceformat as _,
            )
        }
    }
    pub unsafe fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: Option<*mut REFERENCE_TIME>,
        phnsminimumdeviceperiod: Option<*mut REFERENCE_TIME>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDevicePeriod)(
                windows_core::Interface::as_raw(self),
                phnsdefaultdeviceperiod.unwrap_or(core::mem::zeroed()) as _,
                phnsminimumdeviceperiod.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn Start(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Start)(windows_core::Interface::as_raw(self))
        }
    }
    pub unsafe fn Stop(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Stop)(windows_core::Interface::as_raw(self))
        }
    }
    pub unsafe fn Reset(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Reset)(windows_core::Interface::as_raw(self))
        }
    }
    pub unsafe fn SetEventHandle(&self, eventhandle: HANDLE) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetEventHandle)(
                windows_core::Interface::as_raw(self),
                eventhandle,
            )
        }
    }
    pub unsafe fn GetService<T>(&self) -> windows_core::Result<T>
    where
        T: windows_core::Interface,
    {
        let mut result__ = core::ptr::null_mut();
        unsafe {
            (windows_core::Interface::vtable(self).GetService)(
                windows_core::Interface::as_raw(self),
                &T::IID,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IAudioClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub Initialize: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AUDCLNT_SHAREMODE,
        DWORD,
        REFERENCE_TIME,
        REFERENCE_TIME,
        *const WAVEFORMATEX,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub GetBufferSize:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT32) -> windows_core::HRESULT,
    pub GetStreamLatency: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT,
    pub GetCurrentPadding:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT32) -> windows_core::HRESULT,
    pub IsFormatSupported: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AUDCLNT_SHAREMODE,
        *const WAVEFORMATEX,
        *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT,
    pub GetMixFormat: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT,
    pub GetDevicePeriod: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut REFERENCE_TIME,
        *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT,
    pub Start: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub Stop: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub Reset: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub SetEventHandle:
        unsafe extern "system" fn(*mut core::ffi::c_void, HANDLE) -> windows_core::HRESULT,
    pub GetService: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const IID,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioClient2,
    IAudioClient2_Vtbl,
    0x726778cd_f60a_4eda_82de_e47610cd78aa
);
impl core::ops::Deref for IAudioClient2 {
    type Target = IAudioClient;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(IAudioClient2, windows_core::IUnknown, IAudioClient);
#[repr(C)]
pub struct IAudioClient2_Vtbl {
    pub base__: IAudioClient_Vtbl,
    IsOffloadCapable: usize,
    SetClientProperties: usize,
    GetBufferSizeLimits: usize,
}
windows_core::imp::define_interface!(
    IAudioClient3,
    IAudioClient3_Vtbl,
    0x7ed4ee07_8e67_4cd4_8c1a_2b7a5987ad42
);
impl core::ops::Deref for IAudioClient3 {
    type Target = IAudioClient2;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(
    IAudioClient3,
    windows_core::IUnknown,
    IAudioClient,
    IAudioClient2
);
impl IAudioClient3 {
    pub unsafe fn GetSharedModeEnginePeriod(
        &self,
        pformat: *const WAVEFORMATEX,
        pdefaultperiodinframes: *mut UINT32,
        pfundamentalperiodinframes: *mut UINT32,
        pminperiodinframes: *mut UINT32,
        pmaxperiodinframes: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetSharedModeEnginePeriod)(
                windows_core::Interface::as_raw(self),
                pformat,
                pdefaultperiodinframes as _,
                pfundamentalperiodinframes as _,
                pminperiodinframes as _,
                pmaxperiodinframes as _,
            )
        }
    }
    pub unsafe fn GetCurrentSharedModeEnginePeriod(
        &self,
        ppformat: *mut *mut WAVEFORMATEX,
        pcurrentperiodinframes: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCurrentSharedModeEnginePeriod)(
                windows_core::Interface::as_raw(self),
                ppformat as _,
                pcurrentperiodinframes as _,
            )
        }
    }
    pub unsafe fn InitializeSharedAudioStream(
        &self,
        streamflags: DWORD,
        periodinframes: UINT32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: Option<LPCGUID>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).InitializeSharedAudioStream)(
                windows_core::Interface::as_raw(self),
                streamflags,
                periodinframes,
                pformat,
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioClient3_Vtbl {
    pub base__: IAudioClient2_Vtbl,
    pub GetSharedModeEnginePeriod: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const WAVEFORMATEX,
        *mut UINT32,
        *mut UINT32,
        *mut UINT32,
        *mut UINT32,
    ) -> windows_core::HRESULT,
    pub GetCurrentSharedModeEnginePeriod: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut WAVEFORMATEX,
        *mut UINT32,
    ) -> windows_core::HRESULT,
    pub InitializeSharedAudioStream: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        DWORD,
        UINT32,
        *const WAVEFORMATEX,
        LPCGUID,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioEndpointVolume,
    IAudioEndpointVolume_Vtbl,
    0x5cdf2c82_841e_4546_9722_0cf74078229a
);
windows_core::imp::interface_hierarchy!(IAudioEndpointVolume, windows_core::IUnknown);
impl IAudioEndpointVolume {
    pub unsafe fn RegisterControlChangeNotify<P0>(&self, pnotify: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioEndpointVolumeCallback>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).RegisterControlChangeNotify)(
                windows_core::Interface::as_raw(self),
                pnotify.param().abi(),
            )
        }
    }
    pub unsafe fn UnregisterControlChangeNotify<P0>(&self, pnotify: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioEndpointVolumeCallback>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).UnregisterControlChangeNotify)(
                windows_core::Interface::as_raw(self),
                pnotify.param().abi(),
            )
        }
    }
    pub unsafe fn GetChannelCount(&self, pnchannelcount: *mut UINT) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetChannelCount)(
                windows_core::Interface::as_raw(self),
                pnchannelcount as _,
            )
        }
    }
    pub unsafe fn SetMasterVolumeLevel(
        &self,
        fleveldb: f32,
        pguideventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetMasterVolumeLevel)(
                windows_core::Interface::as_raw(self),
                fleveldb,
                pguideventcontext,
            )
        }
    }
    pub unsafe fn SetMasterVolumeLevelScalar(
        &self,
        flevel: f32,
        pguideventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetMasterVolumeLevelScalar)(
                windows_core::Interface::as_raw(self),
                flevel,
                pguideventcontext,
            )
        }
    }
    pub unsafe fn GetMasterVolumeLevel(&self, pfleveldb: *mut f32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetMasterVolumeLevel)(
                windows_core::Interface::as_raw(self),
                pfleveldb as _,
            )
        }
    }
    pub unsafe fn GetMasterVolumeLevelScalar(&self, pflevel: *mut f32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetMasterVolumeLevelScalar)(
                windows_core::Interface::as_raw(self),
                pflevel as _,
            )
        }
    }
    pub unsafe fn SetChannelVolumeLevel(
        &self,
        nchannel: UINT,
        fleveldb: f32,
        pguideventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetChannelVolumeLevel)(
                windows_core::Interface::as_raw(self),
                nchannel,
                fleveldb,
                pguideventcontext,
            )
        }
    }
    pub unsafe fn SetChannelVolumeLevelScalar(
        &self,
        nchannel: UINT,
        flevel: f32,
        pguideventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetChannelVolumeLevelScalar)(
                windows_core::Interface::as_raw(self),
                nchannel,
                flevel,
                pguideventcontext,
            )
        }
    }
    pub unsafe fn GetChannelVolumeLevel(
        &self,
        nchannel: UINT,
        pfleveldb: *mut f32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetChannelVolumeLevel)(
                windows_core::Interface::as_raw(self),
                nchannel,
                pfleveldb as _,
            )
        }
    }
    pub unsafe fn GetChannelVolumeLevelScalar(
        &self,
        nchannel: UINT,
        pflevel: *mut f32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetChannelVolumeLevelScalar)(
                windows_core::Interface::as_raw(self),
                nchannel,
                pflevel as _,
            )
        }
    }
    pub unsafe fn SetMute(&self, bmute: bool, pguideventcontext: LPCGUID) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetMute)(
                windows_core::Interface::as_raw(self),
                bmute.into(),
                pguideventcontext,
            )
        }
    }
    pub unsafe fn GetMute(&self, pbmute: *mut windows_core::BOOL) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetMute)(
                windows_core::Interface::as_raw(self),
                pbmute as _,
            )
        }
    }
    pub unsafe fn GetVolumeStepInfo(
        &self,
        pnstep: *mut UINT,
        pnstepcount: *mut UINT,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetVolumeStepInfo)(
                windows_core::Interface::as_raw(self),
                pnstep as _,
                pnstepcount as _,
            )
        }
    }
    pub unsafe fn VolumeStepUp(&self, pguideventcontext: LPCGUID) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).VolumeStepUp)(
                windows_core::Interface::as_raw(self),
                pguideventcontext,
            )
        }
    }
    pub unsafe fn VolumeStepDown(&self, pguideventcontext: LPCGUID) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).VolumeStepDown)(
                windows_core::Interface::as_raw(self),
                pguideventcontext,
            )
        }
    }
    pub unsafe fn QueryHardwareSupport(
        &self,
        pdwhardwaresupportmask: *mut DWORD,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).QueryHardwareSupport)(
                windows_core::Interface::as_raw(self),
                pdwhardwaresupportmask as _,
            )
        }
    }
    pub unsafe fn GetVolumeRange(
        &self,
        pflvolumemindb: *mut f32,
        pflvolumemaxdb: *mut f32,
        pflvolumeincrementdb: *mut f32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetVolumeRange)(
                windows_core::Interface::as_raw(self),
                pflvolumemindb as _,
                pflvolumemaxdb as _,
                pflvolumeincrementdb as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioEndpointVolume_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub RegisterControlChangeNotify: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub UnregisterControlChangeNotify: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetChannelCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub SetMasterVolumeLevel:
        unsafe extern "system" fn(*mut core::ffi::c_void, f32, LPCGUID) -> windows_core::HRESULT,
    pub SetMasterVolumeLevelScalar:
        unsafe extern "system" fn(*mut core::ffi::c_void, f32, LPCGUID) -> windows_core::HRESULT,
    pub GetMasterVolumeLevel:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut f32) -> windows_core::HRESULT,
    pub GetMasterVolumeLevelScalar:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut f32) -> windows_core::HRESULT,
    pub SetChannelVolumeLevel: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        f32,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub SetChannelVolumeLevelScalar: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        f32,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub GetChannelVolumeLevel:
        unsafe extern "system" fn(*mut core::ffi::c_void, UINT, *mut f32) -> windows_core::HRESULT,
    pub GetChannelVolumeLevelScalar:
        unsafe extern "system" fn(*mut core::ffi::c_void, UINT, *mut f32) -> windows_core::HRESULT,
    pub SetMute: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::BOOL,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub GetMute: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::BOOL,
    ) -> windows_core::HRESULT,
    pub GetVolumeStepInfo: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut UINT,
        *mut UINT,
    ) -> windows_core::HRESULT,
    pub VolumeStepUp:
        unsafe extern "system" fn(*mut core::ffi::c_void, LPCGUID) -> windows_core::HRESULT,
    pub VolumeStepDown:
        unsafe extern "system" fn(*mut core::ffi::c_void, LPCGUID) -> windows_core::HRESULT,
    pub QueryHardwareSupport:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DWORD) -> windows_core::HRESULT,
    pub GetVolumeRange: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut f32,
        *mut f32,
        *mut f32,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioEndpointVolumeCallback,
    IAudioEndpointVolumeCallback_Vtbl,
    0x657804fa_d6ad_4496_8a60_352752af4f89
);
windows_core::imp::interface_hierarchy!(IAudioEndpointVolumeCallback, windows_core::IUnknown);
#[repr(C)]
pub struct IAudioEndpointVolumeCallback_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    OnNotify: usize,
}
windows_core::imp::define_interface!(
    IDispatch,
    IDispatch_Vtbl,
    0x00020400_0000_0000_c000_000000000046
);
windows_core::imp::interface_hierarchy!(IDispatch, windows_core::IUnknown);
#[repr(C)]
pub struct IDispatch_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    GetTypeInfoCount: usize,
    GetTypeInfo: usize,
    GetIDsOfNames: usize,
    Invoke: usize,
}
pub type IID = windows_core::GUID;
windows_core::imp::define_interface!(
    IMMDevice,
    IMMDevice_Vtbl,
    0xd666063f_1587_4e43_81f1_b948e807363f
);
windows_core::imp::interface_hierarchy!(IMMDevice, windows_core::IUnknown);
impl IMMDevice {
    pub unsafe fn Activate<T>(
        &self,
        dwclsctx: DWORD,
        pactivationparams: Option<*const PROPVARIANT>,
    ) -> windows_core::Result<T>
    where
        T: windows_core::Interface,
    {
        let mut result__ = core::ptr::null_mut();
        unsafe {
            (windows_core::Interface::vtable(self).Activate)(
                windows_core::Interface::as_raw(self),
                &T::IID,
                dwclsctx,
                pactivationparams.unwrap_or(core::mem::zeroed()) as _,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
    pub unsafe fn OpenPropertyStore(
        &self,
        stgmaccess: DWORD,
        ppproperties: *mut Option<IPropertyStore>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OpenPropertyStore)(
                windows_core::Interface::as_raw(self),
                stgmaccess,
                ppproperties as _,
            )
        }
    }
    pub unsafe fn GetId(&self, ppstrid: *mut windows_core::PWSTR) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetId)(
                windows_core::Interface::as_raw(self),
                ppstrid as _,
            )
        }
    }
    pub unsafe fn GetState(&self, pdwstate: *mut DWORD) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetState)(
                windows_core::Interface::as_raw(self),
                pdwstate as _,
            )
        }
    }
}
#[repr(C)]
pub struct IMMDevice_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub Activate: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const IID,
        DWORD,
        *const PROPVARIANT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub OpenPropertyStore: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        DWORD,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetId: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetState:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DWORD) -> windows_core::HRESULT,
}
pub type INT = i32;
windows_core::imp::define_interface!(
    IPropertyStore,
    IPropertyStore_Vtbl,
    0x886d8eeb_8cf2_4446_8d02_cdba1dbdcf99
);
windows_core::imp::interface_hierarchy!(IPropertyStore, windows_core::IUnknown);
#[repr(C)]
pub struct IPropertyStore_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    GetCount: usize,
    GetAt: usize,
    GetValue: usize,
    SetValue: usize,
    Commit: usize,
}
windows_core::imp::define_interface!(
    ISequentialStream,
    ISequentialStream_Vtbl,
    0x0c733a30_2a1c_11ce_ade5_00aa0044773d
);
windows_core::imp::interface_hierarchy!(ISequentialStream, windows_core::IUnknown);
#[repr(C)]
pub struct ISequentialStream_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    Read: usize,
    Write: usize,
}
windows_core::imp::define_interface!(
    IStorage,
    IStorage_Vtbl,
    0x0000000b_0000_0000_c000_000000000046
);
windows_core::imp::interface_hierarchy!(IStorage, windows_core::IUnknown);
#[repr(C)]
pub struct IStorage_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    CreateStream: usize,
    OpenStream: usize,
    CreateStorage: usize,
    OpenStorage: usize,
    CopyTo: usize,
    MoveElementTo: usize,
    Commit: usize,
    Revert: usize,
    EnumElements: usize,
    DestroyElement: usize,
    RenameElement: usize,
    SetElementTimes: usize,
    SetClass: usize,
    SetStateBits: usize,
    Stat: usize,
}
windows_core::imp::define_interface!(
    IStream,
    IStream_Vtbl,
    0x0000000c_0000_0000_c000_000000000046
);
impl core::ops::Deref for IStream {
    type Target = ISequentialStream;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(IStream, windows_core::IUnknown, ISequentialStream);
#[repr(C)]
pub struct IStream_Vtbl {
    pub base__: ISequentialStream_Vtbl,
    Seek: usize,
    SetSize: usize,
    CopyTo: usize,
    Commit: usize,
    Revert: usize,
    LockRegion: usize,
    UnlockRegion: usize,
    Stat: usize,
    Clone: usize,
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union LARGE_INTEGER {
    pub Anonymous: LARGE_INTEGER_0,
    pub u: LARGE_INTEGER_1,
    pub QuadPart: LONGLONG,
}
impl Default for LARGE_INTEGER {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct LARGE_INTEGER_0 {
    pub LowPart: DWORD,
    pub HighPart: LONG,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct LARGE_INTEGER_1 {
    pub LowPart: DWORD,
    pub HighPart: LONG,
}
pub type LONG = i32;
pub type LONGLONG = i64;
pub type LPCGUID = *const windows_core::GUID;
pub type LPSAFEARRAY = *mut SAFEARRAY;
pub type LPSECURITY_ATTRIBUTES = *mut SECURITY_ATTRIBUTES;
pub type LPVERSIONEDSTREAM = *mut VERSIONEDSTREAM;
pub type LPVOID = *mut core::ffi::c_void;
pub type PROPVARIANT = tagPROPVARIANT;
pub type PROPVAR_PAD1 = u16;
pub type PROPVAR_PAD2 = u16;
pub type PROPVAR_PAD3 = u16;
pub type PVOID = *mut core::ffi::c_void;
pub type REFERENCE_TIME = i64;
#[repr(C, packed(8))]
#[derive(Clone, Copy)]
pub struct SAFEARRAY {
    pub cDims: USHORT,
    pub fFeatures: USHORT,
    pub cbElements: ULONG,
    pub cLocks: ULONG,
    pub pvData: PVOID,
    pub rgsabound: [SAFEARRAYBOUND; 1],
}
impl Default for SAFEARRAY {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct SAFEARRAYBOUND {
    pub cElements: ULONG,
    pub lLbound: LONG,
}
pub type SCODE = i32;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct SECURITY_ATTRIBUTES {
    pub nLength: DWORD,
    pub lpSecurityDescriptor: LPVOID,
    pub bInheritHandle: windows_core::BOOL,
}
pub type SHORT = i16;
pub const S_OK: windows_core::HRESULT = windows_core::HRESULT(0x0_u32 as _);
pub type UCHAR = u8;
pub type UINT = u32;
pub type UINT32 = u32;
#[repr(C)]
#[derive(Clone, Copy)]
pub union ULARGE_INTEGER {
    pub Anonymous: ULARGE_INTEGER_0,
    pub u: ULARGE_INTEGER_1,
    pub QuadPart: ULONGLONG,
}
impl Default for ULARGE_INTEGER {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct ULARGE_INTEGER_0 {
    pub LowPart: DWORD,
    pub HighPart: DWORD,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct ULARGE_INTEGER_1 {
    pub LowPart: DWORD,
    pub HighPart: DWORD,
}
pub type ULONG = u32;
pub type ULONGLONG = u64;
pub type USHORT = u16;
pub type VARIANT_BOOL = i16;
pub type VARTYPE = u16;
#[repr(C, packed(8))]
#[derive(Default)]
pub struct VERSIONEDSTREAM {
    pub guidVersion: windows_core::GUID,
    pub pStream: core::mem::ManuallyDrop<Option<IStream>>,
}
#[repr(C, packed(1))]
#[derive(Clone, Copy, Default)]
pub struct WAVEFORMATEX {
    pub wFormatTag: WORD,
    pub nChannels: WORD,
    pub nSamplesPerSec: DWORD,
    pub nAvgBytesPerSec: DWORD,
    pub nBlockAlign: WORD,
    pub wBitsPerSample: WORD,
    pub cbSize: WORD,
}
pub type WORD = u16;
#[repr(C, packed(8))]
pub struct tagPROPVARIANT {
    pub Anonymous: tagPROPVARIANT_0,
}
impl Clone for tagPROPVARIANT {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
pub union tagPROPVARIANT_0 {
    pub Anonymous: core::mem::ManuallyDrop<tagPROPVARIANT_0_0>,
    pub decVal: DECIMAL,
}
impl Clone for tagPROPVARIANT_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
pub struct tagPROPVARIANT_0_0 {
    pub vt: VARTYPE,
    pub wReserved1: PROPVAR_PAD1,
    pub wReserved2: PROPVAR_PAD2,
    pub wReserved3: PROPVAR_PAD3,
    pub Anonymous: tagPROPVARIANT_0_0_0,
}
impl Clone for tagPROPVARIANT_0_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT_0_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
pub union tagPROPVARIANT_0_0_0 {
    pub cVal: CHAR,
    pub bVal: UCHAR,
    pub iVal: SHORT,
    pub uiVal: USHORT,
    pub lVal: LONG,
    pub ulVal: ULONG,
    pub intVal: INT,
    pub uintVal: UINT,
    pub hVal: LARGE_INTEGER,
    pub uhVal: ULARGE_INTEGER,
    pub fltVal: FLOAT,
    pub dblVal: DOUBLE,
    pub boolVal: VARIANT_BOOL,
    pub __OBSOLETE__VARIANT_BOOL: VARIANT_BOOL,
    pub scode: SCODE,
    pub cyVal: CY,
    pub date: DATE,
    pub filetime: FILETIME,
    pub puuid: *mut CLSID,
    pub pclipdata: *mut CLIPDATA,
    pub bstrVal: core::mem::ManuallyDrop<windows_core::BSTR>,
    pub bstrblobVal: BSTRBLOB,
    pub blob: BLOB,
    pub pszVal: windows_core::PSTR,
    pub pwszVal: windows_core::PWSTR,
    pub punkVal: core::mem::ManuallyDrop<Option<windows_core::IUnknown>>,
    pub pdispVal: core::mem::ManuallyDrop<Option<IDispatch>>,
    pub pStream: core::mem::ManuallyDrop<Option<IStream>>,
    pub pStorage: core::mem::ManuallyDrop<Option<IStorage>>,
    pub pVersionedStream: LPVERSIONEDSTREAM,
    pub parray: LPSAFEARRAY,
    pub cac: CAC,
    pub caub: CAUB,
    pub cai: CAI,
    pub caui: CAUI,
    pub cal: CAL,
    pub caul: CAUL,
    pub cah: CAH,
    pub cauh: CAUH,
    pub caflt: CAFLT,
    pub cadbl: CADBL,
    pub cabool: CABOOL,
    pub cascode: CASCODE,
    pub cacy: CACY,
    pub cadate: CADATE,
    pub cafiletime: CAFILETIME,
    pub cauuid: CACLSID,
    pub caclipdata: CACLIPDATA,
    pub cabstr: CABSTR,
    pub cabstrblob: CABSTRBLOB,
    pub calpstr: CALPSTR,
    pub calpwstr: CALPWSTR,
    pub capropvar: CAPROPVARIANT,
    pub pcVal: *mut CHAR,
    pub pbVal: *mut UCHAR,
    pub piVal: *mut SHORT,
    pub puiVal: *mut USHORT,
    pub plVal: *mut LONG,
    pub pulVal: *mut ULONG,
    pub pintVal: *mut INT,
    pub puintVal: *mut UINT,
    pub pfltVal: *mut FLOAT,
    pub pdblVal: *mut DOUBLE,
    pub pboolVal: *mut VARIANT_BOOL,
    pub pdecVal: *mut DECIMAL,
    pub pscode: *mut SCODE,
    pub pcyVal: *mut CY,
    pub pdate: *mut DATE,
    pub pbstrVal: *mut windows_core::BSTR,
    pub ppunkVal: *mut Option<windows_core::IUnknown>,
    pub ppdispVal: *mut Option<IDispatch>,
    pub pparray: *mut LPSAFEARRAY,
    pub pvarVal: *mut PROPVARIANT,
}
impl Clone for tagPROPVARIANT_0_0_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT_0_0_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
