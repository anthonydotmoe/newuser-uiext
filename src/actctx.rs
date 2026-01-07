use std::fmt;

use bitflags::bitflags;

use windows::{
    Win32::{
        Foundation::HANDLE,
        System::ApplicationInstallationAndServicing::{ACTCTXW, CreateActCtxW, ReleaseActCtx},
    },
    core::PCWSTR
};

#[derive(Debug)]
pub struct ReleaseActCtxGuard {
    actctx: HANDLE,
    _actctxopt: OwnedActCtxOpt,
}

impl Drop for ReleaseActCtxGuard {
    fn drop(&mut self) {
        log::debug!("Calling ReleaseActCtx()");
        unsafe {
            ReleaseActCtx(self.actctx);
        }
    }
}

impl ReleaseActCtxGuard {
    pub fn to_raw(&self) -> HANDLE {
        self.actctx.clone()
    }
}

pub struct ActCtx(HANDLE);

impl Into<HANDLE> for &ActCtx {
    fn into(self) -> HANDLE {
        (*self).0
    }
}

bitflags!{
    #[repr(transparent)]
    #[derive(Debug)]
    struct ActCtxFlags: u32 {
        const PROCESSOR_ARCHITECTURE_VALID = 0x001;
        const LANGID_VALID                 = 0x002;
        const ASSEMBLY_DIRECTORY_VALID     = 0x004;
        const RESOURCE_NAME_VALID          = 0x008;
        const SET_PROCESS_DEFAULT          = 0x010;
        const APPLICATION_NAME_VALID       = 0x020;
        const HMODULE_VALID                = 0x080;
    }
}

#[repr(u16)]
pub enum ProcessorArchitecture {
    AMD64 = 9,
    Arm = 5,
    Arm64 = 12,
    IA64 = 6,
    Intel = 0,
    Unknown = 0xffff,
}

impl From<u16> for ProcessorArchitecture {
    fn from(value: u16) -> Self {
        match value {
            0  => ProcessorArchitecture::Intel,
            5  => ProcessorArchitecture::Arm,
            6  => ProcessorArchitecture::IA64,
            9  => ProcessorArchitecture::AMD64,
            12 => ProcessorArchitecture::Arm64,
            _  => ProcessorArchitecture::Unknown,
        }
    }
}

impl Into<u16> for ProcessorArchitecture {
    fn into(self) -> u16 {
        self as u16
    }
}

pub fn create_act_ctx(opt: CreateActCtxOpt) -> Result<ReleaseActCtxGuard, CreateActCtxError> {
    let owned_opt = OwnedActCtxOpt::try_from(opt)?;
    match unsafe { CreateActCtxW(owned_opt.as_ptr()) } {
        Ok(handle) => {
            Ok(ReleaseActCtxGuard { actctx: handle, _actctxopt: owned_opt })
        },
        Err(e) => Err(CreateActCtxError::WindowsError(e))
    }
}

#[derive(Default)]
pub struct CreateActCtxOpt {
    pub source: Option<String>,
    pub architecture: Option<ProcessorArchitecture>,
    pub lang_id: Option<u16>,
    pub assembly_dir: Option<String>,
    pub resource_name: Option<String>,
    pub app_name: Option<String>,
    pub hmodule: Option<windows::Win32::Foundation::HMODULE>,
}

#[derive(Debug)]
pub enum CreateActCtxError {
    BothSourceAndHModuleDefined,
    NeitherSourceNorHModuleDefined,
    WindowsError(windows::core::Error),
}

impl fmt::Display for CreateActCtxError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            CreateActCtxError::BothSourceAndHModuleDefined => {
                write!(f, "Both source and hmodule are Some. One must be specified")
            }
            CreateActCtxError::NeitherSourceNorHModuleDefined => {
                write!(f, "Both source and hmodule are None. Only one may be specified")
            }
            CreateActCtxError::WindowsError(ref e) => {
                e.fmt(f)
            }
        }
    }
}

impl std::error::Error for CreateActCtxError {}

#[derive(Debug)]
struct OwnedActCtxOpt {
    actctx_struct: Box<ACTCTXW>,
    
    // Allowing unused since we keep the strings so raw pointers work in FFI
    #[allow(unused)]
    source_wstr: Option<Box<Vec<u16>>>,
    #[allow(unused)]
    assembly_dir_wstr: Option<Box<Vec<u16>>>,
    #[allow(unused)]
    resource_name_wstr: Option<Box<Vec<u16>>>,
    #[allow(unused)]
    app_name_wstr: Option<Box<Vec<u16>>>,
}

impl OwnedActCtxOpt {
    fn as_ptr(&self) -> *const ACTCTXW {
        &*self.actctx_struct as *const ACTCTXW
    }
}

impl TryFrom<CreateActCtxOpt> for OwnedActCtxOpt {
    type Error = CreateActCtxError;

    fn try_from(opt: CreateActCtxOpt) -> Result<Self, CreateActCtxError> {
        match (&opt.source, &opt.hmodule) {
            (None, None) => {
                return Err(CreateActCtxError::NeitherSourceNorHModuleDefined);
            }
            (Some(_), Some(_)) => {
                return Err(CreateActCtxError::BothSourceAndHModuleDefined);
            }
            _ => {}
        }
        let source_wstr = opt.source.map(|source| string_to_pcwstr(source));
        let assembly_dir_wstr = opt.assembly_dir.map(|dir| string_to_pcwstr(dir));
        let resource_name_wstr = opt.resource_name.map(|res| string_to_pcwstr(res));
        let app_name_wstr = opt.app_name.map(|app| string_to_pcwstr(app));

        let mut flags: ActCtxFlags = ActCtxFlags::empty();
        let mut actctx: ACTCTXW = unsafe { std::mem::zeroed() };
        
        actctx.cbSize = std::mem::size_of::<ACTCTXW>() as u32;
        
        if let Some(source) = &source_wstr {
            actctx.lpSource = PCWSTR(source.as_ptr());
        }

        if let Some(arch) = opt.architecture {
            flags |= ActCtxFlags::PROCESSOR_ARCHITECTURE_VALID;
            actctx.wProcessorArchitecture = arch.into();
        }
                
        if let Some(langid) = opt.lang_id {
            flags |= ActCtxFlags::LANGID_VALID;
            actctx.wLangId = langid;
        }
                
        if let Some(dir) = &assembly_dir_wstr {
            flags |= ActCtxFlags::ASSEMBLY_DIRECTORY_VALID;
            actctx.lpAssemblyDirectory = PCWSTR(dir.as_ptr());
        }
                
        // TODO: Accept string or resource ID
        if let Some(res) = &resource_name_wstr {
            flags |= ActCtxFlags::RESOURCE_NAME_VALID;
            //actctx.lpResourceName = PCWSTR(res.as_ptr());
            actctx.lpResourceName = PCWSTR(2u16 as _);
        }
                
        if let Some(app) = &app_name_wstr {
            flags |= ActCtxFlags::APPLICATION_NAME_VALID;
            actctx.lpApplicationName = PCWSTR(app.as_ptr());
        }

        if let Some(hmodule) = opt.hmodule {
            flags |= ActCtxFlags::HMODULE_VALID;
            actctx.hModule = hmodule;
        }
        
        actctx.dwFlags = flags.bits();
        
        log::debug!("Created ACTCTXW: {:?}", actctx);
                
        let actctx_struct = Box::new(actctx);
        Ok(Self { actctx_struct, source_wstr, assembly_dir_wstr, resource_name_wstr, app_name_wstr })
    }
}

fn string_to_pcwstr(s: String) -> Box<Vec<u16>> {
    let mut wstr: Vec<u16> = s.encode_utf16().collect();
    wstr.push(0);
    Box::new(wstr)
}
