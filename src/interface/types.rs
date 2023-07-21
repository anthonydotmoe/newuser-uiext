use core::fmt;

use bitflags::bitflags;

use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;

use windows::{
    Win32::{
        Networking::ActiveDirectory::{
            DSA_NEWOBJ_CTX_PRECOMMIT,
            DSA_NEWOBJ_CTX_COMMIT,
            DSA_NEWOBJ_CTX_POSTCOMMIT,
            DSA_NEWOBJ_CTX_CLEANUP,
        }, UI::{
            Controls::HPROPSHEETPAGE,
            WindowsAndMessaging::{HICON, HMENU}, Shell::CMINVOKECOMMANDINFOEX
        }, Foundation::{
            BOOL, LPARAM
        }, System::Com::{
            FORMATETC, STGMEDIUM            
        },
    }, core::PCWSTR};

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComLPFNSVADDPROPSHEETPAGE(
    pub unsafe extern "stdcall" fn(param0: HPROPSHEETPAGE, param1: LPARAM) -> BOOL
);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComLPARAM(pub LPARAM);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComFORMATETC(pub FORMATETC);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComHMENU(pub HMENU);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput, intercom::ExternOutput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComSTGMEDIUM(pub STGMEDIUM);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComINVOKECOMMANDINFO(pub CMINVOKECOMMANDINFOEX);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComPCWSTR(pub PCWSTR);

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct LPCWSTR(pub *const u16);

impl std::fmt::Display for LPCWSTR {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        unsafe {
            let mut len = 0;
            while *self.0.offset(len) != 0 {
                len += 1;
            }
            let slice = std::slice::from_raw_parts(self.0, len as usize);
            let os_string = OsString::from_wide(slice);
            write!(f, "{}", os_string.to_string_lossy())
        }
    }
}

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(C)]
pub struct DsaNewObjDispInfo {
    pub size: u32,
    pub class_icon: HICON,
    pub wiz_title: LPCWSTR,
    pub container_display_name: LPCWSTR,
}

// TODO: DsaNewObjCtx seems to work, need to find out how.

#[derive(Debug)]
#[repr(C)]
pub enum DsaNewObjCtx {
    Precommit,
    Commit,
    PostCommit,
    Cleanup,
    Unknown(u32),
}

#[derive(intercom::ExternType, intercom::ForeignType, intercom::ExternInput)]
#[allow(non_camel_case_types)]
#[repr(transparent)]
pub struct ComDsaNewObjCtx(pub DsaNewObjCtx);

impl From<u32> for ComDsaNewObjCtx {
    fn from(value: u32) -> Self {
        let context = match value {
            DSA_NEWOBJ_CTX_PRECOMMIT => DsaNewObjCtx::Precommit,            
            DSA_NEWOBJ_CTX_COMMIT => DsaNewObjCtx::Commit,
            DSA_NEWOBJ_CTX_POSTCOMMIT => DsaNewObjCtx::PostCommit,
            DSA_NEWOBJ_CTX_CLEANUP => DsaNewObjCtx::Cleanup,
            _ => DsaNewObjCtx::Unknown(value)
        };
        ComDsaNewObjCtx(context)        
    }
}

impl Into<u32> for ComDsaNewObjCtx {
    fn into(self) -> u32 {
        match self.0 {
            DsaNewObjCtx::Precommit => DSA_NEWOBJ_CTX_PRECOMMIT,
            DsaNewObjCtx::Commit => DSA_NEWOBJ_CTX_PRECOMMIT,
            DsaNewObjCtx::PostCommit => DSA_NEWOBJ_CTX_PRECOMMIT,
            DsaNewObjCtx::Cleanup => DSA_NEWOBJ_CTX_PRECOMMIT,
            DsaNewObjCtx::Unknown(n) => n,
        }
    }
}

bitflags! {
    #[repr(transparent)]
    #[derive(Debug)]
    pub struct CtxMenuInvokeFlags: u32 {
        const ICON           = 0x00000010;        
        const HOTKEY         = 0x00000020;
        const NO_UI          = 0x00000400;
        const NO_CONSOLE     = 0x00008000;
        const UNICODE        = 0x00010000;
      //const HASLINKNAME    = 0x00010000;
      //const HASTITLE       = 0x00020000;
        const FLAG_SEP_VDM   = 0x00040000;
        const ASYNCOK        = 0x00100000;
        const NOZONECHECKS   = 0x00800000;
        const FLAG_LOG_USAGE = 0x04000000;
        const SHIFT_DOWN     = 0x10000000;
        const PTINVOKE       = 0x20000000;
        const CONTROL_DOWN   = 0x40000000;
    }
}