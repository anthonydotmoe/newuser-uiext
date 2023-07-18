use intercom::BString;
use intercom::Variant;
use intercom::prelude::*;
use intercom::IUnknown;

use super::{
    ITypeInfo,
    types::ComLPARAM,
};

#[com_interface(com_iid = "00020400-0000-0000-C000-000000000046")]
pub trait IDispatch: IUnknown {
    fn get_type_info_count(&self) -> ComResult<u32>;
    fn get_type_info(&self, type_info: u32, lcid: u32) -> ComResult<ComRc<dyn ITypeInfo>>;
    fn get_ids_of_names(&self, _riid: intercom::GUID, names: BString, count: u32, lcid: u32) -> ComResult<i32>;    
    fn invoke(&self, dispid: i32, _riid: intercom::GUID, lcid: u32, flags: u16, disp_params: ComLPARAM) -> ComResult<(Variant, u32, u32)>;
}