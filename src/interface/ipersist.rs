use intercom::{ prelude::*, IUnknown };

#[com_interface(com_iid = "0000010C-0000-0000-C000-000000000046")]
pub trait IPersist: IUnknown {
    fn get_clsid(&self) -> ComResult<u32>;
}