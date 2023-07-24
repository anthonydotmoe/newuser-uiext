use intercom::{ IUnknown, prelude::* };

use super::{
    IMoniker,
    types::{ ComFORMATETC, ComSTGMEDIUM },
};

#[com_interface(com_iid = "00000150-0000-0000-C000-000000000046")]
pub trait IAdviseSink: IUnknown {
    fn on_data_change(&self, format: *const ComFORMATETC, medium: *const ComSTGMEDIUM) -> ComResult<()>;
    fn on_view_change(&self, aspect: u32, index: i32) -> ComResult<()>;
    fn on_rename(&self, moniker: ComRc<dyn IMoniker>) -> ComResult<()>;
}
