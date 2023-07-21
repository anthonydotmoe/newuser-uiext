use intercom::{ prelude::*, IUnknown };

use super::types::{ ComFORMATETC, ComSTGMEDIUM };

#[com_interface(com_iid = "0000010E-0000-0000-C000-000000000046")]
pub trait IDataObject: IUnknown {
    fn get_data(&self, ) -> ComResult<()>;
    fn get_data_here(&self, pformatetc: *const ComFORMATETC, pmedium: *mut ComSTGMEDIUM) -> ComResult<()>;
    fn query_get_data(&self, ) -> ComResult<()>;
    fn get_canonical_format(&self, ) -> ComResult<()>;
    fn set_data(&self, ) -> ComResult<()>;
    fn enum_format_etc(&self, ) -> ComResult<()>;
    fn d_advise(&self, ) -> ComResult<()>;
    fn d_unadvise(&self, ) -> ComResult<()>;
    fn enum_d_advise(&self) -> ComResult<()>;
}