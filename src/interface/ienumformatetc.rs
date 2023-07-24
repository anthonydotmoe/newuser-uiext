use intercom::{ IUnknown, prelude::* };

use super::types::ComFORMATETC;

#[com_interface(com_iid = "00000103-0000-0000-C000-000000000046")]
pub trait IEnumFORMATETC: IUnknown {
    fn next(&self, celt: u32) -> ComResult<(ComFORMATETC, i32)>;
    fn skip(&self, celt: u32) -> ComResult<()>;
    fn reset(&self) -> ComResult<()>;
    fn clone(&self) -> ComResult<ComRc<dyn IEnumFORMATETC>>;
}
