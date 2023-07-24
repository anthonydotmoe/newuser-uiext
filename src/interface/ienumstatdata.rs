use intercom::{ IUnknown, prelude::* };

use super::types::ComSTATDATA;

#[com_interface(com_iid = "00000105-0000-0000-C000-000000000046")]
pub trait IEnumSTATDATA: IUnknown {
    fn next(&self, celt: u32) -> ComResult<(ComSTATDATA, i32)>;
    fn skip(&self, celt: u32) -> ComResult<()>;
    fn reset(&self) -> ComResult<()>;
    fn clone(&self) -> ComResult<ComRc<dyn IEnumSTATDATA>>;
}
