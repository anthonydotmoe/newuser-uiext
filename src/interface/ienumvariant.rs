use intercom::{ IUnknown, prelude::*, Variant };

#[com_interface(com_iid = "00020404-0000-0000-C000-000000000046")]
pub trait IEnumVARIANT: IUnknown {
    fn next(&self, celt: u32) -> ComResult<(Variant, i32)>;
    fn skip(&self, celt: u32) -> ComResult<()>;
    fn reset(&self) -> ComResult<()>;
    fn clone(&self) -> ComResult<ComRc<dyn IEnumVARIANT>>;
}
