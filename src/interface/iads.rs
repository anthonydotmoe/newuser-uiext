use intercom::{prelude::*, BString, Variant};

use super::IDispatch;

#[com_interface(com_iid = "FD8256D0-FD15-11CE-ABC4-02608C9E7553")]
pub trait IADs: IDispatch {
    fn name(&self) -> ComResult<BString>;
    fn class(&self) -> ComResult<BString>;
    fn guid(&self) -> ComResult<BString>;
    fn ads_path(&self) -> ComResult<BString>;
    fn parent(&self) -> ComResult<BString>;
    fn schema(&self) -> ComResult<BString>;
    fn get_info(&self) -> ComResult<()>;
    fn set_info(&self) -> ComResult<()>;
    fn get(&self, name: BString) -> ComResult<Variant>;
    fn put(&self, name: BString, prop: Variant) -> ComResult<()>;
    fn get_ex(&self, name: BString) -> ComResult<Variant>;
    fn put_ex(&self, name: BString, prop: Variant) -> ComResult<()>;
    fn get_info_ex(&self, properties: Variant, _rsvd: i32) -> ComResult<()>;
}