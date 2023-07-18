use intercom::{prelude::*, Variant, BString};

use super::{
    IDispatch,
    IEnumVARIANT,
};

#[com_interface(com_iid = "001677D0-FD16-11CE-ABC4-02608C9E7553")]
pub trait IADsContainer: IDispatch {
    fn get_count(&self) -> ComResult<i32>;

    #[allow(non_snake_case)]
    fn get___new_enum(&self) -> ComResult<ComRc<dyn IEnumVARIANT>>;

    fn get_filter(&self) -> ComResult<Variant>;
    fn put_filter(&self, filter: Variant) -> ComResult<()>;
    fn get_hints(&self) -> ComResult<Variant>;
    fn put_hints(&self, hints: Variant) -> ComResult<()>;
    fn get_object(&self, class_name: BString, rel_name: BString) -> ComResult<ComRc<dyn IDispatch>>;
    fn create(&self, class_name: BString, rel_name: BString) -> ComResult<ComRc<dyn IDispatch>>;
    fn delete(&self, class_name: BString, rel_name: BString) -> ComResult<()>;
    fn copy_here(&self, source: BString, dest: BString) -> ComResult<ComRc<dyn IDispatch>>;
    fn move_here(&self, source: BString, dest: BString) -> ComResult<ComRc<dyn IDispatch>>;
}