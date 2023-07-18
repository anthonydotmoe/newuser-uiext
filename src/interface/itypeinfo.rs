use intercom::{IUnknown, prelude::*};

#[com_interface(com_iid = "00020401-0000-0000-c000-000000000046")]
pub trait ITypeInfo: IUnknown {
    fn get_type_attr(&self) -> ComResult<()>;
    fn get_type_alloc(&self) -> ComResult<()>;
    fn get_func_desc(&self) -> ComResult<()>;
    fn get_var_desc(&self) -> ComResult<()>;
    fn get_names(&self) -> ComResult<()>;
    fn get_ref_type_of_impl_type(&self) -> ComResult<()>;
    fn get_impl_type_flags(&self) -> ComResult<()>;
    fn get_ids_of_names(&self) -> ComResult<()>;
    fn invoke(&self) -> ComResult<()>;
    fn get_documentation(&self) -> ComResult<()>;
    fn get_dll_entry(&self) -> ComResult<()>;
    fn get_ref_type_info(&self) -> ComResult<()>;
    fn address_of_member(&self) -> ComResult<()>;
    fn create_instance(&self) -> ComResult<()>;
    fn get_mops(&self) -> ComResult<()>;
    fn get_containing_type_lib(&self) -> ComResult<()>;
    fn release_type_attr(&self) -> ComResult<()>;
    fn release_func_desc(&self) -> ComResult<()>;
    fn release_var_desc(&self) -> ComResult<()>;
}