use intercom::{ prelude::*, IUnknown };

#[com_interface(com_iid = "0c733a30-2a1c-11ce-ade5-00aa0044773d")]
pub trait ISequentialStream: IUnknown {
    fn read(&self, ) -> ComResult<()>;
    fn write(&self, ) -> ComResult<()>;
}