use intercom::prelude::*;

use super::ISequentialStream;

#[com_interface(com_iid = "0000000c-0000-0000-C000-000000000046")]
pub trait IStream: ISequentialStream {
    fn seek(&self, ) -> ComResult<()>;
    fn set_size(&self, ) -> ComResult<()>;
    fn copy_to(&self, ) -> ComResult<()>;
    fn commit(&self, ) -> ComResult<()>;
    fn revert(&self, ) -> ComResult<()>;
    fn lock_region(&self, ) -> ComResult<()>;
    fn unlock_region(&self, ) -> ComResult<()>;
    fn stat(&self, ) -> ComResult<()>;
    fn clone(&self, ) -> ComResult<()>;
}