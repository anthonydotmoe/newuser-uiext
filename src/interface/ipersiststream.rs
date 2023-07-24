use intercom::prelude::*;

use super::{
    IStream,
    IPersist,
};

#[com_interface(com_iid = "00000109-0000-0000-C000-000000000046")]
pub trait IPersistStream: IPersist {
    fn is_dirty(&self) -> ComResult<()>;
    fn load(&self, stream: ComRc<dyn IStream>) -> ComResult<()>;
    fn save(&self, stream: ComRc<dyn IStream>, clear_dirty: bool) -> ComResult<()>;
    fn get_size_max(&self) -> ComResult<u64>;
}