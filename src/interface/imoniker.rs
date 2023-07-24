use intercom::prelude::*;

use super::IPersistStream;

#[com_interface(com_iid = "0000000F-0000-0000-C000-000000000046")]
pub trait IMoniker: IPersistStream {
    fn bind_object(&self) -> ComResult<()>;
    fn bind_to_storage(&self) -> ComResult<()>;
    fn reduce(&self) -> ComResult<()>;
    fn compose_with(&self) -> ComResult<()>;
    fn get_enum(&self) -> ComResult<()>;
    fn is_equal(&self) -> ComResult<()>;
    fn hash(&self) -> ComResult<()>;
    fn is_running(&self) -> ComResult<()>;
    fn get_time_of_last_change(&self) -> ComResult<()>;
    fn inverse(&self) -> ComResult<()>;
    fn common_prefix_with(&self) -> ComResult<()>;
    fn relative_path_to(&self) -> ComResult<()>;
    fn get_display_name(&self) -> ComResult<()>;
    fn parse_display_name(&self) -> ComResult<()>;
    fn is_system_moniker(&self) -> ComResult<()>;
}