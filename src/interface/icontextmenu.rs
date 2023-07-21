use intercom::{ IUnknown, prelude::* };

use super::types::{ ComHMENU, LPCWSTR, ComINVOKECOMMANDINFO };

#[com_interface(com_iid = "000214E4-0000-0000-C000-000000000046")]
pub trait IContextMenu: IUnknown {
    fn query_context_menu(&self, hmenu: ComHMENU, index_menu: u32, id_cmd_first: u32, id_cmd_last: u32, flags: u32) -> ComResult<()>;
    fn invoke_command(&self, ici: *mut ComINVOKECOMMANDINFO) -> ComResult<()>;
    fn get_command_string(&self, id_cmd: u32, flags: u32, _rsvd: usize, name: LPCWSTR, name_max: u32) -> ComResult<()>;
}