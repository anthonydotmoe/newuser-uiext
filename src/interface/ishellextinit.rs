use intercom::{IUnknown, prelude::* };

use super::IDataObject;

#[com_interface(com_iid = "000214E8-0000-0000-C000-000000000046")]
pub trait IShellExtInit: IUnknown {
    /// *pidlFolder* - Points to an ITEMIDLIST structure
    /// 
    /// *dataobj* - Points to an IDataObject interface
    /// 
    /// *hkeyProgID* - Registry key for the file object or folder type
    /// 
    /// In an ADSI context menu extension, the pidlFolder and hkeyProgID
    /// parameters aren't used.
    fn initialize(&self, _pidlfolder: isize, dataobj: ComRc<dyn IDataObject>, _hkeyprogid: isize) -> ComResult<()>;
}