use intercom::{ IUnknown, prelude::* };

#[com_interface(com_iid = "F2573587-E6FC-11D2-82AF-00C04F68928B")]
pub trait IDsAdminNewObj: IUnknown {
    fn set_buttons(&self, curr_index: u32, valid: bool) -> ComResult<()>;
    fn get_page_counts(&self) -> ComResult<(u32, u32)>;
}