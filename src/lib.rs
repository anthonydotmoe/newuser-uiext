mod interface;
use actctx::ReleaseActCtxGuard;
use interface::{
    IADs,
    IADsContainer,
    IDsAdminNewObj,
    IContextMenu,
    IDsAdminNewObjExt,
    IShellExtInit,
    
    types::{
        LPCWSTR, ComDsaNewObjCtx, ComLPFNSVADDPROPSHEETPAGE, DsaNewObjDispInfo, ComLPARAM, ComHMENU, ComINVOKECOMMANDINFO, CtxMenuInvokeFlags,
    }
};

use rand::Rng;

// mod propsheet;
mod wm;

mod actctx;

mod resource_consts;

use intercom::{ prelude::*, BString, Variant, raw::E_INVALIDARG };

use windows::{Win32::{UI::{Controls::{PROPSHEETPAGEW, PROPSHEETPAGEW_0, PROPSHEETPAGEW_1, PROPSHEETPAGEW_2, PSP_DEFAULT, CreatePropertySheetPageW, PSP_HIDEHEADER, InitCommonControlsEx, INITCOMMONCONTROLSEX, ICC_STANDARD_CLASSES}, WindowsAndMessaging::{WM_INITDIALOG, GWL_USERDATA, WM_DESTROY, DefWindowProcW, InsertMenuItemW, MENUITEMINFOW, MIIM_STRING, MENU_ITEM_STATE, MIIM_ID, GetDlgItem, STM_SETICON, SendMessageW, WM_SETTEXT, WM_COMMAND, BN_CLICKED}}, Foundation::{HMODULE, LPARAM, HANDLE, HWND, WPARAM}, Graphics::Gdi::HBITMAP }, core::PCWSTR, w};
use windows::Win32::System::LibraryLoader::{GetModuleHandleExW, GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT};
use windows::Win32::UI::WindowsAndMessaging::HICON;
use windows::Win32::UI::WindowsAndMessaging::SetWindowLongPtrW;

com_library!(
    on_load = on_load,
    class MyNewUserWizard,
);


fn on_load() {
    // Set up logging to project directory
    use log::LevelFilter;
    simple_logging::log_to_file(
        &format!("{}\\debug.log", env!("CARGO_MANIFEST_DIR")),
        LevelFilter::Trace,
    )
    .unwrap();
    let test = generate_password();
    log::debug!("This is a password: {}", test);
}

#[com_class(clsid = "1af9e4e5-d2fc-41e6-9a7e-3fbdb38a34b4", IDsAdminNewObjExt, IShellExtInit, IContextMenu)]
#[derive(Default)]
struct MyNewUserWizard {
    
    debug: i32,
    obj_container: std::cell::RefCell<Option<ComRc<dyn IADsContainer>>>,    
    copy_source: std::cell::RefCell<Option<ComRc<dyn IADs>>>,
    new_object: Option<ComRc<dyn IADs>>,
    wizard: Option<ComRc<dyn IDsAdminNewObj>>,
    wiz_icon: HICON,
    actctx: Option<ReleaseActCtxGuard>,
    obj_dest_str: String,

}

impl IDsAdminNewObjExt for MyNewUserWizard {
    fn initialize(
        &mut self,
        obj_container: &ComItf<dyn IADsContainer>,
        copy_source: Option<&ComItf<dyn IADs>>,
        class_name: LPCWSTR,
        adminnewobj: &ComItf<dyn IDsAdminNewObj>, // Used for primary new object extensions
        disp_info: *const DsaNewObjDispInfo
    ) -> ComResult<()> {
        log::info!("Initialize called. objectClass: {}", class_name);
        
        self.debug = 42069;
        
        unsafe { log::info!("Display info: {}, {}", (*disp_info).wiz_title, (*disp_info).container_display_name) };
        self.obj_dest_str = unsafe { (*disp_info).container_display_name.to_string() };
        self.wiz_icon = unsafe { (*disp_info).class_icon };
        log::debug!("Got class icon: {:?}", self.wiz_icon);
        
        match class_name.to_string().as_str() {
            "user" => {
                log::info!("We support this!")
            }
            _ => {
                log::error!("This is not correct. Bail!");
                return Err(ComError::E_INVALIDARG)
            }
        }

        log::info!("Keeping copies of obj_container and copy_source");
        
        match copy_source {
            Some(interface) => {
                *self.copy_source.borrow_mut() = Some(interface.to_owned());
            }
            None => {
                log::info!("No copy source");
            }
        }
        
        *self.obj_container.borrow_mut() = Some(obj_container.to_owned());
        self.wizard = Some(adminnewobj.to_owned());
        
        let maybe_itf = self.copy_source.borrow();
        if let Some(ref itf) = *maybe_itf {
            match itf.get("mail".into()) {
                Ok(name) => {
                    match name {
                        Variant::String(s) => {
                            log::info!("The name of clone is {}", String::try_from(s).unwrap());
                        }
                        _ => {
                            log::error!("Didn't expect that type. {:?}", name)
                        }
                    }
                }
                Err(e) => {
                    log::error!("Got an error trying to use the copy's IADs::Get(\"distinguishedName\"): {}", e.to_string());
                }
            }
        }

        Ok(())
    }
    
    fn add_pages(&mut self, addpagefn: ComLPFNSVADDPROPSHEETPAGE, param: ComLPARAM) -> ComResult<()> {
        log::debug!("add_pages called");

        let hinstance = get_dll_hinstance();
        log::debug!("got hInstance: {:?}", &hinstance);

        // TODO:
        let mut handle = HANDLE(0);
        
        // Create activation context (common controls)
        
        let actctx = actctx::create_act_ctx(actctx::CreateActCtxOpt{
            hmodule: Some(hinstance),
            resource_name: Some("Nice".to_owned()),
            ..Default::default()
        });
        
        match actctx {
            Ok(actctx) => {
                handle = actctx.to_raw();
                self.actctx = Some(actctx);
            }
            Err(e) => {
                log::error!("Got error from CreateActCtx: {}", e);
                return Err(ComError::new_hr(intercom::raw::HRESULT { hr: 0x7FFF0001 }));
            }
        }
        
        let sex = INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_STANDARD_CLASSES,
        };
        
        match unsafe { InitCommonControlsEx(&sex as *const INITCOMMONCONTROLSEX).into() } {
            true => {
                log::debug!("InitCommonControlsEx returned true");
            }
            false => {
                log::error!("InitCommonControlsEx returned false!");
                return Err(ComError::new_hr(intercom::raw::HRESULT { hr: 0x7FFF0001 }));
            }
        }
        
        // Very unsafe! Getting a reference to self to use in the windowproc
        let this_lparam = LPARAM(self as *mut _ as isize);
        
        let pages = vec![resource_consts::DIALOG_1, resource_consts::DIALOG_2];
        for page in pages.iter() {
                
            let mut propsheet: PROPSHEETPAGEW = PROPSHEETPAGEW {
                dwSize: std::mem::size_of::<PROPSHEETPAGEW>() as u32,
                dwFlags: PSP_DEFAULT | PSP_HIDEHEADER,
                hInstance: hinstance,
                Anonymous1: PROPSHEETPAGEW_0 { pszTemplate: make_int_resource(*page) },
                Anonymous2: PROPSHEETPAGEW_1 { pszIcon: PCWSTR::null() },
                pszTitle: PCWSTR::null(),
                pfnDlgProc: Some(dlgproc),
                lParam: this_lparam,
                pfnCallback: None,
                pcRefParent: std::ptr::null_mut(),
                pszHeaderTitle: PCWSTR::null(),
                pszHeaderSubTitle: PCWSTR::null(),
                hActCtx: handle,
                Anonymous3: PROPSHEETPAGEW_2 { pszbmHeader: PCWSTR::null() },
            };
            
            log::debug!("Created page object: {:p}", &propsheet);
            
            let page1_h = unsafe { CreatePropertySheetPageW(&mut propsheet as _) };
            
            log::debug!("Got HPROPSHEETPAGE: {:?}", &page1_h);
            
            unsafe {
                match addpagefn.0(page1_h, param.0).as_bool() {
                    true => {
                        log::info!("addpagefn returned true");
                    }
                    false => {
                        log::warn!("addpagefn returned false");
                    }
                }
            }
        }
        
        Ok(())
    }

    fn set_object(&mut self, ad_obj: ComRc<dyn IADs>) -> ComResult<()> {
        log::info!("Keeping copy of new_object {:?}", ad_obj);

        match self.new_object {
            None => {
                log::debug!("There is not currently an object in self.new_object");
            },
            Some(_) => {
                log::warn!("There is an object in self.new_object, overwriting!");
                self.new_object = Some(ad_obj)
            }
        }

        Ok(())
        
    }

    fn write_data(&self, _hwnd: ComLPARAM, ctx: ComDsaNewObjCtx) -> ComResult<()> {
        log::error!("OnError called! {:?}", ctx.0);
        log::error!("WriteData called!");
        Err(ComError { hresult: E_INVALIDARG, error_info: None })
    }

    fn on_error(&self, _hwnd: ComLPARAM, _hr: intercom::raw::HRESULT, ctx: ComDsaNewObjCtx) -> ComResult<()> {
        log::error!("OnError called! {:?}", ctx.0);
        Err(ComError { hresult: E_INVALIDARG, error_info: None })
    }

    fn get_summary_info(&self) -> ComResult<BString> {
        let result_string: &'static str = "This is the result string";
        
        Ok(BString::from(result_string))
    }
}

impl IShellExtInit for MyNewUserWizard {
    fn initialize(&self, _pidlfolder: isize, _dataobj: ComRc<dyn interface::IDataObject>, _hkeyprogid: isize) -> ComResult<()> {
        Ok(())
    }
}

impl IContextMenu for MyNewUserWizard {
    fn query_context_menu(&self, hmenu: ComHMENU, index_menu:u32,id_cmd_first:u32,id_cmd_last:u32,flags:u32) -> ComResult<()> {
        log::debug!("QueryContextMenu(): indexMenu: {}, idCmdFirst: {}, idCmdLast: {}, uFlags: {}", index_menu, id_cmd_first, id_cmd_last, flags);
        
        if flags == 0 {
            let hmenu = hmenu.0;
            
            let text = w!("Offboard...");
            
            let menuitem = MENUITEMINFOW {
                cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                fMask: MIIM_STRING | MIIM_ID,
                fType: windows::Win32::UI::WindowsAndMessaging::MENU_ITEM_TYPE(0),
                fState: MENU_ITEM_STATE(0),
                wID: id_cmd_first,
                hSubMenu: windows::Win32::UI::WindowsAndMessaging::HMENU(0),
                hbmpChecked: HBITMAP(0),
                hbmpUnchecked: HBITMAP(0),
                dwItemData: 0,
                dwTypeData: windows::core::PWSTR(text.as_ptr() as *mut u16),
                cch: 0,
                hbmpItem: HBITMAP(0),
            };
            
            let ret: bool = unsafe { InsertMenuItemW(hmenu, id_cmd_first, false, &menuitem).into() };
            
            match ret {
                true => {
                    log::debug!("InsertMenuItemW returned true");
                    Err(ComError::new_hr(intercom::raw::HRESULT { hr: 1 }))
                },
                false => {
                    log::error!("InsertMenuItemW returned false");
                    Err(ComError::new_hr(intercom::raw::HRESULT { hr: 0x7FFF0001 }))
                }
            }
        } else {
            log::debug!("Exiting QueryContextMenu() flag is inappropriate.");
            Err(ComError::new_hr(intercom::raw::HRESULT { hr: 0x7FFF0001 }))
        }
        
    }
    
    fn get_command_string(&self, _id_cmd:u32, _flags: u32, _rsvd: usize, _name: LPCWSTR, _name_max: u32) -> ComResult<()> {
        Ok(())
    }
    
    fn invoke_command(&self, ici: *mut ComINVOKECOMMANDINFO) -> ComResult<()> {
        let flags: CtxMenuInvokeFlags = unsafe { std::mem::transmute::<u32, CtxMenuInvokeFlags>((*ici).0.fMask) };
        unsafe { log::debug!("InvokeCommand(): cbSize: {}, fMask: {:?}", (*ici).0.cbSize, flags) };
        Ok(())
    }
}

fn get_dll_hinstance() -> HMODULE {
    let mut hmodule: HMODULE = HMODULE::default();
    let address = get_dll_hinstance as *const _;

    let result = unsafe { GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT | GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
        PCWSTR::from_raw(address),
        &mut hmodule
    ) };

    match result.as_bool() {
        true => {
            hmodule
        }
        false => {
            log::error!("Wow, getting our HINSTANCE didn't work, oh well.");
            hmodule
        }
    }
}

const fn make_int_resource(resource_id: u16) -> PCWSTR {
    PCWSTR(resource_id as _)
}

unsafe extern "system" fn dlgproc(hwnd_dlg: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> isize {
    match message {
        WM_INITDIALOG => {
            log::debug!("WM_INITDIALOG received by dlgproc");

            // Getting our reference back from the `lparam`. Seems unsafe again.
            // let this = lparam.0 as *mut MyNewUserWizard;
            let this_propsheet = lparam.0 as *const PROPSHEETPAGEW;
            let this = (*this_propsheet).lParam.0 as *mut MyNewUserWizard;
            let this = &mut *this;
            
            let maybe_itf = &this.wizard;
            if let Some(itf) = maybe_itf {
                if let Ok((total_pages, first_page)) = itf.get_page_counts() {
                    log::info!("Got total_pages: {}, first_page: {}", &total_pages, &first_page);
                    for i in [0, 1] {
                        log::info!("Set next button to disabled on page {}", 0);
                        match itf.set_buttons(i, true.into()) {
                            Ok(_) => {
                                log::info!("Setting next button to disabled worked");
                            }
                            Err(e) => {
                                log::error!("Setting next button to disabled didn't work: {}", e.to_string())
                            }
                        }
                    }
                } else {
                    log::error!("Couldn't get total_pages, first_page from itf->get_page_counts()");
                }
            } else {
                log::error!("Couldn't get itf from (*this).wizard.borrow()");
            }

            // Store the pointer in a GWL_USERDATA so we can retrieve it later
            SetWindowLongPtrW(hwnd_dlg, GWL_USERDATA, this as *mut MyNewUserWizard as isize);

            // To get the reference back from GWL_USERDATA pointer:
            // `let this: &mut MyNewUserWizard = &mut *(GetWindowLongPtrW(hwnd_dlg, GWL_USERDATA) as *mut MyNewUserWizard);`
            
            // Change the icon of the propsheet
            change_icon(hwnd_dlg, resource_consts::WIZARD_ICON, this.wiz_icon);
            
            // Change the "Create in:" text
            let mut this_is_dumb: Vec<u16> = this.obj_dest_str.encode_utf16().collect();
            this_is_dumb.push(0);
            let this_is_dumb = Box::new(this_is_dumb);
            set_text(hwnd_dlg, resource_consts::WIZARD_CREATEIN, LPCWSTR(this_is_dumb.as_ptr()));
            true.into()
        }
        
        WM_COMMAND => {
            // | Message Source |                wparam (High word) |      wparam (Low word) |                       lparam |
            // |----------------|-----------------------------------|------------------------|------------------------------|
            // |           Menu |                                 0 |        Menu identifier |                            0 |
            // |    Accelerator |                                 1 | Accelerator identifier |                            0 |
            // |        Control | Control-defined notification code |     Control identifier | Handle to the control window |
                
            let low_word = (wparam.0 as u32 & 0xffff) as u16;
            let hi_word  = (wparam.0 as u32 >> 16 ) & 0xffff;
            
            log::debug!("Got WM_COMMAND: low: {}, hi: {}", low_word, hi_word);
            
            match hi_word {
                BN_CLICKED => {
                    match low_word {
                        resource_consts::IN_NEWPASSWORD => {
                            let mut new_text: Vec<u16> = generate_password().encode_utf16().collect();
                            new_text.push(0);
                            let new_text = Box::new(new_text);
                            set_text(hwnd_dlg, resource_consts::OUT_PASSWORD, LPCWSTR(new_text.as_ptr()));
                            return true.into();
                        }
                        _ => {}
                    }
                }
                _ => {}
            }

            DefWindowProcW(hwnd_dlg, message, wparam, lparam);
            true.into()
        }

        WM_DESTROY => {
            log::debug!("WM_DESTROY received by dlgproc");
            true.into()
        }
        

        _ => {
            log::debug!("{:?} received by dlgproc", wm::WindowMessage::from((message & 0xffff) as u16));
            DefWindowProcW(hwnd_dlg, message, wparam, lparam);
            true.into()
        },
    }
}

fn change_icon(hwnd: HWND, control: u16, new_icon: HICON) {
    let icon_control = unsafe { GetDlgItem(hwnd, control.into()) };
    if !(icon_control.0 == -1 || icon_control.0 == 0) {
        unsafe { SendMessageW(icon_control, STM_SETICON, std::mem::transmute::<HICON,WPARAM>(new_icon), LPARAM(0)) };
    }
}

fn set_text(hwnd: HWND, control: u16, text: LPCWSTR) {
    let text_control = unsafe { GetDlgItem(hwnd, control.into()) };
    if !(text_control.0 == -1 || text_control.0 == 0) {
        unsafe { SendMessageW(text_control, WM_SETTEXT, WPARAM(0), std::mem::transmute::<LPCWSTR,LPARAM>(text)) };
    }
}

// TODO: This is so bad please make a better function.
fn generate_password() -> String {
    let mut rng = rand::thread_rng();
    let uppercase_letter = | rng: &mut rand::rngs::ThreadRng | rng.gen_range('A'..='Z') as char;
    let lowercase_letter = | rng: &mut rand::rngs::ThreadRng | rng.gen_range('a'..='z') as char;
    let special_char = |rng: &mut rand::rngs::ThreadRng | match rng.gen_range(0..=5) {
        0 => '!',
        1 => '@',
        2 => '#',
        3 => '$',
        4 => '%',
        _ => '&',
    };
    let digit = | rng: &mut rand::rngs::ThreadRng | rng.gen_range('0'..='9') as char;

    format!(
        "{}{}{}{}{}{}{}{}{}",
        uppercase_letter(&mut rng),
        lowercase_letter(&mut rng),
        lowercase_letter(&mut rng),
        special_char(&mut rng),
        digit(&mut rng),
        digit(&mut rng),
        digit(&mut rng),
        digit(&mut rng),
        digit(&mut rng),
    )
}