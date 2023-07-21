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

// mod propsheet;

mod actctx;

mod resource_consts;

use intercom::{ prelude::*, BString, Variant, raw::E_INVALIDARG };

use windows::{Win32::{UI::{Controls::{PROPSHEETPAGEW, PROPSHEETPAGEW_0, PROPSHEETPAGEW_1, PROPSHEETPAGEW_2, PSP_DEFAULT, CreatePropertySheetPageW, PSP_HIDEHEADER, InitCommonControlsEx, INITCOMMONCONTROLSEX, ICC_STANDARD_CLASSES}, WindowsAndMessaging::{WM_INITDIALOG, GWL_USERDATA, WM_DESTROY, GetWindowLongPtrW, DefWindowProcW, InsertMenuItemW, MENUITEMINFOW, MIIM_STRING, MENU_ITEM_STATE, MIIM_ID}}, Foundation::{HMODULE, LPARAM, HANDLE, HWND, WPARAM}, Graphics::Gdi::HBITMAP }, core::PCWSTR, w};
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
}

#[com_class(clsid = "1af9e4e5-d2fc-41e6-9a7e-3fbdb38a34b4", IDsAdminNewObjExt, IShellExtInit, IContextMenu)]
#[derive(Default)]
struct MyNewUserWizard {
    
    obj_container: std::cell::RefCell<Option<ComRc<dyn IADsContainer>>>,    
    copy_source: std::cell::RefCell<Option<ComRc<dyn IADs>>>,
    new_object: std::cell::RefCell<Option<ComRc<dyn IADs>>>,
    wizard: Option<ComRc<dyn IDsAdminNewObj>>,
    wiz_icon: HICON,
    actctx: Option<ReleaseActCtxGuard>,

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
        
        unsafe { log::info!("Display info: {}, {}", (*disp_info).wiz_title, (*disp_info).container_display_name) };
        self.wiz_icon = unsafe { (*disp_info).class_icon };
        
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
        
        match unsafe { InitCommonControlsEx(&sex).into() } {
            true => {
                log::debug!("InitCommonControlsEx returned true");
            }
            false => {
                log::error!("InitCommonControlsEx returned false!");
                return Err(ComError::new_hr(intercom::raw::HRESULT { hr: 0x7FFF0001 }));
            }
        }
        
        // Very unsafe! Getting a reference to self to use in the windowproc
        /*
        let this_clone = 
        let this_ptr = Box::into_raw(Box::new(this_clone)) as *mut _;
        let this_param = unsafe { std::mem::transmute::<*mut MyNewUserWizard, LPARAM>(this_ptr) };
        */
        let this_param = LPARAM(0);
        
        let pages = vec![resource_consts::DIALOG_1, resource_consts::DIALOG_2];
        for page in pages.iter() {
                
            let mut propsheet: PROPSHEETPAGEW = PROPSHEETPAGEW {
                dwSize: std::mem::size_of::<PROPSHEETPAGEW>() as u32,
                dwFlags: PSP_DEFAULT | PSP_HIDEHEADER,
                hInstance: hinstance,
                Anonymous1: PROPSHEETPAGEW_0 { pszTemplate: make_int_resource(*page) },
                Anonymous2: PROPSHEETPAGEW_1 { pszIcon: PCWSTR::null() },
                pszTitle: PCWSTR::null(),
                //pfnDlgProc: Some(dlgproc),
                pfnDlgProc: None,
                lParam: this_param,
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

    fn set_object(&self, ad_obj: &ComItf<dyn IADs>) -> ComResult<()> {
        /*
        match *self.new_object.borrow_mut() {
            None => {
                log::info!("Keeping copy of new_object {:p}", ad_obj);
                *self.new_object.borrow_mut() = Some(ad_obj.to_owned());
                Ok(())
            },
            Some(_) => Ok(())
        }
        */
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
            let this = std::mem::transmute::<LPARAM, *mut MyNewUserWizard>(lparam);
            
            let maybe_itf = &this.as_ref().unwrap().wizard;
            if let Some(itf) = maybe_itf {
                if let Ok((total_pages, first_page)) = itf.get_page_counts() {
                    log::info!("Got total_pages: {}, first_page: {}", &total_pages, &first_page);
                    log::info!("Set next button to disabled on page {}", 0);
                    match itf.set_buttons(0, false) {
                        Ok(_) => {
                            log::info!("Setting next button to disabled worked");
                        }
                        Err(e) => {
                            log::error!("Setting next button to disabled didn't work: {}", e.to_string())
                        }
                    }
                } else {
                    log::error!("Couldn't get total_pages, first_page from itf->get_page_counts()");
                }
            } else {
                log::error!("Couldn't get itf from (*this).wizard.borrow()");
            }

            // Store the pointer in a GWL_USERDATA so we can retrieve it later
            SetWindowLongPtrW(hwnd_dlg, GWL_USERDATA, this as _);
            true.into()
        }
        WM_DESTROY => {
            log::debug!("WM_DESTROY received by dlgproc");
            let this = GetWindowLongPtrW(hwnd_dlg, GWL_USERDATA) as *mut MyNewUserWizard;
            drop(Box::from_raw(this)); // Re-box the pointer to deallocate it.
            true.into()
        }
        _ => {
            log::debug!("{} received by dlgproc", message);
            DefWindowProcW(hwnd_dlg, message, wparam, lparam);
            true.into()
        },
    }
}