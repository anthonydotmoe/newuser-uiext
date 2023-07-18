mod interface;
use interface::{
    IADs,
    IADsContainer,
    IDsAdminNewObj,
    IDsAdminNewObjExt,
    
    types::{
        LPCWSTR, ComDsaNewObjCtx, ComLPFNSVADDPROPSHEETPAGE, DsaNewObjDispInfo, ComLPARAM,
    }
};

mod resource_consts;

use intercom::{ prelude::*, BString, Variant, raw::E_INVALIDARG };

use windows::{Win32::{UI::{Controls::{PROPSHEETPAGEW, PROPSHEETPAGEW_0, PROPSHEETPAGEW_1, PROPSHEETPAGEW_2, PSP_DEFAULT, CreatePropertySheetPageW, PSP_USEHICON}, WindowsAndMessaging::{WM_INITDIALOG, GWL_USERDATA, WM_DESTROY, GetWindowLongPtrW, DefWindowProcW}}, Foundation::{HMODULE, LPARAM, HANDLE, HWND, WPARAM} }, core::PCWSTR};
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

#[com_class(clsid = "1af9e4e5-d2fc-41e6-9a7e-3fbdb38a34b4", IDsAdminNewObjExt)]
#[derive(Default)]
struct MyNewUserWizard {
    
    obj_container: std::cell::RefCell<Option<ComRc<dyn IADsContainer>>>,    
    copy_source: std::cell::RefCell<Option<ComRc<dyn IADs>>>,
    new_object: std::cell::RefCell<Option<ComRc<dyn IADs>>>,
    wizard: Option<ComRc<dyn IDsAdminNewObj>>,
    wiz_icon: HICON,

}

/**
# Implementing IDsAdminNewObjExt

1. Initialize()

Initialize supplies the extension with information about the container
about the container the object is being created in, the class name of
the new object and information about the wizard itself.
      
2. AddPages

After the extension is initialized, AddPages is called. The extensions adds
the page or pages to the wizard during this method. Create wizard page by
filling in `PROPSHEETPAGE` structure and passing it to the
`CreatePropertySheetPage` function. The page is then added to the wizard by
calling the callback function that is passed to AddPages in the addpagefn
parameter.

3. SetObject

Before the extension page is displayed, SetObject is called. This supplies
the extension with an IADs interface pointer for the object being created.

While the wizard page is displayed, the page should handle and respond to
any necessary wizard notification messages, such as `PSN_SETACTIVE` and
`PSN_WIZNEXT`

4. GetSummaryInfo

When the user completes all of the wizard pages, the wizard will display a
"Finish" page that provides a summary of the data entered. The wizard
obtains this data by calling the GetSummaryInfo method for each of the
extensions. The GetSummaryInfo method provides a BSTR that contains the text
data displayed in the "Finish" page. An object creation extension does not
have to supply summary data, GetSummaryInfo should return E_NOTIMPL if so.

5. WriteData

When the user clicks the finish button, the wizard calls each of the
extension's WriteData methods with the DSA_NEWOBJ_CTX_PRECOMMIT context.
When this occurs the extension should write the gathered data into the
appropriate properties using the IADs::Put or IADs::PutEx method. The IADs
interface is provided to the extension in the SetObject method. The
extension should not commit the cached properties by calling IADs::SetInfo.
When all the properties are written, the primary object creation extension
commits the changes by calling IADs::SetInfo. hwnd is given to be used as a
parent window for possible error messages.

6. OnError

If an error occurs, the extension will be notified of the error and during
which operation it occured when the OnError method is called. hwnd is given to
be used as a parent window for possible error messages.

# Implementing a Primary Object Creation Wizard

The implementation of a primary object creation wizard is identical to a
secondary object creation wizard, except that a primary object creation
wizard must perform a few more steps.

1. Prior to the first page being dismissed, the object creation wizard must
create the temporary directory object. Call
IDsAdminNewObjPrimarySite::CreateNew. The interface pointer is obtained by
calling QueryInterface on the IDsAdminNewObj interface passed to Initialize.
The CreateNew method creates a new temporary object and calls
IDsAdminNewObjExt::SetObject for each extension.
*/

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
    
    fn add_pages(&self,
        addpagefn: ComLPFNSVADDPROPSHEETPAGE,
        param: ComLPARAM
    ) -> ComResult<()> {
        log::debug!("add_pages called");

        let hinstance = get_dll_hinstance();
        log::debug!("got hInstance: {:?}", &hinstance);
        
        // Very unsafe! Getting a reference to self to use in the windowproc
        let this_clone = self.clone();
        let this_ptr = Box::into_raw(Box::new(this_clone)) as *mut _;
        let this_param = unsafe { std::mem::transmute::<*mut MyNewUserWizard, LPARAM>(this_ptr) };
        
        let pages = vec![resource_consts::DIALOG_1, resource_consts::DIALOG_2];
        for page in pages.iter() {
                
            let mut propsheet: PROPSHEETPAGEW = PROPSHEETPAGEW {
                dwSize: std::mem::size_of::<PROPSHEETPAGEW>() as u32,
                dwFlags: PSP_DEFAULT | PSP_USEHICON,
                hInstance: hinstance,
                Anonymous1: PROPSHEETPAGEW_0 { pszTemplate: make_int_resource(*page) },
                Anonymous2: PROPSHEETPAGEW_1 { pszIcon: PCWSTR::null() },
                pszTitle: PCWSTR::null(),
                pfnDlgProc: Some(dlgproc),
                lParam: this_param,
                pfnCallback: None,
                pcRefParent: std::ptr::null_mut(),
                pszHeaderTitle: PCWSTR::null(),
                pszHeaderSubTitle: PCWSTR::null(),
                hActCtx: HANDLE(0),
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
        match *self.new_object.borrow_mut() {
            None => {
                log::info!("Keeping copy of new_object {:p}", ad_obj);
                *self.new_object.borrow_mut() = Some(ad_obj.to_owned());
                Ok(())
            },
            Some(_) => Ok(())
        }
        
    }

    fn write_data(&self, hwnd: ComLPARAM, ctx: ComDsaNewObjCtx) -> ComResult<()> {
        log::error!("OnError called! {:?}", ctx.0);
        log::error!("WriteData called!");
        Err(ComError { hresult: E_INVALIDARG, error_info: None })
    }

    fn on_error(&self, hwnd: ComLPARAM, hr: intercom::raw::HRESULT, ctx: ComDsaNewObjCtx) -> ComResult<()> {
        log::error!("OnError called! {:?}", ctx.0);
        Err(ComError { hresult: E_INVALIDARG, error_info: None })
    }

    fn get_summary_info(&self) -> ComResult<BString> {
        let result_string: &'static str = "This is the result string";
        
        Ok(BString::from(result_string))
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

unsafe extern "system" fn dlgproc(hwndDlg: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> isize {
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
            SetWindowLongPtrW(hwndDlg, GWL_USERDATA, this as _);
            true.into()
        }
        WM_DESTROY => {
            log::debug!("WM_DESTROY received by dlgproc");
            let this = GetWindowLongPtrW(hwndDlg, GWL_USERDATA) as *mut MyNewUserWizard;
            drop(Box::from_raw(this)); // Re-box the pointer to deallocate it.
            true.into()
        }
        _ => {
            log::debug!("{} received by dlgproc", message);
            DefWindowProcW(hwndDlg, message, wparam, lparam);
            true.into()
        },
    }
}