use intercom::{ IUnknown, prelude::*, BString };

mod interface;
use interface::*;
use windows::{Win32::{UI::Controls::{PROPSHEETPAGEW, PROPSHEETPAGEW_0, PROPSHEETPAGEW_1, PROPSHEETPAGEW_2, PSP_DEFAULT, CreatePropertySheetPageW}, Foundation::{HMODULE, LPARAM, HANDLE} }, core::PCWSTR, w};
use windows::Win32::System::LibraryLoader::{GetModuleHandleExW, GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT};

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

}

/// # Implementing IDsAdminNewObjExt
/// 
/// 1. Initialize()
///
/// Initialize supplies the extension with information about the container
/// about the container the object is being created in, the class name of
/// the new object and information about the wizard itself.
///       
/// 2. AddPages
/// 
/// After the extension is initialized, AddPages is called. The extensions adds
/// the page or pages to the wizard during this method. Create wizard page by
/// filling in `PROPSHEETPAGE` structure and passing it to the
/// `CreatePropertySheetPage` function. The page is then added to the wizard by
/// calling the callback function that is passed to AddPages in the addpagefn
/// parameter.
/// 
/// 3. SetObject
/// 
/// Before the extension page is displayed, SetObject is called. This supplies
/// the extension with an IADs interface pointer for the object being created.
/// 
/// While the wizard page is displayed, the page should handle and respond to
/// any necessary wizard notification messages, such as `PSN_SETACTIVE` and
/// `PSN_WIZNEXT`
/// 
/// 4. GetSummaryInfo
/// 
/// When the user completes all of the wizard pages, the wizard will display a
/// "Finish" page that provides a summary of the data entered. The wizard
/// obtains this data by calling the GetSummaryInfo method for each of the
/// extensions. The GetSummaryInfo method provides a BSTR that contains the text
/// data displayed in the "Finish" page. An object creation extension does not
/// have to supply summary data. GetSummaryInfo should return E_NOTIMPL.
/// 
/// 5. WriteData
/// 
/// When the user clicks the finish button, the wizard calls each of the
/// extension's WriteData methods with the DSA_NEWOBJ_CTX_PRECOMMIT context.
/// When this occurs the extension should write the gathered data into the
/// appropriate properties using the IADs::Put or IADs::PutEx method. The IADs
/// interface is provided to the extension in the SetObject method. The
/// extension should not commit the cached properties by calling IADs::SetInfo.
/// When all the properties are written, the primary object creation extension
/// commits the changes by calling IADs::SetInfo.
///
/// 6. OnError
/// 
/// If an error occurs, the extension will be notified of the error and during
/// which operation it occured when the OnError method is called.
/// 
/// # Implementing a Primary Object Creation Wizard
/// 
/// The implementation of a primary object creation wizard is identical to a
/// secondary object creation wizard, except that a primary object creation
/// wizard must perform a few more steps.
/// 
/// 1. Prior to the first page being dismissed, the object creation wizard must
/// create the temporary directory object. Call
/// IDsAdminNewObjPrimarySite::CreateNew. The interface pointer is obtained by
/// calling QueryInterface on the IDsAdminNewObj interface passed to Initialize.
/// The CreateNew method creates a new temporary object and calls
/// IDsAdminNewObjExt::SetObject for each extension.

impl IDsAdminNewObjExt for MyNewUserWizard {
    fn initialize(&self, iadscontainer: &ComItf<dyn IADsContainer>, iads: &ComItf<dyn IADs>, class_name: ComPCWSTR, adminnewobj: &ComItf<dyn IDsAdminNewObj>, disp_info: usize) -> ComResult<()> {
        log::info!("Initialize called.");
        
        iadscontainer.as_raw_iunknown().add_ref();
        iads.as_raw_iunknown().add_ref();
        adminnewobj.as_raw_iunknown().add_ref();
        //log::info!("Got {:?}", (iadscontainer, iads, class_name.0, adminnewobj, disp_info));
        Ok(())
    }
    
    fn add_pages(&self,addpagefn: *const ComLPFNSVADDPROPSHEETPAGE,param:ComLPARAM) -> ComResult<()> {

        let hinstance = get_dll_hinstance();
        
        let mut page1: PROPSHEETPAGEW = PROPSHEETPAGEW {
            dwSize: std::mem::size_of::<PROPSHEETPAGEW>() as u32,
            dwFlags: PSP_DEFAULT,
            hInstance: hinstance,
            Anonymous1: PROPSHEETPAGEW_0 { pszTemplate: make_int_resource(1) },
            Anonymous2: PROPSHEETPAGEW_1 { pszIcon: make_int_resource(101) },
            pszTitle: w!("Title"),
            pfnDlgProc: None,
            lParam: LPARAM(0),
            pfnCallback: None,
            pcRefParent: std::ptr::null_mut(),
            pszHeaderTitle: w!("HeaderTitle"),
            pszHeaderSubTitle: w!("HeaderSubTitle"),
            hActCtx: HANDLE(0),
            Anonymous3: PROPSHEETPAGEW_2 { pszbmHeader: make_int_resource(1001) },
        };
        
        let page1_h = unsafe { CreatePropertySheetPageW(&mut page1) };
        
        unsafe {
            match (*addpagefn).0 {
                Some(function) => {
                    match function(page1_h, param.0).as_bool() {
                        true => {
                            log::info!("addpagefn returned true");
                        }
                        false => {
                            log::warn!("addpagefn returned false");
                        }

                    }
                }
                _ => {
                    log::error!("How could this happen.");
                    panic!()
                }
            }
        }
        
        Ok(())
    }

    fn set_object(&self,ad_obj: &ComItf<dyn IUnknown>) -> ComResult<()> {
        log::error!("SetObject called!");
        todo!()
    }

    fn write_data(&self,) -> ComResult<()> {
        log::error!("WriteData called!");
        todo!()
    }

    fn on_error(&self,) -> ComResult<()> {
        log::error!("OnError called!");
        todo!()
    }

    fn get_summary_info(&self,) -> ComResult<BString> {
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