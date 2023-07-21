use intercom::{IUnknown, prelude::*, BString};

use super::{
    IADs,
    IADsContainer,
    IDsAdminNewObj,
    types::{
        ComLPFNSVADDPROPSHEETPAGE, ComLPARAM, ComDsaNewObjCtx, LPCWSTR, DsaNewObjDispInfo
    },
};

#[com_interface(com_iid = "6088EAE2-E7BF-11D2-82AF-00C04F68928B")]
pub trait IDsAdminNewObjExt: IUnknown {
    /// Initializes an object creation wizard extension
    /// 
    /// # Inputs
    /// 
    /// `iadscontainer` - Pointer to the IADsContainer interface of an existing
    /// container where the object is created. This parameter must not be
    /// `NULL`. If this object is to be kept beyond the scope of this method,
    /// the reference count must be incremented by calling `IUnknown::AddRef` or
    /// `IUnknown::QueryInterface`.
    /// 
    /// `iads` - Pointer to the `IADs` interface of the object from which a copy
    /// is made. If the new object is not copied from another object, this
    /// parameter is `NULL`. For more information about copy operations, see the
    /// Remarks section. If this object is to be kept beyond the scope of this
    /// method, the reference count must be incremented by calling
    /// `IUnknown::AddRef` or `IUnknown::QueryInterface`.
    /// 
    /// `class_name` - Pointer to a `WCHAR` string containing the LDAP name of
    /// the object class to be created. This parameter must not be `NULL`.
    /// Supported values are "user", "computer", "printQueue", "group", and
    /// "contact".
    /// 
    /// `iadsadminnewobj` - Pointer to an `IDsAdminNewObj` interface that
    /// contains additional data about the wizard. You can also obtain the
    /// `IDsAdminNewObjPrimarySite` interface of the primary extension by
    /// calling `QueryInterface` with `IID_IDsAdminNewObjPrimarySite` on this
    /// interface. If this object is to be kept beyond the scope of this method,
    /// the reference count must be incremented by calling `IUnkown::AddRef` or
    /// `IUnknown::QueryInterface`.
    /// 
    /// `disp_info` - Pointer to a `DSA_NEWOBJ_DISPINFO` structure that contains
    /// additional data about the object creation wizard.
    /// 
    /// # Returns
    /// 
    /// Returns `S_OK` if successful or an OLE-defined error code otherwise.
    /// 
    /// # Remarks
    /// 
    /// An object in Active Directory Domain Services can either be created from
    /// nothing or copied from an existing object. If the new object is created
    /// from an existing object, *iads* will contain a pointer to the object
    /// from which the copy is made. If the new object is not being copied from
    /// another object, *iads* will be `NULL`. The copy operation if only
    /// supported for user objects.
    fn initialize(
        &mut self,
        iadscontainer: &ComItf<dyn IADsContainer>,
        iads: Option<&ComItf<dyn IADs>>,
        class_name: LPCWSTR,
        adminnewobj: &ComItf<dyn IDsAdminNewObj>,
        disp_info: *const DsaNewObjDispInfo
    ) -> ComResult<()>;
    
    /// Called to enable the object creation wizard extension to add the desired
    /// pages to the wizard.
    /// 
    /// # Inputs
    /// 
    /// `addpagefn` - Pointer to a function that the object creation wizard
    /// extension calls to add a page to the wizard. This function takes the
    /// following format.
    /// 
    /// ```cpp
    /// BOOL fnAddPage(HPROPSHEETPAGE hPage, LPARAM lParam);
    /// ```
    /// 
    /// *hPage* contains the handle of the wizard page created by calling
    /// `CreatePropertySheetPage`.
    /// 
    /// *lParam* is the *param* value passed to `add_pages`.
    /// 
    /// `param` - Contains data that is private to the administrative snap-in.
    /// This value is passed as the second parameter to `addpagefn`.
    /// 
    /// # Returns
    /// 
    /// Returns `S_OK` if successful or an OLE-defined error code otherwise.
    /// 
    /// # Remarks
    /// 
    /// For each page, the wizard extension adds to the wizard, the extension
    /// fills in a `PROPSHEETPAGE` structure, calls the
    /// `CreatePropertySheetPage` function to create the page handle and then
    /// calls the `addpagefn` function with the page handle and `param`.
    /// 
    /// This method is identical in format and operation to the
    /// `IShellPropSheetExt::AddPages` method.
    fn add_pages(
        &mut self,
        addpagefn: ComLPFNSVADDPROPSHEETPAGE,
        param: ComLPARAM
    ) -> ComResult<()>;
    
    /// Provides the object creation extension with a pointer to the created
    /// directory object
    /// 
    /// # Inputs
    /// 
    /// `ad_obj` - Pointer to an IADs interface for the object. This parameter
    /// may be `NULL`. If this object is to be kept beyond the scope of this
    /// method, the reference count must be incremented by calling
    /// `IUnknown::AddRef` or `IUnknown::QueryInterface`.
    /// 
    /// # Returns
    /// 
    /// The method should always return S_OK
    fn set_object(&self, ad_obj: &ComItf<dyn IADs>) -> ComResult<()>;
    
    /// Called to enable the object creation wizard extension to write its data
    /// into an object in Active Directory Domain Services.
    /// 
    /// # Inputs
    /// 
    /// `hWnd` - The window handle used as the parent window for possible error
    /// messages.
    /// 
    /// `uContext` - Specifies the context in which WriteData is called. This
    /// will be one of the following values:
    /// 
    /// - **DSA_NEW_OBJ_CTX_PRECOMMIT**
    /// 
    /// WriteData is called prior to the new object committed to persistent
    /// storage. This is the context during which a secondary object creation
    /// extension should write its data.
    /// 
    /// - **DSA_NEWOBJ_CTX_POSTCOMMIT
    /// 
    /// WriteData is called after the new object has been committed to
    /// persistent storage.
    /// 
    /// - **DSA_NEWOBJ_CTX_CLEANUP
    /// 
    /// There has been a failure during the write process of the temporary
    /// object and the temporary object is recreated.
    /// 
    /// # Returns
    /// 
    /// Returns `S_OK` if successful or an OLE-defined error code otherwise.
    fn write_data(&self, hwnd: ComLPARAM, ctx: ComDsaNewObjCtx) -> ComResult<()>;
    
    /// Called when an error has occurred in the wizard pages.
    /// 
    /// # Inputs
    /// 
    /// `hWnd` - The window handle used as the parent window for possible error
    /// messages.
    /// 
    /// `uContext` - Specifies the context in which OnError is called. This will
    /// be one of the following values:
    /// 
    /// - **DSA_NEW_OBJ_CTX_PRECOMMIT**
    /// 
    /// An error occurred prior to the new object committed to persistent
    /// storage.
    /// 
    /// - **DSA_NEWOBJ_CTX_COMMIT**
    /// 
    /// An error occurred while the new object was committed to persistent
    /// storage.
    ///
    /// - **DSA_NEWOBJ_CTX_POSTCOMMIT**
    /// 
    /// An error occurred after the new object was committed to persistent
    /// storage.
    /// 
    /// - **DSA_NEWOBJ_CTX_CLEANUP**
    /// 
    /// An error occurred while the new object was committed to persistent
    /// storage.
    /// 
    /// # Returns
    /// 
    /// A primary creation extension returns `S_OK` to indicate that the error
    /// was handled by the extension or an OLE-defined error code to cause the
    /// system to display an error message. The return value is ignored for a
    /// secondary creation extension.
    fn on_error(&self, hwnd: ComLPARAM, hr: intercom::raw::HRESULT, ctx: ComDsaNewObjCtx) -> ComResult<()>;
    
    /// Called to obtain a string that contains a summary of the data gathered
    /// by the new object wizard extension page. This string is displayed in the
    /// wizard Finish page.
    /// 
    /// # Returns
    /// 
    /// `BString` - Pointer to a BSTR value that receives the summary text.
    fn get_summary_info(&self) -> ComResult<BString>;
}