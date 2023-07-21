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