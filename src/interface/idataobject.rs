use intercom::{ prelude::*, IUnknown };

use super::{
    IAdviseSink,
    IEnumFORMATETC,
    IEnumSTATDATA,
    types::{ ComFORMATETC, ComSTGMEDIUM }
};

#[com_interface(com_iid = "0000010E-0000-0000-C000-000000000046")]
pub trait IDataObject: IUnknown {
    /// Called by a data consumer to obtain sata from a source data object.
    /// Renders the data in the `FORMATETC` structure and transfers it through
    /// the specified `STGMEDIUM` structure. The caller assumes responsibility
    /// for releasing the `STGMEDIUM` structure.
    fn get_data(&self, format: *const ComFORMATETC) -> ComResult<ComSTGMEDIUM>;
    
    /// Called by a data consumer to obtain data from a source data object. This
    /// method differs from the `get_data` method in that the caller must allocate
    /// and free the specified storage medium.
    fn get_data_here(&self, format: *const ComFORMATETC, medium: *mut ComSTGMEDIUM) -> ComResult<()>;
    
    /// Determines whether the data object is capable of rendering the data as
    /// specified. Objects attempting a paste or drop operation can call this
    /// method before calling `get_data` to get an indication of whether the
    /// operation may be successful.
    fn query_get_data(&self, format: *const ComFORMATETC) -> ComResult<()>;
    
    /// Provides a potentially different but logically equivilent `FORMATETC`
    /// structure. You use this method to determine whether two different
    /// `FORMATETC` structures would return the same data, removing the need
    /// for duplicate rendering.
    fn get_canonical_format(&self, format: *const ComFORMATETC) -> ComResult<ComFORMATETC>;
    
    /// Called by an object containing a data source to transfer data to the
    /// object that implements this method.
    fn set_data(&self, format: *const ComFORMATETC, medium: *const ComSTGMEDIUM, release: bool) -> ComResult<()>;
    
    /// Creates an object to enumerate the formats supported by a data object.
    fn enum_format_etc(&self, direction: u32) -> ComResult<ComRc<dyn IEnumFORMATETC>>;
    
    /// Called by an object supporting an advise sink to create a connection
    /// between a data object and the advise sink. This enables the advise sink
    /// to be notified of changes in the data of the object.
    fn d_advise(&self, format: *const ComFORMATETC, advf: u32, advsink: ComRc<dyn IAdviseSink>) -> ComResult<u32>;
    
    /// Destroys a notification connection that had been previously set up.
    fn d_unadvise(&self, connection: u32) -> ComResult<()>;
    
    /// Creates an object that can be used to enumerate the current advisory
    /// connections.
    fn enum_d_advise(&self) -> ComResult<ComRc<dyn IEnumSTATDATA>>;
}