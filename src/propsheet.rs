// TODO: some cool macro that generates a resource file/property sheet struct
// for making pages.

use bitflags::bitflags;

bitflags!{
    #[repr(transparent)]    
    #[derive(Debug)]
    struct PropSheetPageFlags: u32 {
        const DEFAULT           = 0x0000_0000;
        const DLGINDIRECT       = 0x0000_0001;
        const USEHICON          = 0x0000_0002;
        const USEICONID         = 0x0000_0004;
        const USETITLE          = 0x0000_0008;
        const RTLREADING        = 0x0000_0010;
        const HASHELP           = 0x0000_0020;
        const USEREFPARENT      = 0x0000_0040;
        const USECALLBACK       = 0x0000_0080;
        const PREMATURE         = 0x0000_0400;
        const HIDEHEADER        = 0x0000_0800;
        const USEHEADERTITLE    = 0x0000_1000;
        const USEHEADERSUBTITLE = 0x0000_2000;
        const USEFUSIONCONTEXT  = 0x0000_4000;
    }
}

struct RawPropSheetPage {

}

pub struct PropSheetPage(RawPropSheetPage);

impl PropSheetPage {
    pub fn new(opts: PropSheetPageOpts) -> Self {
        Self(
            RawPropSheetPage {  }
        )
    }
}