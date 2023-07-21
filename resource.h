#define IDD_FORMVIEW                    101
#define IDC_STATIC                      -1

#define DS_SETFONT                      0x40L
#define DS_FIXEDSYS                     0x0008L
#define DS_SHELLFONT                    (DS_SETFONT | DS_FIXEDSYS)

#define BS_PUSHBUTTON       0x00000000L
#define BS_DEFPUSHBUTTON    0x00000001L
#define BS_CHECKBOX         0x00000002L
#define BS_AUTOCHECKBOX     0x00000003L
#define BS_RADIOBUTTON      0x00000004L
#define BS_3STATE           0x00000005L
#define BS_AUTO3STATE       0x00000006L
#define BS_GROUPBOX         0x00000007L
#define BS_USERBUTTON       0x00000008L
#define BS_AUTORADIOBUTTON  0x00000009L
#define BS_PUSHBOX          0x0000000AL
#define BS_OWNERDRAW        0x0000000BL
#define BS_TYPEMASK         0x0000000FL
#define BS_LEFTTEXT         0x00000020L
#define BS_TEXT             0x00000000L
#define BS_ICON             0x00000040L
#define BS_BITMAP           0x00000080L
#define BS_LEFT             0x00000100L
#define BS_RIGHT            0x00000200L
#define BS_CENTER           0x00000300L
#define BS_TOP              0x00000400L
#define BS_BOTTOM           0x00000800L
#define BS_VCENTER          0x00000C00L
#define BS_PUSHLIKE         0x00001000L
#define BS_MULTILINE        0x00002000L
#define BS_NOTIFY           0x00004000L
#define BS_FLAT             0x00008000L
#define BS_RIGHTBUTTON      BS_LEFTTEXT

#define CBS_SIMPLE            0x0001L
#define CBS_DROPDOWN          0x0002L
#define CBS_DROPDOWNLIST      0x0003L
#define CBS_OWNERDRAWFIXED    0x0010L
#define CBS_OWNERDRAWVARIABLE 0x0020L
#define CBS_AUTOHSCROLL       0x0040L
#define CBS_OEMCONVERT        0x0080L
#define CBS_SORT              0x0100L
#define CBS_HASSTRINGS        0x0200L
#define CBS_NOINTEGRALHEIGHT  0x0400L
#define CBS_DISABLENOSCROLL   0x0800L
#define CBS_UPPERCASE         0x2000L
#define CBS_LOWERCASE         0x4000L

#define ES_LEFT             0x0000L
#define ES_CENTER           0x0001L
#define ES_RIGHT            0x0002L
#define ES_MULTILINE        0x0004L
#define ES_UPPERCASE        0x0008L
#define ES_LOWERCASE        0x0010L
#define ES_PASSWORD         0x0020L
#define ES_AUTOVSCROLL      0x0040L
#define ES_AUTOHSCROLL      0x0080L
#define ES_NOHIDESEL        0x0100L
#define ES_OEMCONVERT       0x0400L
#define ES_READONLY         0x0800L
#define ES_WANTRETURN       0x1000L
#define ES_NUMBER           0x2000L

#define SS_LEFT             0x00000000L
#define SS_CENTER           0x00000001L
#define SS_RIGHT            0x00000002L
#define SS_ICON             0x00000003L
#define SS_BLACKRECT        0x00000004L
#define SS_GRAYRECT         0x00000005L
#define SS_WHITERECT        0x00000006L
#define SS_BLACKFRAME       0x00000007L
#define SS_GRAYFRAME        0x00000008L
#define SS_WHITEFRAME       0x00000009L
#define SS_USERITEM         0x0000000AL
#define SS_SIMPLE           0x0000000BL
#define SS_LEFTNOWORDWRAP   0x0000000CL
#define SS_OWNERDRAW        0x0000000DL
#define SS_BITMAP           0x0000000EL
#define SS_ENHMETAFILE      0x0000000FL
#define SS_ETCHEDHORZ       0x00000010L
#define SS_ETCHEDVERT       0x00000011L
#define SS_ETCHEDFRAME      0x00000012L
#define SS_TYPEMASK         0x0000001FL
#define SS_REALSIZECONTROL  0x00000040L
#define SS_NOPREFIX         0x00000080L /* Don't do "&" character translation */
#define SS_NOTIFY           0x00000100L
#define SS_CENTERIMAGE      0x00000200L
#define SS_RIGHTJUST        0x00000400L
#define SS_REALSIZEIMAGE    0x00000800L
#define SS_SUNKEN           0x00001000L
#define SS_EDITCONTROL      0x00002000L
#define SS_ENDELLIPSIS      0x00004000L
#define SS_PATHELLIPSIS     0x00008000L
#define SS_WORDELLIPSIS     0x0000C000L
#define SS_ELLIPSISMASK     0x0000C000L

#define WS_BORDER	0x00800000L
#define WS_CAPTION	0x00C00000L
#define WS_CHILD	0x40000000L
#define WS_CHILDWINDOW	0x40000000L
#define WS_CLIPCHILDREN	0x02000000L
#define WS_CLIPSIBLINGS	0x04000000L
#define WS_DISABLED	0x08000000L
#define WS_DLGFRAME	0x00400000L
#define WS_GROUP	0x00020000L
#define WS_HSCROLL	0x00100000L
#define WS_ICONIC	0x20000000L
#define WS_MAXIMIZE	0x01000000L
#define WS_MAXIMIZEBOX	0x00010000L
#define WS_MINIMIZE	0x20000000L
#define WS_MINIMIZEBOX	0x00020000L
#define WS_OVERLAPPED	0x00000000L
#define WS_SYSMENU	0x00080000L
#define WS_POPUP	0x80000000L
#define WS_SIZEBOX	0x00040000L
#define WS_TABSTOP	0x00010000L
#define WS_THICKFRAME	0x00040000L
#define WS_TILED	0x00000000L
#define WS_VISIBLE	0x10000000L
#define WS_VSCROLL	0x00200000L
#define WS_POPUPWINDOW	(WS_POPUP | WS_BORDER | WS_SYSMENU)
#define WS_OVERLAPPEDWINDOW	(WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX)
#define WS_TILEDWINDOW	(WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX)

#define LANG_NEUTRAL                    0x00
#define SUBLANG_NEUTRAL                 0x00

#define RT_MANIFEST                         24

// RUST

#define ISOLATIONAWARE_MANIFEST_RESOURCE_ID  2

#define ICON_1                          1
#define DIALOG_1                        1001
#define DIALOG_2                        1002

#define IN_FIRSTNAME                    201
#define IN_LASTNAME                     202
#define IN_USERNAME                     203
#define IN_JOBTITLE                     204
#define IN_PASSWORD                     205
#define IN_W2KUSERNAME                  206
#define IN_UPNSUFFIX                    207
#define IN_NEWPASSWORD                  208
#define IN_ACCTDISABLED                 209

#define OUT_CREATEIN                    301
#define OUT_NETBIOSDC                   302
#define OUT_PASSWORD                    303
#define OUT_MUSTCHANGEPASS              304
