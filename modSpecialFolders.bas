Attribute VB_Name = "modSpecialFolders"
Option Compare Database
Option Explicit

'   ********** Code Start **********
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
'
'   The following table outlines the different DLL versions,
'   and how they were distributed.
'
'   Version     DLL             Distribution Platform
'   4.00          All               Microsoft® Windows® 95/Windows NT® 4.0.
'   4.70          All               Microsoft® Internet Explorer 3.x.
'   4.71          All               Microsoft® Internet Explorer 4.0
'   4.72          All               Microsoft® Internet Explorer 4.01 and Windows® 98
'   5.00          Shlwapi.dll  Microsoft® Internet Explorer 5
'   5.00          Shell32.dll   Microsoft® Windows® 2000.
'   5.80          Comctl32.dll Microsoft® Internet Explorer 5
'   5.81          Comctl32.dll Microsoft® Windows 2000
'
'

'   © Microsoft. Information copied from Microsoft's
'   Platform SDK Documentation in MSDN
'   (http://msdn.microsoft.com)
'
'   If a special folder does not exist, you can force it to be
'   created by using the following special CSIDL:
'   (Version 5.0)
Public Const CSIDL_FLAG_CREATE = &H8000
 
'   Combine this CSIDL with any of the CSIDLs listed below
'   to force the creation of the associated folder.

'   The remaining CSIDLs correspond to either file system or virtual folders.
'   Where the CSIDL identifies a file system folder, a commonly used path
'   is given as an example. Other paths may be used. Some CSIDLs can be
'   mapped to an equivalent %VariableName% environment variable.
'   CSIDLs are much more reliable, however, and should be used if at all possible.

'   File system directory that is used to store administrative tools for an individual user.
'   The Microsoft Management Console will save customized consoles to
'   this directory and it will roam with the user.
'   (Version 5.0)
Public Const CSIDL_ADMINTOOLS = &H30

'   File system directory that corresponds to the user's
'   nonlocalized Startup program group.
Public Const CSIDL_ALTSTARTUP = &H1D

'   File system directory that serves as a common repository for application-specific
'   data. A typical path is C:\Documents and Settings\username\Application Data.
'   This CSIDL is supported by the redistributable ShFolder.dll for systems that do
'   not have the Internet Explorer 4.0 integrated shell installed.
'   (Version 4.71)
Public Const CSIDL_APPDATA = &H1A

'   Virtual folder containing the objects in the user's Recycle Bin.
Public Const CSIDL_BITBUCKET = &HA

'   File system directory containing containing administrative tools
'   for all users of the computer.
'   Version 5
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F

'   File system directory that corresponds to the nonlocalized Startup program
'   group for all users. Valid only for Windows NT® systems.
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E

'   Application data for all users. A typical path is
'   C:\Documents and Settings\All Users\Application Data.
'   Version 5
Public Const CSIDL_COMMON_APPDATA = &H23

'   File system directory that contains files and folders that appear on the
'   desktop for all users. A typical path is C:\Documents and Settings\All Users\Desktop.
'   Valid only for Windows NT® systems.
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19

'   File system directory that contains documents that are common to all users.
'   A typical path is C:\Documents and Settings\All Users\Documents.
'   Valid for Windows NT® systems and Windows 95 and Windows 98
'   systems with Shfolder.dll installed.
Public Const CSIDL_COMMON_DOCUMENTS = &H2E

'   File system directory that serves as a common repository for all users' favorite items.
'   Valid only for Windows NT® systems.
Public Const CSIDL_COMMON_FAVORITES = &H1F

'   File system directory that contains the directories for the common program
'   groups that appear on the Start menu for all users. A typical path is
'   C:\Documents and Settings\All Users\Start Menu\Programs.
'   Valid only for Windows NT® systems.
Public Const CSIDL_COMMON_PROGRAMS = &H17

'   File system directory that contains the programs and folders that appear on
'   the Start menu for all users. A typical path is
'   C:\Documents and Settings\All Users\Start Menu.
'   Valid only for Windows NT® systems.
Public Const CSIDL_COMMON_STARTMENU = &H16

'   File system directory that contains the programs that appear in the
'   Startup folder for all users. A typical path is
'   C:\Documents and Settings\All Users\Start Menu\Programs\Startup.
'   Valid only for Windows NT® systems.
Public Const CSIDL_COMMON_STARTUP = &H18

'   File system directory that contains the templates that are available to all users.
'   A typical path is C:\Documents and Settings\All Users\Templates.
'   Valid only for Windows NT® systems.
Public Const CSIDL_COMMON_TEMPLATES = &H2D

'   Virtual folder containing icons for the Control Panel applications.
Public Const CSIDL_CONTROLS = &H3

'   File system directory that serves as a common repository for Internet cookies.
'   A typical path is C:\Documents and Settings\username\Cookies.
Public Const CSIDL_COOKIES = &H21

'   Windows Desktop—virtual folder that is the root of the namespace..
Public Const CSIDL_DESKTOP = &H0

'   File system directory used to physically store file objects on the desktop
'   (not to be confused with the desktop folder itself).
'   A typical path is C:\Documents and Settings\username\Desktop
Public Const CSIDL_DESKTOPDIRECTORY = &H10

'   My Computer—virtual folder containing everything on the local computer:
'   storage devices, printers, and Control Panel. The folder may
'   also contain mapped network drives.
Public Const CSIDL_DRIVES = &H11

'   File system directory that serves as a common repository for the user's
'   favorite items. A typical path is C:\Documents and Settings\username\Favorites.
Public Const CSIDL_FAVORITES = &H6

'   Virtual folder containing fonts. A typical path is C:\WINNT\Fonts.
Public Const CSIDL_FONTS = &H14

'   File system directory that serves as a common repository for
'   Internet history items.
Public Const CSIDL_HISTORY = &H22

'   Virtual folder representing the Internet.
Public Const CSIDL_INTERNET = &H1

'   File system directory that serves as a common repository for
'   temporary Internet files. A typical path is
'   C:\Documents and Settings\username\Temporary Internet Files.
Public Const CSIDL_INTERNET_CACHE = &H20

'   File system directory that serves as a data repository for local
'   (non-roaming) applications. A typical path is
'   C:\Documents and Settings\username\Local Settings\Application Data.
'   Version 5
Public Const CSIDL_LOCAL_APPDATA = &H1C

'   My Pictures folder. A typical path is
'   C:\Documents and Settings\username\My Documents\My Pictures.
'   Version 5
Public Const CSIDL_MYPICTURES = &H27

'   A file system folder containing the link objects that may exist in the
'   My Network Places virtual folder. It is not the same as CSIDL_NETWORK,
'   which represents the network namespace root. A typical path is
'   C:\Documents and Settings\username\NetHood.
Public Const CSIDL_NETHOOD = &H13

'   Network Neighborhood—virtual folder representing the
'   root of the network namespace hierarchy.
Public Const CSIDL_NETWORK = &H12

'   File system directory that serves as a common repository for documents.
'   A typical path is C:\Documents and Settings\username\My Documents.
Public Const CSIDL_PERSONAL = &H5

'   Virtual folder containing installed printers.
Public Const CSIDL_PRINTERS = &H4

'   File system directory that contains the link objects that may exist in the
'   Printers virtual folder. A typical path is
'   C:\Documents and Settings\username\PrintHood.
Public Const CSIDL_PRINTHOOD = &H1B

'   User's profile folder.
'   Version 5
Public Const CSIDL_PROFILE = &H28

'   Program Files folder. A typical path is C:\Program Files.
'   Version 5
Public Const CSIDL_PROGRAM_FILES = &H2A
 
'   A folder for components that are shared across applications. A typical path
'   is C:\Program Files\Common.
'   Valid only for Windows NT® and Windows® 2000 systems.
'   Version 5
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B

'   Program Files folder that is common to all users for x86 applications
'   on RISC systems. A typical path is C:\Program Files (x86)\Common.
'   Version 5
Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C

'   Program Files folder for x86 applications on RISC systems. Corresponds
'   to the %PROGRAMFILES(X86)% environment variable.
'   A typical path is C:\Program Files (x86).
'   Version 5
Public Const CSIDL_PROGRAM_FILESX86 = &H2A

'   File system directory that contains the user's program groups (which are
'   also file system directories). A typical path is
'   C:\Documents and Settings\username\Start Menu\Programs.
Public Const CSIDL_PROGRAMS = &H2

'   File system directory that contains the user's most recently used documents.
'   A typical path is C:\Documents and Settings\username\Recent.
'   To create a shortcut in this folder, use SHAddToRecentDocs. In addition to
'   creating the shortcut, this function updates the shell's list of recent documents
'   and adds the shortcut to the Documents submenu of the Start menu.
Public Const CSIDL_RECENT = &H8

'   File system directory that contains Send To menu items. A typical path is
'   C:\Documents and Settings\username\SendTo.
Public Const CSIDL_SENDTO = &H9

'   File system directory containing Start menu items.
'   A typical path is C:\Documents and Settings\username\Start Menu.
Public Const CSIDL_STARTMENU = &HB

'   File system directory that corresponds to the user's Startup program group.
'   The system starts these programs whenever any user logs onto Windows NT® or
'   starts Windows® 95. A typical path is
'   C:\Documents and Settings\username\Start Menu\Programs\Startup.
Public Const CSIDL_STARTUP = &H7

'   System folder. A typical path is C:\WINNT\SYSTEM32.
'   Version 5
Public Const CSIDL_SYSTEM = &H25

'   System folder for x86 applications on RISC systems.
'   A typical path is C:\WINNT\SYS32X86.
'   Version 5
Public Const CSIDL_SYSTEMX86 = &H29

'   File system directory that serves as a common repository
'   for document templates.
Public Const CSIDL_TEMPLATES = &H15

'   Version 5.0. Windows directory or SYSROOT. This corresponds to the %windir%
'   or %SYSTEMROOT% environment variables. A typical path is C:\WINNT.
Public Const CSIDL_WINDOWS = &H24

Public Const NOERROR = 0

'   Retrieves a pointer to the ITEMIDLIST structure of a special folder.
Private Declare Function apiSHGetSpecialFolderLocation Lib "shell32" _
    Alias "SHGetSpecialFolderLocation" _
    (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    ppidl As Long) _
    As Long

'   Converts an item identifier list to a file system path.
Private Declare Function apiSHGetPathFromIDList Lib "shell32" _
    Alias "SHGetPathFromIDList" _
    (pidl As Long, _
    ByVal pszPath As String) _
    As Long

'   Frees a block of task memory previously allocated through a call to
'   the CoTaskMemAlloc or CoTaskMemRealloc function.
Private Declare Sub sapiCoTaskMemFree Lib "ole32" _
    Alias "CoTaskMemFree" _
    (ByVal pv As Long)
    
Private Const MAX_PATH = 260

Function fGetSpecialFolderLocation(ByVal lngCSIDL As Long) As String
'
'   Returns path to a special folder on the machine
'   without a trailing backslash.
'
'   Refer to the comments in declarations for OS and
'   IE dependent CSIDL values.
'
Dim lngRet As Long
Dim strLocation As String
Dim pidl As Long

    '   retrieve a PIDL for the specified location
    lngRet = apiSHGetSpecialFolderLocation(hWndAccessApp, lngCSIDL, pidl)
    If lngRet = NOERROR Then
        strLocation = Space$(MAX_PATH)
        '  convert the pidl to a physical path
        lngRet = apiSHGetPathFromIDList(ByVal pidl, strLocation)
        If Not lngRet = 0 Then
            '   if successful, return the location
            fGetSpecialFolderLocation = Left$(strLocation, _
                                InStr(strLocation, vbNullChar) - 1)
        End If
        '   calling application is responsible for freeing the allocated memory
        '   for pidl when calling SHGetSpecialFolderLocation. We have to
        '   call IMalloc::Release, but to get to IMalloc, a tlb is required.
        '
        '   According to Kraig Brockschmidt in Inside OLE,   CoTaskMemAlloc,
        '   CoTaskMemFree, and CoTaskMemRealloc take the same parameters
        '   as the interface functions and internally call CoGetMalloc, the
        '   appropriate IMalloc function, and then IMalloc::Release.
        Call sapiCoTaskMemFree(pidl)
    End If
End Function
'   ********** Code End **********

