Attribute VB_Name = "mod_SelectFolder"
Option Explicit

'   Copyright © 2001 DonkBuilt Software
'   Written by Allen S. Donker
'   All rights reserved.

'***************************************************************
'   Opens a common dialog window to browse for a folder
'   Returns the path to the folder selected as a string
'***************************************************************


'***************************************************************
'   Browse Dialog Constants
'***************************************************************
Public Type BROWSEINFO
    hOwner           As Long         'Handle to window's owner
    pidlRoot         As Long         'Pointer to an item identifier list
    pszDisplayName   As String       'Pointer to a buffer that receives the display name of the folder selected
    lpszTitle        As String       'Pointer to a null-terminated string that is displayed above the tree view control in the dialog box
    ulFlags          As Long         'Value specifying the types of folders to be listed in the dialog box as well as other options
    lpfn             As Long         'Address an application-defined function that the dialog box calls when events occur
    lParam           As Long         'Application-defined value that the dialog box passes to the callback function (if one is specified).
    iImage           As Long         'Variable that receives the image associated with the selected folder. The image is specified as an index to the system image list.
End Type



'***************************************************************
'   Browse Dialog Flags & Constants
'***************************************************************
Public Const BIF_RETURNONLYFSDIRS = &H1         'Only returns file system directories
Public Const BIF_DONTGOBELOWDOMAIN = &H2        'Does not include network folders below the domain level
Public Const BIF_STATUSTEXT = &H4               'Includes a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box.
Public Const BIF_RETURNFSANCESTORS = &H8        'Only returns file system ancestors
Public Const BIF_BROWSEFORCOMPUTER = &H1000     'Only returns computers
Public Const BIF_BROWSEFORPRINTER = &H2000      'Only returns (network) printers

Public Const MAX_PATH = 255


'***************************************************************
'   Browse Dialog API Declarations
'***************************************************************
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                                (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                                (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)


'***************************************************************
'   opens the Browse Folder window and returns the folder selected
'   as a string, or an empty string if canceled
'***************************************************************
Public Function SelectFolder(frm As Form, _
                            Optional sDialTitle As String = "Select a folder") As String

  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer

            'Fill the BROWSEINFO structure with the needed data.
    With bi
            
        .hOwner = frm.hWnd
        .pidlRoot = 0&                      'Root folder to browse from, or desktop if Null
        .lpszTitle = sDialTitle             'Message to display in dialog
        .ulFlags = BIF_RETURNONLYFSDIRS     'the type of folder to return
  
    End With

            'show the browse for folders dialog
    pidl = SHBrowseForFolder(bi)
 
        'the dialog has closed, so parse & display the user's
        'returned folder selection contained in pidl
    path = Space$(MAX_PATH)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
        pos = InStr(path, Chr$(0))
        SelectFolder = Left(path, pos - 1)
    Else
        SelectFolder = ""
    End If

    Call CoTaskMemFree(pidl)

End Function

