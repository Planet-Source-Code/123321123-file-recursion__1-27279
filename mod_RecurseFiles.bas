Attribute VB_Name = "mod_RecurseFiles"
Option Explicit

'   Copyright © 2001 DonkBuilt Software
'   Written by Allen S. Donker
'   All rights reserved.

'************************************************************
'   Takes a valid path to a folder and, using recursion,
'   fills a listbox with the complete path and filename
'   of each file in the folder passed in
'************************************************************


'************************************************************
'   Folder/File attributes constants
'************************************************************
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Const INVALID_HANDLE_VALUE = -1
Const ERROR_NO_MORE_FILES = 18&
Const MAX_PATH = 255


'************************************************************
'   Struct to hold file data
'************************************************************
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


'************************************************************
'   Struct to hold data returned from API calls
'************************************************************
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


'************************************************************
'   Struct to hold data returned from API calls
'************************************************************
Private gFileData As WIN32_FIND_DATA


'************************************************************
'   Windows API declarations
'************************************************************
Private Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" _
                                            (ByVal lpFileName As String, _
                                            lpFindFileData As WIN32_FIND_DATA)
Private Declare Function FindNextFile& Lib "kernel32" Alias "FindNextFileA" _
                                            (ByVal hFindFile As Long, _
                                            lpFindFileData As WIN32_FIND_DATA)
Private Declare Function FindClose& Lib "kernel32" (ByVal hFindFile As Long)





Public Sub RecurseFiles(ByVal sFolderPath As String, objListbox As ListBox)

Dim hFindFile As Long
Dim ReturnValue As Long
Dim Filename As String
  
            'Get First Directory Entry. File value returned will end in
            'a null-terminated string.
    hFindFile = FindFirstFile(sFolderPath, gFileData)
  
            'Exit if there was an Error Getting First Entry
    If hFindFile = INVALID_HANDLE_VALUE Then
        FindClose (hFindFile)
        Exit Sub
    End If
  
            'Initialize ReturnValue
    ReturnValue = 1
  
    Do While ReturnValue <> 0
    
                'Remove the null charactor from Returned Filename
    Filename = StripNulls(gFileData.cFileName)
  
                ' If it is a Directory but NOT the "." or ".." directories
                ' get all the Files in that directory (starts the recursion)
    If Filename <> "." And Filename <> ".." And _
                    gFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then

        RecurseFiles Mid(sFolderPath, 1, Len(sFolderPath) - 3) & Filename & "\*.*", objListbox


    Else        'Keep getting files in current psStartingPath
        
        If Filename <> "." And Filename <> ".." And Filename <> "" And Not _
                    gFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
          
            
                    'Now that we have the file, do something with it,
                    'in this case, add it to a listbox
            objListbox.AddItem Mid(sFolderPath, 1, Len(sFolderPath) - 3) & Filename
            
            DoEvents
        
        End If
      
    End If
      
                ' Get Next Entry
    ReturnValue = FindNextFile(hFindFile, gFileData)
    
    If ReturnValue = 0 Then

        Filename = ""
        Exit Do

    End If
            
    DoEvents
      
  Loop

                                ' Close Handle
  FindClose (hFindFile)
  
End Sub


Private Function StripNulls(ByVal FileWithNulls As String) As String

  Dim NullPos As Integer
  
  NullPos = InStr(1, FileWithNulls, vbNullChar, 0)
  
  If NullPos <> 0 Then
    
        StripNulls = Left(FileWithNulls, NullPos - 1)
  
  End If

End Function



