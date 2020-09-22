VERSION 5.00
Begin VB.Form frm_DirRecurse 
   Caption         =   "Directory Recursing Test"
   ClientHeight    =   5565
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5565
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   9480
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecurse 
      Caption         =   "List Files"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstFileList 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frm_DirRecurse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Copyright © 2001 DonkBuilt Software
'   Written by Allen S. Donker
'   All rights reserved.

'***************************************************************
'   Demonstrates the use of mod_SelectFolder which
'   uses the windows API's to open a browse folder
'   dialog box and mod_RecurseFolders which uses
'   recursion to browse thru the folder to
'   list the files contained in the folder selected.
'***************************************************************

Dim colSelectedPaths As New Collection


Private Sub cmdRecurse_Click()
On Error GoTo ErrH

Dim sPath As String

    MousePointer = vbHourglass

            '   Open the Browse Folder dialog box and
            '   return the folder selected
    sPath = SelectFolder(Me, "Select folder")
  
  
            '   If no folder was selected, exit here
    If Len(sPath) = 0 Then Exit Sub
  
            
            '   Make sure the path ends with a \ and
            '   add *.* to retreive all file types
    If Right$(sPath, 1) <> "\" Then
        sPath = sPath & "\*.*"
    Else
        sPath = sPath & "*.*"
    End If


            '   Call NewPath to either 1)add the path
            '   to the collection of folders selected
            '   and then get the files within that folder
            '   using recursion or 2)return False if the
            '   folder has already been selected and
            '   don't do anything
    If NewPath(sPath) Then
        RecurseFiles sPath, lstFileList
    End If


    MousePointer = vbDefault

Exit Sub
    
ErrH:
    MousePointer = vbDefault
    MsgBox Err.Number & Chr(10) & Err.Description
End Sub



'************************************************************
'   To prevent the same files being listed more than once,
'   put each path (folder) selected into a collection using
'   the path as both the item and the key. If a folder is
'   selected a second time, trying to add it to the
'   collection will raise an error, so this returns False
'************************************************************
Private Function NewPath(sNewPath As String) As Boolean
On Error GoTo ErrH

    colSelectedPaths.Add sNewPath, sNewPath
    NewPath = True

Exit Function
ErrH:
    NewPath = False
End Function


'************************************************************
'   Clear list and collection to start over
'************************************************************
Private Sub cmdClear_Click()
On Error GoTo ErrH

    Set colSelectedPaths = Nothing
    Set colSelectedPaths = New Collection
    lstFileList.Clear
    
Exit Sub
ErrH:
    Resume Next
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub
