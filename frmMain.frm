VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Explorer:"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   12690
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2040
      ScaleHeight     =   8
      ScaleMode       =   0  'User
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ListView FileList 
      Height          =   8175
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   14420
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Attributes"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ext"
         Object.Width           =   1323
      EndProperty
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059A
            Key             =   "cldfolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06F4
            Key             =   "cldfoldera"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":084E
            Key             =   "opnfoldera"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A8
            Key             =   "drvcd"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F42
            Key             =   "opnfolder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":109C
            Key             =   "drvremove"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1636
            Key             =   "drvfixed"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BD0
            Key             =   "drvremote"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":216A
            Key             =   "mycomputer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2704
            Key             =   "drvunknown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C9E
            Key             =   "drvmemory"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Explorer 
      Height          =   8145
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   14367
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8280
      Width           =   12615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private lIndexCounter As Long
Private sFileWithIcon() As String * 1

'***********************************************************************************
'Form Events.
'***********************************************************************************
'Form_Load.
Private Sub Form_Load()
    
    Dim sComputerName As String * 255
    Dim lAPIReturn As Long
    Dim cDrives As cDrives
        Set cDrives = New cDrives
    
    'Use API to retrieve ComputerName, strip NullChar and Hold ComputerName in.-----
    'Private Variable.
    lAPIReturn& = GetComputerName(sComputerName$, Len(sComputerName$))
        
    mVariables.sComputerName = mProcFunc.ftnStripNullChar(sComputerName$)
    '-------------------------------------------------------------------------------
    
    'Iterate all Local Drives into Explorer Tree.-----------------------------------
    Call cDrives.subLoadTreeView(Explorer)
    '-------------------------------------------------------------------------------
    
    'Expand first Node in Explorer Tree.--------------------------------------------
    Explorer.Nodes(1).Expanded = True
    '-------------------------------------------------------------------------------
    
    'Set timer going.---------------------------------------------------------------
    Call SetTimer(Me.hwnd, 0, 10, AddressOf TIMERPROC)
    '-------------------------------------------------------------------------------
    
End Sub

'Form_Resize.
Private Sub Form_Resize()

    'Explorer.
    With Explorer
        .Height = Me.Height - 700
        .Width = Me.Width / 2
    End With
    
    'FileList.
    With FileList
        .Left = Explorer.Width + 50
        .Height = Explorer.Height + 20
        .Width = Me.Width - Explorer.Width - 180
    End With
    
    'Holds the value of how many Files are visible in frmMain.FileList.-------------
    mVariables.VisibleFileCount = SendMessage(FileList.hwnd, LVM_GETCOUNTPERPAGE, 0&, 0&) ' -1
    '-------------------------------------------------------------------------------

End Sub

'Form_Unload.
Private Sub Form_Unload(Cancel As Integer)
    KillTimer Me.hwnd, 0&
    Unload Me
    Set frmMain = Nothing
End Sub


'***********************************************************************************
'Explorer Events.
'***********************************************************************************
'Explorer_Expand.
Private Sub Explorer_Expand(ByVal Node As MSComctlLib.Node)

    DoEvents
    Dim x As Long
    
    'Branch all sub Folders.--------------------------------------------------------
    Me.MousePointer = 11
    For x = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
        Explorer_NodeClick Explorer.Nodes(x)
    Next x
    Me.MousePointer = 0
    '-------------------------------------------------------------------------------
    
End Sub

'Explorer_NodeClick.
Private Sub Explorer_NodeClick(ByVal Node As MSComctlLib.Node)
   
    Dim sNodePath As String                 'ExplorerTree selectedpath.
        sNodePath$ = mProcFunc.ftnReturnNodePath(Node.FullPath)
    Dim liItem As ListImage                 'ListImage.

    'If Not Children list Folders.--------------------------------------------------
    If Not Node.Children > 0 Then
        mExplorerTree.subShowFolderList List1, Explorer, sNodePath$, Node.Index
    End If
    '-------------------------------------------------------------------------------
    
    'List Files if Node is selected.------------------------------------------------
    If Node.Selected = True And Node.Index > 1 Then
        
        With FileList

            'Kill the timer, not needed when were loading Files.--------------------
            KillTimer .hwnd, 0&
            '-----------------------------------------------------------------------
            
            .Visible = False
            .ListItems.Clear
                
            'Keep a record of selected path.----------------------------------------
            mVariables.sExplorerTreePath = sNodePath$
            '-----------------------------------------------------------------------
            
            'This stops the error, An ImageList cannot be modified etc.-------------
            '.Icons = Nothing
            .SmallIcons = Nothing
            ImageList2.ListImages.Clear
            Set liItem = ImageList2.ListImages.Add(1, , Pic1.Image)
            '-----------------------------------------------------------------------
                
            'List the Files in the selected path.-----------------------------------
            Call mFileList.subFileList(FileList, sNodePath$)
            '-----------------------------------------------------------------------
        
            'Display selected path in caption.--------------------------------------
            Label1.Caption = sNodePath$
            '-----------------------------------------------------------------------
                
            'Retrieve the first visible File top index in the FileList.-------------
            lSendMsgRet(0) = 0 ' SendMessage(.hwnd, LVM_GETTOPINDEX, 0&, 0&)
            '-----------------------------------------------------------------------
    
            'Holds the value of how many Files are visible in frmMain.FileList.-------------
            mVariables.VisibleFileCount = SendMessage(FileList.hwnd, LVM_GETCOUNTPERPAGE, 0&, 0&) '- 1
            '-------------------------------------------------------------------------------
            
            'Keep a track of how many Files are to have icons added. This will stop-
            'an index error when adding icons to each File.
            If .ListItems.Count > 41 Then
                mVariables.lListItemCount = mVariables.VisibleFileCount
            Else
                mVariables.lListItemCount = .ListItems.Count
            End If
            '-----------------------------------------------------------------------

            'This variable is used instead of looping through FileList.ListItems.---
            ReDim sFileWithIcon(.ListItems.Count)
            sFileWithIcon(0) = "I"
            '-----------------------------------------------------------------------

            'Keep a track how how many Files are in FileList and how many Files have
            'had icons added to them.
            lListItemCounter&(1) = .ListItems.Count
            lListItemCounter&(0) = 0
            '-----------------------------------------------------------------------
            
            'Add icons to each visible File in the FileList.------------------------
            Call subAddIcons(sNodePath$)
            '-----------------------------------------------------------------------
    
            .Visible = True
    
            'Set timer going again after listing the Files.-------------------------
            Call SetTimer(Me.hwnd, 0, 10, AddressOf TIMERPROC)
            '-----------------------------------------------------------------------
    
        End With
    
    End If
    '-------------------------------------------------------------------------------

End Sub


'***********************************************************************************
'FileList Events.
'***********************************************************************************
'FileList_DblClick.
Private Sub FileList_DblClick()
                                                                                                                                                                                        
    'Shell associated application. eg Dbl_Click on .txt file executes NotePad.------
    If FileList.ListItems.Count > 0 Then
        Call mProcFunc.subShellApplication(mProcFunc.ftnReturnNodePath(Explorer.SelectedItem.FullPath), FileList.SelectedItem.Text, Me.hwnd)
    End If
    '-------------------------------------------------------------------------------

End Sub



'***********************************************************************************
'Adds icons to each visible File in the FileList.
'***********************************************************************************
Public Sub subAddIcons(sPath As String)
   
    Dim lSysImHndle As Long                     'System image handle.
    Dim lRet As Long                            'Return value for ImageList_Draw.
    Dim liItem As ListImage                     'ListImage object.
    Dim SHI As SHFILEINFO                       'SHFILEINFO structure.
        
    With FileList
        
        .Visible = False

        For lIndexCounter = lSendMsgRet&(0) + 1 To lSendMsgRet&(0) + mVariables.lListItemCount&

            If sFileWithIcon(lIndexCounter) <> "I" Then

                'I`m using an array rather than checking if the FileList.ListItems(-
                'lIndexCounter).Icon = Nothing. Faster.
                sFileWithIcon(lIndexCounter) = "I"
                '-------------------------------------------------------------------
                
                'Keep a counter of how many Files have been displayed in current----
                'Folder.
                lListItemCounter&(0) = lListItemCounter&(0) + 1
                '-------------------------------------------------------------------
        
                'Retrieve the handle to the system image list containing the the small
                'icon images.
                lSysImHndle& = SHGetFileInfo(sPath$ & .ListItems(lIndexCounter).Text, 0&, SHI, Len(SHI), SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
                '-------------------------------------------------------------------

'                Set Pic1.Picture = LoadPicture("")
       
                'Draw the image in the specified device context.--------------------
                lRet = ImageList_Draw(lSysImHndle, SHI.iIcon, Pic1.hDC, 0, 0, ILD_NORMAL)
                '-------------------------------------------------------------------
       
                'Add image to ImageList2, then add image to FileList.---------------
                With ImageList2
                    
                    Set liItem = .ListImages.Add(.ListImages.Count + 1, FileList.ListItems(lIndexCounter).Text, Pic1.Image)
       
                    FileList.ListItems(lIndexCounter).SmallIcon = .ListImages(.ListImages.Count).Index
                
                End With
                '-------------------------------------------------------------------

                'Debug.Print lSendMsgRet(0) & vbTab & lIndexCounter & vbTab & mVariables.lListItemCount & vbTab & sPath$ & .ListItems(lIndexCounter).Text
            
            End If
   
        Next
    
    .Visible = True

'DoEvents
    End With
    

End Sub
