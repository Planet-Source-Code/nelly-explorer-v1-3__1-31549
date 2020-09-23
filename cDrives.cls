VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Constants for DriveType.-----------------------------------------------------------
Private Enum DRIVE_TYPE
    DRIVE_UNKNOWN = 0
    DRIVE_NO_ROOT_DIR = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5
    DRIVE_RAMDISK = 6
End Enum
'-----------------------------------------------------------------------------------

Public Sub subLoadTreeView(TreeList As TreeView)

    Dim iDriveNum As Integer                    'Key Index in ExplorerTree.
    Dim sDriveType As String                    'Holds DriveType.
    Dim x As Integer                            'Loop through Drives.
    Dim RetStr(1) As String                     'Holds Drive letters.
    Dim nNode As Node                           'Node object for ExplorerTree.
    Dim sComputerName As String                 'Hold Computer Name. The "?" before the computername
                                                'stops the Invalid Key error if the Computer Name is
                                                'a number.
        sComputerName$ = "?" & mVariables.sComputerName
    Dim sDriveInformation As String             'Holds information about Drive.
    Dim bytDriveType As Byte                    'Used to hold DriveType using API.
                        
    'Return Drive structure from XFile.Dll. Removed DLL-----------------------------
    RetStr$(0) = ftnShowDriveList
    '-------------------------------------------------------------------------------

    With TreeList

        'Add Computer name to ExplorerTree------------------------------------------
        Set nNode = .Nodes.Add(, , sComputerName$, Mid(sComputerName$, 2, Len(sComputerName$) - 1), "mycomputer", "mycomputer")
        '---------------------------------------------------------------------------
        
        'Loop through RetStr$(0) to retrieve Drives. eg."ACDEF".--------------------
        For x = 1 To Len(RetStr$(0))
            
            'Strip Driveinfo eg "A"-------------------------------------------------
            RetStr$(1) = Mid(RetStr$(0), x, 1)
            '-----------------------------------------------------------------------
            
            sDriveInformation$ = ftnDriveInformation(RetStr$(1) & ":\", "Null")
                
            'Get DriveType information.---------------------------------------------
            bytDriveType = GetDriveType(RetStr(1) & ":")
            '-----------------------------------------------------------------------
            
            'Used to make unique Key Index in ExplorerTree.-------------------------
            iDriveNum% = x
            '-----------------------------------------------------------------------
            
            'Determine Drive type and add to TreeList.------------------------------
            Select Case bytDriveType
                
                'Unknown Drive.-----------------------------------------------------
                Case 1: sDriveType = DRIVE_UNKNOWN
                    Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, sDriveInformation$ & " (" & RetStr$(1) & ":)", "drvunknown")

                'Removable Drive.---------------------------------------------------
                Case 2: sDriveType = DRIVE_REMOVABLE
                    Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, sDriveInformation$ & " (" & RetStr$(1) & ":)", "drvremove")

                'Fixed Drive.-------------------------------------------------------
                Case 3: sDriveType = DRIVE_FIXED
                    Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, sDriveInformation$ & " (" & RetStr$(1) & ":)", "drvfixed")

                'Network Drive.-----------------------------------------------------
                Case 4: sDriveType = DRIVE_REMOTE
                    Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, sDriveInformation$ & " (" & RetStr$(1) & ":)", "drvremote")
                    
                'CD-Rom.------------------------------------------------------------
                Case 5: sDriveType = DRIVE_CDROM
                    Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, sDriveInformation$ & " (" & RetStr$(1) & ":)", "drvcd")
                    
                'Ram Disk.----------------------------------------------------------
                Case 6: sDriveType = DRIVE_RAMDISK
                    Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, sDriveInformation$ & " (" & RetStr$(1) & ":)", "drvram")
            
            End Select
            '-----------------------------------------------------------------------
            
        Next x

    End With

End Sub


'***********************************************************************************
'Return all Drives installed on the Users Local System using GetLogicalDriveStrings API.
'***********************************************************************************
Private Function ftnShowDriveList() As String
    
    Dim sBuffer As String * 200
    Dim lReturnAPI As Long
    Dim iLoop As Integer
    Dim iSearch As Integer
    Dim iLengthOfBuffer As Integer
    Dim sDriveLetter As String
    
    lReturnAPI = GetLogicalDriveStrings(Len(sBuffer$), sBuffer$)

    iLengthOfBuffer% = InStr(1, sBuffer$, Chr$(0) & Chr$(0))

    For iLoop% = 1 To iLengthOfBuffer%

        iSearch% = InStr(iSearch% + 1, sBuffer$, Chr$(0))

        If Mid(sBuffer$, iSearch% - 1, 1) <> Chr$(0) Then
            sDriveLetter$ = sDriveLetter$ & Mid(sBuffer$, iSearch% - 3, 1)
        End If

    Next iLoop%
    
    ftnShowDriveList$ = sDriveLetter$

End Function
'-----------------------------------------------------------------------------------

'***********************************************************************************
'Returns information about a file system and volume whose root directory is specified.
'***********************************************************************************
Public Function ftnDriveInformation(sCurrentDrive As String, sSource As String) As String

    Dim sVolumeName As String * 255
    Dim lVolumeSerialNo As Long
    Dim sFileSystemName As String * 25
    Dim lAPI As Long

    lAPI& = GetVolumeInformation(sCurrentDrive$, sVolumeName$, 255, lVolumeSerialNo&, 255, 0&, sFileSystemName$, 25)
    
    ftnDriveInformation$ = ftnStripNullChar(sVolumeName$)
    
End Function
'-----------------------------------------------------------------------------------

'***********************************************************************************
'Strips the NullCharacter from sInput.
'***********************************************************************************
Public Function ftnStripNullChar(sInput As String) As String

    Dim x As Integer
    
    x = InStr(1, sInput$, Chr$(0))

    If x > 0 Then
        ftnStripNullChar = Left(sInput$, x - 1)
    End If

End Function

