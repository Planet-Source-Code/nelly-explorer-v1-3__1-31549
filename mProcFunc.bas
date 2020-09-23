Attribute VB_Name = "mProcFunc"
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

'***********************************************************************************
'Return Correct Drive Path when Clicking on Node in Explorer Tree eg. "My Computer\C:\WinNt"
'would return "C:\WinNt".
'***********************************************************************************
Public Function ftnReturnNodePath(sExplorerPath As String) As String

    Dim iSearch(1) As Integer
    Dim sRootPath As String
    
    iSearch%(0) = InStr(1, sExplorerPath$, "(", vbTextCompare)
    iSearch%(1) = InStr(1, sExplorerPath$, ")", vbTextCompare)
    
    If iSearch%(0) > 0 Then
        sRootPath$ = Mid(sExplorerPath$, iSearch%(0) + 1, 2)
    End If

    If iSearch%(1) > 0 Then
        ftnReturnNodePath$ = sRootPath$ & Mid(sExplorerPath$, iSearch%(1) + 1, Len(sExplorerPath$)) & "\"
    End If
    
End Function

'***********************************************************************************
'Determine what Attributes the File or Folder has. We pass the Value of the Attribute.
'***********************************************************************************
Public Function ftnReturnAttributes(lAttribute As Long) As String
   
    If lAttribute& And FILE_ATTRIBUTE_READONLY Then
        ftnReturnAttributes = "r"
    End If
    If lAttribute& And FILE_ATTRIBUTE_ARCHIVE Then
        ftnReturnAttributes = ftnReturnAttributes & ".a"
    End If
    If lAttribute& And FILE_ATTRIBUTE_SYSTEM Then
        ftnReturnAttributes = ftnReturnAttributes & ".s"
    End If
    If lAttribute& And FILE_ATTRIBUTE_HIDDEN Then
        ftnReturnAttributes = ftnReturnAttributes & ".h"
    End If
        
End Function

'***********************************************************************************
'Open associated application from sCorrectPath$ & sFileName$.
'***********************************************************************************
Public Sub subShellApplication(sCorrectPath As String, sFileName As String, lHwnd As Long)
    
    Dim lShellFile As Long
    
    'Shell Application using associated Application.--------------------------------
    lShellFile& = ShellExecute(lHwnd&, "open", sFileName$, vbNullString, sCorrectPath$, SW_SHOWNORMAL)
    '-------------------------------------------------------------------------------
    
    'If an error occurs the value returned by x is < 32, if so, Display Error.------
    If lShellFile& > 32 Then
    Exit Sub
    Else
    '-------------------------------------------------------------------------------

        'Use Select Case to determine Error.----------------------------------------
        Select Case lShellFile&
    
            Case 2

                If Right(sFileName$, 3) <> "htm" Then
                    MsgBox "File not found.", vbCritical + vbOKOnly, "X-File:"
                End If
                Exit Sub
            Case 3
                MsgBox "Path not found.", vbCritical + vbOKOnly, "X-File:"
                Exit Sub
            Case 5
                MsgBox "Access denied.", vbCritical + vbOKOnly, "X-File:"
                Exit Sub
            Case 8
                MsgBox "Out of Memory.", vbCritical + vbOKOnly, "X-File:"
                Exit Sub
            Case 32
                MsgBox "Shell32.dll not found.", vbCritical + vbOKOnly, "X-File:"
                Exit Sub

        End Select
        '---------------------------------------------------------------------------
    
    End If

End Sub


'***********************************************************************************
'Returns the string which is located at a memory address.
'***********************************************************************************
Public Function Ptr2StrU(ByVal pAddr As Long) As String
  
    Dim lStrLen As Long
        lStrLen& = lstrlenW(pAddr)      'Length of string at address.
    
  Ptr2StrU = Space$(lStrLen&)           'Create an area in memory to copy string at
                                        'pointer(pAddr).
  CopyMemory ByVal StrPtr(Ptr2StrU), ByVal pAddr, lStrLen& * 2    'Copy memory from source
                                        'to dest. pAddr is * 2 because of the unicode
                                        'strings, which use 2 bytes for every character.

End Function
'-----------------------------------------------------------------------------------

'***********************************************************************************
'Returns the correct path from sInputPath, eg.
'Local Drive    computer1(c:\ DRIVEC) = c:\
'***********************************************************************************
Public Function ftnSelectedPath(sInputPath As String) As String

    Dim iSearch(0) As Integer
    Dim sDrivePath As String
    Dim sFolderPath As String
                    
    iSearch%(0) = InStr(1, sInputPath$, "(", vbTextCompare)
   
    sDrivePath$ = Mid(sInputPath$, iSearch%(0) + 1, 2)

    iSearch%(0) = InStrRev(sInputPath$, ")", -1, vbTextCompare)
            
    If iSearch%(0) > 0 Then
        sFolderPath$ = Mid(sInputPath$, iSearch%(0) + 2, Len(sInputPath$) - iSearch%(0))
    End If
              
    ftnSelectedPath$ = sDrivePath$ & "\" & sFolderPath$ '& "\"
           
 End Function
