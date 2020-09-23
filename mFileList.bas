Attribute VB_Name = "mFileList"
'***********************************************************************************
'All that were doing here is using the FindFirstFile, FindNextFile and FindClose API
'adding all Files to the ListView as we go. Once we`ve added all the Files, we close
'the search Handle lngReturn&
'***********************************************************************************

Option Explicit
                                                                                    
            
Public Sub subFileList(oFileList As ListView, sFolderPath As String)

    Dim lReturn As Long                    'Search Handle of specified Path.
    Dim lNextFile As Long                  'Search Handle of specified File.
    Dim sPath As String                    'Path to search.
    Dim WFD As WIN32_FIND_DATA             'Set Variable WFD as Structure(VBType) WIN32_FIND_DATA.
    Dim lstItem As ListItem                'lstItem = A ListView ListItem.
    Dim lstSubItem As ListSubItem          'lstSubItem = A ListView ListSubItem.
    Dim sFileName As String                'Filename (WFD.cFileName).
        sPath$ = sFolderPath$ & "*.*"      'Hold Path and Filespec of sFolderPath$.
    Dim SYSTIME As SYSTEMTIME              'Set variable SYSTIME as structure(VBType) SYSTEMTIME.
    Dim sFileExtension As String           'Extension for File.
    
    With oFileList
        
        lReturn& = FindFirstFile(sPath$, WFD) & Chr$(0)
        frmMain.MousePointer = 11
                
        'If their are no Files to list, Exit sub.----------------------------------
        If lReturn& <= 0 Then
            .Visible = True
            frmMain.MousePointer = 0
        Exit Sub
        End If
        '--------------------------------------------------------------------------
        
        .SmallIcons = frmMain.ImageList2
        
        Do
            'If we find a Directory do nothing, else List Files taking off the Chr$(0)
            'Loop until lNextFile& = val(0), no more Files to List
            If Not (WFD.dwFileAttributes And vbDirectory) = vbDirectory Then
        
                sFileName$ = mProcFunc.ftnStripNullChar(WFD.cFileName)
            
                If IsEmpty(sFileName$) = False Then
                        
'                        Set lstItem = .ListItems.Add(, , sFileName$)
                        Set lstItem = .ListItems.Add(, , sFileName$, , frmMain.ImageList2.ListImages(1).Index)
                        
                        Set lstSubItem = lstItem.ListSubItems.Add(, , Format(WFD.nFileSizeLow, "#,0"))
                        
                        Set lstSubItem = lstItem.ListSubItems.Add(, , mProcFunc.ftnReturnAttributes(WFD.dwFileAttributes))
                        
                        FileTimeToSystemTime WFD.ftLastWriteTime, SYSTIME
                        Set lstSubItem = lstItem.ListSubItems.Add(, , SYSTIME.wDay & "." & SYSTIME.wMonth & "." & SYSTIME.wYear)
                        
                        sFileExtension$ = LCase(GetExtension(sFolderPath$ & sFileName$))
                        Set lstSubItem = lstItem.ListSubItems.Add(, , sFileExtension$)
                        
                End If
            
            End If
        
            lNextFile& = FindNextFile(lReturn&, WFD)
        
        Loop Until lNextFile& <= Val(0)
        
        frmMain.MousePointer = 0
    
        'Close Search Handle.-------------------------------------------------------
        lNextFile& = FindClose(lReturn&)
        '---------------------------------------------------------------------------
    
    End With

End Sub


'***********************************************************************************
'Retrieves the extension of a file.
Private Function GetExtension(ByVal sPath As String) As String

    GetExtension$ = mProcFunc.Ptr2StrU(PathFindExtension(StrPtr(sPath$)))
        
End Function
'***********************************************************************************

