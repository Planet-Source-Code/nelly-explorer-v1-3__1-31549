Attribute VB_Name = "mSearchNodes"
Option Explicit

'***********************************************************************************
'Searches for a path in an ExplorerTree, once found, act on specific operation.
'***********************************************************************************
Public Sub subSearchNodes(obExplorerTree As TreeView, sPathToSearchFor As String)

    Dim tnNode As Node

    For Each tnNode In obExplorerTree.Nodes
'MsgBox "1 " & mProcFunc.ftnCorrectPath(sDriveList$, nNode.FullPath, "", sExplorerTree$) & "    " & "2 " & sPathToSearchFor$
        With obExplorerTree
'MsgBox mProcFunc.ftnSelectedPath(tnNode.FullPath) & "      " & sPathToSearchFor
            If mProcFunc.ftnSelectedPath(tnNode.FullPath) = sPathToSearchFor$ Then
            
            MsgBox "found"
        
            End If
            
        End With

    Next
    
End Sub
