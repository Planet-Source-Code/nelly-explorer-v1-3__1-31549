Attribute VB_Name = "mVariables"
Private m_sComputerName As String        'Holds ComputerName.
Private m_sExplorerTreePath As String    'Holds ExplorerTree selected path.
Private m_lListItemCount As Long         'Holds frmMain.FileList.ListItem.Count.
Private m_VisibleFileCount As Integer    'Holds the value of how many Files are visible in frmMain.FileList.

'm_ComputerName. Holds ComputerName.
Property Get sComputerName() As String
    sComputerName = m_sComputerName
End Property
Property Let sComputerName(newValue As String)
    m_sComputerName = newValue
End Property

'm_sExplorerTreePath. Holds ExplorerTree selected path.
Property Get sExplorerTreePath() As String
    sExplorerTreePath = m_sExplorerTreePath
End Property
Property Let sExplorerTreePath(newValue As String)
    m_sExplorerTreePath = newValue
End Property

'm_lListItemCount. Holds frmMain.FileList.ListItem.Count.
Property Get lListItemCount() As Long
    lListItemCount = m_lListItemCount
End Property
Property Let lListItemCount(newValue As Long)
    m_lListItemCount = newValue
End Property

'm_lListItemCount. Holds the value of how many Files are visible in frmMain.FileList.
Property Get VisibleFileCount() As Integer
    VisibleFileCount = m_VisibleFileCount
End Property
Property Let VisibleFileCount(newValue As Integer)
    m_VisibleFileCount = newValue
End Property

