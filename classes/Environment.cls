VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private m_lngEnvironmentId As Long
Private m_strEnvironmentName As String
Private m_strProvider As String
Private m_strServer As String
Private m_strDatabase As String
Private m_strLoginId As String
Private m_strPassword As String

Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection

Public Function Update() As Long
On Error GoTo eh
If oConn.State = 0 Then oConn.Open strConnection

    oCmd.ActiveConnection = oConn
    oCmd.CommandText = "sp_WebAppEnvironment_Update"
    oCmd.CommandType = adCmdStoredProc
    oCmd.Parameters("@lngEnvironmentId").Value = lngEnvironmentId
    oCmd.Parameters("@strEnvironmentName").Value = strEnvironmentName
    oCmd.Parameters("@strProvider").Value = strProvider
    oCmd.Parameters("@strServer").Value = strServer
    oCmd.Parameters("@strLoginId").Value = strLoginId
    oCmd.Parameters("@strDatabase").Value = strDatabase
    oCmd.Parameters("@strPassword").Value = strPassword
    oCmd.Execute
    Update = 0
    Exit Function
eh:
    Update = -999

End Function


Public Function GetAssignedGroups() As Collection
On Error Resume Next

Dim oTemp As New Collection
Dim oGroup As Group
Dim oRs As New ADODB.Recordset

    If oConn.State = adStateClosed Then oConn.Open strConnection
    
    oCmd.ActiveConnection = oConn
    oCmd.CommandText = "sp_webAppEnvironmentsAssignedGroups_Get"
    oCmd.CommandType = adCmdStoredProc
    oCmd.Parameters("@lngEnvironmentId").Value = lngEnvironmentId
    Set oRs = oCmd.Execute
       
    While Not oRs.EOF
        Set oGroup = New Group
        With oGroup
            .lngGroupId = oRs!lngGroupId
            .strGroupName = oRs!strGroupName
        End With
        oTemp.Add oGroup, CStr(oGroup.lngGroupId)
        Set oGroup = Nothing
        oRs.MoveNext
    Wend
    
    If oConn.State = adStateOpen Then oConn.Close
    
    Set GetAssignedGroups = oTemp
    Set oTemp = Nothing

End Function


Public Function GetUnAssignedGroups() As Collection

Dim oTemp As New Collection
Dim oGroup As Group
Dim oRs As New ADODB.Recordset

    If oConn.State = adStateClosed Then oConn.Open strConnection
    
    oCmd.ActiveConnection = oConn
    oCmd.CommandText = "sp_webAppEnvironmentsUnassignedGroups_Get"
    oCmd.CommandType = adCmdStoredProc
    oCmd.Parameters("@lngEnvironmentId").Value = lngEnvironmentId
    Set oRs = oCmd.Execute
       
    While Not oRs.EOF
        Set oGroup = New Group
        With oGroup
            .lngGroupId = oRs!lngGroupId
            .strGroupName = oRs!strGroupName
        End With
        oTemp.Add oGroup, CStr(oGroup.lngGroupId)
        Set oGroup = Nothing
        oRs.MoveNext
    Wend
    
    If oConn.State = adStateOpen Then oConn.Close
    
    Set GetUnAssignedGroups = oTemp
    Set oTemp = Nothing

End Function

Public Function AssignGroup(lngGroupId As Long) As Long
On Error GoTo eh
    If oConn.State = adStateClosed Then oConn.Open strConnection
    
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandText = "sp_webAppGroupsEnvironments_Add"
    oCmd.CommandType = adCmdStoredProc
    oCmd.Parameters("@lngEnvironmentId") = lngEnvironmentId
    oCmd.Parameters("@lngGroupId") = lngGroupId
    oCmd.Execute
    
    If oConn.State = adStateOpen Then oConn.Close

    AssignGroup = 0
Exit Function
eh:
    AssignGroup = 99 'Error
End Function

Public Function UnassignGroup(lngGroupId As Long) As Long
On Error GoTo eh
    If oConn.State = adStateClosed Then oConn.Open strConnection
    
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandText = "sp_webAppGroupsEnvironments_Delete"
    oCmd.CommandType = adCmdStoredProc
    oCmd.Parameters("@lngEnvironmentId") = lngEnvironmentId
    oCmd.Parameters("@lngGroupId") = lngGroupId
    oCmd.Execute
    
    If oConn.State = adStateOpen Then oConn.Close

    UnassignGroup = 0
Exit Function
eh:
    UnassignGroup = 99 'Error
End Function


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Property Let lngEnvironmentId(ByVal vData As Long)
    m_lngEnvironmentId = vData
End Property

Public Property Get lngEnvironmentId() As Long
    lngEnvironmentId = m_lngEnvironmentId
End Property

Public Property Let strEnvironmentName(ByVal vData As String)
    m_strEnvironmentName = vData
End Property

Public Property Get strEnvironmentName() As String
    strEnvironmentName = m_strEnvironmentName
End Property

Public Property Let strProvider(ByVal vData As String)
    m_strProvider = vData
End Property

Public Property Get strProvider() As String
    strProvider = m_strProvider
End Property

Public Property Let strServer(ByVal vData As String)
    m_strServer = vData
End Property

Public Property Get strServer() As String
    strServer = m_strServer
End Property

Public Property Let strDatabase(ByVal vData As String)
    m_strDatabase = vData
End Property

Public Property Get strDatabase() As String
    strDatabase = m_strDatabase
End Property

Public Property Let strLoginId(ByVal vData As String)
    m_strLoginId = vData
End Property

Public Property Get strLoginId() As String
    strLoginId = m_strLoginId
End Property

Public Property Let strPassword(ByVal vData As String)
    m_strPassword = vData
End Property

Public Property Get strPassword() As String
    strPassword = m_strPassword
End Property






