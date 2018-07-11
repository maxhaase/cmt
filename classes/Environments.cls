VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Environment"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private m_Rs As New ADODB.Recordset
Private m_Cmd As New ADODB.Command
Private m_Conn As New ADODB.Connection
Private mvarEnvironment As Environment


Public Property Get Environment() As Environment
    If mvarEnvironment Is Nothing Then
        Set mvarEnvironment = New Environment
    End If


    Set Environment = mvarEnvironment
End Property


Public Property Set Environment(vData As Environment)
    Set mvarEnvironment = vData
End Property


Public Function TestConnection(strProvider As String, strServer As String, strDatabase As String, strLoginId As String, strPassword As String) As Long
Dim strConn
On Error Resume Next

Dim oTest As New ADODB.Connection
TestConnection = 0

    strConn = "Provider=" & strProvider & ";User ID=" & strLoginId & ";Password=" & strPassword & ";Initial Catalog=" & strDatabase & ";Data Source=" & strServer
    
    oTest.Errors.Clear
    
    If oTest.State = 1 Then oTest.Close 'make sure there is no open connetion with this object
     
    oTest.Open strConn
     
    If Not IsNull(Err.Number) Then TestConnection = Err.Number 'return the error collection from ADO
    
    If oTest.State = 1 Then oTest.Close ' close this connection if it did open
    
    Set oTest = Nothing

End Function

Public Function Add(strEnvironmentName As String, strProvider As String, strServer As String, strDatabase As String, strLoginId As String, strPassword As String) As Environment
Dim lngEnvironmentId As Long
    
    If m_Conn.State = adStateClosed Then m_Conn.Open strConnection
    
    Set m_Cmd.ActiveConnection = m_Conn
    m_Cmd.CommandText = "sp_webAppEnvironments_Add"
    m_Cmd.CommandType = adCmdStoredProc
    m_Cmd.Parameters("@strEnvironmentName").Value = Trim(strEnvironmentName)
    m_Cmd.Parameters("@strProvider").Value = Trim(strProvider)
    m_Cmd.Parameters("@strServer").Value = Trim(strServer)
    m_Cmd.Parameters("@strLoginId").Value = Trim(strLoginId)
    m_Cmd.Parameters("@strDatabase").Value = Trim(strDatabase)
    m_Cmd.Parameters("@strPassword").Value = strPassword
    m_Cmd.Execute
    lngEnvironmentId = m_Cmd.Parameters("@lngEnvironmentId").Value

    If m_Conn.State = adStateOpen Then m_Conn.Close

    Dim oTemp As Environment
    On Error Resume Next
    Set oTemp = mCol.Item(CStr(lngEnvironmentId))
    
    If oTemp Is Nothing Then
        Set oTemp = New Environment
        mCol.Add oTemp, CStr(lngEnvironmentId)
    End If
    
    With oTemp
        .lngEnvironmentId = lngEnvironmentId
        .strEnvironmentName = Trim(strEnvironmentName)
        .strProvider = Trim(strProvider)
        .strServer = Trim(strServer)
        .strDatabase = Trim(strDatabase)
        .strLoginId = Trim(strLoginId)
        .strPassword = strPassword
    End With
    
    SetConnectionString oTemp
    
    Set Add = oTemp
    
    'Dim objMember As Environment
    'For Each objMember In mCol
    '    Debug.Print objMember.lngEnvironmentId, objMember.strEnvironmentName, objMember.strProvider, objMember.strDatabase
    'Next objMember
    
End Function

Public Function Delete(lngEnvironmentId As Long) As Long
    If m_Conn.State = adStateClosed Then m_Conn.Open strConnection

    Set m_Cmd.ActiveConnection = m_Conn
    m_Cmd.CommandText = "sp_webAppEnvironments_Delete"
    m_Cmd.CommandType = adCmdStoredProc
    m_Cmd.Parameters("@lngEnvironmentId").Value = lngEnvironmentId
    m_Cmd.Execute
    Delete = m_Cmd.Parameters("@lngResult").Value
  
    
End Function

Public Function Fetch(Optional lngEnvironmentId As Long) As Environment
Dim oEnv As Environment

    Set mCol = Nothing
    Set mCol = New Collection
        
    If (m_Conn.State = adStateClosed) Then m_Conn.Open strConnection

    m_Cmd.ActiveConnection = m_Conn
    m_Cmd.CommandType = adCmdStoredProc
    m_Cmd.CommandText = "sp_webAppEnvironments_Fetch"
    Set m_Rs = m_Cmd.Execute
    
    While Not m_Rs.EOF
        Set oEnv = New Environment
        'lngEnvironmentId, strEnvironmentName, strProvider,
        'strServer, strDatabase, strLoginId, strPassword
        With oEnv
            .lngEnvironmentId = m_Rs!lngEnvironmentId
            If Not IsNull(m_Rs!strEnvironmentName) Then .strEnvironmentName = Trim(m_Rs!strEnvironmentName)
            If Not IsNull(m_Rs!strProvider) Then .strProvider = Trim(m_Rs!strProvider)
            If Not IsNull(m_Rs!strServer) Then .strServer = Trim(m_Rs!strServer)
            If Not IsNull(m_Rs!strDatabase) Then .strDatabase = Trim(m_Rs!strDatabase)
            If Not IsNull(m_Rs!strLoginId) Then .strLoginId = Trim(m_Rs!strLoginId)
            If Not IsNull(m_Rs!strPassword) Then .strPassword = m_Rs!strPassword
        End With
        
        mCol.Add oEnv, CStr(oEnv.lngEnvironmentId)
        m_Rs.MoveNext
    Wend
    m_Rs.Close
    
    If (m_Conn.State = adStateOpen) Then m_Conn.Close
    Set m_Rs = Nothing

    '+++++++++++++++++++++++++++++ return what we asked for
    If lngEnvironmentId <> 0 Then
        Set Fetch = mCol.Item(CStr(lngEnvironmentId))
    Else
        Set Fetch = New Environment
    End If
    
  '  Dim objMember As Environment
  '  For Each objMember In mCol
  '     Debug.Print objMember.lngEnvironmentId, objMember.strEnvironmentName, objMember.strProvider, objMember.strDatabase
  ' Next objMember
End Function

Public Property Get Item(lngEnvironmentId As Long) As Environment
  Set Item = mCol(CStr(lngEnvironmentId))
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Connect(strPortalId As String)
    strConnection = FetchConnectionString(strPortalId)
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  Set mvarEnvironment = Nothing
    Set mCol = Nothing
End Sub


