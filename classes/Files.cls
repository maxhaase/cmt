VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"File"
Attribute VB_Ext_KEY = "Member0" ,"File"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


' *****************************************************************************
' SMART WEB API 1.0
' Universal language e-commerce library
' Programmer:   Max Haase   maxhaase@gmail.com
' November 2000
' *****************************************************************************


Private mCol As Collection
Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mvarlngError As Long
Private mvarstrConnection As String

Public Function Update(ByVal lngFileId As Long, ByVal strFileName As String) As Long
On Error GoTo EH:

        If oConn.State = 0 Then oConn.Open strConnection
            
            With oCmd
                Set .ActiveConnection = oConn
                .CommandType = adCmdStoredProc
                .CommandText = "Files_update"
                .Parameters("@strFileName") = strFileName
                .Parameters("@lngFileId") = lngFileId
                .Execute
                Update = 0
                If oConn.State = 1 Then oConn.Close
                Exit Function
            End With
EH:
    Me.lngError = 5
    Update = 5
    If oConn.State = 1 Then oConn.Close
End Function

Public Function Delete(ByVal lngFileId As Long) As Long
On Error GoTo EH:

        If oConn.State = 0 Then oConn.Open strConnection
        
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Files_delete"
            .Parameters("@lngFileId") = lngFileId
            Set oRs = .Execute
            mCol.Remove CStr(lngFileId)
            Delete = 0
         End With
         If oConn.State = 1 Then oConn.Close
         Exit Function
EH:
    Delete = 5
    If oConn.State = 1 Then oConn.Close
End Function

Public Sub Fetch(Optional lngFileId As Long = 0)
       
Set mCol = Nothing
Set mCol = New Collection
Dim oTemp As File 'create a new object
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Files_Fetch"
        .Parameters("@lngFileId") = lngFileId
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New File
        oTemp.lngFileId = oRs("lngFileId")
        oTemp.strFileName = oRs("strFileName")
        mCol.Add oTemp, CStr(oTemp.lngFileId)
        oRs.MoveNext
        Set oTemp = Nothing
    Wend
    
    Set oRs = Nothing
    
If oConn.State = 1 Then oConn.Close

End Sub

Public Function Add(ByVal strFileName As String) As Long
    
    Dim oTemp As File 'create a new object
    Set oTemp = New File
    
    oTemp.strFileName = strFileName
    
    If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Files_Add"
            .Parameters("@strFileName") = oTemp.strFileName
            .Execute
            If Not IsNull(.Parameters("@lngFileId")) Then
                oTemp.lngFileId = .Parameters("@lngFileId")
                mCol.Add oTemp, CStr(oTemp.lngFileId)
            Else
                oTemp.lngFileId = 0
                Me.lngError = 5
            End If
        End With
        
    If oConn.State = 1 Then oConn.Close
    
    Add = oTemp.lngFileId  'return the file id created
    Set oTemp = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As File
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Public Property Let lngError(ByVal vData As Long)
    mvarlngError = vData
End Property

Public Property Get lngError() As Long
    lngError = mvarlngError
End Property

Public Property Let strConnection(ByVal vData As String)
    mvarstrConnection = vData
End Property

Public Property Get strConnection() As String
    strConnection = mvarstrConnection
End Property
