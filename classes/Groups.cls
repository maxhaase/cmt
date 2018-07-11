VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Group"
Attribute VB_Ext_KEY = "Member0" ,"Group"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


' *****************************************************************************
' SMART WEB API 1.0
' Universal language e-commerce library
' Programmer:   Max Haase   maxhaase@gmail.com
' November 2000
' *****************************************************************************

Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mCol As Collection
Private mvarstrConnection As String
Private mvarlngError As Long

Public Property Let lngError(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngError = 5
    mvarlngError = vData
End Property


Public Property Get lngError() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngError
    lngError = mvarlngError
End Property



Public Property Let strConnection(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strConnection = 5
    mvarstrConnection = vData
End Property


Public Property Get strConnection() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strConnection
    strConnection = mvarstrConnection
End Property

Public Function Update(lngGroupId As Long, strGroupName As String) As Long
  On Error GoTo EH:
   
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Groups_Update"
        .Parameters("@lngGroupName") = strGroupName
        .Parameters("@lngGroupId") = lngGroupId
        .Execute
    End With

    If oConn.State = 1 Then oConn.Close
    Update = 0
    Exit Function
    
EH:
    Update = 5
    If oConn.State = 1 Then oConn.Close
End Function

Public Sub Fetch()

Dim oTemp As Group
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Groups_Fetch"
         Set oRs = .Execute
    End With

    While Not oRs.EOF
        Set oTemp = New Group
        oTemp.lngGroupId = oRs("lngGroupId")
        oTemp.strGroupName = oRs("strGroupName")
        mCol.Add oTemp, oTemp.lngGroupId
        oRs.MoveNext
        Set oTemp = Nothing
    Wend
    
End Sub

Public Function Add(ByVal strGroupName As String) As Long

    On Error GoTo EH:
    
    Dim oTemp As Group
    Set oTemp = New Group
    
    oTemp.strGroupName = strGroupName
   
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Groups_Add"
        .Parameters("@strGroupName") = oTemp.strGroupName
        .Execute
        oTemp.lngGroupId = .Parameters("@lngGroupId")
    End With
    
    mCol.Add oTemp, CStr(oTemp.lngGroupId)
    
    'return the id created
    Add = oTemp.lngGroupId
    Set oTemp = Nothing
    If oConn.State = 1 Then oConn.Close
    Exit Function
    
EH:
    Add = 0
    If oConn.State = 1 Then oConn.Close
End Function

Public Property Get Item(vntIndexKey As Variant) As Group
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


Public Function Remove(lngGroupId As Long) As Long
On Error GoTo EH:

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Groups_Delete"
        .Parameters("@lngGroupId") = lngGroupId
        .Execute
    End With

    mCol.Remove CStr(lngGroupId)
    Remove = 0
    If oConn.State = 1 Then oConn.Close
    Exit Function
EH:
    Remove = 5
    If oConn.State = 1 Then oConn.Close
End Function


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

