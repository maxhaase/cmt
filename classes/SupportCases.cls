VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupportCases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"SupportCase"
Attribute VB_Ext_KEY = "Member0" ,"SupportCase"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' *****************************************************************************
' SMART WEB API 1.0
' Universal language e-commerce library
' Programmer:   Max Haase   maxhaase@gmail.com
' December 2000
' *****************************************************************************

Private mCol As Collection
Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mvarstrConnection As String

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

Public Sub Fetch(lngCustomerId As Long)
       
Set mCol = Nothing
Set mCol = New Collection
Dim oTemp As SupportCase  'create a new object
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Cases_Fetch"
        .Parameters("@lngCustomerId") = lngCustomerId
        Set oRs = .Execute
    End With
            
    While Not oRs.EOF
        Set oTemp = New SupportCase
        oTemp.lngCaseId = oRs("lngCaseId")
        oTemp.strCase = oRs("strCase")
        mCol.Add oTemp, CStr(oTemp.lngCaseId)
        oRs.MoveNext
        Set oTemp = Nothing
    Wend

End Sub
Public Function Add(strCase As String, lngCustomerId As Long) As Long
    'create a new object
    Dim oTemp As SupportCase
    Set oTemp = New SupportCase

    'set the properties passed into the method
    oTemp.strCase = strCase
    
       If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Cases_Add"
            .Parameters("@strCase") = oTemp.strCase
            .Execute
            oTemp.lngCaseId = .Parameters("@lngCaseId")
            .CommandText = "CustomersCases_Add"
            .Parameters("@lngCaseId") = oTemp.lngCaseId
            .Parameters("@lngCustomerId") = lngCustomerId
            .Execute
        End With
     
        mCol.Add oTemp, CStr(oTemp.lngCaseId)
   
        'return the case id created
        Add = oTemp.lngCaseId
        Set oTemp = Nothing
        
        If oConn.State = 1 Then oConn.Close

End Function

Public Property Get Item(vntIndexKey As Variant) As SupportCase
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


Public Sub Remove(lngCaseId As Long)

     If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Cases_Delete"
            .Parameters("@lngCaseId") = lngCaseId
            .Execute
        End With

    mCol.Remove CStr(lngCaseId)
End Sub


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

