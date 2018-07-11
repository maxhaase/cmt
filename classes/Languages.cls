VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Languages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Language"
Attribute VB_Ext_KEY = "Member0" ,"Language"
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
Private mvarstrConnection As Variant


Public Property Let strConnection(ByVal vData As Variant)
    mvarstrConnection = vData
End Property

Public Property Set strConnection(ByVal vData As Variant)
    Set mvarstrConnection = vData
End Property

Public Property Get strConnection() As Variant
    If IsObject(mvarstrConnection) Then
        Set strConnection = mvarstrConnection
    Else
        strConnection = mvarstrConnection
    End If
End Property

Public Sub Fetch()
       
Set mCol = Nothing
Set mCol = New Collection
Dim oTemp As Language 'create a new object
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Languages_Fetch"
        Set oRs = .Execute
    End With
            
    While Not oRs.EOF
        Set oTemp = New Language
        oTemp.lngLanguageId = oRs("lngLanguageId")
        oTemp.lngPhraseId = oRs("lngPhraseId")
        mCol.Add oTemp, CStr(oTemp.lngLanguageId)
        oRs.MoveNext
        Set oTemp = Nothing
    Wend


End Sub


Public Function Add(ByVal lngPhraseId As Long, Optional ByVal lngPhraseType As Long) As Long
    
    Dim oTemp As Language   'create a new object
    Set oTemp = New Language
        
    oTemp.lngPhraseId = lngPhraseId
    
     If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Languages_add"
            .Parameters("@lngPhraseId") = oTemp.lngPhraseId
            .Execute
            If Not IsNull(.Parameters("@lngLanguageId")) Then
                oTemp.lngLanguageId = .Parameters("@lngLanguageId")
                mCol.Add oTemp, CStr(oTemp.lngLanguageId)
            Else
                oTemp.lngLanguageId = 0
                Me.lngError = 5
            End If
        End With
        
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Text_Fetch"
            .Parameters("@lngLanguageId") = 1
            Set oRs = .Execute
        End With
        
        While Not oRs.EOF
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Text_add"
            .Parameters("@lngLanguageId") = oTemp.lngLanguageId
            .Parameters("@strPhrase") = oRs("strPhrase")
            .Parameters("@lngPhraseId") = oRs("lngPhraseId")
            .Execute
        End With
            oRs.MoveNext
        Wend
        
        
        
        
    If oConn.State = 1 Then oConn.Close
   
    Me.Fetch
   
    Add = oTemp.lngLanguageId     'return the id created
    Set oTemp = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Language
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


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
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



Public Property Let lngError(ByVal vData As Long)
    mvarlngError = vData
End Property

Public Property Get lngError() As Long
    lngError = mvarlngError
End Property

