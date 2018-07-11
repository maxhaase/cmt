VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Phrases"
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
Public FilesPhrases As Collection
Private mvarlngFileId As Long
Private mvarstrFileName As String
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

Public Function GetPhrases(ByVal lngLanguageId As Long) As Collection

Set FilesPhrases = New Collection

Dim oTemp As Phrase 'create a new phrase object
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "FilesPhrases_Fetch"
        .Parameters("@lngFileId") = Me.lngFileId
        .Parameters("@lngLanguageId") = lngLanguageId
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF 'populate the collection with the recordset
        Set oTemp = New Phrase
        oTemp.lngPhraseId = oRs("lngPhraseId")
        oTemp.strPhrase = oRs("strPhrase")
        FilesPhrases.Add oTemp, CStr(oTemp.lngPhraseId)
        oRs.MoveNext
        Set oTemp = Nothing
    Wend
    If oConn.State = 1 Then oConn.Close
    Set GetPhrases = FilesPhrases
    
End Function

Public Property Let strFileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strFileName = 5
    mvarstrFileName = vData
End Property


Public Property Get strFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strFileName
    strFileName = mvarstrFileName
End Property



Public Property Let lngFileId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngFileId = 5
    mvarlngFileId = vData
End Property


Public Property Get lngFileId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngFileId
    lngFileId = mvarlngFileId
End Property

Private Sub Class_Terminate()
    Set FilesPhrases = Nothing
End Sub
