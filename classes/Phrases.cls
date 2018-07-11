VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Phrase"
Attribute VB_Ext_KEY = "Member0" ,"Phrase"
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
Public strConnection As String

Public Function Add(ByVal strPhrase As String, ByVal lngLanguageId As Long, ByVal lngFileId As Long) As Long
'adds a phrase to the phrase table and a text to the text table, returns the phrase id of the new phrase

    Dim oTemp As Phrase, oLanguage As Language, oLanguages As Languages
    Set oTemp = New Phrase
 
    If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Phrases_add"
            .Parameters("@lngLanguageId") = lngLanguageId
            .Parameters("@strPhrase") = strPhrase
            .Parameters("@lngFileId") = lngFileId
            .Execute
                        
            If Not IsNull(.Parameters("@lngPhraseId")) Then
                oTemp.lngPhraseId = .Parameters("@lngPhraseId")
                mCol.Add oTemp, CStr(oTemp.lngPhraseId)
            Else 'there was an error
                oTemp.lngPhraseId = 0
                Me.lngError = 5 '5 is the phrase id number for the word error
                Exit Function
            End If
                    
        
            Set oLanguages = New Languages   'Add the phrase in all existing languages
            oLanguages.strConnection = strConnection
            oLanguages.Fetch
        
            For Each oLanguage In oLanguages
                If Not lngLanguageId = oLanguage.lngLanguageId Then
                    .CommandType = adCmdStoredProc
                    .CommandText = "Text_add"
                    .Parameters("@strPhrase") = oTemp.strPhrase
                    .Parameters("@lngPhraseId") = oTemp.lngPhraseId
                    .Parameters("@lngLanguageId") = oLanguage.lngLanguageId
                    .Execute
                End If
            Next
        
        End With
        
    If oConn.State = 1 Then oConn.Close
    
    Add = oTemp.lngPhraseId  'return the phrase id created
    Set oTemp = Nothing

End Function

Public Property Get Item(lngPhraseId As Long) As Phrase
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(CStr(lngPhraseId))
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Function Update(ByVal lngPhraseId As Long, ByVal strPhrase As String, ByVal lngLanguageId As Long) As Long

    If oConn.State = 0 Then oConn.Open strConnection
        
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Phrases_update"
            .Parameters("@strPhrase") = strPhrase
            .Parameters("@lngLanguageId") = lngLanguageId
            .Parameters("@lngPhraseId") = lngPhraseId
            .Execute
            Update = 0
        End With
   

End Function

Public Function Delete(ByVal lngPhraseId As Long) As Long
' deletes a phrase from the database, returns 0 if successful
        
        If oConn.State = 0 Then oConn.Open strConnection
       
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Phrases_delete"
            .Parameters("@lngPhraseId") = lngPhraseId
            .Execute
            mCol.Remove CStr(lngPhraseId)
            Delete = 0
        End With
        
End Function


Public Sub Fetch(ByVal lngLanguageId As Long)  'populates the phrases collection
       
Set mCol = Nothing
Set mCol = New Collection
Dim oTemp As Phrase 'create a new phrase object
    
    If oConn.State = 0 Then oConn.Open strConnection
    
            With oCmd
                Set .ActiveConnection = oConn
                .CommandType = adCmdStoredProc
                .CommandText = "Phrases_Fetch"
                .Parameters("@lngLanguageId") = lngLanguageId
                Set oRs = .Execute
            End With
            
    While Not oRs.EOF 'populate the collection with the recordset
        Set oTemp = New Phrase
        oTemp.lngPhraseId = oRs("lngPhraseId")
        oTemp.strPhrase = oRs("strPhrase")
        mCol.Add oTemp, CStr(oTemp.lngPhraseId)
        oRs.MoveNext
        Set oTemp = Nothing
    Wend

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



