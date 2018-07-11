VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Location"
Attribute VB_Ext_KEY = "Member0" ,"Location"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' *****************************************************************************
' SMART WEB API 1.0
' Universal language e-commerce library
' Programmer:   Max Haase   maxhaase@gmail.com
' December 2000
' *****************************************************************************

Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mCol As Collection
Private mvarstrConnection As String

Public Function Update(lngLocationId As Long, strLocation As String, strNotes As String, Optional lngLanguageId As Long = 1) As Long

    Dim lngPhraseId As Long


    If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Locations_Update"
            .Parameters("@strLocation") = strLocation
            .Parameters("@lngLocationId") = lngLocationId
            .Execute
           
            .CommandText = "LocationsPhrases_Fetch"
            .Parameters("@lngLocationId") = lngLocationId
            .Execute
             lngPhraseId = .Parameters("@lngPhraseId")
            
            .CommandText = "Phrases_update"
            .Parameters("@strPhrase") = strNotes
            .Parameters("@lngLanguageId") = lngLanguageId
            .Parameters("@lngPhraseId") = lngPhraseId
            .Execute
            
        End With
        
    If oConn.State = 1 Then oConn.Close


    'return the id created
    Update = 0

End Function


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
Public Sub FetchRentals(Optional lngLanguageId As Long = 1, Optional lngCategories As Long = 0)

Dim oTemp As Location
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Locations_Fetch_Rentals"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngCategories") = lngCategories
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Location
        With oTemp
            .lngLocationId = oRs("lngLocationId")
            .strLocation = oRs("strLocation")
            .strNotes = oRs("strPhrase")
            .lngPhraseId = oRs("lngPhraseId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngLocationId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub

Public Sub FetchResale(Optional lngLanguageId As Long = 1, Optional lngCategories As Long = 0)

Dim oTemp As Location
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Locations_Fetch_Resale"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngCategories") = lngCategories
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Location
        With oTemp
            .lngLocationId = oRs("lngLocationId")
            .strLocation = oRs("strLocation")
            .strNotes = oRs("strPhrase")
            .lngPhraseId = oRs("lngPhraseId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngLocationId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub
Public Sub FetchNew(Optional lngLanguageId As Long = 1, Optional lngCategories As Long = 0)

Dim oTemp As Location
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Locations_Fetch_New"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngCategories") = lngCategories
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Location
        With oTemp
            .lngLocationId = oRs("lngLocationId")
            .strLocation = oRs("strLocation")
            .strNotes = oRs("strPhrase")
            .lngPhraseId = oRs("lngPhraseId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngLocationId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub

Public Sub Fetch(Optional lngLanguageId As Long = 1, Optional lngCategories As Long = 0)

Dim oTemp As Location
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Locations_Fetch_"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngCategories") = lngCategories
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Location
        With oTemp
            .lngLocationId = oRs("lngLocationId")
            .strLocation = oRs("strLocation")
            .strNotes = oRs("strPhrase")
            .lngPhraseId = oRs("lngPhraseId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngLocationId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub

Public Function Add(strLocation As String, strNotes As String, Optional lngLanguageId As Long = 1) As Long

    Dim oTemp As Location, lngPhraseId As Long, oLanguage As Language, oLanguages As Languages
    Set oTemp = New Location
    
    oTemp.strLocation = strLocation
    oTemp.strNotes = strNotes
    
    If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Locations_Add"
            .Parameters("@strLocation") = oTemp.strLocation
            .Execute
            oTemp.lngLocationId = .Parameters("@lngLocationId")
                        
            .CommandText = "Phrases_add" 'add the text about the category
            .Parameters("@lngLanguageId") = lngLanguageId
            .Parameters("@strPhrase") = oTemp.strNotes
            .Parameters("@lngFileId") = 2 'file number 2 is used for translating location texts
            .Execute
            lngPhraseId = .Parameters("@lngPhraseId")
            
            .CommandText = "LocationsPhrases_Add"
            .Parameters("@lngLocationId") = oTemp.lngLocationId
            .Parameters("@lngPhraseId") = lngPhraseId
            .Execute
            Set oLanguages = New Languages   'Add the phrase in all existing languages
            oLanguages.strConnection = strConnection
            oLanguages.Fetch
            
       
            For Each oLanguage In oLanguages
                If Not lngLanguageId = oLanguage.lngLanguageId Then
                    .CommandType = adCmdStoredProc
                    .CommandText = "Text_add"
                    .Parameters("@strPhrase") = oTemp.strNotes
                    .Parameters("@lngPhraseId") = lngPhraseId
                    .Parameters("@lngLanguageId") = oLanguage.lngLanguageId
                    .Execute
                End If
            Next
         
            
        End With
        
    mCol.Add oTemp, CStr(oTemp.lngLocationId)
    
    If oConn.State = 1 Then oConn.Close


    'return the id created
    Add = oTemp.lngLocationId
    Set oTemp = Nothing


End Function

Public Property Get Item(lngLocationId As Long) As Location
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(CStr(lngLocationId))
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Function Delete(lngLocationId As Long) As Long
Dim lngPhraseId As Long
On Error GoTo EH:

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        
        .CommandText = "LocationsPhrases_Fetch"
        .Parameters("@lngLocationId") = lngLocationId
        .Execute
         lngPhraseId = .Parameters("lngPhraseId")
        
        .CommandText = "FilesPhrases_delete"
        .Parameters("@lngFileId") = 2
        .Parameters("lngPhraseId") = lngPhraseId
        .Execute
        
        .CommandText = "Locations_Delete"
        .Parameters("@lngLocationId") = lngLocationId
        .Execute
        
    End With

    mCol.Remove CStr(lngLocationId)
    Delete = 0
    If oConn.State = 1 Then oConn.Close
    Exit Function
EH:
    Delete = 5
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

