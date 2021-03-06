VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Categories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Category"
Attribute VB_Ext_KEY = "Member0" ,"Category"
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

Public Function Update(lngCategoryId As Long, strCategory As String, lngPriceFrom As Long, lngPriceTo As Long, strNotes As String) As Long
Dim lngPhraseId As Long
On Error GoTo EH:
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Categories_Update"
        .Parameters("@lngCategoryId") = lngCategoryId
        .Parameters("@strCategory") = strCategory
        .Parameters("@lngPriceFrom") = lngPriceFrom
        .Parameters("@lngPriceTo") = lngPriceTo
        .Execute
        
        .CommandText = "CategoriesPhrases_Fetch"
        .Parameters("@lngCategoryId") = lngCategoryId
        .Execute
         lngPhraseId = .Parameters("@lngPhraseId")
        
        .CommandText = "Phrases_update"
        .Parameters("@strPhrase") = strNotes
        .Parameters("@lngLanguageId") = 1
        .Parameters("@lngPhraseId") = lngPhraseId
        .Execute
        
        
    End With
    
    If oConn.State = 1 Then oConn.Close

    Update = 0
    Exit Function
    
EH:
    Update = 5
    If oConn.State = 1 Then oConn.Close
    
End Function
Public Sub FetchRentals(Optional lngLocationId As Long = 0, Optional lngLanguageId As Long = 1)

Dim oTemp As Category
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Categories_Fetch_Rentals"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngLocationId") = lngLocationId
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Category
        With oTemp
            .lngCategoryId = oRs("lngCategoryId")
            .lngPriceFrom = oRs("lngPriceFrom")
            .lngPriceTo = oRs("lngPriceTo")
            .strCategory = oRs("strCategory")
            .strNotes = oRs("strPhrase")
            .strLocation = oRs("strLocation")
            .lngLocationId = oRs("lngLocationId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngCategoryId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub

Public Sub FetchResale(Optional lngLocationId As Long = 0, Optional lngLanguageId As Long = 1)

Dim oTemp As Category
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Categories_Fetch_Resale"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngLocationId") = lngLocationId
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Category
        With oTemp
            .lngCategoryId = oRs("lngCategoryId")
            .lngPriceFrom = oRs("lngPriceFrom")
            .lngPriceTo = oRs("lngPriceTo")
            .strCategory = oRs("strCategory")
            .strNotes = oRs("strPhrase")
            .strLocation = oRs("strLocation")
            .lngLocationId = oRs("lngLocationId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngCategoryId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub

Public Sub FetchNew(Optional lngLocationId As Long = 0, Optional lngLanguageId As Long = 1)

Dim oTemp As Category
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Categories_Fetch_New"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngLocationId") = lngLocationId
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Category
        With oTemp
            .lngCategoryId = oRs("lngCategoryId")
            .lngPriceFrom = oRs("lngPriceFrom")
            .lngPriceTo = oRs("lngPriceTo")
            .strCategory = oRs("strCategory")
            .strNotes = oRs("strPhrase")
            .strLocation = oRs("strLocation")
            .lngLocationId = oRs("lngLocationId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngCategoryId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub

Public Sub Fetch(Optional lngLocationId As Long = 0, Optional lngLanguageId As Long = 1)

Dim oTemp As Category
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Categories_Fetch_"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngLocationId") = lngLocationId
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Category
        With oTemp
            .lngCategoryId = oRs("lngCategoryId")
            .lngPriceFrom = oRs("lngPriceFrom")
            .lngPriceTo = oRs("lngPriceTo")
            .strCategory = oRs("strCategory")
            .strNotes = oRs("strPhrase")
            .strLocation = oRs("strLocation")
            .lngLocationId = oRs("lngLocationId")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngCategoryId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Sub

Public Function Add(lngLocationId As Long, strCategory As String, lngPriceFrom As Long, lngPriceTo As Long, strNotes As String, Optional lngLanguageId As Long = 1) As Long
    
    Dim oLanguage As Language, oLanguages As Languages
    Dim oTemp As Category, lngPhraseId As Long
    Set oTemp = New Category
    
        If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Phrases_add" 'add the text about the category
            .Parameters("@lngLanguageId") = lngLanguageId
            .Parameters("@strPhrase") = strNotes
            .Parameters("@lngFileId") = 3 'file number 3 is used for translating CATEGORIES texts
            .Execute
             lngPhraseId = .Parameters("@lngPhraseId")
    

            Set oLanguages = New Languages   'Add the phrase in all existing languages
            oLanguages.strConnection = strConnection
            oLanguages.Fetch
            
            For Each oLanguage In oLanguages
                If Not lngLanguageId = oLanguage.lngLanguageId Then
                    .CommandType = adCmdStoredProc
                    .CommandText = "Text_add"
                    .Parameters("@strPhrase") = strNotes
                    .Parameters("@lngPhraseId") = lngPhraseId
                    .Parameters("@lngLanguageId") = oLanguage.lngLanguageId
                    .Execute
                End If
            Next
            
    End With
   
    
    With oTemp
        .strCategory = strCategory
        .lngPriceFrom = lngPriceFrom
        .lngPriceTo = lngPriceTo
        .lngPhraseId = lngPhraseId
        .strNotes = strNotes
        .lngLocationId = lngLocationId
    End With
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Categories_Add"
        .Parameters("@lngLocationId") = lngLocationId
        .Parameters("@strCategory") = strCategory
        .Parameters("@lngPriceFrom") = lngPriceFrom
        .Parameters("@lngPriceTo") = lngPriceTo
         .Execute
        oTemp.lngCategoryId = .Parameters("@lngCategoryId")
        
        .CommandText = "CategoriesPhrases_Add"
        .Parameters("@lngCategoryId") = oTemp.lngCategoryId
        .Parameters("@lngPhraseId") = lngPhraseId
        .Execute
        
    End With
    
    If oConn.State = 1 Then oConn.Close
     
    mCol.Add oTemp, CStr(oTemp.lngCategoryId)

    'return the id created
    Add = oTemp.lngCategoryId
    Set oTemp = Nothing

End Function

Public Property Get Item(lngCategoryId As Long) As Category
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(CStr(lngCategoryId))
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Function Delete(ByVal lngCategoryId As Long) As Long
Dim lngPhraseId As Long
On Error GoTo EH:

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "CategoriesPhrases_Fetch"
        .Parameters("@lngCategoryId") = lngCategoryId
        .Execute
         lngPhraseId = .Parameters("lngPhraseId")
        
        .CommandText = "FilesPhrases_delete"
        .Parameters("@lngFileId") = 3
        .Parameters("lngPhraseId") = lngPhraseId
        .Execute
        
        .CommandText = "Categories_Delete"
        .Parameters("@lngCategoryId") = lngCategoryId
        .Execute
        
    End With

    mCol.Remove CStr(lngCategoryId)
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

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

