VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarstrCategory As String
Private mvarlngPriceFrom As Long
Private mvarlngPriceTo As Long
Private mvarstrNotes As String
Private mvarlngCategoryId As Long
Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mvarstrConnection As String
Private mvarlngPhraseId As Long
Private mvarlngLocationId As Long
Private mvarstrLocation As String

Public Property Let strLocation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strLocation = 5
    mvarstrLocation = vData
End Property


Public Property Get strLocation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strLocation
    strLocation = mvarstrLocation
End Property



Public Property Let lngLocationId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngLocationId = 5
    mvarlngLocationId = vData
End Property


Public Property Get lngLocationId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngLocationId
    lngLocationId = mvarlngLocationId
End Property




Public Property Let lngPhraseId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngPhraseId = 5
    mvarlngPhraseId = vData
End Property

Public Property Get lngPhraseId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngPhraseId
    lngPhraseId = mvarlngPhraseId
End Property

Public Function AddImage(strImage As String) As Long
  Dim oTemp As Image
    Set oTemp = New Image
    
    oTemp.strImage = strImage
    
    If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Images_Add"
            .Parameters("@strImage") = oTemp.strImage
            .Execute
             oTemp.lngImageId = .Parameters("@lngImageId")
            .CommandText = "CategoriesImages_Add"
            .Parameters("@lngImageId") = oTemp.lngImageId
            .Parameters("@lngCategoryId") = Me.lngCategoryId
            .Execute
        End With
        
    If oConn.State = 1 Then oConn.Close
    
    AddImage = oTemp.lngImageId  'return the Image id created
    Set oTemp = Nothing


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

Public Function GetImages() As Collection
Dim oTemp, oCol
Set oCol = New Collection

  If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Images_Fetch"
        .Parameters("@lngCategoryId") = Me.lngCategoryId
        .Parameters("@lngLocationId") = 0
        .Parameters("@lngHouseId") = 0
        Set oRs = .Execute
        
        While Not oRs.EOF
            Set oTemp = New Image
            oTemp.lngImageId = .Parameters("@lngImageId")
            oTemp.strImage = .Parameters("@strImage")
            oTemp.lngPhotoId = .Parameters("@lngPhotoId")
            oTemp.strPhoto = .Parameters("@strPhoto")
            oCol.Add oTemp, CStr(oTemp.lngImageId)
            oRs.MoveNext
            Set oTemp = Nothing
        Wend
                 
    End With
        
    If oConn.State = 1 Then oConn.Close
    Set GetImages = oCol
    Set oCol = Nothing


End Function


Public Property Let lngCategoryId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCategoryId = 5
    mvarlngCategoryId = vData
End Property


Public Property Get lngCategoryId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCategoryId
    lngCategoryId = mvarlngCategoryId
End Property



Public Property Let strNotes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strNotes = 5
    mvarstrNotes = vData
End Property


Public Property Get strNotes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strNotes
    strNotes = mvarstrNotes
End Property



Public Property Let lngPriceTo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngPriceTo = 5
    mvarlngPriceTo = vData
End Property


Public Property Get lngPriceTo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngPriceTo
    lngPriceTo = mvarlngPriceTo
End Property



Public Property Let lngPriceFrom(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngPriceFrom = 5
    mvarlngPriceFrom = vData
End Property


Public Property Get lngPriceFrom() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngPriceFrom
    lngPriceFrom = mvarlngPriceFrom
End Property



Public Property Let strCategory(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strCategory = 5
    mvarstrCategory = vData
End Property


Public Property Get strCategory() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strCategory
    strCategory = mvarstrCategory
End Property



