VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "House"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' *****************************************************************************
' SMART WEB API 1.0
' Universal language e-commerce library
' Programmer:   Max Haase   maxhaase@gmail.com
' November 2000
' *****************************************************************************

Private mvarlngHouseId As Long
Private mvarlngCategoryId As Long
Private mvarlngLocationId As Long
Private mvarlngTypeId As Long
Private mvarstrTerms As String
Private mvarlngPrice As Long
Private mvarlngArea As Long
Private mvarlngBedrooms As Long
Private mvarlngBathrooms As Long
Private mvarblGarage As Boolean
Private mvarblRoofTerrace As Boolean
Private mvarblSwimmingPool As Boolean
Private mvarlngPatioArea As Long
Private mvarlngDistanceBeach As Long
Private mvarlngDistanceGolf As Long
Private mvarlngDistanceAirport As Long
Private mvarlngDistanceCentrum As Long
Private mvarlngYearBuilt As Long
Private mvarstrNotes As String
Private mvarstrAddress As String
Private mvarstrCity As String
Private mvarstrProvince As String
Private mvarstrCountry As String
Private mvarstrOwnerName As String
Private mvarstrOwnerLastName As String
Private mvarstrOwnerPhone As String
Private mvarstrOwnerFax As String
Private mvarstrOwnerEmail As String
Private mvarblFurniture As Boolean

Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mvarstrConnection As String
Private mvarlngPhraseId As Long
Private mvarlngFloor As Long
Private mvarlngFloorCount As Long
Public Property Let lngFloorCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngFloorCount = 5
    mvarlngFloorCount = vData
End Property


Public Property Get lngFloorCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngFloorCount
    lngFloorCount = mvarlngFloorCount
End Property



Public Property Let lngFloor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngFloor = 5
    mvarlngFloor = vData
End Property


Public Property Get lngFloor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngFloor
    lngFloor = mvarlngFloor
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
            .CommandText = "HousesImages_Add"
            .Parameters("@lngImageId") = oTemp.lngImageId
            .Parameters("@lngHouseId") = Me.lngHouseId
            .Execute
        End With
        
    If oConn.State = 1 Then oConn.Close
    
    AddImage = oTemp.lngImageId  'return the Image id created
    Set oTemp = Nothing


End Function

Public Function GetImages() As Collection
Dim oTemp, oCol
Set oCol = New Collection

  If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Images_Fetch"
        .Parameters("@lngCategoryId") = 0
        .Parameters("@lngLocationId") = 0
        .Parameters("@lngHouseId") = Me.lngHouseId
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

Public Property Let blFurniture(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blFurniture = 5
    mvarblFurniture = vData
End Property


Public Property Get blFurniture() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blFurniture
    blFurniture = mvarblFurniture
End Property




Public Property Let strOwnerEmail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOwnerEmail = 5
    mvarstrOwnerEmail = vData
End Property


Public Property Get strOwnerEmail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOwnerEmail
    strOwnerEmail = mvarstrOwnerEmail
End Property



Public Property Let strOwnerFax(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOwnerFax = 5
    mvarstrOwnerFax = vData
End Property


Public Property Get strOwnerFax() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOwnerFax
    strOwnerFax = mvarstrOwnerFax
End Property



Public Property Let strOwnerPhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOwnerPhone = 5
    mvarstrOwnerPhone = vData
End Property


Public Property Get strOwnerPhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOwnerPhone
    strOwnerPhone = mvarstrOwnerPhone
End Property



Public Property Let strOwnerLastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOwnerLastName = 5
    mvarstrOwnerLastName = vData
End Property


Public Property Get strOwnerLastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOwnerLastName
    strOwnerLastName = mvarstrOwnerLastName
End Property



Public Property Let strOwnerName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOwnerName = 5
    mvarstrOwnerName = vData
End Property


Public Property Get strOwnerName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOwnerName
    strOwnerName = mvarstrOwnerName
End Property



Public Property Let strCountry(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strCountry = 5
    mvarstrCountry = vData
End Property


Public Property Get strCountry() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strCountry
    strCountry = mvarstrCountry
End Property



Public Property Let strProvince(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strProvince = 5
    mvarstrProvince = vData
End Property


Public Property Get strProvince() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strProvince
    strProvince = mvarstrProvince
End Property



Public Property Let strCity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strCity = 5
    mvarstrCity = vData
End Property


Public Property Get strCity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strCity
    strCity = mvarstrCity
End Property



Public Property Let strAddress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strAddress = 5
    mvarstrAddress = vData
End Property


Public Property Get strAddress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strAddress
    strAddress = mvarstrAddress
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



Public Property Let lngYearBuilt(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngYearBuilt = 5
    mvarlngYearBuilt = vData
End Property


Public Property Get lngYearBuilt() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngYearBuilt
    lngYearBuilt = mvarlngYearBuilt
End Property



Public Property Let lngDistanceCentrum(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngDistanceCentrum = 5
    mvarlngDistanceCentrum = vData
End Property


Public Property Get lngDistanceCentrum() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngDistanceCentrum
    lngDistanceCentrum = mvarlngDistanceCentrum
End Property



Public Property Let lngDistanceAirport(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngDistanceAirport = 5
    mvarlngDistanceAirport = vData
End Property


Public Property Get lngDistanceAirport() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngDistanceAirport
    lngDistanceAirport = mvarlngDistanceAirport
End Property



Public Property Let lngDistanceGolf(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngDistanceGolf = 5
    mvarlngDistanceGolf = vData
End Property


Public Property Get lngDistanceGolf() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngDistanceGolf
    lngDistanceGolf = mvarlngDistanceGolf
End Property



Public Property Let lngDistanceBeach(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngDistanceBeach = 5
    mvarlngDistanceBeach = vData
End Property


Public Property Get lngDistanceBeach() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngDistanceBeach
    lngDistanceBeach = mvarlngDistanceBeach
End Property



Public Property Let lngPatioArea(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngPatioArea = 5
    mvarlngPatioArea = vData
End Property


Public Property Get lngPatioArea() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngPatioArea
    lngPatioArea = mvarlngPatioArea
End Property



Public Property Let blSwimmingPool(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blSwimmingPool = 5
    mvarblSwimmingPool = vData
End Property


Public Property Get blSwimmingPool() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blSwimmingPool
    blSwimmingPool = mvarblSwimmingPool
End Property



Public Property Let blRoofTerrace(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blRoofTerrace = 5
    mvarblRoofTerrace = vData
End Property


Public Property Get blRoofTerrace() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blRoofTerrace
    blRoofTerrace = mvarblRoofTerrace
End Property



Public Property Let blGarage(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blGarage = 5
    mvarblGarage = vData
End Property


Public Property Get blGarage() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blGarage
    blGarage = mvarblGarage
End Property



Public Property Let lngBathrooms(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngBathrooms = 5
    mvarlngBathrooms = vData
End Property


Public Property Get lngBathrooms() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngBathrooms
    lngBathrooms = mvarlngBathrooms
End Property



Public Property Let lngBedrooms(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngBedrooms = 5
    mvarlngBedrooms = vData
End Property


Public Property Get lngBedrooms() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngBedrooms
    lngBedrooms = mvarlngBedrooms
End Property



Public Property Let lngArea(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngArea = 5
    mvarlngArea = vData
End Property


Public Property Get lngArea() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngArea
    lngArea = mvarlngArea
End Property



Public Property Let lngPrice(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngPrice = 5
    mvarlngPrice = vData
End Property


Public Property Get lngPrice() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngPrice
    lngPrice = mvarlngPrice
End Property



Public Property Let strTerms(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strTerms = 5
    mvarstrTerms = vData
End Property


Public Property Get strTerms() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strTerms
    strTerms = mvarstrTerms
End Property



Public Property Let lngTypeId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngTypeId = 5
    mvarlngTypeId = vData
End Property


Public Property Get lngTypeId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngTypeId
    lngTypeId = mvarlngTypeId
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



Public Property Let lngHouseId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngHouseId = 5
    mvarlngHouseId = vData
End Property


Public Property Get lngHouseId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngHouseId
    lngHouseId = mvarlngHouseId
End Property



