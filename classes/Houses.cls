VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Houses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"House"
Attribute VB_Ext_KEY = "Member0" ,"House"
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

Public Function Fetch(Optional lngLanguageId As Long = 1, Optional lngCategoryId As Long = -1, Optional lngLocationId As Long = -1, Optional lngTypeId As Long = -1, Optional strTerms As String = "-1", Optional lngBedrooms As Long = -1, Optional lngPrice As Long = -1)

Set mCol = Nothing
Set mCol = New Collection
Dim oTemp As House
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Houses_Fetch"
        .Parameters("@lngLanguageId") = lngLanguageId
        .Parameters("@lngCategoryId") = lngCategoryId
        .Parameters("@lngLocationId") = lngLocationId
        .Parameters("@lngTypeId") = lngTypeId
        .Parameters("@strTerms") = strTerms
        .Parameters("@lngBedrooms") = lngBedrooms
        .Parameters("@lngPrice") = lngPrice
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New House
        With oTemp
            .lngHouseId = oRs("lngHouseId")
            .lngCategoryId = oRs("lngCategoryId")
            .lngLocationId = oRs("lngLocationId")
            .lngTypeId = oRs("lngTypeId")
            .strTerms = oRs("strTerms")
            .lngPrice = oRs("lngPrice")
            .lngArea = oRs("lngArea")
            .lngBedrooms = oRs("lngBedrooms")
            .lngBathrooms = oRs("lngBathrooms")
            .blGarage = oRs("blGarage")
            .blRoofTerrace = oRs("blRoofTerrace")
            .blSwimmingPool = oRs("blSwimmingPool")
            .lngPatioArea = oRs("lngPatioArea")
            .lngDistanceBeach = oRs("lngDistanceBeach")
            .lngDistanceGolf = oRs("lngDistanceGolf")
            .lngDistanceAirport = oRs("lngDistanceAirport")
            .lngDistanceCentrum = oRs("lngDistanceCentrum")
            .lngYearBuilt = oRs("lngYearBuilt")
            .strNotes = oRs("strPhrase")
            .strAddress = oRs("strAddress")
            .strCity = oRs("strCity")
            .strProvince = oRs("strProvince")
            .strCountry = oRs("strCountry")
            .strOwnerName = oRs("strOwnerName")
            .strOwnerLastName = oRs("strOwnerLastName")
            .strOwnerPhone = oRs("strOwnerPhone")
            .strOwnerFax = oRs("strOwnerFax")
            .strOwnerEmail = oRs("strOwnerEmail")
            .blFurniture = oRs("blFurniture")
            .lngFloor = oRs("lngFloor")
            .lngFloorCount = oRs("lngFloorCount")
            
        End With

        mCol.Add oTemp, CStr(oTemp.lngHouseId)
        oRs.MoveNext
        Set oTemp = Nothing
    Wend
    
    Set oRs = Nothing
    
If oConn.State = 1 Then oConn.Close


End Function

Public Function Update(lngHouseId As Long, lngCategoryId As Long, lngLocationId As Long, lngTypeId As Long, strTerms As String, lngPrice As Long, lngArea As Long, lngBedrooms As Long, lngBathrooms As Long, blGarage As Boolean, blRoofTerrace As Boolean, blSwimmingPool As Boolean, lngPatioArea As Long, lngDistanceBeach As Long, lngDistanceGolf As Long, lngDistanceAirport As Long, lngDistanceCentrum As Long, lngYearBuilt As Long, strNotes As String, strAddress As String, strCity As String, strProvince As String, strCountry As String, strOwnerName As String, strOwnerLastName As String, strOwnerPhone As String, strOwnerFax As String, strOwnerEmail As String, blFurniture As Boolean, lngFloorCount As Long, lngFloor As Long) As Long
Dim lngPhraseId As Long
On Error GoTo EH:
    
        If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
                        
            .CommandText = "HousesPhrases_Fetch" 'Get the phraseid
            .Parameters("@lngHouseId") = lngHouseId
            .Execute
            lngPhraseId = .Parameters("@lngPhraseId")
            
            .CommandText = "Phrases_update"
            .Parameters("@strPhrase") = strNotes
            .Parameters("@lngLanguageId") = 1
            .Parameters("@lngPhraseId") = lngPhraseId
            .Execute
            
            .CommandText = "Houses_update"
            .Parameters("@lngHouseId") = lngHouseId
            .Parameters("@blGarage") = blGarage
            .Parameters("@blRoofTerrace") = blRoofTerrace
            .Parameters("@blSwimmingPool") = blSwimmingPool
            .Parameters("@lngArea") = lngArea
            .Parameters("@lngBathrooms") = lngBathrooms
            .Parameters("@lngBedrooms") = lngBedrooms
            .Parameters("@lngCategoryId") = lngCategoryId
            .Parameters("@lngDistanceAirport") = lngDistanceAirport
            .Parameters("@lngDistanceBeach") = lngDistanceBeach
            .Parameters("@lngDistanceCentrum") = lngDistanceCentrum
            .Parameters("@lngDistanceGolf") = lngDistanceGolf
            .Parameters("@lngLocationId") = lngLocationId
            .Parameters("@strTerms") = strTerms
            .Parameters("@lngPatioArea") = lngPatioArea
            .Parameters("@lngPrice") = lngPrice
            .Parameters("@lngTypeId") = lngTypeId
            .Parameters("@lngYearBuilt") = lngYearBuilt
            .Parameters("@strAddress") = strAddress
            .Parameters("@strCity") = strCity
            .Parameters("@strCountry") = strCountry
            .Parameters("@strNotes") = strNotes
            .Parameters("@strOwnerEmail") = strOwnerEmail
            .Parameters("@strOwnerFax") = strOwnerFax
            .Parameters("@strOwnerLastName") = strOwnerLastName
            .Parameters("@strOwnerName") = strOwnerName
            .Parameters("@strOwnerPhone") = strOwnerPhone
            .Parameters("@strProvince") = strProvince
            .Parameters("@blFurniture") = blFurniture
            .Parameters("@lngFloorCount") = lngFloorCount
            .Parameters("@lngFloor") = lngFloor
            .Execute
            Update = 0
        End With
  
    If oConn.State = 1 Then oConn.Close
    Exit Function
    
EH:
Update = 5
 
End Function
Public Function Add(lngCategoryId As Long, lngLocationId As Long, lngTypeId As Long, strTerms As String, lngPrice As Long, lngArea As Long, lngBedrooms As Long, lngBathrooms As Long, blGarage As Boolean, blRoofTerrace As Boolean, blSwimmingPool As Boolean, lngPatioArea As Long, lngDistanceBeach As Long, lngDistanceGolf As Long, lngDistanceAirport As Long, lngDistanceCentrum As Long, lngYearBuilt As Long, strNotes As String, strAddress As String, strCity As String, strProvince As String, strCountry As String, strOwnerName As String, strOwnerLastName As String, strOwnerPhone As String, strOwnerFax As String, strOwnerEmail As String, blFurniture As Boolean, lngFloorCount As Long, lngFloor As Long, Optional lngHouseId As Long, Optional lngLanguageId As Long = 1) As Long
    
    Dim oTemp As House, lngPhraseId As Long, oLanguage As Language, oLanguages As Languages
    Set oTemp = New House
    
    With oTemp
        .lngHouseId = lngHouseId
        .lngCategoryId = lngCategoryId
        .lngLocationId = lngLocationId
        .lngTypeId = lngTypeId
        .strTerms = strTerms
        .lngPrice = lngPrice
        .lngArea = lngArea
        .lngBedrooms = lngBedrooms
        .lngBathrooms = lngBathrooms
        .blGarage = blGarage
        .blRoofTerrace = blRoofTerrace
        .blSwimmingPool = blSwimmingPool
        .lngPatioArea = lngPatioArea
        .lngDistanceBeach = lngDistanceBeach
        .lngDistanceGolf = lngDistanceGolf
        .lngDistanceAirport = lngDistanceAirport
        .lngDistanceCentrum = lngDistanceCentrum
        .lngYearBuilt = lngYearBuilt
        .strNotes = strNotes
        .strAddress = strAddress
        .strCity = strCity
        .strProvince = strProvince
        .strCountry = strCountry
        .strOwnerName = strOwnerName
        .strOwnerLastName = strOwnerLastName
        .strOwnerPhone = strOwnerPhone
        .strOwnerFax = strOwnerFax
        .strOwnerEmail = strOwnerEmail
        .blFurniture = blFurniture
        .lngFloorCount = lngFloorCount
        .lngFloor = lngFloor
    End With
    '****************************************************** DATABASE
    
        If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Phrases_add" 'add the text about the house
            .Parameters("@lngLanguageId") = 1
            .Parameters("@strPhrase") = strNotes
            .Parameters("@lngFileId") = 1 'file number 1 is used for translating house texts
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
             
            .CommandText = "Houses_Add" 'add the house
            .Parameters("@blGarage") = oTemp.blGarage
            .Parameters("@blRoofTerrace") = oTemp.blRoofTerrace
            .Parameters("@blSwimmingPool") = oTemp.blSwimmingPool
            .Parameters("@lngArea") = oTemp.lngArea
            .Parameters("@lngBathrooms") = oTemp.lngBathrooms
            .Parameters("@lngBedrooms") = oTemp.lngBedrooms
            .Parameters("@lngCategoryId") = oTemp.lngCategoryId
            .Parameters("@lngDistanceAirport") = oTemp.lngDistanceAirport
            .Parameters("@lngDistanceBeach") = oTemp.lngDistanceBeach
            .Parameters("@lngDistanceCentrum") = oTemp.lngDistanceCentrum
            .Parameters("@lngDistanceGolf") = oTemp.lngDistanceGolf
            .Parameters("@lngLocationId") = oTemp.lngLocationId
            .Parameters("@strTerms") = oTemp.strTerms
            .Parameters("@lngPatioArea") = oTemp.lngPatioArea
            .Parameters("@lngPrice") = oTemp.lngPrice
            .Parameters("@lngTypeId") = oTemp.lngTypeId
            .Parameters("@lngYearBuilt") = oTemp.lngYearBuilt
            .Parameters("@strAddress") = oTemp.strAddress
            .Parameters("@strCity") = oTemp.strCity
            .Parameters("@strCountry") = oTemp.strCountry
            .Parameters("@strNotes") = oTemp.strNotes
            .Parameters("@strOwnerEmail") = oTemp.strOwnerEmail
            .Parameters("@strOwnerFax") = oTemp.strOwnerFax
            .Parameters("@strOwnerLastName") = oTemp.strOwnerLastName
            .Parameters("@strOwnerName") = oTemp.strOwnerName
            .Parameters("@strOwnerPhone") = oTemp.strOwnerPhone
            .Parameters("@strProvince") = oTemp.strProvince
            .Parameters("@blFurniture") = oTemp.blFurniture
            .Parameters("@lngFloorCount") = oTemp.lngFloorCount
            .Parameters("@lngFloor") = oTemp.lngFloor
            .Execute
             oTemp.lngHouseId = .Parameters("@lngHouseId")
            .CommandText = "HousesPhrases_Add" 'Assign the phrase to the house
            .Parameters("@lngHouseId") = oTemp.lngHouseId
            .Parameters("@lngPhraseId") = lngPhraseId
            .Execute
            
        End With
        
    If oConn.State = 1 Then oConn.Close
    
           
    '****************************************************************
    
    
    mCol.Add oTemp, CStr(oTemp.lngHouseId)
    
    'return the house id created
    Add = oTemp.lngHouseId
    Set oTemp = Nothing


End Function

Public Property Get Item(lngHouseId As Long) As House
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(CStr(lngHouseId))
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Function Delete(lngHouseId As Long) As Long
Dim lngPhraseId As Long
'On Error GoTo EH:

   If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        
        .CommandText = "HousesPhrases_Fetch" 'Get the phraseid
        .Parameters("@lngHouseId") = lngHouseId
        .Execute
        lngPhraseId = .Parameters("@lngPhraseId")
        
     
        .CommandText = "HousesPhrases_Delete"
        .Parameters("@lngHouseId") = lngHouseId
        .Parameters("@lngPhraseId") = lngPhraseId
        .Execute
        
        .CommandText = "FilesPhrases_delete"
        .Parameters("@lngFileId") = 1
        .Parameters("@lngPhraseId") = lngPhraseId
        .Execute
       
       
        .CommandText = "HousesImages_delete"
        .Parameters("@lngHouseId") = lngHouseId
        .Execute
       
        .CommandText = "Houses_Delete"
        .Parameters("@lngHouseId") = lngHouseId
        .Execute
        
    End With

    mCol.Remove CStr(lngHouseId)
    If oConn.State = 1 Then oConn.Close
    Delete = 0
    Exit Function
    
'EH:
   ' Delete = 5
   ' If oConn.State = 1 Then oConn.Close
        
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

