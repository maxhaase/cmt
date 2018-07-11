VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarlngCustomerId As Long
Private mvarstrFirstName As String
Private mvarstrLastName As String
Private mvarstrCompany As String
Private mvarstrAddress As String
Private mvarstrCity As String
Private mvarstrState As String
Private mvarstrZip As String
Private mvarstrPhone As String
Private mvarstrFax As String
Private mvarstrEmail As String
Private mvarstrOrganizationNr As String
Private mvarstrCountry As String
Private mvarstrNotes As String
Private mvarstrWebSite As String

Public Property Let strWebSite(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strWebSite = 5
    mvarstrWebSite = vData
End Property


Public Property Get strWebSite() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strWebSite
    strWebSite = mvarstrWebSite
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



Public Property Let strOrganizationNr(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strOrganizationNr = 5
    mvarstrOrganizationNr = vData
End Property


Public Property Get strOrganizationNr() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strOrganizationNr
    strOrganizationNr = mvarstrOrganizationNr
End Property



Public Property Let strEmail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strEmail = 5
    mvarstrEmail = vData
End Property


Public Property Get strEmail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strEmail
    strEmail = mvarstrEmail
End Property



Public Property Let strFax(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strFax = 5
    mvarstrFax = vData
End Property


Public Property Get strFax() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strFax
    strFax = mvarstrFax
End Property



Public Property Let strPhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strPhone = 5
    mvarstrPhone = vData
End Property


Public Property Get strPhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strPhone
    strPhone = mvarstrPhone
End Property



Public Property Let strZip(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strZip = 5
    mvarstrZip = vData
End Property


Public Property Get strZip() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strZip
    strZip = mvarstrZip
End Property



Public Property Let strState(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strState = 5
    mvarstrState = vData
End Property


Public Property Get strState() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strState
    strState = mvarstrState
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



Public Property Let strCompany(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strCompany = 5
    mvarstrCompany = vData
End Property


Public Property Get strCompany() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strCompany
    strCompany = mvarstrCompany
End Property



Public Property Let strLastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strLastName = 5
    mvarstrLastName = vData
End Property


Public Property Get strLastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strLastName
    strLastName = mvarstrLastName
End Property



Public Property Let strFirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strFirstName = 5
    mvarstrFirstName = vData
End Property


Public Property Get strFirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strFirstName
    strFirstName = mvarstrFirstName
End Property



Public Property Let lngCustomerId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCustomerId = 5
    mvarlngCustomerId = vData
End Property


Public Property Get lngCustomerId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCustomerId
    lngCustomerId = mvarlngCustomerId
End Property



