VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Customer"
Attribute VB_Ext_KEY = "Member0" ,"Customer"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mCol As Collection
Private mvarstrConnection As String


Public Function Update(ByVal lngCustomerId As Long, strFirstName As String, strLastName As String, strCompany As String, strAddress As String, strCity As String, strState As String, strZip As String, strPhone As String, strFax As String, strEmail As String, strOrganizationNr As String, strCountry As String, strWebSite As String, strNotes As String) As Long
    On Error GoTo EH:
    If oConn.State = 0 Then oConn.Open strConnection
   
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Customers_Update"
            .Parameters("@lngCustomerId") = lngCustomerId
            .Parameters("@strFirstName") = strFirstName
            .Parameters("@strLastName") = strLastName
            .Parameters("@strCompany") = strCompany
            .Parameters("@strAddress") = strAddress
            .Parameters("@strCity") = strCity
            .Parameters("@strState") = strState
            .Parameters("@strZip") = strZip
            .Parameters("@strPhone") = strPhone
            .Parameters("@strFax") = strFax
            .Parameters("@strEmail") = strEmail
            .Parameters("@strOrganizationNr") = strOrganizationNr
            .Parameters("@strCountry") = strCountry
            .Parameters("@strWebSite") = strWebSite
            .Parameters("@strNotes") = strNotes
            .Execute
        End With
    
    Update = 0
    Exit Function
EH:
    Update = 5 'the smart error message

End Function
Public Function Fetch(Optional lngCustomerId As Long = -1)
    
Dim oTemp As Customer
Set mCol = Nothing
Set mCol = New Collection

    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
    Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Customers_Fetch"
        .Parameters("@lngCustomerId") = lngCustomerId
        Set oRs = .Execute
    End With
    
    While Not oRs.EOF
        Set oTemp = New Customer
        With oTemp
            .lngCustomerId = oRs("lngCustomerId")
            .strFirstName = oRs("strFirstName")
            .strLastName = oRs("strLastName")
            .strCompany = oRs("strCompany")
            .strAddress = oRs("strAddress")
            .strCity = oRs("strCity")
            .strState = oRs("strState")
            .strZip = oRs("strZip")
            .strPhone = oRs("strPhone")
            .strFax = oRs("strFax")
            .strEmail = oRs("strEmail")
            .strNotes = oRs("strNotes")
            .strOrganizationNr = oRs("strOrganizationNr")
        End With
        
        mCol.Add oTemp, CStr(oTemp.lngCustomerId)
        oRs.MoveNext
        Set oTemp = Nothing
        
    Wend

If oConn.State = 1 Then oConn.Close

End Function
Public Function Add(strFirstName As String, strLastName As String, strCompany As String, strAddress As String, strCity As String, strState As String, strZip As String, strPhone As String, strFax As String, strEmail As String, strOrganizationNr As String, strCountry As String, strWebSite As String, strNotes As String) As Customer
    
    Dim lngCustomerId As Long
    Dim oTemp As Customer
    Set oTemp = New Customer
    
    With oTemp
        .strFirstName = strFirstName
        .strLastName = strLastName
        .strCompany = strCompany
        .strAddress = strAddress
        .strCity = strCity
        .strState = strState
        .strZip = strZip
        .strPhone = strPhone
        .strFax = strFax
        .strEmail = strEmail
        .strOrganizationNr = strOrganizationNr
        .strCountry = strCountry
        .strWebSite = strWebSite
        .strNotes = strNotes
    End With

  If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Customers_Add"
            .Parameters("@strFirstName") = oTemp.strFirstName
            .Parameters("@strLastName") = oTemp.strLastName
            .Parameters("@strCompany") = oTemp.strCompany
            .Parameters("@strAddress") = oTemp.strAddress
            .Parameters("@strCity") = oTemp.strCity
            .Parameters("@strState") = oTemp.strState
            .Parameters("@strZip") = oTemp.strZip
            .Parameters("@strPhone") = oTemp.strPhone
            .Parameters("@strFax") = oTemp.strFax
            .Parameters("@strEmail") = oTemp.strEmail
            .Parameters("@strOrganizationNr") = oTemp.strOrganizationNr
            .Parameters("@strCountry") = oTemp.strCountry
            .Parameters("@strWebSite") = oTemp.strWebSite
            .Parameters("@strNotes") = oTemp.strNotes
            .Execute
            
            If Not IsNull(.Parameters("@lngCustomerId")) Then
                oTemp.lngCustomerId = .Parameters("@lngCustomerId")
                mCol.Add oTemp, CStr(oTemp.lngCustomerId)
            Else
                oTemp.lngCustomerId = 0
            End If
        End With
        
    If oConn.State = 1 Then oConn.Close
   
    'return the object created
    Set Add = oTemp
    Set oTemp = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Customer
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


Public Function Remove(ByVal lngCustomerId As Long) As Long
On Error GoTo EH:
  If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Customers_Delete"
        .Parameters("@lngCustomerId") = lngCustomerId
        .Execute
    End With
    mCol.Remove CStr(lngCustomerId)
    Remove = 0
    Exit Function
EH:
    Remove = 5
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

