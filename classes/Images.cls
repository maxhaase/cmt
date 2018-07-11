VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Images"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Image"
Attribute VB_Ext_KEY = "Member0" ,"Image"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


' *****************************************************************************
' SMART WEB API 1.0
' Universal language e-commerce library
' Programmer:   Max Haase   maxhaase@gmail.com
' December 2000
' *****************************************************************************


Private mCol As Collection
Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mvarlngError As Long
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

Public Function Fetch(Optional lngHouseId As Long = 0, Optional lngLocationId As Long = 0, Optional lngCategoryId As Long = 0) As Long
    
    Dim oTemp As Image
    Set mCol = Nothing
    Set mCol = New Collection
    
    If oConn.State = 0 Then oConn.Open strConnection
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .CommandText = "Images_Fetch"
        .Parameters("@lngCategoryId") = lngCategoryId
        .Parameters("@lngLocationId") = lngLocationId
        .Parameters("@lngHouseId") = lngHouseId
        Set oRs = .Execute
        
        While Not oRs.EOF
            Set oTemp = New Image
            If Not IsNull(oRs("lngImageId")) Then oTemp.lngImageId = oRs("lngImageId")
            If Not IsNull(oRs("strImage")) Then oTemp.strImage = oRs("strImage")
            If Not IsNull(oRs("lngPhotoId")) Then oTemp.lngPhotoId = oRs("lngPhotoId")
            If Not IsNull(oRs("strPhoto")) Then oTemp.strPhoto = oRs("strPhoto")
            mCol.Add oTemp, CStr(oTemp.lngImageId)
            oRs.MoveNext
            Set oTemp = Nothing
        Wend
                 
    End With
        
    If oConn.State = 1 Then oConn.Close

End Function

Public Function Update(ByVal lngImageId As Long, ByVal strImage As String) As Long
On Error GoTo EH:

        If oConn.State = 0 Then oConn.Open strConnection
            
            With oCmd
                Set .ActiveConnection = oConn
                .CommandType = adCmdStoredProc
                .CommandText = "Images_update"
                .Parameters("@strImage") = strImage
                .Parameters("@lngImageId") = lngImageId
                Set oRs = .Execute
                Update = 0
            End With
   
        If oConn.State = 1 Then oConn.Close
        Exit Function

EH:
        Update = 5
        If oConn.State = 1 Then oConn.Close

End Function

Public Function Delete(ByVal lngImageId As Long, Optional lngPhotoId As Long = 0) As Long
On Error GoTo EH:

        If oConn.State = 0 Then oConn.Open strConnection
        
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Images_delete"
            .Parameters("@lngImageId") = lngImageId
            .Parameters("@lngPhotoId") = lngPhotoId
            Set oRs = .Execute
            mCol.Remove CStr(lngImageId)
            Delete = 0
         End With
         If oConn.State = 1 Then oConn.Close
         Exit Function
EH:
    Delete = 5
    If oConn.State = 1 Then oConn.Close
End Function

Public Property Let lngError(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngError = 5
    mvarlngError = vData
End Property


Public Property Get lngError() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngError
    lngError = mvarlngError
End Property




Public Function Add(strImage As String, Optional lngHouseId As Long = 0, Optional lngLocationId As Long = 0, Optional lngCategoryId As Long = 0) As Long
    
    Dim oTemp As Image 'create a new object
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
            mCol.Add oTemp, CStr(oTemp.lngImageId)
            
            If lngHouseId <> 0 Then 'this image belongs to a house
                .CommandText = "HousesImages_Add"
                .Parameters("@lngImageId") = oTemp.lngImageId
                .Parameters("@lngHouseId") = lngHouseId
                .Execute
            End If
            
            If lngLocationId <> 0 Then 'this image belongs to a location
                .CommandText = "LocationsImages_Add"
                .Parameters("@lngImageId") = oTemp.lngImageId
                .Parameters("@lngLocationId") = lngLocationId
                .Execute
            End If
            
            If lngCategoryId <> 0 Then 'this image belongs to a category
                .CommandText = "CategoriesImages_Add"
                .Parameters("@lngImageId") = oTemp.lngImageId
                .Parameters("@lngCategoryId") = lngCategoryId
                .Execute
            End If
           
        End With
        
    If oConn.State = 1 Then oConn.Close
    
    Add = oTemp.lngImageId  'return the Image id created
    Set oTemp = Nothing

End Function

Public Property Get Item(lngImageId As Long) As Image
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(CStr(lngImageId))
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

