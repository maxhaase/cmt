VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Image"
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
' December 2000
' *****************************************************************************

Private mvarlngImageId As Long
Private mvarstrImage As String
Private mvarlngPhotoId As Long
Private mvarstrPhoto As String

Private oRs As New ADODB.Recordset
Private oCmd As New ADODB.Command
Private oConn As New ADODB.Connection
Private mvarlngError As Long
Private mvarstrConnection As String

Public Function AddPhoto(strPhoto As String) As Long
       
    If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Photos_Add"
            .Parameters("@strPhoto") = strPhoto
            .Execute
             Me.lngPhotoId = .Parameters("@lngPhotoId")
            .CommandText = "PhotosImages_Add"
            .Parameters("@lngPhotoId") = Me.lngPhotoId
            .Parameters("@lngImageId") = Me.lngImageId
            .Execute
           
        End With
        
    If oConn.State = 1 Then oConn.Close
    
    AddPhoto = Me.lngPhotoId  'return the Photo id created
    
End Function

Public Function RemovePhoto() As Long

    If oConn.State = 0 Then oConn.Open strConnection
    
        With oCmd
            Set .ActiveConnection = oConn
            .CommandType = adCmdStoredProc
            .CommandText = "Photos_Delete"
            .Parameters("@lngPhotoId") = Me.lngPhotoId
            .Execute
            .CommandText = "PhotosImages_Delete"
            .Parameters("@lngPhotoId") = Me.lngPhotoId
            .Parameters("@lngImageId") = Me.lngImageId
            .Execute
        End With
        
    If oConn.State = 1 Then oConn.Close
    
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

Public Property Let strPhoto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strPhoto = 5
    mvarstrPhoto = vData
End Property

Public Property Get strPhoto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strPhoto
    strPhoto = mvarstrPhoto
End Property

Public Property Let lngPhotoId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngPhotoId = 5
    mvarlngPhotoId = vData
End Property

Public Property Get lngPhotoId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngPhotoId
    lngPhotoId = mvarlngPhotoId
End Property

Public Property Let strImage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strImage = 5
    mvarstrImage = vData
End Property

Public Property Get strImage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strImage
    strImage = mvarstrImage
End Property

Public Property Let lngImageId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngImageId = 5
    mvarlngImageId = vData
End Property

Public Property Get lngImageId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngImageId
    lngImageId = mvarlngImageId
End Property



