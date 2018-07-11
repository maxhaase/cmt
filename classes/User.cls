VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
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


Private mvarlngUserId As Long
Private mvarstrUserName As String
Private mvarstrFirstName As String
Private mvarstrLastName As String
Private mvarstrEmail As String
Private mvarstrPassword As String

Public Property Let strPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strPassword = 5
    mvarstrPassword = vData
End Property


Public Property Get strPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strPassword
    strPassword = mvarstrPassword
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



Public Property Let strUserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strUserName = 5
    mvarstrUserName = vData
End Property


Public Property Get strUserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strUserName
    strUserName = mvarstrUserName
End Property



Public Property Let lngUserId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngUserId = 5
    mvarlngUserId = vData
End Property


Public Property Get lngUserId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngUserId
    lngUserId = mvarlngUserId
End Property



