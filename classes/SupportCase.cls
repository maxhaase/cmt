VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupportCase"
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

Private mvarstrCase As String
Private mvarlngCaseId As Long

Public Property Let lngCaseId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCaseId = 5
    mvarlngCaseId = vData
End Property


Public Property Get lngCaseId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCaseId
    lngCaseId = mvarlngCaseId
End Property



Public Property Let strCase(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strCase = 5
    mvarstrCase = vData
End Property


Public Property Get strCase() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strCase
    strCase = mvarstrCase
End Property



