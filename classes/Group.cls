VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
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


Private mvarlngGroupId As Long
Private mvarstrGroupName As String
Public Property Let strGroupName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strGroupName = 5
    mvarstrGroupName = vData
End Property


Public Property Get strGroupName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strGroupName
    strGroupName = mvarstrGroupName
End Property



Public Property Let lngGroupId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngGroupId = 5
    mvarlngGroupId = vData
End Property


Public Property Get lngGroupId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngGroupId
    lngGroupId = mvarlngGroupId
End Property



