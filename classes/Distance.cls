VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Distance"
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

Private mvardblDistance As Double
Private mvarstrUnit As String
Public Property Let strUnit(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strUnit = 5
    mvarstrUnit = vData
End Property


Public Property Get strUnit() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strUnit
    strUnit = mvarstrUnit
End Property



Public Property Let dblDistance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dblDistance = 5
    mvardblDistance = vData
End Property


Public Property Get dblDistance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dblDistance
    dblDistance = mvardblDistance
End Property



