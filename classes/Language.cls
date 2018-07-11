VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Language"
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

Private mvarlngLanguageId As Long
Private mvarlngPhraseId As Long

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



Public Property Let lngLanguageId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngLanguageId = 5
    mvarlngLanguageId = vData
End Property


Public Property Get lngLanguageId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngLanguageId
    lngLanguageId = mvarlngLanguageId
End Property



