VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phrase"
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

Private mvarlngPhraseId As Long
Private mvarstrPhrase As String

Public Property Let strPhrase(ByVal vData As String)
    mvarstrPhrase = vData
End Property

Public Property Get strPhrase() As String
    strPhrase = mvarstrPhrase
End Property

Public Property Let lngPhraseId(ByVal vData As Long)
    mvarlngPhraseId = vData
End Property

Public Property Get lngPhraseId() As Long
    lngPhraseId = mvarlngPhraseId
End Property



