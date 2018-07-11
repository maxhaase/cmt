VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarlngInvoiceId As Long
Private mvarstrName As String
Private mvarstrPhone As String
Private mvarstrFax As String
Private mvarstrEmail As String
Private mvarstrOrganizationNumber As String
Private mvarlngTotalAmount As Currency
Private mvarbitValid As Boolean

Public Property Let bitValid(ByVal vData As Boolean)
    mvarbitValid = vData
End Property

Public Property Get bitValid() As Boolean
    bitValid = mvarbitValid
End Property

Public Property Let lngTotalAmount(ByVal vData As Currency)
    mvarlngTotalAmount = vData
End Property

Public Property Get lngTotalAmount() As Currency
    lngTotalAmount = mvarlngTotalAmount
End Property

Public Property Let strOrganizationNumber(ByVal vData As String)
    mvarstrOrganizationNumber = vData
End Property

Public Property Get strOrganizationNumber() As String
    strOrganizationNumber = mvarstrOrganizationNumber
End Property

Public Property Let strEmail(ByVal vData As String)
    mvarstrEmail = vData
End Property


Public Property Get strEmail() As String
    strEmail = mvarstrEmail
End Property

Public Property Let strFax(ByVal vData As String)
    mvarstrFax = vData
End Property

Public Property Get strFax() As String
    strFax = mvarstrFax
End Property

Public Property Let strPhone(ByVal vData As String)
    mvarstrPhone = vData
End Property
Public Property Get strPhone() As String
    strPhone = mvarstrPhone
End Property



Public Property Let strName(ByVal vData As String)
    mvarstrName = vData
End Property


Public Property Get strName() As String
    strName = mvarstrName
End Property



Public Property Let lngInvoiceId(ByVal vData As Long)
    mvarlngInvoiceId = vData
End Property


Public Property Get lngInvoiceId() As Long
    lngInvoiceId = mvarlngInvoiceId
End Property



