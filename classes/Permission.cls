VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Permission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private mvarlngPermissionId As Long
Private mvarstrPermission As String
Public Property Let strPermission(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strPermission = 5
    mvarstrPermission = vData
End Property


Public Property Get strPermission() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strPermission
    strPermission = mvarstrPermission
End Property



Public Property Let lngPermissionId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngPermissionId = 5
    mvarlngPermissionId = vData
End Property


Public Property Get lngPermissionId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngPermissionId
    lngPermissionId = mvarlngPermissionId
End Property



