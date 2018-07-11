VERSION 5.00
Begin VB.Form Test 
   Caption         =   "Test Container"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Fix Houses"
      Height          =   495
      Left            =   8160
      TabIndex        =   28
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Rentals"
      Height          =   495
      Left            =   8160
      TabIndex        =   27
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "UsedHouses"
      Height          =   495
      Left            =   8160
      TabIndex        =   26
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NewHouses"
      Height          =   495
      Left            =   8160
      TabIndex        =   25
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "House"
      Height          =   495
      Left            =   6120
      TabIndex        =   24
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Category"
      Height          =   495
      Left            =   6120
      TabIndex        =   23
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Images"
      Height          =   495
      Left            =   6120
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Distancias"
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Locations"
      Height          =   495
      Left            =   6120
      TabIndex        =   20
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton GetFilePhrases 
      Caption         =   "FILE PHRASES"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton btnPhrasesDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton btnPhrasesUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnPhrasesAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnPhrasesFetch 
      Caption         =   "Fetch"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame framePhrases 
      Caption         =   "Phrases"
      Height          =   2295
      Left            =   0
      TabIndex        =   12
      Top             =   4080
      Width           =   5295
      Begin VB.TextBox strPhrase 
         Height          =   1215
         Left            =   240
         TabIndex        =   17
         Text            =   "strPhrase"
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lbPhraseId 
         Caption         =   "lngPhraseId"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Add_Language 
      Caption         =   "Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton FetchLanguage 
      Caption         =   "Fetch"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox lngFileID 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Text            =   "lngFileId"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox strFileName 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Text            =   "strFileName"
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton FetchFile 
      Caption         =   "Fetch"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton DeleteFile 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton UpdateFile 
      Caption         =   "Update"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton AddFile 
      Caption         =   "Add"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame FrameFiles 
      Caption         =   "Files"
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5295
   End
   Begin VB.Frame FrameLanguage 
      Caption         =   "Languages"
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   5295
      Begin VB.TextBox lngPhraseId 
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Text            =   "lngPhraseId"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox lngLanguageId 
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Text            =   "lngLanguageId"
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strConnection As String

Private Sub Add_Language_Click()

Dim oLanguages As New Languages, oLanguage As Language, oPhrases As New Phrases, oPhrase As Phrase
Dim lngPhraseId As Long, lngResult As Long
oLanguages.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
oPhrases.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
oPhrases.Fetch 1

oLanguages.Fetch

lngPhraseId = oPhrases.Add("Spanish", 1, 1)

lngResult = oLanguages.Add(lngPhraseId)

MsgBox lngResult

End Sub

Private Sub btnPhrasesAdd_Click()
Dim oPhrases As New Phrases
Dim lngPhraseId As Long

oPhrases.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

lngPhraseId = oPhrases.Add(Test.strPhrase.Text, 1, 1)

MsgBox lngPhraseId

End Sub

Private Sub btnPhrasesDelete_Click()

Dim oPhrases As New Phrases
Dim lngResult As Long

oPhrases.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oPhrases.Fetch 1

lngResult = oPhrases.Delete(7)

MsgBox lngResult


End Sub

Private Sub btnPhrasesFetch_Click()
Dim oPhrases As New Phrases
Dim oPhrase As Phrase

oPhrases.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=INVESTCASABLANCA;Data Source=SUPERNOVA"

oPhrases.Fetch 1

For Each oPhrase In oPhrases
    MsgBox oPhrase.strPhrase
Next

MsgBox oPhrases.Item(7).strPhrase

End Sub

Private Sub btnPhrasesUpdate_Click()

Dim oPhrases As New Phrases
Dim lngResult As Long

oPhrases.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oPhrases.Fetch 1

lngResult = oPhrases.Update(7, Test.strPhrase.Text, 2)

MsgBox lngResult

End Sub



Private Sub Command1_Click()
Dim oLocation As Location, lngResult As Long, oImages As Collection
Dim oLocations As New Locations
oLocations.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
oLocations.FetchRentals
'oLocations.Item(9).strConnection = oLocations.strConnection

'Set oImages = oLocations.Item(9).GetImages

MsgBox oLocations.Count




End Sub

Private Sub Command2_Click()
Dim strDIstances
Dim oDistances As New SMART.Distances, oDistance As SMART.Distance
oDistances.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oDistances.Fetch

For Each oDistance In oDistances
    strDIstances = strDIstances & oDistance.lngDistance & vbCr
Next

MsgBox strDIstances



End Sub

Private Sub Command3_Click()
Dim oImages As New SMARTHOUSES.Images, oImage As SMARTHOUSES.Image, lngImageId As Long
oImages.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=FRANS;Data Source=supernova"
'lngImageId = oImages.Add("House1_Image_57845.16.jpg", 3)
'MsgBox lngResult
'oImages.Fetch 37

oImages.Fetch , , 1


For Each oImage In oImages
    strstring = strstring & oImage.strImage & "-" & oImage.strPhoto & vbCr
Next

MsgBox strstring



End Sub



Private Sub Command4_Click()
Dim oCategory As Category, oCategories As New Categories
Set oLanguages = New Languages
oCategories.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"


oCategories.Fetch 4, 1

For Each oCategory In oCategories

MsgBox oCategory.strCategory

Next




End Sub

Private Sub Command5_Click()
Dim oHouse As House, oHouses As Houses
Set oHouses = New Houses


oHouses.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=InvestCasablanca;Data Source=SUPERNOVA"

lngHouseId = oHouses.Add(34, 1, 2, 8, 258429, 200, 4, 2, 1, 1, 0, 30, 0, 20, 60, 1, 1, " ", "3rd turn left after the villas from Poli, La Bahia, opposite the ceramic shop, after pharmacy.", "", "", " ", "", " ", "", "", "", 1, 2, 2)

MsgBox lngHouseId














End Sub

Private Sub Command6_Click()
Dim oLocation As Location, lngResult As Long, oImages As Collection
Dim oLocations As New Locations
oLocations.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
oLocations.FetchNew 1, 1
'oLocations.Item(9).strConnection = oLocations.strConnection

'Set oImages = oLocations.Item(9).GetImages

MsgBox oLocations.Count


End Sub

Private Sub Command7_Click()
Dim oLocation As Location, lngResult As Long, oImages As Collection
Dim oLocations As New Locations
oLocations.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
oLocations.FetchResale 1, 1
'oLocations.Item(9).strConnection = oLocations.strConnection

'Set oImages = oLocations.Item(9).GetImages

MsgBox oLocations.Count
End Sub

Private Sub Command8_Click()
Dim oLocation As Location, lngResult As Long, oImages As Collection
Dim oLocations As New Locations
oLocations.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
oLocations.FetchRentals 1, 1
'oLocations.Item(9).strConnection = oLocations.strConnection

'Set oImages = oLocations.Item(9).GetImages

MsgBox oLocations.Count
End Sub

Private Sub Command9_Click()

Dim oHouse As House, oHouses As Houses, oConn As New ADODB.Connection
Set oHouses = New Houses


oHouses.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
oHouses.Fetch 1
oConn.ConnectionString = oHouses.strConnection

Exit Sub 'Dont ever do this again!



oConn.Open


For Each oHouse In oHouses

    oConn.Execute "INSERT INTO HousesCategories (lngHouseId, lngCategoryId) values (" & oHouse.lngHouseId & "," & oHouse.lngCategoryId & ")"

Next

For Each oHouse In oHouses

    oConn.Execute "INSERT INTO HousesLocations (lngHouseId, lngLocationId) values (" & oHouse.lngHouseId & "," & oHouse.lngLocationId & ")"

Next

oConn.Close

MsgBox "Done"


End Sub

Private Sub FetchLanguage_Click()

Dim oLanguage As Language, oLanguages
Set oLanguages = New Languages


oLanguages.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oLanguages.Fetch

For Each oLanguage In oLanguages
    MsgBox oLanguage.lngLanguageId & " - " & oLanguage.lngPhraseId
Next



End Sub




Private Sub UpdateFile_Click()

Dim oFile, oFiles, ret
Set oFiles = New Files
oFiles.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oFiles.Fetch
ret = oFiles.Update(Test.lngFileID, Test.strFileName)
If ret <> 0 Then MsgBox "Error"
End Sub

Private Sub AddFile_Click()
Dim oFile, oFiles, strFileName As String
Set oFiles = New Files
oFiles.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oFiles.Fetch

strFileName = Test.strFileName.Text
Set oFile = New File
oFile.strFileName = oFiles.Add(strFileName)
For Each oFile In oFiles
    MsgBox oFile.lngFileID
Next

End Sub

Private Sub DeleteFile_Click()
Dim oFile, oFiles, ret
Set oFiles = New Files
oFiles.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oFiles.Fetch
ret = oFiles.Delete(Test.lngFileID)
If ret <> 0 Then MsgBox "Error"
End Sub

Private Sub FetchFile_Click()
Dim oFile, oFiles
Set oFiles = New Files
oFiles.strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"

oFiles.Fetch

For Each oFile In oFiles
    MsgBox oFile.lngFileID & vbTab & oFile.strFileName
Next



End Sub

Private Sub Form_Load()
strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User" & _
    "ID=sa;Password=;Initial Catalog=INVESTCASABLANCA;Data Source=LOCAL"
End Sub

