VERSION 5.00
Object = "{B5A05027-B630-44D0-AEA7-B7A3CB76105C}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12675
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFadeExit 
      Enabled         =   0   'False
      Left            =   9720
      Top             =   2760
   End
   Begin VB.Frame Frame1 
      Caption         =   "root"
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   10335
      Begin VB.ListBox lstResult 
         Height          =   2595
         ItemData        =   "frmmain.frx":628A
         Left            =   5160
         List            =   "frmmain.frx":628C
         TabIndex        =   27
         Top             =   840
         Width           =   4695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtDir 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "C:\"
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2775
      End
      Begin VB.Timer tmrUpdate 
         Interval        =   20
         Left            =   6600
         Top             =   360
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   270
         Left            =   2400
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "&Stop"
         Height          =   270
         Left            =   3720
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "Only &Hidden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkFolders 
         Caption         =   "Only F&olders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox chkFiles 
         Caption         =   "Only &Files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "Only &Archive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "Only &Read Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Only &System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblPath 
         Caption         =   "C:\CRAP"
         Height          =   1095
         Left            =   4800
         TabIndex        =   26
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblFilesSearched 
         Caption         =   "Files Searched"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   9495
      End
      Begin VB.Label lblFilesFound 
         Caption         =   "Files Found"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   9495
      End
      Begin VB.Label lblCurpath 
         Caption         =   "Current path"
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   9495
      End
   End
   Begin VB.PictureBox mgui 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      Picture         =   "frmmain.frx":628E
      ScaleHeight     =   5535
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin cantSee_v1.GUI_Rollover min 
         Height          =   255
         Left            =   6200
         TabIndex        =   30
         Top             =   430
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   450
         Enabled         =   0   'False
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":8D3CC
         ImageHover      =   "frmmain.frx":8F65D
         ImageDown       =   "frmmain.frx":919B1
         ImageDisabled   =   "frmmain.frx":93D05
         ImageMask       =   "frmmain.frx":95F96
         ImageSelected   =   "frmmain.frx":98227
         ImageSelectedHover=   "frmmain.frx":9A4B8
      End
      Begin cantSee_v1.GUI_Rollover xit 
         Height          =   270
         Left            =   6600
         TabIndex        =   29
         Top             =   420
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   476
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":9C749
         ImageHover      =   "frmmain.frx":9ED87
         ImageDown       =   "frmmain.frx":A1403
         ImageDisabled   =   "frmmain.frx":A3A7F
         ImageMask       =   "frmmain.frx":A60BD
         ImageSelected   =   "frmmain.frx":A86FB
         ImageSelectedHover=   "frmmain.frx":AAD39
      End
      Begin vkUserContolsXP.vkCommand cmdunhideall 
         Height          =   375
         Left            =   6240
         TabIndex        =   28
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Unhide All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         Enabled         =   0   'False
         BorderColor     =   12632256
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand cmdUnhide 
         Height          =   375
         Left            =   5365
         TabIndex        =   9
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Unhide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         Enabled         =   0   'False
         BorderColor     =   12632256
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCheck ckRead 
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Read Only"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCheck ckArchive 
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   4560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Archive"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCheck ckSystem 
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "System"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCheck ckHidden 
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Hidden"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin vkUserContolsXP.vkCommand cmdscan 
         Height          =   375
         Left            =   4490
         TabIndex        =   4
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   14737632
         BackColorPushed1=   12632256
         BackColorPushed2=   11645361
         Caption         =   "Scan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   12632256
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15332854
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkListBox vkListBox2 
         Height          =   2535
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4471
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         Sorted          =   0
         ListType        =   1
      End
      Begin vkUserContolsXP.vkListBox vkListBox1 
         Height          =   2535
         Left            =   480
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   4471
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         Sorted          =   0
         ListType        =   3
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Select a Dirve and then press Scan to start searching..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblmnu1 
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   32
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblmnu 
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   31
         Top             =   885
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   5400
         Left            =   25
         Picture         =   "frmmain.frx":AD377
         Top             =   25
         Width           =   7560
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------
'====================================================================================================
'//////////////Design       :   Ratul Ahmed /////////////////////////////////////////////////////////
'//////////////Code         :   Ratul Ahmed /////////////////////////////////////////////////////////
'//////////////Copyright    :   © Ratul Ahmed ///////////////////////////////////////////////////////
'//////////////Thanx to     :   Mathieu Chartier(Searching Module)///////////////////////////////////
'=====================================================================I Love my Bangladesh===========
'----------------------------------------------------------------------------------------------------

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim IngSuccess As Long
Dim tempmem1 As String
Dim FormLoad As Boolean
Dim AlreadyQueried As Boolean
Dim transcount As Integer
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not AlreadyQueried Then 'make sure it only fades once
        AlreadyQueried = True
        Cancel = True          'cancels the unload
        ExitFadeOut            'calls the fadeout method
    End If
End Sub

Private Sub ExitFadeOut()
    transcount = 255
    tmrFadeExit.Interval = 5
    tmrFadeExit.Enabled = True 'enables the timer that does the fade out
End Sub

Private Sub tmrFadeExit_Timer()
    If transcount > 3 Then
        transcount = transcount - 4 'increases the percentage of transparency
        MakeTransparent frmmain.hWnd, transcount
    Else
        Unload frmmain
    End If
End Sub
Private Sub cmdscan_Click()
If cmdscan.Caption = "Stop" Then cmdStop_Click
If ckHidden.value = vbChecked Then chkHidden.value = 1 Else chkHidden.value = 0
If ckArchive.value = vbChecked Then chkArchive.value = 1 Else chkArchive.value = 0
If ckSystem.value = vbChecked Then chkSystem.value = 1 Else chkSystem.value = 0
If ckRead.value = vbChecked Then chkReadOnly.value = 1 Else chkReadOnly.value = 0
cmdscan.Caption = "Stop"
cmdSearch_Click
End Sub

Private Sub cmdUnhide_Click()
Dim Attr As VbFileAttribute
On Error Resume Next
chkReadOnly.value = 0
chkHidden.value = 0
chkSystem.value = 0
chkArchive.value = 0
If chkReadOnly Then Attr = Attr Or vbReadOnly
If chkHidden Then Attr = Attr Or vbHidden
If chkSystem Then Attr = Attr Or vbSystem
If chkArchive Then Attr = Attr Or vbArchive
Call SetAttr(lblPath, Attr)
vkListBox2.RemoveItem (vkListBox2.ListIndex)
If vkListBox2.ListCount > 0 Then lblPath = vkListBox2.List(vkListBox2.ListIndex + 1) Else Exit Sub
End Sub

Private Sub cmdunhideall_Click()
Dim Attr As VbFileAttribute
On Error Resume Next
Dim i As Integer
chkReadOnly.value = 0
chkHidden.value = 0
chkSystem.value = 0
chkArchive.value = 0
If chkReadOnly Then Attr = Attr Or vbReadOnly
If chkHidden Then Attr = Attr Or vbHidden
If chkSystem Then Attr = Attr Or vbSystem
If chkArchive Then Attr = Attr Or vbArchive
For i = 0 To lstResult.ListCount
lblPath = lstResult.List(i)
Call SetAttr(lblPath, Attr)
vkLabel1.Caption = "File passed : " & i & " of : " & FilesFound
vkListBox2.RemoveItem (i)
Next i
vkListBox2.Clear
lstResult.Clear
vkListBox2.Refresh
lstResult.Refresh

End Sub

Private Sub Form_Load()
Dim WindowRegion As Long
FormLoad = True
mgui.ScaleMode = vbPixels
mgui.AutoRedraw = True
mgui.AutoSize = True
mgui.BorderStyle = vbBSNone
Me.BorderStyle = vbBSNone
Me.Width = mgui.Width
Me.Height = mgui.Height
WindowRegion = MakeRegion(mgui)
SetWindowRgn Me.hWnd, WindowRegion, True
StayOnTop frmmain
vkListBox2.Clear
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
lblmnu(0).ForeColor = &H80000012
lblmnu1(1).ForeColor = &H80000012
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub mnuFileBrowse_Click()
    If Right$(lstResult.Text, 1) = "\" Then
        StartDoc lstResult.Text
    Else
        StartDoc UpOne(lstResult.Text)
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim X As Integer
vkListBox2.Clear
frmmain.MousePointer = vbHourglass
    Call FileSearch(lstResult, txtDir, txtFilter, , , CBool(chkFiles), CBool(chkFolders), _
    CBool(chkReadOnly), CBool(chkArchive), CBool(chkHidden), CBool(chkSystem))
   
For X = 0 To lstResult.ListCount - 1
  vkListBox2.AddItem lstResult.List(X)
Next X
vkListBox2.Refresh
cmdscan.Caption = "Scan"
frmmain.MousePointer = vbDefault
cmdUnhide.Enabled = True
cmdunhideall.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Abort = True
    cmdscan.Caption = "Scan"
End Sub

Private Sub cmdUpOne_Click()
    txtDir = UpOne(txtDir)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lblmnu_Click(Index As Integer)
frmhelp.Visible = True
End Sub

Private Sub lblmnu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnu(0).ForeColor = &H40C0&
End Sub

Private Sub lblmnu1_Click(Index As Integer)
frmabut.Visible = True
End Sub

Private Sub lblmnu1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmnu1(1).ForeColor = &H40C0&
End Sub

Private Sub lstResult_DblClick()
    StartDoc lstResult.Text
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        If Right$(lstResult.Text, 1) = "\" Then
            If DeleteFolder(lstResult.Text) Then
                lstResult.RemoveItem lstResult.ListIndex
            Else
                MsgBox "Error deleting folder"
            End If
        Else
            If DeleteFile(lstResult.Text) Then
                lstResult.RemoveItem lstResult.ListIndex
            Else
                MsgBox "Error deleting file"
            End If
        End If
    End Select
End Sub


Private Sub min_OnMouseClick()
frmmain.Windowstate = 1
End Sub

Private Sub tmrUpdate_Timer()
    lblFilesFound = "Files Found: " & FilesFound
    lblFilesSearched = "Total Files Searched: " & FileSearchCount
    lblCurpath = CurrentName
vkLabel1.Caption = "Files Found : " & FilesFound & "  ||  Files Searched : " & FileSearchCount

End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Attr As VbFileAttribute
    If chkReadOnly Then Attr = Attr Or vbReadOnly
    If chkHidden Then Attr = Attr Or vbHidden
    If chkSystem Then Attr = Attr Or vbSystem
    If chkArchive Then Attr = Attr Or vbArchive
    Call SetAttr(lblPath, Attr)
   
End Sub

Sub SetFileName(FileName As String)
    Dim Attr As Long
    lblPath = FileName
    Attr = GetAttr(FileName)
    chkReadOnly = -((Attr And vbReadOnly) <> 0)
    chkHidden = -((Attr And vbHidden) <> 0)
    chkSystem = -((Attr And vbSystem) <> 0)
    chkArchive = -((Attr And vbArchive) <> 0)
End Sub

Private Sub vkListBox1_ItemClick(Item As vkUserContolsXP.vkListItem)
txtDir = Item.tagString1
End Sub

Private Sub vkListBox1_ItemDblClick(Item As vkUserContolsXP.vkListItem)
On Error Resume Next
If Item.Text = vbNullString Then
        MsgBox "Please select an item first", vbExclamation, "can'tSee"
Exit Sub
End If
End Sub

Private Sub vkListBox2_ItemClick(Item As vkUserContolsXP.vkListItem)
On Error Resume Next
If Item.Text = vbNullString Then
        MsgBox "Please select an item first", vbExclamation, "can'tSee"
Exit Sub
End If
SetFileName (Item.Text)
If chkHidden.value = 1 Then ckHidden.value = vbChecked Else ckHidden.value = vbUnchecked

If chkArchive.value = 1 Then ckArchive.value = vbChecked Else ckArchive.value = vbUnchecked

If chkSystem.value = 1 Then ckSystem.value = vbChecked Else ckSystem.value = vbUnchecked

If chkReadOnly.value = 1 Then ckRead.value = vbChecked Else ckRead.value = vbUnchecked
lblPath = Item.Text
End Sub

Private Sub vkListBox2_ItemDblClick(Item As vkUserContolsXP.vkListItem)
On Error Resume Next
If Item.Text = vbNullString Then
        MsgBox "Please select an item first", vbExclamation, "can'tSee"
Exit Sub
End If
End Sub

Private Sub xit_OnMouseClick()
Unload Me
End Sub
