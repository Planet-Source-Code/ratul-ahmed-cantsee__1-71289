VERSION 5.00
Object = "{B5A05027-B630-44D0-AEA7-B7A3CB76105C}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmhelp 
   BorderStyle     =   0  'None
   Caption         =   "SuperUnhider Help"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   Icon            =   "frmhelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkScrollContainer vkScrollContainer1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5953
      BorderWidth     =   0
      Begin vkUserContolsXP.vkCommand vkCommand1 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "OK"
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
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         BorderStyle     =   1
         BorderColor     =   16576
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "HELP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "6. or press Unhide all to Unhide all listed files."
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
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "5. Then press Unhide button."
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
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "4. Select your desired file to make Unhide."
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
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "3. Press the Scan button and wait."
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
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "1. Select a Drive form the Drive Menu."
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
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "2. Select your filters from the bottom."
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
      End
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
StayOnTop frmhelp
End Sub

Private Sub vkCommand1_Click()
Unload frmhelp
End Sub
