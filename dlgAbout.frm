VERSION 5.00
Begin VB.Form dlgAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SizerBox ActiveX Control"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "dlgAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1785
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   3855
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Author: Daniel Morgato Martin"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label Label4 
         Caption         =   "      This ActiveX is absolutely FREE. You can use it as many as you want."
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2505
      TabIndex        =   2
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1665
      TabIndex        =   1
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SizerBox ActiveX Control"
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1057
      TabIndex        =   0
      Top             =   360
      Width           =   2790
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label3 = CStr(App.Major) + "." + CStr(App.Minor)
End Sub


