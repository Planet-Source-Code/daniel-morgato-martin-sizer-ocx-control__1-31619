VERSION 5.00
Object = "{9BF28D13-D1B8-49CF-8B24-07B5A38E95E7}#4.0#0"; "sizer.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   3000
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin SizerControl.SizerBox SizerBox1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _extentx        =   4048
      _extenty        =   2990
      font            =   "frmtest1.frx":0000
      picture         =   "frmtest1.frx":002C
      scaleheight     =   97
      scalemode       =   0
      scalewidth      =   169
      loked           =   -1  'True
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
