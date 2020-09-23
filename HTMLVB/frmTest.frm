VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin HTMLButton.cltMarquee cltMarquee2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Text            =   "Bouncing"
      Behavior        =   2
      ScrollAmount    =   "7"
      ScrollDelay     =   "7"
      BGColor         =   "008FF0"
      FontColor       =   "FFFFF0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HTMLButton.cltMarquee cltMarquee1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1085
      Text            =   "HTML marquee used as VB text animation...."
      ScrollAmount    =   "3"
      BGColor         =   "0080C0"
      FontColor       =   "C0C0C0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HTMLButton.cltHTMLButton cltHTMLButton4 
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BGCOLOR         =   "000000"
      BUTTON_HOVER_COLOR=   "008000"
      BUTTON_COLOR    =   "0FF000"
      CAPTION         =   "EXIT"
      FontColor       =   "000000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HTMLButton.cltHTMLButton cltHTMLButton2 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BGCOLOR         =   "400000"
      BUTTON_HOVER_COLOR=   "804000"
      BUTTON_COLOR    =   "F08040"
      CAPTION         =   "Help"
      BorderStyle     =   4
      BorderColor     =   "FFFFF0"
      BorderWidth     =   "7"
      FontColor       =   "FFFFF0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HTMLButton.cltMarquee cltMarquee3 
      Height          =   3015
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5318
      Text            =   "For more information, click Help"
      Direction       =   3
      ScrollAmount    =   "2"
      ScrollDelay     =   "3"
      BGColor         =   "000080"
      FontColor       =   "0FFFF0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cltHTMLButton2_Click()
Call Shell("explorer.exe " & App.Path & "\info.html", 0)
End Sub


Private Sub cltHTMLButton4_Click()
End
End Sub
