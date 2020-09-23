VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl cltMarquee 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtMFont 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Text            =   "14px Arial"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgColors3 
      Left            =   2520
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser webMarquee 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      ExtentX         =   4683
      ExtentY         =   2355
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtMarqueeColor 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDirection 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtBehavior 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "cltMarquee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum MarqueeColorConst
    Choose = 0
    Fontcolor_ = 1
    BGColor_ = 2
End Enum

Public Enum BehaviorConst
    Scroll = 0
    slide = 1
    Alternate = 2
End Enum

Public Enum DirectionConst
    Left_ = 0
    Right_ = 1
    Up_ = 3
End Enum

Public Enum LoopConst
    Repeat = -1
    Once = 1
End Enum

'Default Property Values:
Const m_def_Loop = -1
Const m_def_Colors = 0
Const m_def_BGColor = "FF0000"
Const m_def_FontColor = "000000"
Const m_def_ScrollAmount = 10
Const m_def_ScrollDelay = 10
Const m_def_Direction = 0
Const m_def_Behavior = 0
Const m_def_Text = "Default Text"
'Property Variables:
Dim m_Loop As LoopConst
Dim m_Colors As Variant
Dim m_BGColor As Variant
Dim m_FontColor As Variant
Dim m_ScrollAmount As Variant
Dim m_ScrollDelay As Variant
Dim m_Direction As DirectionConst
Dim m_Behavior As BehaviorConst
Dim m_Text As Variant
Dim itl As String
Dim bld As String

Public Sub WriteHTML2()
    
    'Color codes
    txtMarqueeColor.MaxLength = 6
    If Colors = BGColor_ Then
        dlgColors3.ShowColor
        'Makes common dialog box oclor to HTML color
        txtMarqueeColor.Text = Right(StrReverse(Hex(dlgColors3.Color)), Len(Hex(dlgColors3.Color)) - 1) & "000000"
        Colors = Choose
        BGColor = txtMarqueeColor.Text
    End If
    If Colors = Fontcolor_ Then
        dlgColors3.ShowColor
        txtMarqueeColor.Text = Right(StrReverse(Hex(dlgColors3.Color)), Len(Hex(dlgColors3.Color)) - 1) & "000000"
        Colors = Choose
        FontColor = txtMarqueeColor.Text
    End If
       
    'Behavior code
    If Behavior = Alternate Then txtBehavior = "alternate"
    If Behavior = Scroll Then txtBehavior = "scroll"
    If Behavior = slide Then txtBehavior = "slide"
       
    'Direction code
    If Direction = Left_ Then txtDirection.Text = "left"
    If Direction = Right_ Then txtDirection.Text = "right"
    If Direction = Up_ Then txtDirection.Text = "up"
       
    'Font code
    If UserControl.FontItalic = True Then
        itl = "italic"
    Else
        itl = ""
    End If
    If UserControl.FontBold = True Then
        bld = "bold"
    Else
    bld = ""
    End If
    txtMFont.Text = itl & " " & bld & " " & UserControl.FontSize & "px" & " " & UserControl.FontName
           
           
    'webButton must use navigate first before using .doucment.Open,.doucment.write,
    'and .doucment.close and the html file "blank.html" must also exist to work properly.
    'Although it sometimes wiil work without using navigate first. I advice you to use it.
    webMarquee.Navigate App.Path & "\blank.html"
    DoEvents
    
    webMarquee.Left = -200
    webMarquee.Top = -375
    webMarquee.Width = ScaleWidth + 600
    webMarquee.Height = ScaleHeight + 700
    DoEvents
           
    webMarquee.Document.Open
    webMarquee.Document.write "<html>"
    webMarquee.Document.write "<body bgcolor=#" & BGColor & ">"
    webMarquee.Document.write "<table align=center width=100% height=100%><tr><td valign=center>"
    webMarquee.Document.write "<DIV style='font:" & txtMFont.Text & ";color:#" & FontColor & "'>"
    webMarquee.Document.write "<MARQUEE LOOP=" & Loop_ & " DIRECTION=" & txtDirection.Text & " SCROLLDELAY=" & ScrollDelay & " SCROLLAMOUNT=" & ScrollAmount & " BEHAVIOR=" & txtBehavior.Text & ">"
    webMarquee.Document.write Text & "</MARQUEE></DIV></td></tr></table>"
    webMarquee.Document.write "</body>"
    webMarquee.Document.write "</html>"
    webMarquee.Document.Close
    DoEvents
    
End Sub

Public Sub UserControl_Initialize()
    Call WriteHTML2
End Sub

Public Sub UserControl_Resize()
    Call WriteHTML2
End Sub

Public Sub UserControl_Show()
    Call WriteHTML2
End Sub

Public Property Get Text() As Variant
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As Variant)
    m_Text = New_Text
    PropertyChanged "Text"
    Call WriteHTML2
End Property


Private Sub UserControl_InitProperties()
    m_Text = m_def_Text
    m_Behavior = m_def_Behavior
    m_Direction = m_def_Direction
    m_ScrollAmount = m_def_ScrollAmount
    m_ScrollDelay = m_def_ScrollDelay
    m_BGColor = m_def_BGColor
    m_FontColor = m_def_FontColor
    m_Colors = m_def_Colors
    Set UserControl.Font = Ambient.Font
    m_Loop = m_def_Loop
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_Behavior = PropBag.ReadProperty("Behavior", m_def_Behavior)
    m_Direction = PropBag.ReadProperty("Direction", m_def_Direction)
    m_ScrollAmount = PropBag.ReadProperty("ScrollAmount", m_def_ScrollAmount)
    m_ScrollDelay = PropBag.ReadProperty("ScrollDelay", m_def_ScrollDelay)
    m_BGColor = PropBag.ReadProperty("BGColor", m_def_BGColor)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_Colors = PropBag.ReadProperty("Colors", m_def_Colors)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Loop = PropBag.ReadProperty("Loop", m_def_Loop)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Behavior", m_Behavior, m_def_Behavior)
    Call PropBag.WriteProperty("Direction", m_Direction, m_def_Direction)
    Call PropBag.WriteProperty("ScrollAmount", m_ScrollAmount, m_def_ScrollAmount)
    Call PropBag.WriteProperty("ScrollDelay", m_ScrollDelay, m_def_ScrollDelay)
    Call PropBag.WriteProperty("BGColor", m_BGColor, m_def_BGColor)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("Colors", m_Colors, m_def_Colors)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Loop", m_Loop, m_def_Loop)
End Sub

Public Property Get Behavior() As BehaviorConst
    Behavior = m_Behavior
End Property

Public Property Let Behavior(ByVal New_Behavior As BehaviorConst)
    m_Behavior = New_Behavior
    PropertyChanged "Behavior"
    Call WriteHTML2
End Property

Public Property Get Direction() As DirectionConst
    Direction = m_Direction
End Property

Public Property Let Direction(ByVal New_Direction As DirectionConst)
    m_Direction = New_Direction
    PropertyChanged "Direction"
    Call WriteHTML2
End Property

Public Property Get ScrollAmount() As Variant
    ScrollAmount = m_ScrollAmount
End Property

Public Property Let ScrollAmount(ByVal New_ScrollAmount As Variant)
    m_ScrollAmount = New_ScrollAmount
    PropertyChanged "ScrollAmount"
    Call WriteHTML2
End Property

Public Property Get ScrollDelay() As Variant
    ScrollDelay = m_ScrollDelay
End Property

Public Property Let ScrollDelay(ByVal New_ScrollDelay As Variant)
    m_ScrollDelay = New_ScrollDelay
    PropertyChanged "ScrollDelay"
    Call WriteHTML2
End Property

Public Property Get BGColor() As Variant
    BGColor = m_BGColor
End Property

Public Property Let BGColor(ByVal New_BGColor As Variant)
    m_BGColor = New_BGColor
    PropertyChanged "BGColor"
    Call WriteHTML2
End Property

Public Property Get FontColor() As Variant
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As Variant)
    m_FontColor = New_FontColor
    PropertyChanged "FontColor"
    Call WriteHTML2
End Property

Public Property Get Colors() As MarqueeColorConst
    Colors = m_Colors
End Property

Public Property Let Colors(ByVal New_Colors As MarqueeColorConst)
    m_Colors = New_Colors
    PropertyChanged "Colors"
    Call WriteHTML2
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call WriteHTML2
End Property

Public Property Get Loop_() As LoopConst
    Loop_ = m_Loop
End Property

Public Property Let Loop_(ByVal New_Loop As LoopConst)
    m_Loop = New_Loop
    PropertyChanged "Loop"
    Call WriteHTML2
End Property

