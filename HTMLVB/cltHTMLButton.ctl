VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl cltHTMLButton 
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   EditAtDesignTime=   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   240
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtColors 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtBFont 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "14px Arial"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser webButton 
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   240
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "cltHTMLButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum StyleButtonColorConst
    Choose = 0
    BGColor_ = 1
    Button_Color_ = 2
    Button_Hover_Color_ = 3
    Border_Color = 4
    Font_Color = 5
End Enum

Public Enum AlphaStyleConst
    NoAlpha = 0
    Alpha1 = 1
    Alpha2 = 2
    Alpha3 = 3
End Enum

Public Enum BorderConst
    None = 0
    Solid = 1
    Double_ = 2
    Dotted = 3
    Dashed = 4
    Inset = 5
    Outset = 6
End Enum

'Default Property Values:
Const m_def_Opacity = 100
Const m_def_FinishOpacity = 30
Const m_def_FontColor = "FFFFFF"
Const m_def_AlphaStyle = 3
Const m_def_BorderWidth = 2
Const m_def_BorderColor = "008FF0"
Const m_def_BorderStyle = 6
Const m_def_Caption = "HTML BUTTON"
Const m_def_Enabled = 0
Const m_def_BGColor = "FFFFFF"
Const m_def_Colors = 0
Const m_def_Button_Hover_Color = "FFF000"
Const m_def_Button_Color = "FF0000"
'Property Variables:
Dim m_Opacity As Variant
Dim m_FinishOpacity As Variant
Dim m_FontColor As Variant
Dim m_AlphaStyle As Variant
Dim m_BorderWidth As Variant
Dim m_BorderColor As Variant
Dim m_Caption As String
Dim m_Enabled As Boolean
Dim m_BGColor As Variant
Dim m_Button_Hover_Color As Variant
Dim m_Button_Color As Variant
Dim m_Colors As StyleButtonColorConst
Dim m_BorderStyle As BorderConst
Dim m_BorderType As String
Dim itl As String
Dim bld As String
'Event Declarations:
Event Click()

Public Sub WriteHTML()

On Error Resume Next
    
    txtColors.MaxLength = 6
    'Color codes
    If Colors = BGColor_ Then
        dlgColors.ShowColor
        'makes common dialog color to HTML color
        txtColors.Text = Right(StrReverse(Hex(dlgColors.Color)), Len(Hex(dlgColors.Color)) - 1) & "000000"
        Colors = Choose
        BGColor = txtColors.Text
    End If
    If Colors = Button_Color_ Then
        dlgColors.ShowColor
        txtColors.Text = Right(StrReverse(Hex(dlgColors.Color)), Len(Hex(dlgColors.Color)) - 1) & "000000"
        Colors = Choose
        Button_Color = txtColors.Text
    End If
    If Colors = Button_Hover_Color_ Then
        dlgColors.ShowColor
        txtColors.Text = Right(StrReverse(Hex(dlgColors.Color)), Len(Hex(dlgColors.Color)) - 1) & "000000"
        Colors = Choose
        Button_Hover_Color = txtColors.Text
    End If
    If Colors = Border_Color Then
        dlgColors.ShowColor
        txtColors.Text = Right(StrReverse(Hex(dlgColors.Color)), Len(Hex(dlgColors.Color)) - 1) & "000000"
        Colors = Choose
         BorderColor = txtColors.Text
    End If
    If Colors = Font_Color Then
        dlgColors.ShowColor
        txtColors.Text = Right(StrReverse(Hex(dlgColors.Color)), Len(Hex(dlgColors.Color)) - 1) & "000000"
        Colors = Choose
        FontColor = txtColors.Text
    End If
    'Border codes
    If BorderStyle = None Then
        m_BorderType = "none"
    End If
    If BorderStyle = Inset Then
        m_BorderType = "inset"
    End If
    If BorderStyle = Outset Then
        m_BorderType = "outset"
    End If
    If BorderStyle = Dotted Then
        m_BorderType = "dotted"
    End If
    If BorderStyle = Dashed Then
        m_BorderType = "dashed"
    End If
    If BorderStyle = Solid Then
        m_BorderType = "solid"
    End If
    If BorderStyle = Double_ Then
        m_BorderType = "double"
    End If
    If BorderStyle = Dotted Then
        m_BorderType = "dotted"
    End If
             
    'Font codes
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
    txtBFont.Text = itl & " " & bld & " " & UserControl.FontSize & "px" & " " & UserControl.FontName
      
    'webButton must use navigate first before using .doucment.Open,.doucment.write,
    'and .doucment.close and the html file "blank.html" must also exist to work properly.
    'Although it sometimes wiil work without using navigate first. I advice you to use it.
    webButton.Navigate App.Path & "\blank.html"
    DoEvents
    
    webButton.Left = -22
    webButton.Top = -22
    webButton.Width = ScaleWidth + 50
    webButton.Height = ScaleHeight + 50
    DoEvents
      
    Call HTMLCode
    
End Sub

'Code of the button in HTML
Private Sub HTMLCode()
    If AlphaStyle = NoAlpha Then
        webButton.Document.Open
        webButton.Document.write "<html>"
        'The javascript makes the Click() event work.
        'See - "webButton_StatusTextChange(ByVal Text As String)" sub
        webButton.Document.write "<script Language=JavaScript>"
        webButton.Document.write "function clickdown(){"
        webButton.Document.write "window.status='Click'}"
        webButton.Document.write "function clickup(){"
        webButton.Document.write "window.status='Clicks'}"
        webButton.Document.write "</script>"
        webButton.Document.write "<body bgcolor=#" & BGColor & ">"
        webButton.Document.write "<button id=button1 style='position:absolute;left:20;top:20;background:#" & Button_Color & ";width:" & webButton.Width - 50 & ";height:" & webButton.Height - 50 & ";cursor:hand"
        webButton.Document.write ";font:" & txtBFont.Text & ";color:#" & FontColor & ";border:" & BorderWidth & " " & m_BorderType & " #" & BorderColor & "'"
        webButton.Document.write " onmouseover=button1.style.background='#" & Button_Hover_Color & "';clickup() onmouseout=button1.style.background='#" & Button_Color & "' onclick=clickdown()>"
        webButton.Document.write Caption & "</button>"
        webButton.Document.write "</body>"
        webButton.Document.write "</html>"
        webButton.Document.Close
        DoEvents
    Else
        webButton.Document.Open
        webButton.Document.write "<html>"
        'The javascript makes the Click() event work.
        webButton.Document.write "<script Language=JavaScript>"
        webButton.Document.write "function clickdown(){"
        webButton.Document.write "window.status='Click'}"
        webButton.Document.write "function clickup(){"
        webButton.Document.write "window.status='Clicks';}"
        webButton.Document.write "</script>"
        webButton.Document.write "<body bgcolor=#" & BGColor & ">"
        webButton.Document.write "<button id=button1 style='position:absolute;left:20;top:20;background:#" & Button_Color & ";width:" & webButton.Width - 50 & ";height:" & webButton.Height - 50 & ";cursor:hand;"
        webButton.Document.write "filter:Alpha(opacity=" & Opacity & ",finishopacity=" & FinishOpacity & ",style=" & AlphaStyle & ");font:" & txtBFont.Text & ";color:#" & FontColor & ";border:" & BorderWidth & " " & m_BorderType & " #" & BorderColor & "'"
        webButton.Document.write " onmouseover=button1.style.background='#" & Button_Hover_Color & "';clickup() onmouseout=button1.style.background='#" & Button_Color & "' onclick=clickdown() onmouseup=clickup()>"
        webButton.Document.write Caption & "</button>"
        webButton.Document.write "</body>"
        webButton.Document.write "</html>"
        webButton.Document.Close
        DoEvents
    End If
End Sub

Public Property Get BGColor() As Variant
    BGColor = m_BGColor
End Property

Public Property Let BGColor(ByVal New_BGColor As Variant)
    m_BGColor = New_BGColor
    PropertyChanged "BGCOLOR"
    Call WriteHTML
End Property

Public Property Get Colors() As StyleButtonColorConst
    Colors = m_Colors
End Property

Public Property Let Colors(ByVal New_BGColor As StyleButtonColorConst)
    m_Colors = New_BGColor
    PropertyChanged "BGCOLOR"
    Call WriteHTML
End Property

Public Property Get Button_Hover_Color() As Variant
    Button_Hover_Color = m_Button_Hover_Color
End Property

Public Property Let Button_Hover_Color(ByVal New_BUTTON_HOVER_COLOR As Variant)
    m_Button_Hover_Color = New_BUTTON_HOVER_COLOR
    PropertyChanged "BUTTON_HOVER_COLOR"
    Call WriteHTML
End Property

Public Property Get Button_Color() As Variant
    Button_Color = m_Button_Color
End Property

Public Property Let Button_Color(ByVal New_BUTTON_COLOR As Variant)
    m_Button_Color = New_BUTTON_COLOR
    PropertyChanged "BUTTON_COLOR"
    Call WriteHTML
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
       Call WriteHTML
End Sub

Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_BGColor = m_def_BGColor
    m_Button_Hover_Color = m_def_Button_Hover_Color
    m_Button_Color = m_def_Button_Color
    m_Caption = m_def_Caption
    m_BorderStyle = m_def_BorderStyle
    m_BorderColor = m_def_BorderColor
    m_BorderColor = m_def_BorderColor
    m_BorderWidth = m_def_BorderWidth
    m_AlphaStyle = m_def_AlphaStyle
    m_FontColor = m_def_FontColor
    m_Opacity = m_def_Opacity
    m_FinishOpacity = m_def_FinishOpacity
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BGColor = PropBag.ReadProperty("BGCOLOR", m_def_BGColor)
    m_Button_Hover_Color = PropBag.ReadProperty("BUTTON_HOVER_COLOR", m_def_Button_Hover_Color)
    m_Button_Color = PropBag.ReadProperty("BUTTON_COLOR", m_def_Button_Color)
    m_Caption = PropBag.ReadProperty("CAPTION", m_def_Caption)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderWidth = PropBag.ReadProperty("BorderWidth", m_def_BorderWidth)
    m_AlphaStyle = PropBag.ReadProperty("AlphaStyle", m_def_AlphaStyle)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_Opacity = PropBag.ReadProperty("Opacity", m_def_Opacity)
    m_FinishOpacity = PropBag.ReadProperty("FinishOpacity", m_def_FinishOpacity)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_Resize()
    Call WriteHTML
End Sub

Private Sub UserControl_Show()
    Call WriteHTML
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BGCOLOR", m_BGColor, m_def_BGColor)
    Call PropBag.WriteProperty("BUTTON_HOVER_COLOR", m_Button_Hover_Color, m_def_Button_Hover_Color)
    Call PropBag.WriteProperty("BUTTON_COLOR", m_Button_Color, m_def_Button_Color)
    Call PropBag.WriteProperty("CAPTION", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderWidth", m_BorderWidth, m_def_BorderWidth)
    Call PropBag.WriteProperty("AlphaStyle", m_AlphaStyle, m_def_AlphaStyle)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Opacity", m_Opacity, m_def_Opacity)
    Call PropBag.WriteProperty("FinishOpacity", m_FinishOpacity, m_def_FinishOpacity)
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_CAPTION As String)
    m_Caption = New_CAPTION
    PropertyChanged "CAPTION"
    Call WriteHTML
End Property

Public Property Get BorderStyle() As BorderConst
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderConst)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call WriteHTML
End Property

Public Property Get BorderColor() As Variant
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Variant)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    Call WriteHTML
End Property

Public Property Get BorderWidth() As Variant
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Variant)
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    Call WriteHTML
End Property

Public Property Get AlphaStyle() As AlphaStyleConst
    AlphaStyle = m_AlphaStyle
End Property

Public Property Let AlphaStyle(ByVal New_AlphaStyle As AlphaStyleConst)
    m_AlphaStyle = New_AlphaStyle
    PropertyChanged "AlphaStyle"
    Call WriteHTML
End Property

'This part makes the Click() event to work
Private Sub webButton_StatusTextChange(ByVal Text As String)
    If Text = "Click" Then
        RaiseEvent Click
        Call HTMLCode
  End If
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get FontColor() As Variant
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As Variant)
    m_FontColor = New_FontColor
    PropertyChanged "FontColor"
    Call WriteHTML
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Attribute Font.VB_Description = "Returns a Font object."
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call WriteHTML
End Property

Public Property Get Opacity() As Variant
    Opacity = m_Opacity
End Property

Public Property Let Opacity(ByVal New_Opacity As Variant)
    m_Opacity = New_Opacity
    PropertyChanged "Opacity"
    Call WriteHTML
End Property

Public Property Get FinishOpacity() As Variant
    FinishOpacity = m_FinishOpacity
End Property

Public Property Let FinishOpacity(ByVal New_FinishOpacity As Variant)
    m_FinishOpacity = New_FinishOpacity
    PropertyChanged "FinishOpacity"
    Call WriteHTML
End Property


