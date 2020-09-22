VERSION 5.00
Begin VB.UserControl SizerBox 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2130
   ControlContainer=   -1  'True
   PropertyPages   =   "sizer.ctx":0000
   ScaleHeight     =   104
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   142
   ToolboxBitmap   =   "sizer.ctx":0035
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   7
      Left            =   1920
      MousePointer    =   9  'Size W E
      Picture         =   "sizer.ctx":0567
      Top             =   960
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   6
      Left            =   1920
      MousePointer    =   6  'Size NE SW
      Picture         =   "sizer.ctx":09E1
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   5
      Left            =   1920
      MousePointer    =   7  'Size N S
      Picture         =   "sizer.ctx":0E5B
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   4
      Left            =   1920
      MousePointer    =   8  'Size NW SE
      Picture         =   "sizer.ctx":12D5
      Top             =   600
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   3
      Left            =   1920
      MousePointer    =   9  'Size W E
      Picture         =   "sizer.ctx":174F
      Top             =   480
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   2
      Left            =   1920
      MousePointer    =   6  'Size NE SW
      Picture         =   "sizer.ctx":1BC9
      Top             =   360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   1
      Left            =   1920
      MousePointer    =   7  'Size N S
      Picture         =   "sizer.ctx":2043
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image WinNode 
      Height          =   105
      Index           =   0
      Left            =   1920
      MousePointer    =   8  'Size NW SE
      Picture         =   "sizer.ctx":24BD
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image ImgOn 
      Height          =   105
      Left            =   360
      Picture         =   "sizer.ctx":2937
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image ImgOff 
      Height          =   105
      Left            =   120
      Picture         =   "sizer.ctx":2DB1
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "SizerBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_CurrentX = 0
Const m_def_CurrentY = 0
Const m_def_MousePointer = 0
Const m_def_SizeEdit = 0
Const m_def_Loked = 0
Const m_def_Moveable = -1
'Const m_def_VirtualBorder = 0
'Property Variables:
Dim m_SizeEdit As Integer
Dim m_Loked As Boolean
Dim m_Moveable As Variant
'Dim m_VirtualBorder As Integer
'Event Declarations:
Event AfterSizeEdit(StartingNode As SizeNodeConst)
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
'Event AfterSizeEdit()
Event BeforeSizeEdit(StartingNode As SizeNodeConst)

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    dlgAbout.Show vbModal
    Unload dlgAbout
    Set dlgAbout = Nothing
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
Attribute ActiveControl.VB_Description = "Returns the control that has focus."
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As BackStyleConst
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleConst)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
Attribute Controls.VB_Description = "A collection whose elements represent each control on a form, including elements of control arrays. "
    Set Controls = UserControl.Controls
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As DrawModeConstants
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
    DrawMode = UserControl.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As DrawModeConstants)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As DrawStyleConstants
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As DrawStyleConstants)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "SizerPropertyPage"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillStyle
Public Property Get FillStyle() As FillStyleConstants
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = UserControl.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleConstants)
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = UserControl.Image
End Property

Private Sub UserControl_GotFocus()
Dim i As Integer
    If SizeEdit = sed_Automatic Then
        For i = 0 To 7
            WinNode(i).Visible = True
        Next i
        
        Call UserControl_Resize
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub UserControl_LostFocus()
Dim i As Integer
    If SizeEdit = sed_Automatic Then
        For i = 0 To 7
            WinNode(i).Visible = False
        Next i
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Moveable = True And Locked = False Then
        ReleaseCapture
        SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
        UserControl_Resize
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    PropertyChanged "MousePointer"
    UserControl.MousePointer = New_MousePointer
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_Resize
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(X As Single, Y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
    Point = UserControl.Point(X, Y)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
Dim i As Integer, j As Integer, oldScale As ScaleModeConstants
    If Moveable = False Then
        For i = -2 To 2 Step 1
            j = (8 + i) Mod 8
            WinNode(j) = ImgOff
        Next i
    Else
        For i = -2 To 2 Step 1
            j = (8 + i) Mod 8
            WinNode(j) = ImgOn
        Next i
    End If
    
    If Locked = True Then
        For i = 0 To 7
            WinNode(i) = ImgOff
        Next i
    End If

    
    oldScale = UserControl.ScaleMode
    UserControl.ScaleMode = 3
    WinNode(0).Top = 0
    WinNode(0).Left = 0
    WinNode(1).Top = 0
    WinNode(1).Left = (Extender.ScaleWidth - 7) / 2
    WinNode(2).Top = 0
    WinNode(2).Left = Extender.ScaleWidth - 7
    WinNode(3).Top = (Extender.ScaleHeight - 7) / 2
    WinNode(3).Left = Extender.ScaleWidth - 7
    WinNode(4).Top = Extender.ScaleHeight - 7
    WinNode(4).Left = Extender.ScaleWidth - 7
    WinNode(5).Top = Extender.ScaleHeight - 7
    WinNode(5).Left = (Extender.ScaleWidth - 7) / 2
    WinNode(6).Top = Extender.ScaleHeight - 7
    WinNode(6).Left = 0
    WinNode(7).Top = (Extender.ScaleHeight - 7) / 2
    WinNode(7).Left = 0
    UserControl.ScaleMode = oldScale
    
    RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SizeEdit() As SizeEditConst
Attribute SizeEdit.VB_Description = "Select sizing mode."
Attribute SizeEdit.VB_ProcData.VB_Invoke_Property = "SizerPropertyPage"
    SizeEdit = m_SizeEdit
End Property

Public Property Let SizeEdit(ByVal New_SizeEdit As SizeEditConst)
    m_SizeEdit = New_SizeEdit
    PropertyChanged "SizeEdit"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Loked() As Boolean
Attribute Loked.VB_Description = "Turn On/Off the possibility of being resized."
Attribute Loked.VB_ProcData.VB_Invoke_Property = "SizerPropertyPage"
    Loked = m_Loked
End Property

Public Property Let Loked(ByVal New_Loked As Boolean)
    m_Loked = New_Loked
    
    If m_Locked = True Then
        For i = 0 To 7
            WinNode(i) = ImgOff
        Next i
    Else
        For i = 0 To 7 Step 1
            WinNode(i) = ImgOn
        Next i
    End If

    PropertyChanged "Loked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,-1
Public Property Get Moveable() As Boolean
Attribute Moveable.VB_ProcData.VB_Invoke_Property = "SizerPropertyPage"
    Moveable = m_Moveable
End Property

Public Property Let Moveable(ByVal New_Moveable As Boolean)
Dim i As Integer, j As Integer
    m_Moveable = New_Moveable
    If Moveable = False Then
        For i = -2 To 2 Step 1
            j = (8 + i) Mod 8
            WinNode(j) = ImgOff
            WinNode(j).Tag = True
        Next i
    Else
        For i = -2 To 2 Step 1
            j = (8 + i) Mod 8
            WinNode(j) = ImgOn
            WinNode(j).Tag = False
        Next i
    End If
    PropertyChanged "Moveable"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    UserControl.MousePointer = m_def_MousePointer
    m_SizeEdit = m_def_SizeEdit
    m_Loked = m_def_Loked
    m_Moveable = m_def_Moveable
    UserControl.CurrentX = m_def_CurrentX
    UserControl.CurrentY = m_def_CurrentY
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Integer, j As Integer

    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 2310)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 3)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 3750)
    m_SizeEdit = PropBag.ReadProperty("SizeEdit", m_def_SizeEdit)
    m_Loked = PropBag.ReadProperty("Loked", m_def_Loked)
    m_Moveable = PropBag.ReadProperty("Moveable", m_def_Moveable)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.ClipControls = PropBag.ReadProperty("ClipControls", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", m_def_CurrentX)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", m_def_CurrentY)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    
    If Moveable = False Then
        For i = -2 To 2 Step 1
            j = (8 + i) Mod 8
            WinNode(j) = ImgOff
        Next i
    Else
        For i = -2 To 2 Step 1
            j = (8 + i) Mod 8
            WinNode(j) = ImgOn
        Next i
    End If
    
    If Locked = True Then
        For i = 0 To 7
            WinNode(i) = ImgOff
        Next i
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 2310)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 3)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 3750)
    Call PropBag.WriteProperty("SizeEdit", m_SizeEdit, m_def_SizeEdit)
    Call PropBag.WriteProperty("Loked", m_Loked, m_def_Loked)
    Call PropBag.WriteProperty("Moveable", m_Moveable, m_def_Moveable)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("ClipControls", UserControl.ClipControls, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, m_def_CurrentX)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, m_def_CurrentY)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ClipControls
Public Property Get ClipControls() As Boolean
Attribute ClipControls.VB_Description = "Determines whether graphics methods in Paint events repaint an entire object or newly exposed areas."
Attribute ClipControls.VB_ProcData.VB_Invoke_Property = "SizerPropertyPage"
    ClipControls = UserControl.ClipControls
End Property

Public Property Let ClipControls(ByVal New_ClipControls As Boolean)
    UserControl.ClipControls() = New_ClipControls
    PropertyChanged "ClipControls"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderConst
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = "SizerPropertyPage"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderConst)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,2,0
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_MemberFlags = "400"
    CurrentX = UserControl.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    If Ambient.UserMode = False Then Err.Raise 387
    UserControl.CurrentX = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,2,0
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_MemberFlags = "400"
    CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    If Ambient.UserMode = False Then Err.Raise 387
    UserControl.CurrentY = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As AppearanceConst
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConst)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Private Sub WinNode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Locked = False Then
        ReleaseCapture
        Select Case Index
            Case 0
                If Moveable = True Then
                    RaiseEvent BeforeSizeEdit(CVar(Index))
                    SendMessage hWnd, WM_NCLBUTTONDOWN, HTTOPLEFT, 0&
                    RaiseEvent AfterSizeEdit(CVar(Index))
                End If
            Case 1
                If Moveable = True Then
                    RaiseEvent BeforeSizeEdit(CVar(Index))
                    SendMessage hWnd, WM_NCLBUTTONDOWN, HTTOP, 0&
                    RaiseEvent AfterSizeEdit(CVar(Index))
                End If
            Case 2
                If Moveable = True Then
                    RaiseEvent BeforeSizeEdit(CVar(Index))
                    SendMessage hWnd, WM_NCLBUTTONDOWN, HTTOPRIGHT, 0&
                    RaiseEvent AfterSizeEdit(CVar(Index))
                End If
            Case 3
                RaiseEvent BeforeSizeEdit(CVar(Index))
                SendMessage hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
                RaiseEvent AfterSizeEdit(CVar(Index))
            Case 4
                RaiseEvent BeforeSizeEdit(CVar(Index))
                SendMessage hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
                RaiseEvent AfterSizeEdit(CVar(Index))
            Case 5
                RaiseEvent BeforeSizeEdit(CVar(Index))
                SendMessage hWnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0&
                RaiseEvent AfterSizeEdit(CVar(Index))
            Case 6
                If Moveable = True Then
                    RaiseEvent BeforeSizeEdit(CVar(Index))
                    SendMessage hWnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, 0&
                    RaiseEvent AfterSizeEdit(CVar(Index))
                End If
            Case 7
                If Moveable = True Then
                    RaiseEvent BeforeSizeEdit(CVar(Index))
                    SendMessage hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
                    RaiseEvent AfterSizeEdit(CVar(Index))
                End If
        End Select
        UserControl_Resize
    End If
End Sub


