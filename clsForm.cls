#If Win64 Then
    Private Declare ptrsafe Function FindWindow Lib "User32" _
        Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
        
    Private Declare ptrsafe Function GetWindowLong Lib "User32" _
        Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long
        
    Private Declare ptrsafe Function SetWindowLong Lib "User32" _
        Alias "SetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
        
    Private Declare ptrsafe Function DrawMenuBar Lib "User32" ( _
        ByVal hwnd As Long) As Long

#Else
    Private Declare Function FindWindow Lib "User32" _
        Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
        
    Private Declare Function GetWindowLong Lib "User32" _
        Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long
        
    Private Declare Function SetWindowLong Lib "User32" _
        Alias "SetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
        
    Private Declare Function DrawMenuBar Lib "User32" ( _
        ByVal hwnd As Long) As Long
#End If

Private myForm As Object
Private mButtons() As clsForm
Public Main As clsForm
Private iMenuWidth As Integer
Public Wellcome As MSForms.Frame

Private WithEvents TitleBar As MSForms.Frame
Public WithEvents TitleText As MSForms.Label
Private WithEvents CloseButton As MSForms.Label
Private WithEvents MinButton As MSForms.Label
Public WithEvents SideNav As MSForms.Frame

Public WithEvents btnMenu As MSForms.Label
Public WithEvents frameMenu As MSForms.Frame
Public WithEvents Display As MSForms.Multipage


Dim iX, iY As Long
Dim SkalaX As Double
Dim SkalaY As Double
Dim isMin As Boolean

Private ControlSkala(100) As clsForm
Public CtrSkala As MSForms.Control
Public LeftSkala As Boolean
Public TopSkala As Boolean
Public RightSkala As Boolean
Public BottomSkala As Boolean

Private TotalSkala As Integer
Public Property Get MenuWidth() As Integer
    If iMenuWidth < 100 Then
        MenuWidth = 100
    Else
        MenuWidth = iMenuWidth
    End If
    
End Property
Public Property Let MenuWidth(value As Integer)
    iMenuWidth = value
End Property

Public Sub Sync(UserForm As Object, Optional Multipage As MSForms.Multipage = Nothing, Optional FullScreen As Boolean = True)
    Set myForm = UserForm
    Set Display = Multipage
    SkalaX = Display.Width
    SkalaY = Display.Height
    
    Application.DisplayFullScreen = True
    
    If FullScreen = True Then
        myForm.Move 0, 0, Application.Width - 15, Application.Height - 15
    End If
     
    RemoveTitleBar myForm
    Call AddTitleBar
    Call AddTitleText
    Call AddCloseButton
   If FullScreen Then Call AddMinButton
    Call SetSideNav
    Call SetDisplay
End Sub
Private Sub SetSideNav()
    Set SideNav = myForm.Controls.Add("forms.Frame.1")
    With SideNav
        .Move 0, 30, MenuWidth, myForm.Height - 30
        .SpecialEffect = fmSpecialEffectFlat
        
        If Display.BackColor = &H8000000F Then .BackColor = &HFFBD75
        .BackColor = Display.BackColor
                  
        .ForeColor = myForm.BorderColor
         
    End With
    ReDim mButtons(Display.Pages.Count)
    For i = 0 To Display.Pages.Count - 1
        AddMenu Display.Pages(i).Caption, Display.Pages(i)
    Next
End Sub
Private Sub AddMenu(MenuName As String, Page As Object)
    Dim fm As MSForms.Frame
    Set fm = SideNav.Controls.Add("forms.frame.1")
    With fm
        .Move 0, (Page.Index * 30) + 10, SideNav.Width, 30
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    Dim mn As MSForms.Label
    Set mn = fm.Controls.Add("forms.label.1")
    With mn
        .Tag = Page.Index
        .Caption = Chr(160) & MenuName
        .Picture = Page.Picture
        .Move 5, 5, SideNav.Width
        .ForeColor = myForm.ForeColor
        .Font.Name = Display.Font.Name
        Page.Picture = Nothing
        .PicturePosition = fmPicturePositionLeftCenter
        .Font.Size = 12
        
        .AutoSize = True
        .AutoSize = False
        .Width = .Width + 4
        .BackStyle = fmBackStyleTransparent
        
        
    End With
    
    Set mButtons(Page.Index) = New clsForm
    With mButtons(Page.Index)
        Set .Main = Me
        Set .btnMenu = mn
        Set .frameMenu = fm
    End With
End Sub

Private Sub frameMenu_Click()
    Call btnMenu_Click
End Sub
Private Sub btnMenu_Click()
    Main.Display.value = btnMenu.Tag
    Main.TitleText.Caption = Main.TitleText.Tag & " - " & btnMenu.Caption
    
    For Each ctr In Main.SideNav.Controls
        ctr.BackColor = Main.SideNav.BackColor
    Next
    frameMenu.BackColor = Main.SideNav.ForeColor
    Main.Wellcome.Visible = False
End Sub

Private Sub SetDisplay()
    With Display
        .Style = fmTabStyleNone
        .Move MenuWidth, 30, TitleBar.Width - MenuWidth, myForm.Height - TitleBar.Height
        SkalaX = .Width - SkalaX
        SkalaY = .Height - SkalaY
        Call SetSkala
    End With
    Set Wellcome = myForm.Controls.Add("forms.Frame.1")
    With Wellcome
        '.Style = fmTabStyleNone
        .Move MenuWidth, 30, TitleBar.Width - MenuWidth, myForm.Height - TitleBar.Height
        .PictureSizeMode = fmPictureSizeModeStretch
        .Picture = myForm.Picture
    End With
    
End Sub

Sub SetSkala()
    
    For i = 0 To TotalSkala - 1
        
        With ControlSkala(i)
            If .RightSkala And .LeftSkala Then
                .CtrSkala.Width = .CtrSkala.Width + SkalaX
            ElseIf .RightSkala And .LeftSkala = False Then
                .CtrSkala.Left = .CtrSkala.Left + SkalaX
            ElseIf .RightSkala = False And .LeftSkala = False Then
                .CtrSkala.Left = .CtrSkala.Left + (SkalaX / 2)
            End If
           
            If .BottomSkala And .TopSkala Then
                .CtrSkala.Height = .CtrSkala.Height + SkalaY + 17
            ElseIf .BottomSkala And .TopSkala = False Then
                .CtrSkala.Top = .CtrSkala.Top + SkalaY + 17
            ElseIf .BottomSkala = False And .TopSkala = False Then
                .CtrSkala.Top = .CtrSkala.Top + (SkalaY / 2)
            End If
        End With
    Next
    
End Sub
Public Sub LayoutX(ListControls As Variant)
    For Each c In ListControls
        Layout c, , , , False
    Next
End Sub
Public Sub LayoutY(ListControls As Variant)
    For Each c In ListControls
        Layout c, , , False, True
    Next
End Sub
Public Sub LayoutXY(ListControl As Variant)
    
    For Each c In ListControl
       Layout c
    Next
End Sub
Public Sub LayoutBottomLeft(ListControl As Variant)
    
    For Each c In ListControl
        Layout c, , False, False
    Next
End Sub
Public Sub LayoutBottomRight(ListControl As Variant)
    For Each c In ListControl
        Layout c, False, False
    Next
End Sub
Public Sub LayoutTopRight(ListControl As Variant)
    For Each c In ListControl
        Layout c, False, , True, False
    Next
End Sub
Public Sub Layout(Control As Variant, Optional Left As Boolean = True, Optional Top As Boolean = True, Optional Right As Boolean = True, Optional Bottom As Boolean = True)
       ' MsgBox Len(ControlSkala)
       TotalSkala = TotalSkala + 1
       
       Set ControlSkala(TotalSkala - 1) = New clsForm
       With ControlSkala(TotalSkala - 1)
           Set .CtrSkala = Control
                .LeftSkala = Left
                .TopSkala = Top
                .RightSkala = Right
                .BottomSkala = Bottom
       End With
        
End Sub

Sub RemoveTitleBar(frm As Object)
    Dim lStyle          As Long
    Dim hMenu           As Long
    Dim mhWndForm       As Long
     
    If Val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", frm.Caption) 'for Office 97 version
    Else
        mhWndForm = FindWindow("ThunderDFrame", frm.Caption) 'for office 2000 or above
    End If
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
End Sub

Private Sub AddTitleBar()
    
    Set TitleBar = myForm.Controls.Add("forms.Frame.1")
    With TitleBar
        .Move 0, 0, myForm.InsideWidth, 30
        .SpecialEffect = fmSpecialEffectFlat
        
        .MousePointer = fmMousePointerSizeAll
        If myForm.BackColor = &H8000000F Then myForm.BackColor = &HA86604
        .BackColor = myForm.BackColor
    End With
End Sub

Private Sub AddTitleText()
    Set TitleText = TitleBar.Controls.Add("forms.label.1")
    With TitleText
        Set .Font = myForm.Font
        .Move 5, 5, TitleBar.Width
        .Caption = myForm.Caption
        .Tag = myForm.Caption
        .Font.Size = 14
        .ForeColor = &HFFFFFF
        .MousePointer = fmMousePointerArrow
    End With
End Sub

Private Sub AddCloseButton()
    Set CloseButton = TitleBar.Controls.Add("forms.label.1")
    With CloseButton
       .Caption = "X"
       .Move TitleBar.Width - 30, 0, 20
       .ForeColor = &HFFFFFF
       .BackColor = &HC0&
       .TextAlign = fmTextAlignCenter
       .Font.Size = 12
       .Font.Name = "Arial"
       .Font.Bold = True
       .MousePointer = fmMousePointerArrow
    End With
End Sub
Private Sub AddMinButton()
    Set MinButton = TitleBar.Controls.Add("forms.label.1")
    With MinButton
       .Caption = "_"
       .Move TitleBar.Width - 53, 0, 20
       .ForeColor = &HFFFFFF
       .BackColor = TitleBar.BackColor
       
       .TextAlign = fmTextAlignCenter
       .Font.Size = 12
       .Font.Name = "Arial"
       .Font.Bold = True
       .MousePointer = fmMousePointerArrow
    End With
End Sub
Private Sub CloseButton_Click()
    Unload myForm
End Sub

Private Sub CloseButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CloseButton.BackColor = &HFF&
End Sub

Private Sub MinButton_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub MinButton_Click()
    isMin = True
    With myForm
        .Width = 200
        .Height = 30
        .Left = (Application.Width - .Width) / 2
    End With
    With TitleText
        .Width = 190
        .TextAlign = fmTextAlignCenter
    End With
End Sub

Private Sub MinButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MinButton.BackColor = &HE0E0E0
End Sub

Private Sub TitleBar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        TitleText.ControlTipText = "DoubleClick Untuk Reset Posisi"
        If isMin Then
            myForm.Move myForm.Left + (X - iX), myForm.Left + (Y - iY)
        Else
            myForm.Move myForm.Left + (X - iX), myForm.Top + (Y - iY)
        End If
    Else
        iX = X
        iY = Y
    End If
    CloseButton.BackColor = &HC0&
    MinButton.BackColor = TitleBar.BackColor
End Sub

Private Sub TitleText_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TitleText_Click()
    If isMin Then
        isMin = False
         myForm.Move 1, 1, Application.Width - 15, Application.Height - 15
         With TitleText
        .Width = myForm.Width
        .TextAlign = fmTextAlignLeft
    End With
    End If
End Sub

Private Sub TitleText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If isMin Then Exit Sub
    TitleText.ControlTipText = ""
    myForm.Move 1, 1
End Sub



