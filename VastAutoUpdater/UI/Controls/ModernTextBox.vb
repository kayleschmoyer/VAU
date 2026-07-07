Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices

''' <summary>
''' A modern input: borderless TextBox inside a rounded outline, with a
''' magenta underline that animates out from the center on focus, native
''' placeholder text, and an optional show/hide password toggle.
''' </summary>
Public Class ModernTextBox
    Inherits Panel

    Private WithEvents InnerBox As New TextBox()
    Private ReadOnly _animTimer As New Timer() With {.Interval = 15}
    Private _focusProgress As Single = 0.0F
    Private _focused As Boolean = False
    Private _isPassword As Boolean = False
    Private _revealed As Boolean = False
    Private _cueText As String = String.Empty
    Private _overEye As Boolean = False

    Private Const EYE_ZONE_WIDTH As Integer = 30
    Private Const EM_SETCUEBANNER As Integer = &H1501

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
    End Function

    ''' <summary>Text of the inner textbox.</summary>
    Public Overrides Property Text As String
        Get
            Return InnerBox.Text
        End Get
        Set(value As String)
            InnerBox.Text = value
        End Set
    End Property

    ''' <summary>Placeholder shown while the field is empty.</summary>
    Public Property CueText As String
        Get
            Return _cueText
        End Get
        Set(value As String)
            _cueText = If(value, String.Empty)
            ApplyCueBanner()
        End Set
    End Property

    ''' <summary>Masks input and shows the reveal (eye) toggle.</summary>
    Public Property IsPassword As Boolean
        Get
            Return _isPassword
        End Get
        Set(value As Boolean)
            _isPassword = value
            InnerBox.UseSystemPasswordChar = value AndAlso Not _revealed
            LayoutInnerBox()
            Invalidate()
        End Set
    End Property

    ''' <summary>The hosted TextBox, for wiring events like KeyDown.</summary>
    Public ReadOnly Property TextBox As TextBox
        Get
            Return InnerBox
        End Get
    End Property

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        BackColor = Color.White
        Size = New Size(280, 36)
        Cursor = Cursors.IBeam

        InnerBox.BorderStyle = BorderStyle.None
        InnerBox.Font = New Font("Segoe UI", 10.5F)
        InnerBox.ForeColor = UiTheme.Charcoal
        InnerBox.BackColor = Color.White
        Controls.Add(InnerBox)
        LayoutInnerBox()

        AddHandler _animTimer.Tick, AddressOf AnimTick
    End Sub

    Private Sub ApplyCueBanner()
        If InnerBox.IsHandleCreated Then
            SendMessage(InnerBox.Handle, EM_SETCUEBANNER, New IntPtr(1), _cueText)
        End If
    End Sub

    Private Sub InnerBox_HandleCreated(sender As Object, e As EventArgs) Handles InnerBox.HandleCreated
        ApplyCueBanner()
    End Sub

    Private Sub LayoutInnerBox()
        Dim rightPad As Integer = If(_isPassword, EYE_ZONE_WIDTH + 4, 12)
        InnerBox.Location = New Point(12, (Height - InnerBox.Height) \ 2)
        InnerBox.Width = Math.Max(10, Width - 12 - rightPad)
    End Sub

    Protected Overrides Sub OnResize(eventargs As EventArgs)
        MyBase.OnResize(eventargs)
        LayoutInnerBox()
    End Sub

    Private Sub AnimTick(sender As Object, e As EventArgs)
        Dim target As Single = If(_focused, 1.0F, 0.0F)
        Dim delta As Single = target - _focusProgress
        If Math.Abs(delta) < 0.05F Then
            _focusProgress = target
            _animTimer.Stop()
        Else
            _focusProgress += delta * 0.25F
        End If
        Invalidate()
    End Sub

    Private Sub InnerBox_GotFocus(sender As Object, e As EventArgs) Handles InnerBox.GotFocus
        _focused = True
        _animTimer.Start()
    End Sub

    Private Sub InnerBox_LostFocus(sender As Object, e As EventArgs) Handles InnerBox.LostFocus
        _focused = False
        _animTimer.Start()
    End Sub

    Private Sub InnerBox_TextChanged(sender As Object, e As EventArgs) Handles InnerBox.TextChanged
        OnTextChanged(EventArgs.Empty)
    End Sub

    Private Function EyeZone() As Rectangle
        Return New Rectangle(Width - EYE_ZONE_WIDTH - 2, 0, EYE_ZONE_WIDTH, Height)
    End Function

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Dim over As Boolean = _isPassword AndAlso EyeZone().Contains(e.Location)
        If over <> _overEye Then
            _overEye = over
            Cursor = If(over, Cursors.Hand, Cursors.IBeam)
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        If _isPassword AndAlso EyeZone().Contains(e.Location) Then
            _revealed = Not _revealed
            InnerBox.UseSystemPasswordChar = Not _revealed
            Invalidate()
        Else
            InnerBox.Focus()
        End If
    End Sub

    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e)
        InnerBox.Focus()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' Corners blend into whatever surface hosts the field (usually a card)
        Dim hostColor As Color = If(Parent IsNot Nothing, Parent.BackColor, Color.White)
        e.Graphics.Clear(hostColor)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        Dim rect As New Rectangle(0, 0, Width - 1, Height - 1)
        Using path As GraphicsPath = UiTheme.RoundedPath(rect, 6)
            Using fill As New SolidBrush(Color.White)
                e.Graphics.FillPath(fill, path)
            End Using
            Using pen As New Pen(UiTheme.InputBorder, 1.0F)
                e.Graphics.DrawPath(pen, path)
            End Using
        End Using

        ' Focus underline grows out from the center
        If _focusProgress > 0.01F Then
            Dim underlineWidth As Integer = CInt((Width - 12) * _focusProgress)
            Dim x As Integer = (Width - underlineWidth) \ 2
            Using pen As New Pen(UiTheme.Magenta, 2.0F) With {.StartCap = LineCap.Round, .EndCap = LineCap.Round}
                e.Graphics.DrawLine(pen, x, Height - 3, x + underlineWidth, Height - 3)
            End Using
        End If

        If _isPassword Then
            DrawEye(e.Graphics)
        End If
    End Sub

    ''' <summary>
    ''' Draw the show/hide password eye with GDI+ (no icon-font dependency).
    ''' A slash across the eye means the password is currently revealed.
    ''' </summary>
    Private Sub DrawEye(g As Graphics)
        Dim zone As Rectangle = EyeZone()
        Dim cx As Integer = zone.X + zone.Width \ 2
        Dim cy As Integer = zone.Y + zone.Height \ 2
        Dim eyeColor As Color = If(_overEye, UiTheme.Charcoal, UiTheme.TextMuted)

        Using pen As New Pen(eyeColor, 1.6F)
            ' Almond outline: two arcs
            Dim w As Integer = 16
            Dim h As Integer = 12
            Dim eyeRect As New Rectangle(cx - w \ 2, cy - h \ 2, w, h)
            g.DrawArc(pen, eyeRect, 200, 140)
            g.DrawArc(pen, eyeRect, 20, 140)
            ' Pupil
            Using brush As New SolidBrush(eyeColor)
                g.FillEllipse(brush, cx - 2.5F, cy - 2.5F, 5.0F, 5.0F)
            End Using
            ' Slash when the password is visible
            If _revealed Then
                g.DrawLine(pen, cx - 9, cy + 8, cx + 9, cy - 8)
            End If
        End Using
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            _animTimer.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

End Class
