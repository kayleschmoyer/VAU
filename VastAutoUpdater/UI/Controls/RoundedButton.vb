Imports System.Drawing.Drawing2D

''' <summary>
''' Flat rounded button with a smoothly animated hover state (the base color
''' eases toward the dark accent instead of snapping). Implements
''' IButtonControl so it works as a form's AcceptButton / CancelButton.
''' </summary>
Public Class RoundedButton
    Inherits Control
    Implements IButtonControl

    Private ReadOnly _animTimer As New Timer() With {.Interval = 15}
    Private _hoverProgress As Single = 0.0F
    Private _hovered As Boolean = False
    Private _pressed As Boolean = False
    Private _cornerRadius As Integer = 10
    Private _accent As Color = UiTheme.Magenta
    Private _accentDark As Color = UiTheme.MagentaDark
    Private _dialogResult As DialogResult = DialogResult.None

    Public Property CornerRadius As Integer
        Get
            Return _cornerRadius
        End Get
        Set(value As Integer)
            _cornerRadius = value
            Invalidate()
        End Set
    End Property

    ''' <summary>Base fill color (resting state).</summary>
    Public Property AccentColor As Color
        Get
            Return _accent
        End Get
        Set(value As Color)
            _accent = value
            Invalidate()
        End Set
    End Property

    ''' <summary>Fill color the button eases toward on hover.</summary>
    Public Property AccentDarkColor As Color
        Get
            Return _accentDark
        End Get
        Set(value As Color)
            _accentDark = value
            Invalidate()
        End Set
    End Property

    Public Property DialogResult As DialogResult Implements IButtonControl.DialogResult
        Get
            Return _dialogResult
        End Get
        Set(value As DialogResult)
            _dialogResult = value
        End Set
    End Property

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or
                 ControlStyles.Selectable Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        ForeColor = Color.White
        Font = UiTheme.Semibold(10.5F)
        Cursor = Cursors.Hand
        Size = New Size(160, 44)
        AddHandler _animTimer.Tick, AddressOf AnimTick
    End Sub

    Public Sub NotifyDefault(value As Boolean) Implements IButtonControl.NotifyDefault
        Invalidate()
    End Sub

    Public Sub PerformClick() Implements IButtonControl.PerformClick
        If Enabled AndAlso Visible Then
            OnClick(EventArgs.Empty)
        End If
    End Sub

    Private Sub AnimTick(sender As Object, e As EventArgs)
        Dim target As Single = If(_hovered, 1.0F, 0.0F)
        Dim delta As Single = target - _hoverProgress
        If Math.Abs(delta) < 0.05F Then
            _hoverProgress = target
            _animTimer.Stop()
        Else
            _hoverProgress += delta * 0.25F
        End If
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        _hovered = True
        _animTimer.Start()
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        _hovered = False
        _pressed = False
        _animTimer.Start()
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If e.Button = MouseButtons.Left Then
            _pressed = True
            Focus()
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _pressed = False
        Invalidate()
    End Sub

    Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
        MyBase.OnKeyUp(e)
        If e.KeyCode = Keys.Space OrElse e.KeyCode = Keys.Enter Then
            PerformClick()
        End If
    End Sub

    Protected Overrides Sub OnClick(e As EventArgs)
        MyBase.OnClick(e)
        If _dialogResult <> DialogResult.None Then
            Dim owner As Form = FindForm()
            If owner IsNot Nothing AndAlso owner.Modal Then
                owner.DialogResult = _dialogResult
            End If
        End If
    End Sub

    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' Paint the parent's background behind the rounded corners
        If Parent IsNot Nothing Then
            Using bg As New SolidBrush(Parent.BackColor)
                e.Graphics.FillRectangle(bg, ClientRectangle)
            End Using
        End If
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        Dim fillColor As Color
        If Not Enabled Then
            fillColor = Color.FromArgb(196, 196, 202)
        ElseIf _pressed Then
            fillColor = UiTheme.Blend(_accentDark, Color.Black, 0.12F)
        Else
            fillColor = UiTheme.Blend(_accent, _accentDark, _hoverProgress)
        End If

        Dim rect As New Rectangle(0, 0, Width - 1, Height - 1)
        Using path As GraphicsPath = UiTheme.RoundedPath(rect, _cornerRadius)
            Using brush As New SolidBrush(fillColor)
                e.Graphics.FillPath(brush, path)
            End Using
        End Using

        Dim textColor As Color = If(Enabled, ForeColor, Color.FromArgb(240, 240, 244))
        TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, textColor,
            TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            _animTimer.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

End Class
