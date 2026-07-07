Imports System.Drawing.Drawing2D

''' <summary>
''' Small status pill (e.g. "Update available", "Up to date"). Auto-sizes to
''' its text; colors follow the light-tint / dark-text pattern.
''' </summary>
Public Class PillBadge
    Inherits Control

    Private _fillColor As Color = UiTheme.PinkPale
    Private _textColor As Color = UiTheme.MagentaDark

    Public Property FillColor As Color
        Get
            Return _fillColor
        End Get
        Set(value As Color)
            _fillColor = value
            Invalidate()
        End Set
    End Property

    Public Property TextColor As Color
        Get
            Return _textColor
        End Get
        Set(value As Color)
            _textColor = value
            Invalidate()
        End Set
    End Property

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or
                 ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        Font = UiTheme.Semibold(8.25F)
        Height = 22
    End Sub

    ''' <summary>
    ''' Set text and colors in one call, resizing the pill to fit.
    ''' </summary>
    Public Sub SetState(text As String, fill As Color, textColor As Color)
        _fillColor = fill
        _textColor = textColor
        Me.Text = text
        ResizeToText()
        Invalidate()
    End Sub

    Private Sub ResizeToText()
        Dim measured As Size = TextRenderer.MeasureText(Text, Font)
        Width = measured.Width + 22
    End Sub

    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        ResizeToText()
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        If Parent IsNot Nothing Then
            Using bg As New SolidBrush(Parent.BackColor)
                e.Graphics.FillRectangle(bg, ClientRectangle)
            End Using
        End If
        If String.IsNullOrEmpty(Text) Then Return

        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Dim rect As New Rectangle(0, 0, Width - 1, Height - 1)
        Using path As GraphicsPath = UiTheme.RoundedPath(rect, Height \ 2)
            Using brush As New SolidBrush(_fillColor)
                e.Graphics.FillPath(brush, path)
            End Using
        End Using
        TextRenderer.DrawText(e.Graphics, Text, Font, rect, _textColor,
            TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
    End Sub

End Class
