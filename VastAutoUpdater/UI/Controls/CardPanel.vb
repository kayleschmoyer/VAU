Imports System.Drawing.Drawing2D

''' <summary>
''' A panel drawn as a white rounded card with a hairline border, sitting on
''' the light-gray form canvas. Gives the layout depth without new colors.
''' </summary>
Public Class CardPanel
    Inherits Panel

    Private _cornerRadius As Integer = 10

    Public Property CornerRadius As Integer
        Get
            Return _cornerRadius
        End Get
        Set(value As Integer)
            _cornerRadius = value
            Invalidate()
        End Set
    End Property

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        BackColor = Color.White
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' Corners must blend into the canvas behind the card
        Dim canvasColor As Color = If(Parent IsNot Nothing, Parent.BackColor, UiTheme.Canvas)
        e.Graphics.Clear(canvasColor)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        Dim rect As New Rectangle(0, 0, Width - 1, Height - 1)
        Using path As GraphicsPath = UiTheme.RoundedPath(rect, _cornerRadius)
            Using fill As New SolidBrush(Color.White)
                e.Graphics.FillPath(fill, path)
            End Using
            Using pen As New Pen(UiTheme.CardBorder, 1.0F)
                e.Graphics.DrawPath(pen, path)
            End Using
        End Using
    End Sub

End Class
