Imports System.Drawing.Drawing2D

''' <summary>
''' Slim rounded progress bar in brand magenta. The displayed value eases
''' toward the target so progress glides instead of jumping. Replaces the
''' Win32 PBM_SETBARCOLOR hack (which comctl32 v6 visual styles ignore).
''' </summary>
Public Class SmoothProgressBar
    Inherits Control

    Private ReadOnly _animTimer As New Timer() With {.Interval = 15}
    Private _target As Integer = 0
    Private _displayed As Single = 0.0F

    ''' <summary>Target progress value, 0-100. The bar animates toward it.</summary>
    Public Property Value As Integer
        Get
            Return _target
        End Get
        Set(newValue As Integer)
            Dim clamped As Integer = Math.Max(0, Math.Min(100, newValue))
            ' Jumping backwards (a reset) should not animate in reverse
            If clamped < _target AndAlso clamped = 0 Then
                _displayed = 0.0F
            End If
            _target = clamped
            _animTimer.Start()
        End Set
    End Property

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Size = New Size(200, 6)
        AddHandler _animTimer.Tick, AddressOf AnimTick
    End Sub

    Private Sub AnimTick(sender As Object, e As EventArgs)
        Dim delta As Single = _target - _displayed
        If Math.Abs(delta) < 0.4F Then
            _displayed = _target
            _animTimer.Stop()
        Else
            _displayed += delta * 0.16F
        End If
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        If Parent IsNot Nothing Then
            Using bg As New SolidBrush(Parent.BackColor)
                e.Graphics.FillRectangle(bg, ClientRectangle)
            End Using
        End If
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        Dim radius As Integer = Math.Max(1, Height \ 2)
        Dim track As New Rectangle(0, 0, Width - 1, Height - 1)
        Using trackPath As GraphicsPath = UiTheme.RoundedPath(track, radius)
            Using brush As New SolidBrush(UiTheme.PinkTrack)
                e.Graphics.FillPath(brush, trackPath)
            End Using
        End Using

        Dim fillWidth As Integer = CInt(Math.Round((Width - 1) * (_displayed / 100.0F)))
        If fillWidth > 0 Then
            ' Keep at least a full cap so the rounded end never looks clipped
            fillWidth = Math.Max(fillWidth, Height)
            Dim fillRect As New Rectangle(0, 0, Math.Min(fillWidth, Width - 1), Height - 1)
            Using fillPath As GraphicsPath = UiTheme.RoundedPath(fillRect, radius)
                Using brush As New SolidBrush(UiTheme.Magenta)
                    e.Graphics.FillPath(brush, fillPath)
                End Using
            End Using
        End If
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            _animTimer.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

End Class
