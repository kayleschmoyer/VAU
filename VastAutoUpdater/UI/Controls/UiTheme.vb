Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices

''' <summary>
''' Central brand palette, typography helpers, and window-chrome utilities
''' shared by all custom-drawn UI. Keeps every control on the same magenta /
''' charcoal / Segoe UI system.
''' </summary>
Public Module UiTheme

    ' -- Brand palette --
    Public ReadOnly Magenta As Color = Color.FromArgb(237, 1, 127)
    Public ReadOnly MagentaDark As Color = Color.FromArgb(180, 0, 96)
    Public ReadOnly MagentaDeep As Color = Color.FromArgb(150, 0, 80)
    Public ReadOnly Charcoal As Color = Color.FromArgb(51, 51, 51)

    ' -- Neutral surfaces (no new brand colors; neutral grays for depth) --
    Public ReadOnly Canvas As Color = Color.FromArgb(246, 246, 248)
    Public ReadOnly CardBorder As Color = Color.FromArgb(228, 228, 232)
    Public ReadOnly InputBorder As Color = Color.FromArgb(224, 224, 228)
    Public ReadOnly TextMuted As Color = Color.FromArgb(138, 138, 144)

    ' -- Brand tints for tracks and badges --
    Public ReadOnly PinkPale As Color = Color.FromArgb(251, 224, 238)
    Public ReadOnly PinkTrack As Color = Color.FromArgb(243, 224, 235)

    ' -- Status colors (already used by the app) --
    Public ReadOnly SuccessGreen As Color = Color.FromArgb(0, 150, 80)
    Public ReadOnly SuccessPale As Color = Color.FromArgb(227, 244, 236)
    Public ReadOnly ErrorRed As Color = Color.FromArgb(200, 0, 0)
    Public ReadOnly ErrorPale As Color = Color.FromArgb(250, 228, 228)

    Private _mdl2Available As Boolean? = Nothing

    ''' <summary>
    ''' True when the Segoe MDL2 Assets icon font is installed (Windows 10/11).
    ''' </summary>
    Public ReadOnly Property Mdl2Available As Boolean
        Get
            If Not _mdl2Available.HasValue Then
                Try
                    Using f As New Font("Segoe MDL2 Assets", 10.0F)
                        _mdl2Available = f.Name.Equals("Segoe MDL2 Assets", StringComparison.OrdinalIgnoreCase)
                    End Using
                Catch
                    _mdl2Available = False
                End Try
            End If
            Return _mdl2Available.Value
        End Get
    End Property

    ''' <summary>
    ''' Icon font for MDL2 glyphs, falling back to Segoe UI Symbol on older systems.
    ''' </summary>
    Public Function IconFont(sizePt As Single) As Font
        Return New Font(If(Mdl2Available, "Segoe MDL2 Assets", "Segoe UI Symbol"), sizePt)
    End Function

    ''' <summary>
    ''' Segoe UI Semibold if installed, otherwise Segoe UI Bold. Same family, better hierarchy.
    ''' </summary>
    Public Function Semibold(sizePt As Single) As Font
        Dim f As New Font("Segoe UI Semibold", sizePt)
        If f.Name.Equals("Segoe UI Semibold", StringComparison.OrdinalIgnoreCase) Then Return f
        f.Dispose()
        Return New Font("Segoe UI", sizePt, FontStyle.Bold)
    End Function

    ''' <summary>
    ''' Build a rounded-rectangle path. Caller is responsible for disposing it.
    ''' </summary>
    Public Function RoundedPath(rect As Rectangle, radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        If radius <= 0 OrElse rect.Width <= 0 OrElse rect.Height <= 0 Then
            path.AddRectangle(rect)
            Return path
        End If
        Dim d As Integer = Math.Min(radius * 2, Math.Min(rect.Width, rect.Height))
        path.AddArc(rect.X, rect.Y, d, d, 180, 90)
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90)
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90)
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90)
        path.CloseFigure()
        Return path
    End Function

    ''' <summary>
    ''' Linear blend between two colors. t is clamped to [0, 1].
    ''' </summary>
    Public Function Blend(a As Color, b As Color, t As Single) As Color
        Dim clamped As Single = Math.Max(0.0F, Math.Min(1.0F, t))
        Return Color.FromArgb(
            CInt(CInt(a.R) + (CInt(b.R) - CInt(a.R)) * clamped),
            CInt(CInt(a.G) + (CInt(b.G) - CInt(a.G)) * clamped),
            CInt(CInt(a.B) + (CInt(b.B) - CInt(a.B)) * clamped))
    End Function

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Function DwmSetWindowAttribute(hWnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    Private Const DWMWA_WINDOW_CORNER_PREFERENCE As Integer = 33
    Private Const DWMWCP_ROUND As Integer = 2

    ''' <summary>
    ''' Ask DWM for native rounded window corners (Windows 11).
    ''' Silently no-ops on Windows 10 and earlier.
    ''' </summary>
    Public Sub ApplyRoundedCorners(handle As IntPtr)
        Try
            Dim pref As Integer = DWMWCP_ROUND
            DwmSetWindowAttribute(handle, DWMWA_WINDOW_CORNER_PREFERENCE, pref, 4)
        Catch
            ' DWM unavailable (old OS) -- square corners are an acceptable fallback
        End Try
    End Sub

    ''' <summary>
    ''' Make a control (and the form it lives on) draggable by that control,
    ''' for borderless windows.
    ''' </summary>
    Public Sub AttachDrag(form As Form, ctrl As Control)
        Dim dragging As Boolean = False
        Dim start As Point = Point.Empty
        AddHandler ctrl.MouseDown,
            Sub(s, e)
                If e.Button = MouseButtons.Left Then
                    dragging = True
                    start = e.Location
                End If
            End Sub
        AddHandler ctrl.MouseMove,
            Sub(s, e)
                If dragging Then
                    Dim screenPoint As Point = ctrl.PointToScreen(e.Location)
                    form.Location = New Point(screenPoint.X - start.X, screenPoint.Y - start.Y)
                End If
            End Sub
        AddHandler ctrl.MouseUp, Sub(s, e) dragging = False
    End Sub

End Module
