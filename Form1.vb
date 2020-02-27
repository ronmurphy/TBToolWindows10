Imports System.Runtime.InteropServices

Public Class Form1
    Friend Structure WindowCompositionAttributeData
        Public Attribute As WindowCompositionAttribute
        Public Data As IntPtr
        Public SizeOfData As Integer
    End Structure

    Friend Enum WindowCompositionAttribute
        WCA_ACCENT_POLICY = 19
    End Enum

    Friend Enum AccentState
        ACCENT_DISABLED = 0
        ACCENT_ENABLE_GRADIENT = 1
        ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
        ACCENT_ENABLE_BLURBEHIND = 3
        ACCENT_ENABLE_TRANSPARANT = 6
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
    End Enum

    ' <StructLayout(LayoutKind.Sequential)>
    Friend Structure AccentPolicy
        Public AccentState As AccentState
        Public AccentFlags As Integer
        Public GradientColor As Integer
        Public AnimationId As Integer
    End Structure

    Friend Declare Function SetWindowCompositionAttribute Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef data As WindowCompositionAttributeData) As Integer
    Private Declare Auto Function FindWindow Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    Dim ctb As New clsCTB
    Friend Sub EnableTaskbarStyle()
        If ComboBox1.Text = "None" Then Exit Sub

        Dim tskBarClassName As String = "Shell_TrayWnd"
        Dim tskBarHwnd As IntPtr = FindWindow(tskBarClassName, Nothing)
        Dim accent = New AccentPolicy()
        Dim accentStructSize = Marshal.SizeOf(accent)
        If ComboBox1.Text = "Transparent" Then
            ' # Taskbar Style Transparant
            accent.AccentState = AccentState.ACCENT_ENABLE_TRANSPARANT
        End If
        If ComboBox1.Text = "Arcrylic" Then
            ' # Taskbar Style Acrylic
            accent.AccentState = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND
            accent.GradientColor = 10 'Or 16777215
        End If
        If ComboBox1.Text = "Blur" Then
            ' # Taskbar Style Blur
            accent.AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND
        End If

        'Transparent
        'Arcrylic
        'Blur



        Dim accentPtr = Marshal.AllocHGlobal(accentStructSize)
        Marshal.StructureToPtr(accent, accentPtr, False)
        Dim data = New WindowCompositionAttributeData()
        data.Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY
        data.SizeOfData = accentStructSize
        data.Data = accentPtr
        SetWindowCompositionAttribute(tskBarHwnd, data)
        Marshal.FreeHGlobal(accentPtr)
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadSettings()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        EnableTaskbarStyle()
    End Sub



    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            ctb.AddApplicationToStartup()
            UpdateSettings()
            Exit Sub
        Else
            ctb.RemoveApplicationFromStartup()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox2.Checked = True Then
            ctb.Start()
            Exit Sub
        Else
            ctb.ResetAll()

        End If
    End Sub
    Private Sub UpdateSettings()
        My.Settings.AutoCenterTasks = CheckBox2.Checked
        My.Settings.TBType = ComboBox1.Text
        My.Settings.AutoStartWindows = CheckBox1.Checked
        My.Settings.Save()
    End Sub

    Private Sub LoadSettings()
        CheckBox2.Checked = My.Settings.AutoCenterTasks
        ComboBox1.Text = My.Settings.TBType
        EnableTaskbarStyle()
        CheckBox1.Checked = My.Settings.AutoStartWindows
    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ctb.ResetAll()
    End Sub
End Class
