Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Automation

Public Class clsCTB
    Public Const appName As String = "CTB"
    Private Const SWP_NOSIZE As Integer = &H1
    Private Const SWP_NOZORDER As Integer = &H4
    Private Const SWP_SHOWWINDOW As Integer = &H40
    Private Const SWP_ASYNCWINDOWPOS As Integer = &H4000
    Private trayIcon As NotifyIcon
    Private Shared desktop As AutomationElement = AutomationElement.RootElement
    Private Shared MSTaskListWClass As String = "MSTaskListWClass"
    'static String ReBarWindow32 = "ReBarWindow32";
    Private Shared Shell_TrayWnd As String = "Shell_TrayWnd"
    Private Shared Shell_SecondaryTrayWnd As String = "Shell_SecondaryTrayWnd"
    Private lasts As Dictionary(Of AutomationElement, Double) = New Dictionary(Of AutomationElement, Double)()
    Private children As Dictionary(Of AutomationElement, AutomationElement) = New Dictionary(Of AutomationElement, AutomationElement)()
    Private bars As List(Of AutomationElement) = New List(Of AutomationElement)()
    Private activeFramerate As Integer = 60
    Private positionThread As Thread

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As Integer) As Boolean
    End Function

    ' Public Sub New(ByVal args As String())
    ' If args.Length > 0 Then
    '
    '    Try
    '                activeFramerate = Integer.Parse(args(0))
    '                Debug.WriteLine("Active refresh rate: " & activeFramerate)
    '    Catch e As FormatException
    '                Debug.WriteLine(e.Message)
    '    End Try
    '    End If
    '
    '    Dim header As MenuItem = New MenuItem("CTB (" & activeFramerate & " fps)", AddressOf [Exit])
    '        header.Enabled = False
    '    Dim startup As MenuItem = New MenuItem("Start with Windows", AddressOf ToggleStartup)
    '        startup.Checked = IsApplicationInStatup()
    '
    '        ' Setup Tray Icon
    '        trayIcon = New NotifyIcon() With {
    '    .Icon = My.Resources.Icon1, 'Me.Resources.Icon1,
    '    .ContextMenu = New ContextMenu(New MenuItem() {header, New MenuItem("Scan for screens", AddressOf Restart), startup, New MenuItem("Exit", AddressOf [Exit])}),
    '    .Visible = True
    '            }
    '        Start()
    '    End Sub
    Public Sub Main()
        '  If args.Length > 0 Then
        '
        '            Try
        '            activeFramerate = Integer.Parse(args(0))
        '            Debug.WriteLine("Active refresh rate: " & activeFramerate)
        '            Catch e As FormatException
        '            Debug.WriteLine(e.Message)
        '            End Try
        '            End If

        ' Dim header As MenuItem = New MenuItem("CTB (" & activeFramerate & " fps)", AddressOf [Exit])
        ' header.Enabled = False
        ' Dim startup As MenuItem = New MenuItem("Start with Windows", AddressOf ToggleStartup)
        ' startup.Checked = IsApplicationInStatup()
        '
        '        ' Setup Tray Icon
        '        trayIcon = New NotifyIcon() With {
        '        .Icon = My.Resources.Icon1,
        '        .ContextMenu = New ContextMenu(New MenuItem() {header, New MenuItem("Scan for screens", AddressOf Restart), startup, New MenuItem("Exit", AddressOf [Exit])}),
        '        .Visible = True
        '            }
        '        Dim gui As MenuItem = New MenuItem("Gui Config", AddressOf ShowGUI)
        '
        '        Start()
    End Sub
    Public Sub ShowGUI()

    End Sub
    Public Sub ToggleStartup(ByVal sender As Object, ByVal e As EventArgs)
        If IsApplicationInStatup() Then
            RemoveApplicationFromStartup()
            TryCast(sender, MenuItem).Checked = False
        Else
            AddApplicationToStartup()
            TryCast(sender, MenuItem).Checked = True
        End If
    End Sub

    Public Function IsApplicationInStatup() As Boolean
        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            If key Is Nothing Then Return False
            Dim value As Object = key.GetValue(appName)
            If TypeOf value Is String Then Return (TryCast(value, String).StartsWith("""" & Application.ExecutablePath & """"))
            Return False
        End Using
    End Function

    Public Sub AddApplicationToStartup()
        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            key.SetValue(appName, """" & Application.ExecutablePath & """ " & activeFramerate)
        End Using
    End Sub

    Public Sub RemoveApplicationFromStartup()
        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            key.DeleteValue(appName, False)
        End Using
    End Sub

    Private Sub [Exit](ByVal sender As Object, ByVal e As EventArgs)
        ' Hide tray icon, otherwise it will remain shown until user mouses over it
        trayIcon.Visible = False
        ResetAll()
        Application.Exit()
    End Sub

    Private Sub Restart(ByVal sender As Object, ByVal e As EventArgs)
        If positionThread IsNot Nothing Then
            positionThread.Abort()
        End If

        Start()
    End Sub

    Public Sub ResetAll()
        If positionThread IsNot Nothing Then
            positionThread.Abort()
        End If

        For Each trayWnd As AutomationElement In bars
            Reset(trayWnd)
        Next
    End Sub

    Public Sub Reset(ByVal trayWnd As AutomationElement)
        Debug.WriteLine("Begin Reset Calculation")
        Dim tasklist As AutomationElement = trayWnd.FindFirst(TreeScope.Descendants, New PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass))

        If tasklist Is Nothing Then
            Debug.WriteLine("Null values found, aborting reset")
            Return
        End If

        Dim tasklistcontainer As AutomationElement = TreeWalker.ControlViewWalker.GetParent(tasklist)

        If tasklistcontainer Is Nothing Then
            Debug.WriteLine("Null values found, aborting reset")
            Return
        End If

        Dim trayBounds As Rect = trayWnd.Cached.BoundingRectangle
        Dim horizontal As Boolean = trayBounds.Width > trayBounds.Height
        Dim tasklistPtr As IntPtr = CType(tasklist.Current.NativeWindowHandle, IntPtr)
        Dim listBounds As Double = If(horizontal, tasklist.Current.BoundingRectangle.X, tasklist.Current.BoundingRectangle.Y)
        Dim bounds As Rect = tasklist.Current.BoundingRectangle
        Dim newWidth As Integer = bounds.Width
        Dim newHeight As Integer = bounds.Height
        SetWindowPos(tasklistPtr, IntPtr.Zero, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_ASYNCWINDOWPOS)
    End Sub

    Public Sub Start()
        Dim condition As OrCondition = New OrCondition(New PropertyCondition(AutomationElement.ClassNameProperty, Shell_TrayWnd), New PropertyCondition(AutomationElement.ClassNameProperty, Shell_SecondaryTrayWnd))
        Dim cacheRequest As CacheRequest = New CacheRequest()
        cacheRequest.Add(AutomationElement.NameProperty)
        cacheRequest.Add(AutomationElement.BoundingRectangleProperty)
        bars.Clear()
        children.Clear()
        lasts.Clear()

        Using cacheRequest.Activate()
            Dim lists As AutomationElementCollection = desktop.FindAll(TreeScope.Children, condition)

            If lists Is Nothing Then
                Debug.WriteLine("Null values found, aborting")
                Return
            End If

            Debug.WriteLine(lists.Count & " bar(s) detected")
            lasts.Clear()

            For Each trayWnd As AutomationElement In lists
                Dim tasklist As AutomationElement = trayWnd.FindFirst(TreeScope.Descendants, New PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass))

                If tasklist Is Nothing Then
                    Debug.WriteLine("Null values found, aborting")
                    Continue For
                End If

                Automation.AddAutomationPropertyChangedEventHandler(tasklist, TreeScope.Element, New AutomationPropertyChangedEventHandler(AddressOf OnUIAutomationEvent), AutomationElement.BoundingRectangleProperty)
                bars.Add(trayWnd)
                children.Add(trayWnd, tasklist)
            Next
        End Using

        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, desktop, TreeScope.Subtree, New AutomationEventHandler(AddressOf OnUIAutomationEvent))
        Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, desktop, TreeScope.Subtree, New AutomationEventHandler(AddressOf OnUIAutomationEvent))
        [loop]()
    End Sub

    Public Sub OnUIAutomationEvent(ByVal src As Object, ByVal e As AutomationEventArgs)
        If Not positionThread.IsAlive Then
            [loop]()
        End If
    End Sub

    Public Sub [loop]()
        positionThread = New Thread(Sub()
                                        Dim keepGoing As Integer = 0

                                        While keepGoing < activeFramerate / 5

                                            For Each trayWnd As AutomationElement In bars

                                                If Not PositionLoop(trayWnd) Then
                                                    keepGoing += 1
                                                End If
                                            Next

                                            Thread.Sleep(1000 / activeFramerate)
                                        End While

                                        Debug.WriteLine("Thread ended due to inactivity, sleeping")
                                    End Sub)
        positionThread.Start()
    End Sub

    Public Function PositionLoop(ByVal trayWnd As AutomationElement) As Boolean
        Debug.WriteLine("Begin Reposition Calculation")
        Dim tasklist As AutomationElement = children(trayWnd)
        Dim last As AutomationElement = TreeWalker.ControlViewWalker.GetLastChild(tasklist)

        If last Is Nothing Then
            Debug.WriteLine("Null values found for items, aborting")
            Return True
        End If

        Dim trayBounds As Rect = trayWnd.Cached.BoundingRectangle
        Dim horizontal As Boolean = trayBounds.Width > trayBounds.Height
        Dim lastChildPos As Double = If(horizontal, last.Current.BoundingRectangle.Left, last.Current.BoundingRectangle.Top) ' Use the left/top bounds because there is an empty element as the last child with a nonzero width
        Debug.WriteLine("Last child position: " & lastChildPos)

        If lasts.ContainsKey(trayWnd) AndAlso lastChildPos = lasts(trayWnd) Then
            Debug.WriteLine("Size/location unchanged, sleeping")
            Return False
        Else
            Debug.WriteLine("Size/location changed, recalculating center")
            lasts(trayWnd) = lastChildPos
            Dim first As AutomationElement = TreeWalker.ControlViewWalker.GetFirstChild(tasklist)

            If first Is Nothing Then
                Debug.WriteLine("Null values found for first child item, aborting")
                Return True
            End If

            Dim scale As Double = If(horizontal, last.Current.BoundingRectangle.Height / trayBounds.Height, last.Current.BoundingRectangle.Width / trayBounds.Width)
            Debug.WriteLine("UI Scale: " & scale)
            Dim size As Double = (lastChildPos - If(horizontal, first.Current.BoundingRectangle.Left, first.Current.BoundingRectangle.Top)) / scale

            If size < 0 Then
                Debug.WriteLine("Size calculation failed")
                Return True
            End If

            Dim tasklistcontainer As AutomationElement = TreeWalker.ControlViewWalker.GetParent(tasklist)

            If tasklistcontainer Is Nothing Then
                Debug.WriteLine("Null values found for parent, aborting")
                Return True
            End If

            Dim tasklistBounds As Rect = tasklist.Current.BoundingRectangle
            Dim barSize As Double = If(horizontal, trayWnd.Cached.BoundingRectangle.Width, trayWnd.Cached.BoundingRectangle.Height)
            Dim targetPos As Double = Math.Round((barSize - size) / 2) + If(horizontal, trayBounds.X, trayBounds.Y)
            Debug.Write("Bar size: ")
            Debug.WriteLine(barSize)
            Debug.Write("Total icon size: ")
            Debug.WriteLine(size)
            Debug.Write("Target abs " & If(horizontal, "X", "Y") & " position: ")
            Debug.WriteLine(targetPos)
            Dim delta As Double = Math.Abs(targetPos - If(horizontal, tasklistBounds.X, tasklistBounds.Y))

            ' Previous bounds check
            If delta <= 1 Then
                ' Already positioned within margin of error, avoid the unneeded MoveWindow call
                Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call (Delta: " & delta & ")")
                Return False
            End If


            ' Right bounds check
            Dim rightBounds As Integer = sideBoundary(False, horizontal, tasklist)

            If targetPos + size > rightBounds Then
                ' Shift off center when the bar is too big
                Dim extra As Double = targetPos + size - rightBounds
                Debug.WriteLine("Shifting off center, too big and hitting right/bottom boundary (" & targetPos + size & " > " & rightBounds & ") // " & extra)
                targetPos -= extra
            End If


            ' Left bounds check
            Dim leftBounds As Integer = sideBoundary(True, horizontal, tasklist)

            If targetPos <= leftBounds Then
                ' Prevent X position ending up beyond the normal left aligned position
                Debug.WriteLine("Target is more left than left/top aligned default, left/top aligning (" & targetPos & " <= " & leftBounds & ")")
                Reset(trayWnd)
                Return True
            End If

            Dim tasklistPtr As IntPtr = CType(tasklist.Current.NativeWindowHandle, IntPtr)

            If horizontal Then
                SetWindowPos(tasklistPtr, IntPtr.Zero, relativePos(targetPos, horizontal, tasklist), 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_ASYNCWINDOWPOS)
                Debug.Write("Final X Position: ")
                Debug.WriteLine(tasklist.Current.BoundingRectangle.X)
                Debug.Write(If(tasklist.Current.BoundingRectangle.X = targetPos, "Move hit target", "Move missed target"))
                Debug.WriteLine(" (diff: " & Math.Abs(tasklist.Current.BoundingRectangle.X - targetPos) & ")")
            Else
                SetWindowPos(tasklistPtr, IntPtr.Zero, 0, relativePos(targetPos, horizontal, tasklist), 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_ASYNCWINDOWPOS)
                Debug.Write("Final Y Position: ")
                Debug.WriteLine(tasklist.Current.BoundingRectangle.Y)
                Debug.Write(If(tasklist.Current.BoundingRectangle.Y = targetPos, "Move hit target", "Move missed target"))
                Debug.WriteLine(" (diff: " & Math.Abs(tasklist.Current.BoundingRectangle.Y - targetPos) & ")")
            End If

            lasts(trayWnd) = If(horizontal, last.Current.BoundingRectangle.Left, last.Current.BoundingRectangle.Top)
            Return True
        End If
    End Function

    Public Function relativePos(ByVal x As Double, ByVal horizontal As Boolean, ByVal element As AutomationElement) As Integer
        Dim adjustment As Integer = sideBoundary(True, horizontal, element)
        Dim newPos As Double = x - adjustment

        If newPos < 0 Then
            Debug.WriteLine("Relative position < 0, adjusting to 0 (Previous: " & newPos & ")")
            newPos = 0
        End If

        Return newPos
    End Function

    Public Function sideBoundary(ByVal left As Boolean, ByVal horizontal As Boolean, ByVal element As AutomationElement) As Integer
        Dim adjustment As Double = 0
        Dim prevSibling As AutomationElement = TreeWalker.ControlViewWalker.GetPreviousSibling(element)
        Dim nextSibling As AutomationElement = TreeWalker.ControlViewWalker.GetNextSibling(element)
        Dim parent As AutomationElement = TreeWalker.ControlViewWalker.GetParent(element)

        If left AndAlso prevSibling IsNot Nothing Then
            adjustment = If(horizontal, prevSibling.Current.BoundingRectangle.Right, prevSibling.Current.BoundingRectangle.Bottom)
        ElseIf Not left AndAlso nextSibling IsNot Nothing Then
            adjustment = If(horizontal, nextSibling.Current.BoundingRectangle.Left, nextSibling.Current.BoundingRectangle.Top)
        ElseIf parent IsNot Nothing Then

            If horizontal Then
                adjustment = If(left, parent.Current.BoundingRectangle.Left, parent.Current.BoundingRectangle.Right)
            Else
                adjustment = If(left, parent.Current.BoundingRectangle.Top, parent.Current.BoundingRectangle.Bottom)
            End If
        End If

        If horizontal Then
            Debug.WriteLine(If(left, "Left", "Right") & " side boundary calulcated at " & adjustment)
        Else
            Debug.WriteLine(If(left, "Top", "Bottom") & " side boundary calulcated at " & adjustment)
        End If

        Return adjustment
    End Function
End Class
