Option Strict Off
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports NXOpen
Imports NXOpen.UF
Imports System.Runtime.InteropServices
Imports System.Text

Module NxSnipTool
    Dim theSession As Session = Session.GetSession()
    Private capturePrompt As String = "正在截图中"
    Private captionTimer As System.Windows.Forms.Timer
    Private existingWindows As New System.Collections.ArrayList()
    Private hasCancelButton As Boolean
    Private rectangleAnchor() As Double
    Private rectangleView As Tag
    Private rectangleMatrix As Matrix3x3
    Private rectangleError As String
    Const ExportScale As Double = 3.0
    Const MaxExportPixels As Double = 24000000.0

    Sub Main()
        If theSession.Parts.Work Is Nothing OrElse theSession.Parts.Display Is Nothing Then
            MsgBox("请先打开零件。")
            Return
        End If
        Dim uf As UFSession = UFSession.GetUFSession()
        Dim tempFile As String = IO.Path.Combine(IO.Path.GetTempPath(), "NX_snip_" & Guid.NewGuid().ToString("N") & ".png")
        Dim previousCursorView As Integer
        Dim locked As Boolean = False
        uf.Ui.AskCursorView(previousCursorView)
        Try
            ' 原生 NX 选点，不创建 Form，不显示截图预览窗口。
            existingWindows.Clear()
            EnumThreadWindows(GetCurrentThreadId(), AddressOf RememberWindow, IntPtr.Zero)
            captionTimer = New System.Windows.Forms.Timer()
            captionTimer.Interval = 120
            AddHandler captionTimer.Tick, AddressOf UpdateCaptureCaption
            captionTimer.Start()
            uf.Ui.SetStatus("正在截图中：请选择矩形截图范围")
            uf.Ui.SetCursorView(1)
            uf.Ui.LockUgAccess(UFConstants.UF_UI_FROM_CUSTOM)
            locked = True
            Dim first(2) As Double
            Dim second(2) As Double
            Dim firstView As Tag
            Dim secondView As Tag
            Dim response As Integer
            capturePrompt = "正在截图中 - 点击第一个角点"
            uf.Ui.SpecifyScreenPosition("正在截图中：点击第一个角点，取消可退出", Nothing, IntPtr.Zero, first, firstView, response)
            If response <> UFConstants.UF_UI_PICK_RESPONSE Then Return
            rectangleAnchor = CType(first.Clone(), Double())
            rectangleView = firstView
            rectangleMatrix = theSession.Parts.Display.ModelingViews.WorkView.Matrix
            rectangleError = Nothing
            ' NX 原生临时叠加线随鼠标更新；不建立窗口或模型曲线。
            capturePrompt = "正在截图中 - 点击对角点确认矩形范围"
            uf.Ui.SpecifyScreenPosition("正在截图中：移动鼠标显示矩形框，点击对角点确认", AddressOf DrawRectangle, IntPtr.Zero, second, secondView, response)
            If rectangleError IsNot Nothing Then Throw New InvalidOperationException(rectangleError)
            If response <> UFConstants.UF_UI_PICK_RESPONSE Then Return
            If firstView <> secondView Then
                MsgBox("请在同一个图形视图中指定两个角点。")
                Return
            End If
            uf.Ui.UnlockUgAccess(UFConstants.UF_UI_FROM_CUSTOM)
            locked = False

            Dim view As NXOpen.View = theSession.Parts.Display.ModelingViews.WorkView
            If firstView <> view.Tag Then
                MsgBox("请在当前工作视图中选取截图区域。")
                Return
            End If
            Dim clip(3) As Double
            uf.View.AskCurrentXyClip(view.Tag, clip)
            Dim matrix As Matrix3x3 = view.Matrix
            Dim origin As Point3d = view.Origin
            uf.Ui.SetStatus("正在截图中：生成高清图像，请稍候")
            Dim a As PointF = NormalizedPoint(first, matrix, origin, clip)
            Dim b As PointF = NormalizedPoint(second, matrix, origin, clip)
            If Not ExportViewImage(tempFile) Then Return
            Using src As New Bitmap(tempFile)
                Dim left As Integer = Math.Max(0, CInt(Math.Floor(Math.Min(a.X, b.X) * src.Width)))
                Dim top As Integer = Math.Max(0, CInt(Math.Floor(Math.Min(a.Y, b.Y) * src.Height)))
                Dim right As Integer = Math.Min(src.Width, CInt(Math.Ceiling(Math.Max(a.X, b.X) * src.Width)))
                Dim bottom As Integer = Math.Min(src.Height, CInt(Math.Ceiling(Math.Max(a.Y, b.Y) * src.Height)))
                If right - left < 2 OrElse bottom - top < 2 Then
                    MsgBox("截图区域太小，请重新选择。")
                    Return
                End If
                Using cropped As Bitmap = src.Clone(Rectangle.FromLTRB(left, top, right, bottom), src.PixelFormat)
                    Clipboard.SetImage(cropped)
                    uf.Ui.SetStatus("截图完成，已复制到剪贴板，可按 Ctrl+V 粘贴")
                End Using
            End Using
        Catch ex As Exception
            MsgBox("截图失败：" & ex.Message)
        Finally
            If captionTimer IsNot Nothing Then
                captionTimer.Stop()
                RemoveHandler captionTimer.Tick, AddressOf UpdateCaptureCaption
                captionTimer.Dispose()
                captionTimer = Nothing
            End If
            If locked Then uf.Ui.UnlockUgAccess(UFConstants.UF_UI_FROM_CUSTOM)
            uf.Ui.SetCursorView(previousCursorView)
            Try
                If IO.File.Exists(tempFile) Then IO.File.Delete(tempFile)
            Catch
            End Try
        End Try
    End Sub

    Private Delegate Function EnumWindowCallback(ByVal hwnd As IntPtr, ByVal parameter As IntPtr) As Boolean

    <DllImport("user32.dll")> _
    Private Function EnumThreadWindows(ByVal threadId As UInteger, ByVal callback As EnumWindowCallback, ByVal parameter As IntPtr) As Boolean
    End Function
    <DllImport("kernel32.dll")> _
    Private Function GetCurrentThreadId() As UInteger
    End Function
    <DllImport("user32.dll")> _
    Private Function IsWindowVisible(ByVal hwnd As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")> _
    Private Function EnumChildWindows(ByVal hwnd As IntPtr, ByVal callback As EnumWindowCallback, ByVal parameter As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Unicode)> _
    Private Function GetWindowText(ByVal hwnd As IntPtr, ByVal text As StringBuilder, ByVal count As Integer) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Unicode)> _
    Private Function SetWindowText(ByVal hwnd As IntPtr, ByVal text As String) As Boolean
    End Function

    Private Sub UpdateCaptureCaption(ByVal sender As Object, ByVal e As EventArgs)
        ' 仅更新本次选点新建且带取消按钮的对话框。
        EnumThreadWindows(GetCurrentThreadId(), AddressOf SetCaptureCaption, IntPtr.Zero)
    End Sub

    Private Function RememberWindow(ByVal hwnd As IntPtr, ByVal parameter As IntPtr) As Boolean
        If IsWindowVisible(hwnd) Then existingWindows.Add(hwnd)
        Return True
    End Function

    Private Function FindCancelButton(ByVal hwnd As IntPtr, ByVal parameter As IntPtr) As Boolean
        Dim text As New StringBuilder(128)
        GetWindowText(hwnd, text, text.Capacity)
        Dim label As String = text.ToString().Replace("&", "").Trim()
        If label = "取消" OrElse label.Equals("Cancel", StringComparison.OrdinalIgnoreCase) Then
            hasCancelButton = True
        End If
        Return True
    End Function

    Private Function SetCaptureCaption(ByVal hwnd As IntPtr, ByVal parameter As IntPtr) As Boolean
        If IsWindowVisible(hwnd) AndAlso Not existingWindows.Contains(hwnd) Then
            Dim text As New StringBuilder(512)
            GetWindowText(hwnd, text, text.Capacity)
            Dim title As String = text.ToString()
            If title.Length = 0 OrElse title.StartsWith("正在截图中") Then
                hasCancelButton = False
                EnumChildWindows(hwnd, AddressOf FindCancelButton, IntPtr.Zero)
                If hasCancelButton Then SetWindowText(hwnd, capturePrompt)
            End If
        End If
        Return True
    End Function

    Private Sub DrawRectangle(ByVal position() As Double, ByRef data As UFUi.MotionCbData, ByVal clientData As IntPtr)
        Try
            If data.view_tag <> rectangleView OrElse rectangleAnchor Is Nothing Then Return
            Dim dx As Double = position(0) - rectangleAnchor(0)
            Dim dy As Double = position(1) - rectangleAnchor(1)
            Dim dz As Double = position(2) - rectangleAnchor(2)
            Dim alongX As Double = dx * rectangleMatrix.Xx + dy * rectangleMatrix.Xy + dz * rectangleMatrix.Xz
            Dim alongY As Double = dx * rectangleMatrix.Yx + dy * rectangleMatrix.Yy + dz * rectangleMatrix.Yz
            Dim corners(4, 2) As Double
            For i As Integer = 0 To 2
                corners(0, i) = rectangleAnchor(i)
                corners(4, i) = rectangleAnchor(i)
            Next
            Dim xAxis() As Double = {rectangleMatrix.Xx, rectangleMatrix.Xy, rectangleMatrix.Xz}
            Dim yAxis() As Double = {rectangleMatrix.Yx, rectangleMatrix.Yy, rectangleMatrix.Yz}
            For i As Integer = 0 To 2
                corners(1, i) = rectangleAnchor(i) + alongX * xAxis(i)
                corners(2, i) = corners(1, i) + alongY * yAxis(i)
                corners(3, i) = rectangleAnchor(i) + alongY * yAxis(i)
            Next
            UFSession.GetUFSession().Disp.DisplayOgpPolyline(rectangleView, corners, 5)
        Catch ex As Exception
            ' 不从原生鼠标回调抛异常，返回后统一提示。
            rectangleError = ex.Message
        End Try
    End Sub

    Private Function NormalizedPoint(ByVal p() As Double, ByVal m As Matrix3x3, ByVal origin As Point3d, ByVal clip() As Double) As PointF
        Dim x As Double = m.Xx * p(0) + m.Xy * p(1) + m.Xz * p(2) + origin.X
        Dim y As Double = m.Yx * p(0) + m.Yy * p(1) + m.Yz * p(2) + origin.Y
        If clip(1) <= clip(0) OrElse clip(3) <= clip(2) Then Throw New InvalidOperationException("无法读取当前视图范围。")
        Return New PointF(CSng(Math.Max(0.0, Math.Min(1.0, (x - clip(0)) / (clip(1) - clip(0))))),
                          CSng(Math.Max(0.0, Math.Min(1.0, (clip(3) - y) / (clip(3) - clip(2))))))
    End Function

    Function ExportViewImage(ByVal filename As String) As Boolean

        Dim workPart As Part = theSession.Parts.Work
        Dim imageBuilder As NXOpen.Gateway.ImageExportBuilder = Nothing

        ' 备份用户界面显示状态
        Dim wcsVisible As Boolean = theSession.Parts.BaseDisplay.WCS.Visibility
        Dim triadVisible As Boolean = workPart.Preferences.ScreenVisualization.TriadVisibility
        Dim showNames As Boolean = theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewNames
        Dim showBorders As Boolean = theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewBorders

        Try
            ' 隐藏 WCS 坐标系 / 基准三角架 / 视图名 / 视图边框, 让截图更干净
            theSession.Parts.BaseDisplay.WCS.Visibility = False
            workPart.Preferences.ScreenVisualization.TriadVisibility = False
            theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewNames = False
            theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewBorders = False

            ' 若文件已存在则删除
            If IO.File.Exists(filename) Then IO.File.Delete(filename)

            ' NX 自带图像导出 API : ImageExportBuilder (即 文件 -> 导出 -> 图像)
            imageBuilder = workPart.Views.CreateImageExportBuilder()
            imageBuilder.RegionMode = False
            imageBuilder.FileFormat = NXOpen.Gateway.ImageExportBuilder.FileFormats.Png
            imageBuilder.FileName = filename
            imageBuilder.BackgroundOption = NXOpen.Gateway.ImageExportBuilder.BackgroundOptions.CustomColor

            ' 白色背景 (R=1, G=1, B=1)
            Dim bg(2) As Double
            bg(0) = 1.0
            bg(1) = 1.0
            bg(2) = 1.0
            imageBuilder.SetCustomBackgroundColor(bg)

            ' 先读取实际导出比例，再按比例提升设备尺寸，避免拉伸或改变取景。
            imageBuilder.Commit()
            Dim baseWidth As Integer
            Dim baseHeight As Integer
            Using probe As Image = Image.FromFile(filename)
                baseWidth = probe.Width
                baseHeight = probe.Height
            End Using
            Dim factor As Double = Math.Min(ExportScale, Math.Sqrt(MaxExportPixels / (CDbl(baseWidth) * baseHeight)))
            factor = Math.Max(1.0, factor)
            Dim exportWidth As Integer = Math.Max(baseWidth, CInt(Math.Floor(baseWidth * factor)))
            Dim exportHeight As Integer = Math.Max(baseHeight, CInt(Math.Floor(baseHeight * factor)))
            imageBuilder.DeviceWidth = exportWidth
            imageBuilder.DeviceHeight = exportHeight
            ' 保留抗锯齿，避免边线强调模式产生明显锯齿。
            imageBuilder.EnhanceEdges = False
            IO.File.Delete(filename)
            imageBuilder.Commit()

            Return True

        Catch ex As Exception
            MsgBox("导出图像时出错: " & ex.Message)
            Return False

        Finally
            If imageBuilder IsNot Nothing Then imageBuilder.Destroy()
            ' 恢复显示状态
            theSession.Parts.BaseDisplay.WCS.Visibility = wcsVisible
            workPart.Preferences.ScreenVisualization.TriadVisibility = triadVisible
            theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewNames = showNames
            theSession.Parts.BaseDisplay.Preferences.NamesBorderVisualization.ShowModelViewBorders = showBorders
        End Try

    End Function
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return NXOpen.Session.LibraryUnloadOption.Immediately
    End Function
End Module
