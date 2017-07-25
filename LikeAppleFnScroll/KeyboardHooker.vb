' Reference from http://thom.hateblo.jp/entry/2015/02/11/073150

Imports System.Runtime.InteropServices
Public Class KeyboardHooker

    Const WM_NULL As Integer = 0
    Const WM_KEYDOWN As Integer = &H100
    Const WM_KEYUP As Integer = &H101

    Public Sub New()
        hookproc = AddressOf KeybordHookProc
        hHook = SetWindowsHookEx(WH_KEYBOARD_LL, hookproc, GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName), 0)
        If hHook.Equals(0) Then
            MsgBox("SetWindowsHookEx Failed")
        End If
    End Sub

    Dim WH_KEYBOARD_LL As Integer = 13
    Shared hHook As Integer = 0
    Shared KillingModifier As Boolean

    Private hookproc As CallBack

    Public Delegate Function CallBack(
        ByVal nCode As Integer,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Integer

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function SetWindowsHookEx _
          (ByVal idHook As Integer, ByVal HookProc As CallBack,
    ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function GetModuleHandle _
    (ByVal lpModuleName As String) As IntPtr
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function CallNextHookEx _
          (ByVal idHook As Integer, ByVal nCode As Integer,
    ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function UnhookWindowsHookEx _
    (ByVal idHook As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> Public Structure KeyboardLLHookStruct
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    <DllImport("User32.dll")>
    Public Shared Function GetAsyncKeyState(vKey As Keys) As Short
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function

    Protected KeyDownMsg As IntPtr = WM_KEYDOWN
    Protected KeyUpMsg As IntPtr = WM_KEYUP
    Protected HWND_BROADCAST As IntPtr = &HFFFF

    Public Function KeybordHookProc(
        ByVal nCode As Integer,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Integer

        If (nCode < 0) Then
            Return CallNextHookEx(hHook, nCode, wParam, lParam)
        End If

        Dim hookStruct As New KeyboardLLHookStruct()
        hookStruct = CType(Marshal.PtrToStructure(lParam, hookStruct.GetType()), KeyboardLLHookStruct)

        If KeyDownMsg.Equals(wParam) Or KeyUpMsg.Equals(wParam) Then
            Dim _keyup As Boolean
            _keyup = KeyUpMsg.Equals(wParam)
            Dim e As New KeyBoardHookerEventArgs
            e.vkCode = hookStruct.vkCode
            Dim r As Integer
            If e.vkCode = Keys.RControlKey Then
                ' 右Ctrlキーが押されているかどうかを共有変数に代入
                If _keyup Then
                    If KillingModifier Then
                        KillingModifier = False
                    End If
                Else
                    If Not KillingModifier Then
                        KillingModifier = True
                    End If
                End If
                ' 右Ctrlキーは全部握りつぶす
                Return 1
            End If
            If KillingModifier Then
                ' 右Ctrlキーが押されている場合はキーコードに応じて動作を上書き
                ' ここでキーコードを判定しているのは、イベントハンドラは戻り値を返せないため
                Select Case e.vkCode
                    Case Keys.Up
                        If _keyup Then
                            RaiseEvent KeyUp(Me, e)
                        Else
                            RaiseEvent KeyDown(Me, e)
                        End If
                        r = 1
                    Case Keys.Down
                        If _keyup Then
                            RaiseEvent KeyUp(Me, e)
                        Else
                            RaiseEvent KeyDown(Me, e)
                        End If
                        r = 1
                    Case Keys.Left
                        If _keyup Then
                            RaiseEvent KeyUp(Me, e)
                        Else
                            RaiseEvent KeyDown(Me, e)
                        End If
                        r = 1
                    Case Keys.Right
                        If _keyup Then
                            RaiseEvent KeyUp(Me, e)
                        Else
                            RaiseEvent KeyDown(Me, e)
                        End If
                        r = 1
                    Case Else
                        r = 0
                End Select
            End If
            If r = 0 Then
                ' 処理されなかった場合は次のフックを呼ぶ
                ' CallNextHookExに渡すフックハンドルは9xでのみ有効なのでNT系では0でもかまわない
                r = CallNextHookEx(0, nCode, wParam, lParam)
            End If
            Return r
        End If

        Return CallNextHookEx(hHook, nCode, wParam, lParam)
    End Function

    Public Event KeyDown(ByVal sender As Object, ByVal EventArgs As KeyBoardHookerEventArgs)
    Public Event KeyUp(ByVal sender As Object, ByVal EventArgs As KeyBoardHookerEventArgs)

    Public Sub Dispose()
        Dim ret As Boolean = UnhookWindowsHookEx(hHook)
        If ret.Equals(False) Then
        End If
        ' フック解除したことを通知
        ' メッセージループが回らないと解放されないため
        PostMessage(HWND_BROADCAST, WM_NULL, IntPtr.Zero, IntPtr.Zero)
    End Sub
End Class


Public Class KeyBoardHookerEventArgs
    Inherits EventArgs

    Dim _vkCode As Integer

    Public Property vkCode() As Integer
        Get
            Return _vkCode
        End Get
        Set(ByVal value As Integer)
            _vkCode = value
        End Set
    End Property

End Class
