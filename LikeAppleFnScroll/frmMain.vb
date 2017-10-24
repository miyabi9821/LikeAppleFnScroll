Imports System.ComponentModel
Imports Microsoft.Win32

Public Class frmMain
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Keys) As Short
    End Function

    Public Sub New()
        Me.InitializeComponent()
    End Sub

    WithEvents KeyboardHooker1 As New KeyboardHooker
    Sub KeybordHooker1_KeyDown(sender As Object, e As KeyBoardHookerEventArgs) Handles KeyboardHooker1.KeyDown
        'putLog("キーコード：" & CStr(e.vkCode)) ' forDebug
        If StatusToolStripMenuItem.Checked = True Then
            ' 右Ctrlキーのチェックはフック処理で既に行っているのでここではしない
            Select Case e.vkCode
                Case Keys.Up
                    putLog("Ctrl+Up=>PgUp")
                    SendKeys.Send("{PGUP}")
                Case Keys.Down
                    putLog("Ctrl+Down=>PgDown")
                    SendKeys.Send("{PGDN}")
                Case Keys.Left
                    putLog("Ctrl+Left=>Home")
                    SendKeys.Send("{HOME}")
                Case Keys.Right
                    putLog("Ctrl+Right=>End")
                    SendKeys.Send("{END}")
            End Select
        End If

    End Sub

    Private Sub putLog(msg As String)
        txtLog.Text = "[" & System.DateTime.Now & "] " & msg & vbCrLf & txtLog.Text
    End Sub

    '閉じるボタンでは終了しないので、終了メニューから終了させる
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    '有効／無効の切り替え
    Private Sub StatusToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusToolStripMenuItem.Click
        StatusToolStripMenuItem.Checked = Not StatusToolStripMenuItem.Checked
    End Sub

    'ウィンドウとタスクバーに表示
    Private Sub WindowShowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WindowShowToolStripMenuItem.Click
        Me.WindowState = FormWindowState.Normal
        Me.ShowInTaskbar = True
        Me.Visible = True
    End Sub

    '閉じるボタンでウィンドウやタスクバーから消す
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.ShowInTaskbar = False
        Me.Visible = False
        e.Cancel = True
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        'セッション終了通知登録
        AddHandler SystemEvents.SessionEnding, AddressOf SystemEvents_SessionEnding

        Me.ShowInTaskbar = False
        Me.Visible = False
        Me.Location = New Point(0, 0) '最小化で起動するとウィンドウ位置が変になるので強制変更
    End Sub

    Private Sub frmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        'セッション終了通知解除
        RemoveHandler SystemEvents.SessionEnding, AddressOf SystemEvents_SessionEnding
    End Sub

    'Windows終了やログオフ時はアプリケーションを終了する
    Private Sub SystemEvents_SessionEnding(ByVal sender As Object, ByVal e As SessionEndingEventArgs)
        Application.Exit()
    End Sub

    'ダブルクリックでウィンドウの表示／非表示を変更
    Private Sub NotifyIcon_DoubleClick(sender As Object, e As EventArgs) Handles NotifyIcon.DoubleClick
        If Me.Visible = False Then
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            Me.Visible = True
        ElseIf Me.Visible = True Then
            Me.ShowInTaskbar = False
            Me.Visible = False
        End If
    End Sub
End Class
