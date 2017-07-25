Imports System.ComponentModel

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

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub StatusToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusToolStripMenuItem.Click
        StatusToolStripMenuItem.Checked = Not StatusToolStripMenuItem.Checked
    End Sub

    Private Sub WindowShowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WindowShowToolStripMenuItem.Click
        Me.ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
        e.Cancel = True
    End Sub
End Class
