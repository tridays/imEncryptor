Imports System.Windows.Forms

Public Class Jdt

    Public Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Jdt_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If Me.Tag = "Text" Then
            frmText.Enabled = True
            frmText.Select()
        Else
            frmFile.Enabled = True
            frmFile.Select()
        End If
    End Sub

    Private Sub Jdt_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If OK_Button.Enabled = False Then
            If MsgBox("现在关闭将无任何输出，确认吗？", 36) = MsgBoxResult.No Then
                e.Cancel = True
            Else
                Jdt1.Maximum = 0 '引发错误
            End If
        End If
    End Sub
End Class