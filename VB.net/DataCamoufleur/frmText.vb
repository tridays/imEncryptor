Public Class frmText
    Dim BMP As cBitmap

    Private Function Encrypt(ByRef B() As Byte) As Byte()
        Dim i As Integer, t() As Byte, a As Integer = B.Length
        ReDim t(a - 1)
        GetKey(txtPW.Text)
        Rnd(-mKey)

        Dim Start As Date = Now
        Jdt.Show()
        Jdt.Tag = "Text"
        Jdt.Select()
        Me.Enabled = False
        With Jdt
            .Text = "imEncryptor - 正在运行"
            .lblA.Text = "程序正在执行您所要求的运算"
            .OK_Button.Text = "请稍候"
            .OK_Button.Enabled = False
            .Cancel_Button.Enabled = True
            .Jdt1.Value = 0
            .Jdt1.Maximum = a
        End With

        For i = 0 To a - 1
            t(i) = B(i) Xor CByte(Int(Rnd() * 256))
            If GetInputState Or (i Mod 512000 = 0) Then
                Jdt.Jdt1.Value = i
                Jdt.lblP.Text = Format(i / a, "00.00 %")
                My.Application.DoEvents()
            End If
        Next

        With Jdt
            .Text = "imEncryptor - 完成"
            .lblA.Text = "您所要求的命令已经成功执行。"
            .OK_Button.Text = "关闭"
            .OK_Button.Enabled = True
            .Cancel_Button.Enabled = False
            .Jdt1.Value = .Jdt1.Maximum
            .lblP.Text = ("耗时：" & Now.Subtract(Start).ToString).Substring(0, 15)
            .Select()
        End With

        Return t
    End Function

    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text += Application.ProductVersion

        cmbC.SelectedIndex = 0
        cmbP.SelectedIndex = 0

        chkPw.Checked = My.Settings.Item("ShowChar")

        If My.Settings.Item("TextFirst") = False Then
            frmFile.Show()
            Me.Close()
            Exit Sub
        End If

        If inIDE() = True Then
            Me.Show()
            Dim i As Integer, s1 As String = Nothing, s2 As String = My.Resources.test
            For i = 1 To 10
                s1 += s2
            Next
            rtxP.Text = s1

            cmbP.SelectedIndex = 2
            'btnPjm_Click(sender, e)
        End If
    End Sub

    Private Sub btnPdk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPdk.Click
        If DialogOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Debug.Print(IO.File.ReadAllText(DialogOpen.FileName)) 这跟下面那个完全不是一个编码的orz
            rtxP.LoadFile(DialogOpen.FileName, RichTextBoxStreamType.PlainText)
        End If
    End Sub
    Private Sub btnPlc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlc.Click
        If Len(rtxC.Text) Then
            DialogSave.FileName = DialogOpen.SafeFileName
            If DialogSave.ShowDialog = Windows.Forms.DialogResult.OK Then
                rtxP.SaveFile(DialogSave.FileName, RichTextBoxStreamType.PlainText)
            End If
        End If
    End Sub
    Private Sub btnCdk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCdk.Click
        If cmbC.SelectedIndex = 2 Then
            If DialogOpenB.ShowDialog() = Windows.Forms.DialogResult.OK Then
                BMP = New cBitmap
                BMP.LoadFile(DialogOpenB.FileName)
                Pic.Image = BMP.Img
                Pic.Refresh()
            End If
        ElseIf DialogOpen2.ShowDialog() = Windows.Forms.DialogResult.OK Then
            rtxC.LoadFile(DialogOpen2.FileName, RichTextBoxStreamType.PlainText)
        End If
    End Sub
    Private Sub btnClc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClc.Click
        If cmbC.SelectedIndex = 2 Then
            Try
                Dim t As Type
                t = BMP.GetType()
            Catch
                MsgBox("没有可以另存的位图信息，请加密或者打开一张加密过的位图。", 48)
                Exit Sub
            End Try
            DialogSaveB.FileName = DialogOpenB.SafeFileName
            If DialogSaveB.ShowDialog = Windows.Forms.DialogResult.OK Then
                BMP.SaveFile(DialogSaveB.FileName)
            End If
        ElseIf Len(rtxC.Text) Then
            DialogSave2.FileName = DialogOpen2.SafeFileName
            If DialogSave2.ShowDialog = Windows.Forms.DialogResult.OK Then
                rtxP.SaveFile(DialogSave2.FileName, RichTextBoxStreamType.PlainText)
            End If
        End If
    End Sub

    Private Sub btnChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChange.Click
        If cmbC.SelectedIndex = 2 Then Exit Sub

        Dim t As String
        t = rtxP.Text
        rtxP.Text = rtxC.Text
        rtxC.Text = t
    End Sub
    Private Sub btnPqk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPqk.Click
        If Len(rtxP.Text) Then
            If MsgBox("确认清空 明文框 的所有内容吗？", 36) = MsgBoxResult.Yes Then rtxP.Text = Nothing
        End If
    End Sub
    Private Sub btnCqk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCqk.Click
        If cmbC.SelectedIndex = 2 Then
            Try
                Dim t As Type
                t = BMP.GetType
                If MsgBox("确认清空已加载的 位图 吗？", 36) = MsgBoxResult.Yes Then
                    Pic.Image = Nothing
                    Pic.Refresh()
                    BMP = Nothing
                End If
            Catch
            End Try
        ElseIf Len(rtxC.Text) Then
            If MsgBox("确认清空 密文框 的所有内容吗？", 36) = MsgBoxResult.Yes Then rtxC.Text = Nothing
        End If
    End Sub

    Private Sub btnPjm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPjm.Click
        Dim B() As Byte, s As String
        B = StrToBin(rtxP.Text)
        If B.Length = 0 Then Exit Sub '明文为空
        Try
            B = Encrypt(B)
        Catch
            Exit Sub '手动关闭
        End Try
        If cmbP.SelectedIndex = 2 Then
            BMP = New cBitmap
            BMP.FromArray(B)
            Pic.Image = BMP.Img
            Pic.Refresh()
        Else
            If cmbP.SelectedIndex = 0 Then
                s = BinToB64(B)
            Else
                s = BinToHex(B)
            End If
            rtxC.Text = s
        End If
        cmbC.SelectedIndex = cmbP.SelectedIndex
    End Sub
    Private Sub btnCjm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCjm.Click
        Dim B() As Byte
        If cmbC.SelectedIndex = 2 Then
            'Try
            B = BMP.ToArray
            'Catch
            '    MsgBox("没有可以解密的位图信息，请加密或者打开一张加密过的位图。", 48)
            '    Exit Sub
            'End Try
        Else
            Dim s As String
            s = rtxC.Text
            If Len(s) = 0 Then Exit Sub '密文为空
            Try
                If cmbC.SelectedIndex = 0 Then
                    B = B64ToBin(s)
                Else
                    B = HexToBin(s)
                End If
            Catch
                MsgBox("密文格式错误，数据丢失或被篡改！", 16)
                Exit Sub '格式错误
            End Try
        End If
        Try
            B = Encrypt(B)
        Catch
            Exit Sub '手动关闭
        End Try
        rtxP.Text = BinToStr(B)
    End Sub

    Private Sub chkPw_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPw.CheckedChanged
        txtPW.PasswordChar = IIf(chkPw.Checked, "*", Nothing)
        My.Settings.Item("ShowChar") = chkPw.Checked
    End Sub

    Private Sub cmbC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbC.SelectedIndexChanged
        Pic.Visible = (cmbC.SelectedIndex = 2)
    End Sub

    Private Sub lnkToFile_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkToFile.LinkClicked
        My.Settings.Item("TextFirst") = False
        frmFile.Show()
        Me.Close()
    End Sub
End Class
