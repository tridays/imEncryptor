Public Class frmFile
    Dim BMP As New cBitmap

    Private Function Encrypt(ByRef B() As Byte, ByRef Password As String) As Byte()
        Dim i As Integer, t() As Byte, a As Integer = B.Length
        ReDim t(a - 1)
        GetKey(Password)
        Rnd(-mKey)

        Dim Start As Date = Now
        Jdt.Show()
        Jdt.Tag = "File"
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

    Private Sub frmFile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text += Application.ProductVersion

        cmbEn.SelectedIndex = 0
        cmbDe.SelectedIndex = 0

        chkEP.Checked = My.Settings.Item("ShowChar")
        chkDP.Checked = My.Settings.Item("ShowChar")

    End Sub

    Private Sub chkEP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEP.CheckedChanged
        txtEP.PasswordChar = IIf(chkEP.Checked, "*", Nothing)
        chkDP.Checked = chkEP.Checked
        My.Settings.Item("ShowChar") = chkEP.Checked
    End Sub
    Private Sub chkDP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDP.CheckedChanged
        txtDP.PasswordChar = IIf(chkDP.Checked, "*", Nothing)
        chkEP.Checked = chkDP.Checked
        My.Settings.Item("ShowChar") = chkDP.Checked
    End Sub

    Private Sub lnkToText_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkToText.LinkClicked
        My.Settings.Item("TextFirst") = True
        frmText.Show()
        Me.Close()
    End Sub

    Private Sub btnEF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEF.Click
        If DialogOpenE.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim Path As String = DialogOpenE.FileName
            txtEF.Text = Path
            txtEF.SelectionStart = txtEF.TextLength
            If Len(txtET.Text) = 0 Then
                Select Case cmbEn.SelectedIndex
                    Case 0, 1
                        txtET.Text = txtEF.Text & ".txt"
                    Case 2
                        txtET.Text = txtEF.Text & ".bmp"
                    Case 3
                        txtET.Text = txtEF.Text
                End Select
                txtET.SelectionStart = txtEF.TextLength
            End If
            DialogSaveE.FileName = Path
        End If
    End Sub
    Private Sub btnET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnET.Click
        If cmbEn.SelectedIndex = 2 Then
            DialogSaveE.Filter = "位图文件|*.bmp|所有文件|*"
        Else
            DialogSaveE.Filter = "文本文档|*.txt|所有文件|*"
        End If
        If DialogSaveE.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtET.Text = DialogSaveE.FileName
            txtET.SelectionStart = txtET.TextLength
        End If
    End Sub

    Private Sub btnDF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDF.Click
        DialogOpenD.FilterIndex = IIf(cmbDe.SelectedIndex = 2, 2, 1)
        If DialogOpenD.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim Path As String = DialogOpenD.FileName
            Select Case System.IO.Path.GetExtension(Path).ToUpper
                Case ".TXT"
                    If cmbDe.SelectedIndex = 2 Then
                        MsgBox("当前解密方案（来自位图）可能并不适合本文件，" & vbCrLf & vbCrLf & "请确认后再执行解密操作。", 48)
                        cmbDe.Select()
                    End If
                Case ".BMP"
                    cmbDe.SelectedIndex = 2
            End Select


            txtDF.Text = Path
            txtDF.SelectionStart = txtDF.TextLength
            If Len(txtDT.Text) = 0 Then
                txtDT.Text = txtDF.Text
                txtDT.SelectionStart = txtDF.TextLength
            End If
            DialogSaveD.FileName = Path
        End If
    End Sub
    Private Sub btnDT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDT.Click
        If DialogSaveD.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtDT.Text = DialogSaveD.FileName
            txtDT.SelectionStart = txtDT.TextLength
        End If
    End Sub

    Private Sub btnEn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEn.Click
        If IO.File.Exists(txtEF.Text) = False Then
            MsgBox("指定的文件来源不存在！", 48)
            txtEF.Select()
            Exit Sub
        End If
        Try
            Dim Fs As IO.FileStream
            Fs = IO.File.Create(txtET.Text)
            Fs.Close()
        Catch
            MsgBox("输入的目标地址无法访问！", 16)
            txtET.Select()
            Exit Sub
        End Try
        Dim B() As Byte
        B = IO.File.ReadAllBytes(txtEF.Text)
        If B.Length = 0 Then
            MsgBox("待加密文件为空！", 48)
            Exit Sub
        End If
        B = Encrypt(B, txtEP.Text)
        Select Case cmbEn.SelectedIndex
            Case 2
                BMP.FromArray(B)
                BMP.SaveFile(txtET.Text)
            Case 3
                IO.File.WriteAllBytes(txtET.Text, B)
            Case Else
                Dim s As String
                If cmbEn.SelectedIndex() = 0 Then
                    s = BinToB64(B)
                Else
                    s = BinToHex(B)
                End If
                IO.File.WriteAllText(txtET.Text, s)
        End Select

        'Jdt.OK_Button_Click(sender, e)
        If MsgBox("是否打开已生成文件的目录？", 36) = MsgBoxResult.Yes Then
            Debug.Print(IO.Path.GetDirectoryName(txtET.Text))
            Shell("explorer """ & IO.Path.GetDirectoryName(txtET.Text) & """", AppWinStyle.MaximizedFocus)
        End If
    End Sub

    Private Sub btnDe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDe.Click
        If IO.File.Exists(txtDF.Text) = False Then
            MsgBox("指定的文件来源不存在！", 48)
            txtDF.Select()
            Exit Sub
        End If
        Try
            Dim Fs As IO.FileStream
            Fs = IO.File.Create(txtDT.Text)
            Fs.Close()
        Catch
            MsgBox("输入的目标地址无法访问！", 16)
            txtDT.Select()
            Exit Sub
        End Try
        Dim B() As Byte
        Select Case cmbEn.SelectedIndex
            Case 2
                Try
                    BMP.LoadFile(txtDF.Text)
                    B = BMP.ToArray()
                Catch
                    MsgBox("位图格式错误，数据丢失或被篡改！", 16)
                    Exit Sub '格式错误
                End Try
            Case 3
                B = IO.File.ReadAllBytes(txtDF.Text)
            Case Else
                Dim s As String
                s = IO.File.ReadAllText(txtDF.Text)
                Try
                    If cmbEn.SelectedIndex = 0 Then
                        B = B64ToBin(s)
                    Else
                        B = HexToBin(s)
                    End If
                Catch
                    MsgBox("文件格式错误，数据丢失或被篡改！", 16)
                    Exit Sub '格式错误
                End Try
        End Select
        B = Encrypt(B, txtDP.Text)
        IO.File.WriteAllBytes(txtDT.Text, B)
        'Jdt.OK_Button_Click(sender, e)
        If MsgBox("是否打开已生成文件的目录？", 36) = MsgBoxResult.Yes Then
            Debug.Print(IO.Path.GetDirectoryName(txtDT.Text))
            Shell("explorer """ & IO.Path.GetDirectoryName(txtDT.Text) & """", AppWinStyle.MaximizedFocus)
        End If
    End Sub

    Private Sub txtEF_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtEF.DragDrop
        txtEF.Text = e.Data.GetData(e.Data.GetFormats(0)(0))
    End Sub

    Private Sub cmbDe_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDe.SelectedIndexChanged

    End Sub
End Class