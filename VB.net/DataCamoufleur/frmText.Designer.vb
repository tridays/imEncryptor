<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmText
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.rtxP = New System.Windows.Forms.RichTextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnPdk = New System.Windows.Forms.Button()
        Me.btnPlc = New System.Windows.Forms.Button()
        Me.btnPqk = New System.Windows.Forms.Button()
        Me.btnChange = New System.Windows.Forms.Button()
        Me.cmbP = New System.Windows.Forms.ComboBox()
        Me.btnPjm = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnCjm = New System.Windows.Forms.Button()
        Me.cmbC = New System.Windows.Forms.ComboBox()
        Me.btnCqk = New System.Windows.Forms.Button()
        Me.btnClc = New System.Windows.Forms.Button()
        Me.btnCdk = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.rtxC = New System.Windows.Forms.RichTextBox()
        Me.lnkToFile = New System.Windows.Forms.LinkLabel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkPw = New System.Windows.Forms.CheckBox()
        Me.txtPW = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DialogOpen = New System.Windows.Forms.OpenFileDialog()
        Me.DialogSave = New System.Windows.Forms.SaveFileDialog()
        Me.DialogOpenB = New System.Windows.Forms.OpenFileDialog()
        Me.DialogSaveB = New System.Windows.Forms.SaveFileDialog()
        Me.DialogSave2 = New System.Windows.Forms.SaveFileDialog()
        Me.DialogOpen2 = New System.Windows.Forms.OpenFileDialog()
        Me.Pic = New System.Windows.Forms.PictureBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.Pic, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'rtxP
        '
        Me.rtxP.AcceptsTab = True
        Me.rtxP.Location = New System.Drawing.Point(12, 38)
        Me.rtxP.Name = "rtxP"
        Me.rtxP.Size = New System.Drawing.Size(276, 305)
        Me.rtxP.TabIndex = 0
        Me.rtxP.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(12, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(132, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "明文 PlainText"
        '
        'btnPdk
        '
        Me.btnPdk.Location = New System.Drawing.Point(12, 349)
        Me.btnPdk.Name = "btnPdk"
        Me.btnPdk.Size = New System.Drawing.Size(48, 35)
        Me.btnPdk.TabIndex = 1
        Me.btnPdk.Text = "打开"
        Me.btnPdk.UseVisualStyleBackColor = True
        '
        'btnPlc
        '
        Me.btnPlc.Location = New System.Drawing.Point(66, 349)
        Me.btnPlc.Name = "btnPlc"
        Me.btnPlc.Size = New System.Drawing.Size(48, 35)
        Me.btnPlc.TabIndex = 2
        Me.btnPlc.Text = "另存"
        Me.btnPlc.UseVisualStyleBackColor = True
        '
        'btnPqk
        '
        Me.btnPqk.Location = New System.Drawing.Point(120, 349)
        Me.btnPqk.Name = "btnPqk"
        Me.btnPqk.Size = New System.Drawing.Size(48, 35)
        Me.btnPqk.TabIndex = 3
        Me.btnPqk.Text = "清空"
        Me.btnPqk.UseVisualStyleBackColor = True
        '
        'btnChange
        '
        Me.btnChange.Font = New System.Drawing.Font("宋体", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btnChange.Location = New System.Drawing.Point(294, 38)
        Me.btnChange.Name = "btnChange"
        Me.btnChange.Size = New System.Drawing.Size(32, 305)
        Me.btnChange.TabIndex = 6
        Me.btnChange.Text = "↔"
        Me.btnChange.UseVisualStyleBackColor = True
        '
        'cmbP
        '
        Me.cmbP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbP.FormattingEnabled = True
        Me.cmbP.Items.AddRange(New Object() {"输出为BASE64编码", "输出为16进制文本", "生成为加密的位图"})
        Me.cmbP.Location = New System.Drawing.Point(149, 10)
        Me.cmbP.Name = "cmbP"
        Me.cmbP.Size = New System.Drawing.Size(139, 20)
        Me.cmbP.TabIndex = 4
        '
        'btnPjm
        '
        Me.btnPjm.Location = New System.Drawing.Point(174, 349)
        Me.btnPjm.Name = "btnPjm"
        Me.btnPjm.Size = New System.Drawing.Size(114, 35)
        Me.btnPjm.TabIndex = 5
        Me.btnPjm.Text = "单击此处加密"
        Me.btnPjm.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(294, 349)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(32, 35)
        Me.Button1.TabIndex = 9
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnCjm
        '
        Me.btnCjm.Location = New System.Drawing.Point(494, 349)
        Me.btnCjm.Name = "btnCjm"
        Me.btnCjm.Size = New System.Drawing.Size(114, 35)
        Me.btnCjm.TabIndex = 12
        Me.btnCjm.Text = "单击此处解密"
        Me.btnCjm.UseVisualStyleBackColor = True
        '
        'cmbC
        '
        Me.cmbC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbC.FormattingEnabled = True
        Me.cmbC.Items.AddRange(New Object() {"来自于BASE64编码", "来自于16进制文本", "来自加密过的位图"})
        Me.cmbC.Location = New System.Drawing.Point(479, 10)
        Me.cmbC.Name = "cmbC"
        Me.cmbC.Size = New System.Drawing.Size(129, 20)
        Me.cmbC.TabIndex = 11
        '
        'btnCqk
        '
        Me.btnCqk.Location = New System.Drawing.Point(440, 349)
        Me.btnCqk.Name = "btnCqk"
        Me.btnCqk.Size = New System.Drawing.Size(48, 35)
        Me.btnCqk.TabIndex = 10
        Me.btnCqk.Text = "清空"
        Me.btnCqk.UseVisualStyleBackColor = True
        '
        'btnClc
        '
        Me.btnClc.Location = New System.Drawing.Point(386, 349)
        Me.btnClc.Name = "btnClc"
        Me.btnClc.Size = New System.Drawing.Size(48, 35)
        Me.btnClc.TabIndex = 9
        Me.btnClc.Text = "另存"
        Me.btnClc.UseVisualStyleBackColor = True
        '
        'btnCdk
        '
        Me.btnCdk.Location = New System.Drawing.Point(332, 349)
        Me.btnCdk.Name = "btnCdk"
        Me.btnCdk.Size = New System.Drawing.Size(48, 35)
        Me.btnCdk.TabIndex = 8
        Me.btnCdk.Text = "打开"
        Me.btnCdk.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(332, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(141, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "密文 CipherText"
        '
        'rtxC
        '
        Me.rtxC.AcceptsTab = True
        Me.rtxC.Location = New System.Drawing.Point(332, 38)
        Me.rtxC.Name = "rtxC"
        Me.rtxC.Size = New System.Drawing.Size(276, 305)
        Me.rtxC.TabIndex = 7
        Me.rtxC.Text = ""
        '
        'lnkToFile
        '
        Me.lnkToFile.AutoSize = True
        Me.lnkToFile.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.lnkToFile.Location = New System.Drawing.Point(294, 466)
        Me.lnkToFile.Name = "lnkToFile"
        Me.lnkToFile.Size = New System.Drawing.Size(312, 21)
        Me.lnkToFile.TabIndex = 15
        Me.lnkToFile.TabStop = True
        Me.lnkToFile.Text = "单击这里可以切换到""文件加解密""的界面！"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.chkPw)
        Me.GroupBox1.Controls.Add(Me.txtPW)
        Me.GroupBox1.ForeColor = System.Drawing.Color.Green
        Me.GroupBox1.Location = New System.Drawing.Point(13, 390)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(275, 97)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "密钥 Password"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Gray
        Me.Label3.Location = New System.Drawing.Point(12, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(245, 12)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "注意：即使密钥为空，程序仍然可以正常运行"
        '
        'chkPw
        '
        Me.chkPw.AutoSize = True
        Me.chkPw.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPw.Location = New System.Drawing.Point(14, 53)
        Me.chkPw.Name = "chkPw"
        Me.chkPw.Size = New System.Drawing.Size(180, 16)
        Me.chkPw.TabIndex = 14
        Me.chkPw.Text = "使用迷惑性字符代替文本显示"
        Me.chkPw.UseVisualStyleBackColor = True
        '
        'txtPW
        '
        Me.txtPW.BackColor = System.Drawing.Color.FromArgb(CType(CType(220, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtPW.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txtPW.Location = New System.Drawing.Point(14, 22)
        Me.txtPW.Name = "txtPW"
        Me.txtPW.Size = New System.Drawing.Size(248, 26)
        Me.txtPW.TabIndex = 13
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Black
        Me.GroupBox2.Location = New System.Drawing.Point(298, 391)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(309, 75)
        Me.GroupBox2.TabIndex = 19
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "关于 本程序 的一些说明"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Gray
        Me.Label4.Location = New System.Drawing.Point(10, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(293, 48)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "　　本加密程序使用密钥为种子生成的伪随机数实现，" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "同时有创意地制作了生成位图的功能，相信它将非常有" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "趣。我本来是在 VB6.0下完成的它，但不久前将它移植" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & _
    "到了VB.net下，话不多说，Enjoy it."
        '
        'DialogOpen
        '
        Me.DialogOpen.DefaultExt = "txt"
        Me.DialogOpen.Filter = "文本文档|*.txt|所有文件|*"
        '
        'DialogSave
        '
        Me.DialogSave.Filter = "文本文档|*.txt|所有文件|*"
        '
        'DialogOpenB
        '
        Me.DialogOpenB.Filter = "位图文件|*.bmp|所有文件|*"
        '
        'DialogSaveB
        '
        Me.DialogSaveB.Filter = "位图文件|*.bmp"
        '
        'DialogSave2
        '
        Me.DialogSave2.Filter = "文本文档|*.txt|所有文件|*"
        '
        'DialogOpen2
        '
        Me.DialogOpen2.DefaultExt = "txt"
        Me.DialogOpen2.Filter = "文本文档|*.txt|所有文件|*"
        '
        'Pic
        '
        Me.Pic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Pic.Location = New System.Drawing.Point(332, 38)
        Me.Pic.Name = "Pic"
        Me.Pic.Size = New System.Drawing.Size(276, 305)
        Me.Pic.TabIndex = 20
        Me.Pic.TabStop = False
        Me.Pic.Visible = False
        '
        'frmText
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(618, 499)
        Me.Controls.Add(Me.Pic)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lnkToFile)
        Me.Controls.Add(Me.btnCjm)
        Me.Controls.Add(Me.cmbC)
        Me.Controls.Add(Me.btnCqk)
        Me.Controls.Add(Me.btnClc)
        Me.Controls.Add(Me.btnCdk)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.rtxC)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnPjm)
        Me.Controls.Add(Me.cmbP)
        Me.Controls.Add(Me.btnChange)
        Me.Controls.Add(Me.btnPqk)
        Me.Controls.Add(Me.btnPlc)
        Me.Controls.Add(Me.btnPdk)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.rtxP)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmText"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Camoufleur ver."
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.Pic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rtxP As System.Windows.Forms.RichTextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnPdk As System.Windows.Forms.Button
    Friend WithEvents btnPlc As System.Windows.Forms.Button
    Friend WithEvents btnPqk As System.Windows.Forms.Button
    Friend WithEvents btnChange As System.Windows.Forms.Button
    Friend WithEvents cmbP As System.Windows.Forms.ComboBox
    Friend WithEvents btnPjm As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnCjm As System.Windows.Forms.Button
    Friend WithEvents cmbC As System.Windows.Forms.ComboBox
    Friend WithEvents btnCqk As System.Windows.Forms.Button
    Friend WithEvents btnClc As System.Windows.Forms.Button
    Friend WithEvents btnCdk As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents rtxC As System.Windows.Forms.RichTextBox
    Friend WithEvents lnkToFile As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPW As System.Windows.Forms.TextBox
    Friend WithEvents chkPw As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DialogOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents DialogSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents DialogOpenB As System.Windows.Forms.OpenFileDialog
    Friend WithEvents DialogSaveB As System.Windows.Forms.SaveFileDialog
    Friend WithEvents DialogSave2 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents DialogOpen2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Pic As System.Windows.Forms.PictureBox

End Class
