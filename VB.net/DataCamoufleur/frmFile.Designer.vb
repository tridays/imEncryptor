<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFile
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tabEn = New System.Windows.Forms.TabPage()
        Me.cmbEn = New System.Windows.Forms.ListBox()
        Me.chkEP = New System.Windows.Forms.CheckBox()
        Me.btnET = New System.Windows.Forms.Button()
        Me.txtET = New System.Windows.Forms.TextBox()
        Me.btnEF = New System.Windows.Forms.Button()
        Me.btnEn = New System.Windows.Forms.Button()
        Me.txtEP = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtEF = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tabDe = New System.Windows.Forms.TabPage()
        Me.cmbDe = New System.Windows.Forms.ListBox()
        Me.chkDP = New System.Windows.Forms.CheckBox()
        Me.btnDT = New System.Windows.Forms.Button()
        Me.txtDT = New System.Windows.Forms.TextBox()
        Me.btnDF = New System.Windows.Forms.Button()
        Me.btnDe = New System.Windows.Forms.Button()
        Me.txtDP = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtDF = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lnkToText = New System.Windows.Forms.LinkLabel()
        Me.DialogOpenE = New System.Windows.Forms.OpenFileDialog()
        Me.DialogSaveE = New System.Windows.Forms.SaveFileDialog()
        Me.DialogSaveD = New System.Windows.Forms.SaveFileDialog()
        Me.DialogOpenD = New System.Windows.Forms.OpenFileDialog()
        Me.TabControl1.SuspendLayout()
        Me.tabEn.SuspendLayout()
        Me.tabDe.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tabEn)
        Me.TabControl1.Controls.Add(Me.tabDe)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(419, 309)
        Me.TabControl1.TabIndex = 8
        '
        'tabEn
        '
        Me.tabEn.Controls.Add(Me.cmbEn)
        Me.tabEn.Controls.Add(Me.chkEP)
        Me.tabEn.Controls.Add(Me.btnET)
        Me.tabEn.Controls.Add(Me.txtET)
        Me.tabEn.Controls.Add(Me.btnEF)
        Me.tabEn.Controls.Add(Me.btnEn)
        Me.tabEn.Controls.Add(Me.txtEP)
        Me.tabEn.Controls.Add(Me.Label3)
        Me.tabEn.Controls.Add(Me.Label2)
        Me.tabEn.Controls.Add(Me.txtEF)
        Me.tabEn.Controls.Add(Me.Label1)
        Me.tabEn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.tabEn.Location = New System.Drawing.Point(4, 22)
        Me.tabEn.Name = "tabEn"
        Me.tabEn.Padding = New System.Windows.Forms.Padding(3)
        Me.tabEn.Size = New System.Drawing.Size(411, 283)
        Me.tabEn.TabIndex = 0
        Me.tabEn.Text = "加密 Encrypt"
        Me.tabEn.UseVisualStyleBackColor = True
        '
        'cmbEn
        '
        Me.cmbEn.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.cmbEn.FormattingEnabled = True
        Me.cmbEn.ItemHeight = 14
        Me.cmbEn.Items.AddRange(New Object() {"保存为BASE64编码文本", "保存为16进制文本文件", "保存为加密的位图文件", "仅加密保存，不转格式"})
        Me.cmbEn.Location = New System.Drawing.Point(18, 206)
        Me.cmbEn.Name = "cmbEn"
        Me.cmbEn.Size = New System.Drawing.Size(151, 60)
        Me.cmbEn.TabIndex = 20
        '
        'chkEP
        '
        Me.chkEP.AutoSize = True
        Me.chkEP.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.chkEP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEP.Location = New System.Drawing.Point(170, 137)
        Me.chkEP.Name = "chkEP"
        Me.chkEP.Size = New System.Drawing.Size(235, 20)
        Me.chkEP.TabIndex = 19
        Me.chkEP.Text = "使用迷惑性字符代替文本显示"
        Me.chkEP.UseVisualStyleBackColor = True
        '
        'btnET
        '
        Me.btnET.Location = New System.Drawing.Point(354, 101)
        Me.btnET.Name = "btnET"
        Me.btnET.Size = New System.Drawing.Size(45, 26)
        Me.btnET.TabIndex = 18
        Me.btnET.Text = "浏览"
        Me.btnET.UseVisualStyleBackColor = True
        '
        'txtET
        '
        Me.txtET.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtET.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
        Me.txtET.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtET.Location = New System.Drawing.Point(9, 101)
        Me.txtET.Name = "txtET"
        Me.txtET.Size = New System.Drawing.Size(343, 26)
        Me.txtET.TabIndex = 17
        '
        'btnEF
        '
        Me.btnEF.Location = New System.Drawing.Point(354, 42)
        Me.btnEF.Name = "btnEF"
        Me.btnEF.Size = New System.Drawing.Size(45, 26)
        Me.btnEF.TabIndex = 16
        Me.btnEF.Text = "浏览"
        Me.btnEF.UseVisualStyleBackColor = True
        '
        'btnEn
        '
        Me.btnEn.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btnEn.Location = New System.Drawing.Point(192, 215)
        Me.btnEn.Name = "btnEn"
        Me.btnEn.Size = New System.Drawing.Size(193, 45)
        Me.btnEn.TabIndex = 15
        Me.btnEn.Text = "单击此处加密"
        Me.btnEn.UseVisualStyleBackColor = True
        '
        'txtEP
        '
        Me.txtEP.BackColor = System.Drawing.Color.FromArgb(CType(CType(220, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtEP.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txtEP.Location = New System.Drawing.Point(9, 160)
        Me.txtEP.Name = "txtEP"
        Me.txtEP.Size = New System.Drawing.Size(390, 26)
        Me.txtEP.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("微软雅黑", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Green
        Me.Label3.Location = New System.Drawing.Point(13, 130)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(149, 27)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "密钥 Password"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(13, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(185, 27)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "保存到 Destination"
        '
        'txtEF
        '
        Me.txtEF.AllowDrop = True
        Me.txtEF.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.txtEF.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
        Me.txtEF.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEF.Location = New System.Drawing.Point(9, 42)
        Me.txtEF.Name = "txtEF"
        Me.txtEF.Size = New System.Drawing.Size(343, 26)
        Me.txtEF.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(13, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(144, 27)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "来源于 Source"
        '
        'tabDe
        '
        Me.tabDe.Controls.Add(Me.cmbDe)
        Me.tabDe.Controls.Add(Me.chkDP)
        Me.tabDe.Controls.Add(Me.btnDT)
        Me.tabDe.Controls.Add(Me.txtDT)
        Me.tabDe.Controls.Add(Me.btnDF)
        Me.tabDe.Controls.Add(Me.btnDe)
        Me.tabDe.Controls.Add(Me.txtDP)
        Me.tabDe.Controls.Add(Me.Label4)
        Me.tabDe.Controls.Add(Me.Label5)
        Me.tabDe.Controls.Add(Me.txtDF)
        Me.tabDe.Controls.Add(Me.Label6)
        Me.tabDe.ForeColor = System.Drawing.SystemColors.ControlText
        Me.tabDe.Location = New System.Drawing.Point(4, 22)
        Me.tabDe.Name = "tabDe"
        Me.tabDe.Padding = New System.Windows.Forms.Padding(3)
        Me.tabDe.Size = New System.Drawing.Size(411, 283)
        Me.tabDe.TabIndex = 1
        Me.tabDe.Text = "解密 Decrypt"
        Me.tabDe.UseVisualStyleBackColor = True
        '
        'cmbDe
        '
        Me.cmbDe.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.cmbDe.FormattingEnabled = True
        Me.cmbDe.ItemHeight = 14
        Me.cmbDe.Items.AddRange(New Object() {"来自加密的BASE64编码", "来自加密的16进制文本", "来自已生成的位图文件", "来自加密的二进制文件"})
        Me.cmbDe.Location = New System.Drawing.Point(18, 206)
        Me.cmbDe.Name = "cmbDe"
        Me.cmbDe.Size = New System.Drawing.Size(151, 60)
        Me.cmbDe.TabIndex = 31
        '
        'chkDP
        '
        Me.chkDP.AutoSize = True
        Me.chkDP.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.chkDP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDP.Location = New System.Drawing.Point(170, 137)
        Me.chkDP.Name = "chkDP"
        Me.chkDP.Size = New System.Drawing.Size(235, 20)
        Me.chkDP.TabIndex = 30
        Me.chkDP.Text = "使用迷惑性字符代替文本显示"
        Me.chkDP.UseVisualStyleBackColor = True
        '
        'btnDT
        '
        Me.btnDT.Location = New System.Drawing.Point(354, 101)
        Me.btnDT.Name = "btnDT"
        Me.btnDT.Size = New System.Drawing.Size(45, 26)
        Me.btnDT.TabIndex = 29
        Me.btnDT.Text = "浏览"
        Me.btnDT.UseVisualStyleBackColor = True
        '
        'txtDT
        '
        Me.txtDT.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtDT.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
        Me.txtDT.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDT.Location = New System.Drawing.Point(9, 101)
        Me.txtDT.Name = "txtDT"
        Me.txtDT.Size = New System.Drawing.Size(343, 26)
        Me.txtDT.TabIndex = 28
        '
        'btnDF
        '
        Me.btnDF.Location = New System.Drawing.Point(354, 42)
        Me.btnDF.Name = "btnDF"
        Me.btnDF.Size = New System.Drawing.Size(45, 26)
        Me.btnDF.TabIndex = 27
        Me.btnDF.Text = "浏览"
        Me.btnDF.UseVisualStyleBackColor = True
        '
        'btnDe
        '
        Me.btnDe.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btnDe.Location = New System.Drawing.Point(192, 215)
        Me.btnDe.Name = "btnDe"
        Me.btnDe.Size = New System.Drawing.Size(193, 45)
        Me.btnDe.TabIndex = 26
        Me.btnDe.Text = "单击此处解密"
        Me.btnDe.UseVisualStyleBackColor = True
        '
        'txtDP
        '
        Me.txtDP.BackColor = System.Drawing.Color.FromArgb(CType(CType(220, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtDP.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txtDP.Location = New System.Drawing.Point(9, 160)
        Me.txtDP.Name = "txtDP"
        Me.txtDP.Size = New System.Drawing.Size(390, 26)
        Me.txtDP.TabIndex = 25
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("微软雅黑", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Green
        Me.Label4.Location = New System.Drawing.Point(13, 130)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(149, 27)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "密钥 Password"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("微软雅黑", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(13, 71)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(185, 27)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "保存到 Destination"
        '
        'txtDF
        '
        Me.txtDF.AllowDrop = True
        Me.txtDF.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.txtDF.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
        Me.txtDF.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDF.Location = New System.Drawing.Point(9, 42)
        Me.txtDF.Name = "txtDF"
        Me.txtDF.Size = New System.Drawing.Size(343, 26)
        Me.txtDF.TabIndex = 22
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("微软雅黑", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(13, 12)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(144, 27)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "来源于 Source"
        '
        'lnkToText
        '
        Me.lnkToText.AutoSize = True
        Me.lnkToText.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.lnkToText.Location = New System.Drawing.Point(21, 333)
        Me.lnkToText.Name = "lnkToText"
        Me.lnkToText.Size = New System.Drawing.Size(312, 21)
        Me.lnkToText.TabIndex = 18
        Me.lnkToText.TabStop = True
        Me.lnkToText.Text = "单击这里可以切换到""文本加解密""的界面！"
        '
        'DialogOpenE
        '
        Me.DialogOpenE.Filter = "所有文件|*"
        '
        'DialogSaveD
        '
        Me.DialogSaveD.Filter = "所有文件|*"
        '
        'DialogOpenD
        '
        Me.DialogOpenD.Filter = "文本文档|*.txt|位图文件|*.bmp|所有文件|*"
        '
        'frmFile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(444, 373)
        Me.Controls.Add(Me.lnkToText)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmFile"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Camoufleur ver."
        Me.TabControl1.ResumeLayout(False)
        Me.tabEn.ResumeLayout(False)
        Me.tabEn.PerformLayout()
        Me.tabDe.ResumeLayout(False)
        Me.tabDe.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tabEn As System.Windows.Forms.TabPage
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtEF As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tabDe As System.Windows.Forms.TabPage
    Friend WithEvents txtEP As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnEn As System.Windows.Forms.Button
    Friend WithEvents lnkToText As System.Windows.Forms.LinkLabel
    Friend WithEvents btnEF As System.Windows.Forms.Button
    Friend WithEvents btnET As System.Windows.Forms.Button
    Friend WithEvents txtET As System.Windows.Forms.TextBox
    Friend WithEvents chkEP As System.Windows.Forms.CheckBox
    Friend WithEvents DialogOpenE As System.Windows.Forms.OpenFileDialog
    Friend WithEvents DialogSaveE As System.Windows.Forms.SaveFileDialog
    Friend WithEvents DialogSaveD As System.Windows.Forms.SaveFileDialog
    Friend WithEvents DialogOpenD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmbEn As System.Windows.Forms.ListBox
    Friend WithEvents cmbDe As System.Windows.Forms.ListBox
    Friend WithEvents chkDP As System.Windows.Forms.CheckBox
    Friend WithEvents btnDT As System.Windows.Forms.Button
    Friend WithEvents txtDT As System.Windows.Forms.TextBox
    Friend WithEvents btnDF As System.Windows.Forms.Button
    Friend WithEvents btnDe As System.Windows.Forms.Button
    Friend WithEvents txtDP As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDF As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
