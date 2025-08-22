<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EmployeeBlock
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.EmpGrid1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RegLbx1 = New System.Windows.Forms.Label()
        Me.RegLbx2 = New System.Windows.Forms.Label()
        Me.RegLbx3 = New System.Windows.Forms.Label()
        Me.RegLbx4 = New System.Windows.Forms.Label()
        Me.RegLbx5 = New System.Windows.Forms.Label()
        Me.RegLbx6 = New System.Windows.Forms.Label()
        Me.RegLbx7 = New System.Windows.Forms.Label()
        Me.RegTbx1 = New System.Windows.Forms.TextBox()
        Me.RegTbx2 = New System.Windows.Forms.TextBox()
        Me.RegTbx3 = New System.Windows.Forms.TextBox()
        Me.RegCmb4 = New System.Windows.Forms.ComboBox()
        Me.RegTbx5 = New System.Windows.Forms.TextBox()
        Me.RegCmb1 = New System.Windows.Forms.ComboBox()
        Me.RegCmb2 = New System.Windows.Forms.ComboBox()
        Me.RegCmb3 = New System.Windows.Forms.ComboBox()
        Me.RegCmb5 = New System.Windows.Forms.ComboBox()
        Me.RegLbx8 = New System.Windows.Forms.Label()
        Me.RegLbx9 = New System.Windows.Forms.Label()
        Me.RegLbx10 = New System.Windows.Forms.Label()
        Me.EmpBtn2 = New System.Windows.Forms.Button()
        Me.EmpBtn1 = New System.Windows.Forms.Button()
        Me.MaskRegTbx1 = New System.Windows.Forms.MaskedTextBox()
        Me.EmpBtn3 = New System.Windows.Forms.Button()
        CType(Me.EmpGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'EmpGrid1
        '
        Me.EmpGrid1.AllowUserToAddRows = False
        Me.EmpGrid1.AllowUserToDeleteRows = False
        Me.EmpGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.EmpGrid1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5, Me.Col6, Me.Col7})
        Me.EmpGrid1.Location = New System.Drawing.Point(272, 56)
        Me.EmpGrid1.Name = "EmpGrid1"
        Me.EmpGrid1.ReadOnly = True
        Me.EmpGrid1.Size = New System.Drawing.Size(600, 397)
        Me.EmpGrid1.TabIndex = 0
        '
        'Col1
        '
        Me.Col1.HeaderText = "ID Number"
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        '
        'Col2
        '
        Me.Col2.HeaderText = "Nik"
        Me.Col2.Name = "Col2"
        Me.Col2.ReadOnly = True
        '
        'Col3
        '
        Me.Col3.HeaderText = "Name"
        Me.Col3.Name = "Col3"
        Me.Col3.ReadOnly = True
        '
        'Col4
        '
        Me.Col4.HeaderText = "Active"
        Me.Col4.Name = "Col4"
        Me.Col4.ReadOnly = True
        '
        'Col5
        '
        Me.Col5.HeaderText = "Start Since"
        Me.Col5.Name = "Col5"
        Me.Col5.ReadOnly = True
        '
        'Col6
        '
        Me.Col6.HeaderText = "Pay as"
        Me.Col6.Name = "Col6"
        Me.Col6.ReadOnly = True
        '
        'Col7
        '
        Me.Col7.HeaderText = "Astek"
        Me.Col7.Name = "Col7"
        Me.Col7.ReadOnly = True
        '
        'RegLbx1
        '
        Me.RegLbx1.AutoSize = True
        Me.RegLbx1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx1.Location = New System.Drawing.Point(9, 9)
        Me.RegLbx1.Name = "RegLbx1"
        Me.RegLbx1.Size = New System.Drawing.Size(129, 15)
        Me.RegLbx1.TabIndex = 1
        Me.RegLbx1.Text = "Registration Number: "
        '
        'RegLbx2
        '
        Me.RegLbx2.AutoSize = True
        Me.RegLbx2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx2.Location = New System.Drawing.Point(9, 52)
        Me.RegLbx2.Name = "RegLbx2"
        Me.RegLbx2.Size = New System.Drawing.Size(29, 15)
        Me.RegLbx2.TabIndex = 2
        Me.RegLbx2.Text = "Nik:"
        '
        'RegLbx3
        '
        Me.RegLbx3.AutoSize = True
        Me.RegLbx3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx3.Location = New System.Drawing.Point(9, 94)
        Me.RegLbx3.Name = "RegLbx3"
        Me.RegLbx3.Size = New System.Drawing.Size(82, 15)
        Me.RegLbx3.TabIndex = 3
        Me.RegLbx3.Text = "Name/Nama:"
        '
        'RegLbx4
        '
        Me.RegLbx4.AutoSize = True
        Me.RegLbx4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx4.Location = New System.Drawing.Point(9, 135)
        Me.RegLbx4.Name = "RegLbx4"
        Me.RegLbx4.Size = New System.Drawing.Size(82, 15)
        Me.RegLbx4.TabIndex = 4
        Me.RegLbx4.Text = "Date/Tanggal:"
        '
        'RegLbx5
        '
        Me.RegLbx5.AutoSize = True
        Me.RegLbx5.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx5.Location = New System.Drawing.Point(9, 176)
        Me.RegLbx5.Name = "RegLbx5"
        Me.RegLbx5.Size = New System.Drawing.Size(75, 15)
        Me.RegLbx5.TabIndex = 5
        Me.RegLbx5.Text = "Active/Actif:"
        '
        'RegLbx6
        '
        Me.RegLbx6.AutoSize = True
        Me.RegLbx6.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx6.Location = New System.Drawing.Point(9, 217)
        Me.RegLbx6.Name = "RegLbx6"
        Me.RegLbx6.Size = New System.Drawing.Size(73, 15)
        Me.RegLbx6.TabIndex = 6
        Me.RegLbx6.Text = "Type of Pay:"
        '
        'RegLbx7
        '
        Me.RegLbx7.AutoSize = True
        Me.RegLbx7.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx7.Location = New System.Drawing.Point(9, 258)
        Me.RegLbx7.Name = "RegLbx7"
        Me.RegLbx7.Size = New System.Drawing.Size(67, 15)
        Me.RegLbx7.TabIndex = 7
        Me.RegLbx7.Text = "Jamsostek:"
        '
        'RegTbx1
        '
        Me.RegTbx1.Location = New System.Drawing.Point(12, 29)
        Me.RegTbx1.Name = "RegTbx1"
        Me.RegTbx1.Size = New System.Drawing.Size(176, 20)
        Me.RegTbx1.TabIndex = 8
        '
        'RegTbx2
        '
        Me.RegTbx2.Enabled = False
        Me.RegTbx2.Location = New System.Drawing.Point(12, 71)
        Me.RegTbx2.Name = "RegTbx2"
        Me.RegTbx2.Size = New System.Drawing.Size(176, 20)
        Me.RegTbx2.TabIndex = 9
        '
        'RegTbx3
        '
        Me.RegTbx3.Enabled = False
        Me.RegTbx3.Location = New System.Drawing.Point(12, 112)
        Me.RegTbx3.Name = "RegTbx3"
        Me.RegTbx3.Size = New System.Drawing.Size(176, 20)
        Me.RegTbx3.TabIndex = 10
        '
        'RegCmb4
        '
        Me.RegCmb4.BackColor = System.Drawing.SystemColors.Window
        Me.RegCmb4.FormattingEnabled = True
        Me.RegCmb4.Items.AddRange(New Object() {"Nik", "Name"})
        Me.RegCmb4.Location = New System.Drawing.Point(272, 28)
        Me.RegCmb4.Name = "RegCmb4"
        Me.RegCmb4.Size = New System.Drawing.Size(136, 21)
        Me.RegCmb4.TabIndex = 15
        Me.RegCmb4.Text = "Nik"
        '
        'RegTbx5
        '
        Me.RegTbx5.Location = New System.Drawing.Point(423, 30)
        Me.RegTbx5.Name = "RegTbx5"
        Me.RegTbx5.Size = New System.Drawing.Size(281, 20)
        Me.RegTbx5.TabIndex = 16
        '
        'RegCmb1
        '
        Me.RegCmb1.Enabled = False
        Me.RegCmb1.FormattingEnabled = True
        Me.RegCmb1.Items.AddRange(New Object() {"Yes", "No"})
        Me.RegCmb1.Location = New System.Drawing.Point(12, 194)
        Me.RegCmb1.Name = "RegCmb1"
        Me.RegCmb1.Size = New System.Drawing.Size(136, 21)
        Me.RegCmb1.TabIndex = 17
        '
        'RegCmb2
        '
        Me.RegCmb2.Enabled = False
        Me.RegCmb2.FormattingEnabled = True
        Me.RegCmb2.Items.AddRange(New Object() {"BTN", "CASH"})
        Me.RegCmb2.Location = New System.Drawing.Point(12, 235)
        Me.RegCmb2.Name = "RegCmb2"
        Me.RegCmb2.Size = New System.Drawing.Size(136, 21)
        Me.RegCmb2.TabIndex = 18
        '
        'RegCmb3
        '
        Me.RegCmb3.Enabled = False
        Me.RegCmb3.FormattingEnabled = True
        Me.RegCmb3.Location = New System.Drawing.Point(12, 276)
        Me.RegCmb3.Name = "RegCmb3"
        Me.RegCmb3.Size = New System.Drawing.Size(136, 21)
        Me.RegCmb3.TabIndex = 19
        '
        'RegCmb5
        '
        Me.RegCmb5.FormattingEnabled = True
        Me.RegCmb5.Items.AddRange(New Object() {"Yes", "No"})
        Me.RegCmb5.Location = New System.Drawing.Point(710, 30)
        Me.RegCmb5.Name = "RegCmb5"
        Me.RegCmb5.Size = New System.Drawing.Size(136, 21)
        Me.RegCmb5.TabIndex = 20
        Me.RegCmb5.Text = "Yes"
        '
        'RegLbx8
        '
        Me.RegLbx8.AutoSize = True
        Me.RegLbx8.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx8.Location = New System.Drawing.Point(269, 9)
        Me.RegLbx8.Name = "RegLbx8"
        Me.RegLbx8.Size = New System.Drawing.Size(99, 15)
        Me.RegLbx8.TabIndex = 21
        Me.RegLbx8.Text = "Identification As:"
        '
        'RegLbx9
        '
        Me.RegLbx9.AutoSize = True
        Me.RegLbx9.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx9.Location = New System.Drawing.Point(420, 9)
        Me.RegLbx9.Name = "RegLbx9"
        Me.RegLbx9.Size = New System.Drawing.Size(68, 15)
        Me.RegLbx9.TabIndex = 22
        Me.RegLbx9.Text = "Search For:"
        '
        'RegLbx10
        '
        Me.RegLbx10.AutoSize = True
        Me.RegLbx10.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx10.Location = New System.Drawing.Point(710, 12)
        Me.RegLbx10.Name = "RegLbx10"
        Me.RegLbx10.Size = New System.Drawing.Size(45, 15)
        Me.RegLbx10.TabIndex = 23
        Me.RegLbx10.Text = "Active:"
        '
        'EmpBtn2
        '
        Me.EmpBtn2.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.EmpBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EmpBtn2.Location = New System.Drawing.Point(12, 320)
        Me.EmpBtn2.Name = "EmpBtn2"
        Me.EmpBtn2.Size = New System.Drawing.Size(75, 23)
        Me.EmpBtn2.TabIndex = 25
        Me.EmpBtn2.Text = "&Add"
        Me.EmpBtn2.UseVisualStyleBackColor = True
        '
        'EmpBtn1
        '
        Me.EmpBtn1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.EmpBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EmpBtn1.Location = New System.Drawing.Point(90, 320)
        Me.EmpBtn1.Name = "EmpBtn1"
        Me.EmpBtn1.Size = New System.Drawing.Size(75, 23)
        Me.EmpBtn1.TabIndex = 24
        Me.EmpBtn1.Text = "&Save"
        Me.EmpBtn1.UseVisualStyleBackColor = True
        '
        'MaskRegTbx1
        '
        Me.MaskRegTbx1.Enabled = False
        Me.MaskRegTbx1.Location = New System.Drawing.Point(12, 153)
        Me.MaskRegTbx1.Name = "MaskRegTbx1"
        Me.MaskRegTbx1.Size = New System.Drawing.Size(173, 20)
        Me.MaskRegTbx1.TabIndex = 26
        '
        'EmpBtn3
        '
        Me.EmpBtn3.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.EmpBtn3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EmpBtn3.Location = New System.Drawing.Point(90, 320)
        Me.EmpBtn3.Name = "EmpBtn3"
        Me.EmpBtn3.Size = New System.Drawing.Size(75, 23)
        Me.EmpBtn3.TabIndex = 27
        Me.EmpBtn3.Text = "&Save"
        Me.EmpBtn3.UseVisualStyleBackColor = True
        Me.EmpBtn3.Visible = False
        '
        'EmployeeBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(877, 457)
        Me.Controls.Add(Me.EmpBtn3)
        Me.Controls.Add(Me.MaskRegTbx1)
        Me.Controls.Add(Me.EmpBtn2)
        Me.Controls.Add(Me.EmpBtn1)
        Me.Controls.Add(Me.RegLbx10)
        Me.Controls.Add(Me.RegLbx9)
        Me.Controls.Add(Me.RegLbx8)
        Me.Controls.Add(Me.RegCmb5)
        Me.Controls.Add(Me.RegCmb3)
        Me.Controls.Add(Me.RegCmb2)
        Me.Controls.Add(Me.RegCmb1)
        Me.Controls.Add(Me.RegTbx5)
        Me.Controls.Add(Me.RegCmb4)
        Me.Controls.Add(Me.RegTbx3)
        Me.Controls.Add(Me.RegTbx2)
        Me.Controls.Add(Me.RegTbx1)
        Me.Controls.Add(Me.RegLbx7)
        Me.Controls.Add(Me.RegLbx6)
        Me.Controls.Add(Me.RegLbx5)
        Me.Controls.Add(Me.RegLbx4)
        Me.Controls.Add(Me.RegLbx3)
        Me.Controls.Add(Me.RegLbx2)
        Me.Controls.Add(Me.RegLbx1)
        Me.Controls.Add(Me.EmpGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "EmployeeBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "QC ~ Add Employee"
        CType(Me.EmpGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents EmpGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents RegLbx1 As System.Windows.Forms.Label
    Friend WithEvents RegLbx2 As System.Windows.Forms.Label
    Friend WithEvents RegLbx3 As System.Windows.Forms.Label
    Friend WithEvents RegLbx4 As System.Windows.Forms.Label
    Friend WithEvents RegLbx5 As System.Windows.Forms.Label
    Friend WithEvents RegLbx6 As System.Windows.Forms.Label
    Friend WithEvents RegLbx7 As System.Windows.Forms.Label
    Friend WithEvents RegTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents RegTbx2 As System.Windows.Forms.TextBox
    Friend WithEvents RegTbx3 As System.Windows.Forms.TextBox
    Friend WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RegCmb4 As System.Windows.Forms.ComboBox
    Friend WithEvents RegTbx5 As System.Windows.Forms.TextBox
    Friend WithEvents RegCmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents RegCmb2 As System.Windows.Forms.ComboBox
    Friend WithEvents RegCmb3 As System.Windows.Forms.ComboBox
    Friend WithEvents RegCmb5 As System.Windows.Forms.ComboBox
    Friend WithEvents RegLbx8 As System.Windows.Forms.Label
    Friend WithEvents RegLbx9 As System.Windows.Forms.Label
    Friend WithEvents RegLbx10 As System.Windows.Forms.Label
    Friend WithEvents EmpBtn2 As System.Windows.Forms.Button
    Friend WithEvents EmpBtn1 As System.Windows.Forms.Button
    Friend WithEvents MaskRegTbx1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents EmpBtn3 As System.Windows.Forms.Button
End Class
