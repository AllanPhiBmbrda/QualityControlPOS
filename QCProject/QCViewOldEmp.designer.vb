<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EmployeeOldBlock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EmployeeOldBlock))
        Me.ViewGrid1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ViewTbx1 = New System.Windows.Forms.TextBox()
        Me.ViewLb1 = New System.Windows.Forms.Label()
        Me.ViewLb2 = New System.Windows.Forms.Label()
        Me.ViewCmb1 = New System.Windows.Forms.ComboBox()
        Me.VECmb1 = New System.Windows.Forms.CheckBox()
        Me.EmpRadBtn1 = New System.Windows.Forms.RadioButton()
        Me.EmpRadBtn2 = New System.Windows.Forms.RadioButton()
        CType(Me.ViewGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ViewGrid1
        '
        Me.ViewGrid1.AllowUserToAddRows = False
        Me.ViewGrid1.AllowUserToDeleteRows = False
        Me.ViewGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ViewGrid1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3})
        Me.ViewGrid1.Location = New System.Drawing.Point(4, 75)
        Me.ViewGrid1.Name = "ViewGrid1"
        Me.ViewGrid1.ReadOnly = True
        Me.ViewGrid1.Size = New System.Drawing.Size(497, 260)
        Me.ViewGrid1.TabIndex = 0
        '
        'Col1
        '
        Me.Col1.HeaderText = "Nik"
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        '
        'Col2
        '
        Me.Col2.HeaderText = "Name"
        Me.Col2.Name = "Col2"
        Me.Col2.ReadOnly = True
        Me.Col2.Width = 250
        '
        'Col3
        '
        Me.Col3.HeaderText = "Pay"
        Me.Col3.Name = "Col3"
        Me.Col3.ReadOnly = True
        '
        'ViewTbx1
        '
        Me.ViewTbx1.Location = New System.Drawing.Point(124, 27)
        Me.ViewTbx1.Name = "ViewTbx1"
        Me.ViewTbx1.Size = New System.Drawing.Size(246, 20)
        Me.ViewTbx1.TabIndex = 10
        '
        'ViewLb1
        '
        Me.ViewLb1.AutoSize = True
        Me.ViewLb1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ViewLb1.Location = New System.Drawing.Point(1, 8)
        Me.ViewLb1.Name = "ViewLb1"
        Me.ViewLb1.Size = New System.Drawing.Size(102, 15)
        Me.ViewLb1.TabIndex = 9
        Me.ViewLb1.Text = "Idenitification As:"
        '
        'ViewLb2
        '
        Me.ViewLb2.AutoSize = True
        Me.ViewLb2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ViewLb2.Location = New System.Drawing.Point(121, 9)
        Me.ViewLb2.Name = "ViewLb2"
        Me.ViewLb2.Size = New System.Drawing.Size(68, 15)
        Me.ViewLb2.TabIndex = 23
        Me.ViewLb2.Text = "Search For:"
        '
        'ViewCmb1
        '
        Me.ViewCmb1.FormattingEnabled = True
        Me.ViewCmb1.Items.AddRange(New Object() {"Nik", "Name"})
        Me.ViewCmb1.Location = New System.Drawing.Point(4, 26)
        Me.ViewCmb1.Name = "ViewCmb1"
        Me.ViewCmb1.Size = New System.Drawing.Size(102, 21)
        Me.ViewCmb1.TabIndex = 24
        Me.ViewCmb1.Text = "Nik"
        '
        'VECmb1
        '
        Me.VECmb1.AutoSize = True
        Me.VECmb1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VECmb1.Location = New System.Drawing.Point(389, 12)
        Me.VECmb1.Name = "VECmb1"
        Me.VECmb1.Size = New System.Drawing.Size(85, 18)
        Me.VECmb1.TabIndex = 25
        Me.VECmb1.Text = "Auto Search"
        Me.VECmb1.UseVisualStyleBackColor = True
        '
        'EmpRadBtn1
        '
        Me.EmpRadBtn1.AutoSize = True
        Me.EmpRadBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EmpRadBtn1.Location = New System.Drawing.Point(389, 28)
        Me.EmpRadBtn1.Name = "EmpRadBtn1"
        Me.EmpRadBtn1.Size = New System.Drawing.Size(69, 19)
        Me.EmpRadBtn1.TabIndex = 26
        Me.EmpRadBtn1.Text = "Normal "
        Me.EmpRadBtn1.UseVisualStyleBackColor = True
        '
        'EmpRadBtn2
        '
        Me.EmpRadBtn2.AutoSize = True
        Me.EmpRadBtn2.Checked = True
        Me.EmpRadBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EmpRadBtn2.Location = New System.Drawing.Point(389, 50)
        Me.EmpRadBtn2.Name = "EmpRadBtn2"
        Me.EmpRadBtn2.Size = New System.Drawing.Size(47, 19)
        Me.EmpRadBtn2.TabIndex = 27
        Me.EmpRadBtn2.TabStop = True
        Me.EmpRadBtn2.Text = "Fast"
        Me.EmpRadBtn2.UseVisualStyleBackColor = True
        '
        'EmployeeOldBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(506, 340)
        Me.Controls.Add(Me.EmpRadBtn2)
        Me.Controls.Add(Me.EmpRadBtn1)
        Me.Controls.Add(Me.VECmb1)
        Me.Controls.Add(Me.ViewCmb1)
        Me.Controls.Add(Me.ViewLb2)
        Me.Controls.Add(Me.ViewTbx1)
        Me.Controls.Add(Me.ViewLb1)
        Me.Controls.Add(Me.ViewGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EmployeeOldBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Employee List Ver 2.0"
        CType(Me.ViewGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ViewGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents ViewTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents ViewLb1 As System.Windows.Forms.Label
    Friend WithEvents ViewLb2 As System.Windows.Forms.Label
    Friend WithEvents ViewCmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VECmb1 As System.Windows.Forms.CheckBox
    Friend WithEvents EmpRadBtn1 As System.Windows.Forms.RadioButton
    Friend WithEvents EmpRadBtn2 As System.Windows.Forms.RadioButton
End Class
