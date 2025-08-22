<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DedBlock
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.UpDedTbx1 = New System.Windows.Forms.MaskedTextBox()
        Me.DedBtn1 = New Glass.GlassButton()
        Me.OpenDedExcel = New System.Windows.Forms.OpenFileDialog()
        Me.DedUpGrid = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ViewLb2 = New System.Windows.Forms.Label()
        Me.DedPerCmb2 = New System.Windows.Forms.ComboBox()
        Me.DedPerCmb1 = New System.Windows.Forms.ComboBox()
        Me.DedBtn2 = New Glass.GlassButton()
        Me.DedBtn3 = New Glass.GlassButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.UpDedTbx3 = New System.Windows.Forms.MaskedTextBox()
        Me.UpDedTbx2 = New System.Windows.Forms.MaskedTextBox()
        CType(Me.DedUpGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UpDedTbx1
        '
        Me.UpDedTbx1.Location = New System.Drawing.Point(94, 32)
        Me.UpDedTbx1.Name = "UpDedTbx1"
        Me.UpDedTbx1.Size = New System.Drawing.Size(446, 20)
        Me.UpDedTbx1.TabIndex = 86
        '
        'DedBtn1
        '
        Me.DedBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.DedBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DedBtn1.ForeColor = System.Drawing.Color.Black
        Me.DedBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DedBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DedBtn1.Location = New System.Drawing.Point(12, 27)
        Me.DedBtn1.Name = "DedBtn1"
        Me.DedBtn1.Size = New System.Drawing.Size(76, 29)
        Me.DedBtn1.TabIndex = 85
        Me.DedBtn1.Text = "Browse"
        '
        'DedUpGrid
        '
        Me.DedUpGrid.AllowUserToAddRows = False
        Me.DedUpGrid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DedUpGrid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DedUpGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DedUpGrid.Location = New System.Drawing.Point(12, 116)
        Me.DedUpGrid.Name = "DedUpGrid"
        Me.DedUpGrid.ReadOnly = True
        Me.DedUpGrid.Size = New System.Drawing.Size(528, 390)
        Me.DedUpGrid.TabIndex = 87
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(129, 68)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 15)
        Me.Label1.TabIndex = 92
        Me.Label1.Text = "Periode"
        '
        'ViewLb2
        '
        Me.ViewLb2.AutoSize = True
        Me.ViewLb2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ViewLb2.Location = New System.Drawing.Point(10, 69)
        Me.ViewLb2.Name = "ViewLb2"
        Me.ViewLb2.Size = New System.Drawing.Size(86, 15)
        Me.ViewLb2.TabIndex = 91
        Me.ViewLb2.Text = "Periode Range"
        '
        'DedPerCmb2
        '
        Me.DedPerCmb2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DedPerCmb2.FormattingEnabled = True
        Me.DedPerCmb2.Items.AddRange(New Object() {"Periode I", "Periode II"})
        Me.DedPerCmb2.Location = New System.Drawing.Point(132, 87)
        Me.DedPerCmb2.Name = "DedPerCmb2"
        Me.DedPerCmb2.Size = New System.Drawing.Size(92, 23)
        Me.DedPerCmb2.TabIndex = 89
        '
        'DedPerCmb1
        '
        Me.DedPerCmb1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DedPerCmb1.FormattingEnabled = True
        Me.DedPerCmb1.Location = New System.Drawing.Point(12, 87)
        Me.DedPerCmb1.MaxDropDownItems = 5
        Me.DedPerCmb1.Name = "DedPerCmb1"
        Me.DedPerCmb1.Size = New System.Drawing.Size(114, 23)
        Me.DedPerCmb1.TabIndex = 88
        '
        'DedBtn2
        '
        Me.DedBtn2.BackColor = System.Drawing.Color.Gainsboro
        Me.DedBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DedBtn2.ForeColor = System.Drawing.Color.Black
        Me.DedBtn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DedBtn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DedBtn2.Location = New System.Drawing.Point(267, 81)
        Me.DedBtn2.Name = "DedBtn2"
        Me.DedBtn2.Size = New System.Drawing.Size(89, 29)
        Me.DedBtn2.TabIndex = 93
        Me.DedBtn2.Text = "Read"
        '
        'DedBtn3
        '
        Me.DedBtn3.BackColor = System.Drawing.Color.Gainsboro
        Me.DedBtn3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DedBtn3.ForeColor = System.Drawing.Color.Black
        Me.DedBtn3.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DedBtn3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DedBtn3.Location = New System.Drawing.Point(373, 81)
        Me.DedBtn3.Name = "DedBtn3"
        Me.DedBtn3.Size = New System.Drawing.Size(89, 29)
        Me.DedBtn3.TabIndex = 94
        Me.DedBtn3.Text = "Save"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 527)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 15)
        Me.Label2.TabIndex = 96
        Me.Label2.Text = "Read Count"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 553)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 15)
        Me.Label3.TabIndex = 98
        Me.Label3.Text = "Save Count"
        '
        'UpDedTbx3
        '
        Me.UpDedTbx3.Location = New System.Drawing.Point(120, 551)
        Me.UpDedTbx3.Name = "UpDedTbx3"
        Me.UpDedTbx3.Size = New System.Drawing.Size(156, 20)
        Me.UpDedTbx3.TabIndex = 97
        '
        'UpDedTbx2
        '
        Me.UpDedTbx2.Location = New System.Drawing.Point(120, 525)
        Me.UpDedTbx2.Name = "UpDedTbx2"
        Me.UpDedTbx2.Size = New System.Drawing.Size(156, 20)
        Me.UpDedTbx2.TabIndex = 95
        '
        'DedBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(552, 580)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.UpDedTbx3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.UpDedTbx2)
        Me.Controls.Add(Me.DedBtn3)
        Me.Controls.Add(Me.DedBtn2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ViewLb2)
        Me.Controls.Add(Me.DedPerCmb2)
        Me.Controls.Add(Me.DedPerCmb1)
        Me.Controls.Add(Me.DedUpGrid)
        Me.Controls.Add(Me.UpDedTbx1)
        Me.Controls.Add(Me.DedBtn1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "DedBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Deduction Module"
        CType(Me.DedUpGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UpDedTbx1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents DedBtn1 As Glass.GlassButton
    Friend WithEvents OpenDedExcel As System.Windows.Forms.OpenFileDialog
    Friend WithEvents DedUpGrid As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ViewLb2 As System.Windows.Forms.Label
    Friend WithEvents DedPerCmb2 As System.Windows.Forms.ComboBox
    Friend WithEvents DedPerCmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents DedBtn2 As Glass.GlassButton
    Friend WithEvents DedBtn3 As Glass.GlassButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents UpDedTbx3 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents UpDedTbx2 As System.Windows.Forms.MaskedTextBox
End Class
