<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UpEmpBlock
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.OpenExcel = New System.Windows.Forms.OpenFileDialog()
        Me.UpEmpTbx1 = New System.Windows.Forms.MaskedTextBox()
        Me.UPEmpBtn1 = New Glass.GlassButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.UpEmpTbx3 = New System.Windows.Forms.MaskedTextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.UpEmpTbx2 = New System.Windows.Forms.MaskedTextBox()
        Me.UPEmpBtn2 = New Glass.GlassButton()
        Me.UpEmpGrid = New System.Windows.Forms.DataGridView()
        Me.UPEmpBtn3 = New Glass.GlassButton()
        Me.UPEmpBtn6 = New Glass.GlassButton()
        CType(Me.UpEmpGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UpEmpTbx1
        '
        Me.UpEmpTbx1.Location = New System.Drawing.Point(94, 27)
        Me.UpEmpTbx1.Name = "UpEmpTbx1"
        Me.UpEmpTbx1.Size = New System.Drawing.Size(536, 20)
        Me.UpEmpTbx1.TabIndex = 84
        '
        'UPEmpBtn1
        '
        Me.UPEmpBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.UPEmpBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UPEmpBtn1.ForeColor = System.Drawing.Color.Black
        Me.UPEmpBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UPEmpBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UPEmpBtn1.Location = New System.Drawing.Point(7, 22)
        Me.UPEmpBtn1.Name = "UPEmpBtn1"
        Me.UPEmpBtn1.Size = New System.Drawing.Size(76, 29)
        Me.UPEmpBtn1.TabIndex = 83
        Me.UPEmpBtn1.Text = "Browse"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(9, 528)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 15)
        Me.Label3.TabIndex = 107
        Me.Label3.Text = "Save Count"
        '
        'UpEmpTbx3
        '
        Me.UpEmpTbx3.Location = New System.Drawing.Point(85, 526)
        Me.UpEmpTbx3.Name = "UpEmpTbx3"
        Me.UpEmpTbx3.Size = New System.Drawing.Size(156, 20)
        Me.UpEmpTbx3.TabIndex = 106
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 502)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 15)
        Me.Label2.TabIndex = 105
        Me.Label2.Text = "Read Count"
        '
        'UpEmpTbx2
        '
        Me.UpEmpTbx2.Location = New System.Drawing.Point(85, 500)
        Me.UpEmpTbx2.Name = "UpEmpTbx2"
        Me.UpEmpTbx2.Size = New System.Drawing.Size(156, 20)
        Me.UpEmpTbx2.TabIndex = 104
        '
        'UPEmpBtn2
        '
        Me.UPEmpBtn2.BackColor = System.Drawing.Color.Gainsboro
        Me.UPEmpBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UPEmpBtn2.ForeColor = System.Drawing.Color.Black
        Me.UPEmpBtn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UPEmpBtn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UPEmpBtn2.Location = New System.Drawing.Point(7, 57)
        Me.UPEmpBtn2.Name = "UPEmpBtn2"
        Me.UPEmpBtn2.Size = New System.Drawing.Size(118, 29)
        Me.UPEmpBtn2.TabIndex = 102
        Me.UPEmpBtn2.Text = "Read"
        '
        'UpEmpGrid
        '
        Me.UpEmpGrid.AllowUserToAddRows = False
        Me.UpEmpGrid.AllowUserToDeleteRows = False
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.UpEmpGrid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.UpEmpGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.UpEmpGrid.Location = New System.Drawing.Point(7, 88)
        Me.UpEmpGrid.Name = "UpEmpGrid"
        Me.UpEmpGrid.ReadOnly = True
        Me.UpEmpGrid.Size = New System.Drawing.Size(623, 395)
        Me.UpEmpGrid.TabIndex = 101
        '
        'UPEmpBtn3
        '
        Me.UPEmpBtn3.BackColor = System.Drawing.Color.Gainsboro
        Me.UPEmpBtn3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UPEmpBtn3.ForeColor = System.Drawing.Color.Black
        Me.UPEmpBtn3.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UPEmpBtn3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UPEmpBtn3.Location = New System.Drawing.Point(131, 57)
        Me.UPEmpBtn3.Name = "UPEmpBtn3"
        Me.UPEmpBtn3.Size = New System.Drawing.Size(118, 29)
        Me.UPEmpBtn3.TabIndex = 108
        Me.UPEmpBtn3.Text = "Upload"
        '
        'UPEmpBtn6
        '
        Me.UPEmpBtn6.BackColor = System.Drawing.Color.Gainsboro
        Me.UPEmpBtn6.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UPEmpBtn6.ForeColor = System.Drawing.Color.Black
        Me.UPEmpBtn6.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UPEmpBtn6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UPEmpBtn6.Location = New System.Drawing.Point(255, 57)
        Me.UPEmpBtn6.Name = "UPEmpBtn6"
        Me.UPEmpBtn6.Size = New System.Drawing.Size(118, 29)
        Me.UPEmpBtn6.TabIndex = 110
        Me.UPEmpBtn6.Text = "Refresh"
        '
        'UpEmpBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(639, 555)
        Me.Controls.Add(Me.UPEmpBtn6)
        Me.Controls.Add(Me.UPEmpBtn3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.UpEmpTbx3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.UpEmpTbx2)
        Me.Controls.Add(Me.UPEmpBtn2)
        Me.Controls.Add(Me.UpEmpGrid)
        Me.Controls.Add(Me.UpEmpTbx1)
        Me.Controls.Add(Me.UPEmpBtn1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "UpEmpBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "QC ~ Upload Employee"
        CType(Me.UpEmpGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenExcel As System.Windows.Forms.OpenFileDialog
    Friend WithEvents UpEmpTbx1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents UPEmpBtn1 As Glass.GlassButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents UpEmpTbx3 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents UpEmpTbx2 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents UPEmpBtn2 As Glass.GlassButton
    Friend WithEvents UpEmpGrid As System.Windows.Forms.DataGridView
    Friend WithEvents UPEmpBtn3 As Glass.GlassButton
    Friend WithEvents UPEmpBtn6 As Glass.GlassButton
End Class
