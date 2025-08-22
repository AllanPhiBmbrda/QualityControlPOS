<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NoRekBlock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NoRekBlock))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ViewLb2 = New System.Windows.Forms.Label()
        Me.NRCmb2 = New System.Windows.Forms.ComboBox()
        Me.NRCmb1 = New System.Windows.Forms.ComboBox()
        Me.NRChkBox1 = New System.Windows.Forms.CheckBox()
        Me.NRChkBox2 = New System.Windows.Forms.CheckBox()
        Me.UpNRBtn3 = New Glass.GlassButton()
        Me.UpNRBtn1 = New Glass.GlassButton()
        Me.UpNRBtn2 = New Glass.GlassButton()
        Me.UpNRGrid = New System.Windows.Forms.DataGridView()
        CType(Me.UpNRGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(158, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 15)
        Me.Label1.TabIndex = 216
        Me.Label1.Text = "Periode"
        '
        'ViewLb2
        '
        Me.ViewLb2.AutoSize = True
        Me.ViewLb2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ViewLb2.Location = New System.Drawing.Point(12, 9)
        Me.ViewLb2.Name = "ViewLb2"
        Me.ViewLb2.Size = New System.Drawing.Size(86, 15)
        Me.ViewLb2.TabIndex = 215
        Me.ViewLb2.Text = "Periode Range"
        '
        'NRCmb2
        '
        Me.NRCmb2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NRCmb2.FormattingEnabled = True
        Me.NRCmb2.Items.AddRange(New Object() {"Periode I", "Periode II"})
        Me.NRCmb2.Location = New System.Drawing.Point(161, 27)
        Me.NRCmb2.Name = "NRCmb2"
        Me.NRCmb2.Size = New System.Drawing.Size(106, 23)
        Me.NRCmb2.TabIndex = 214
        '
        'NRCmb1
        '
        Me.NRCmb1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NRCmb1.FormattingEnabled = True
        Me.NRCmb1.Location = New System.Drawing.Point(14, 27)
        Me.NRCmb1.MaxDropDownItems = 5
        Me.NRCmb1.Name = "NRCmb1"
        Me.NRCmb1.Size = New System.Drawing.Size(141, 23)
        Me.NRCmb1.TabIndex = 213
        '
        'NRChkBox1
        '
        Me.NRChkBox1.AutoSize = True
        Me.NRChkBox1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NRChkBox1.Location = New System.Drawing.Point(292, 12)
        Me.NRChkBox1.Name = "NRChkBox1"
        Me.NRChkBox1.Size = New System.Drawing.Size(155, 19)
        Me.NRChkBox1.TabIndex = 217
        Me.NRChkBox1.Text = "No Rek (BTN Employee)"
        Me.NRChkBox1.UseVisualStyleBackColor = True
        '
        'NRChkBox2
        '
        Me.NRChkBox2.AutoSize = True
        Me.NRChkBox2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NRChkBox2.Location = New System.Drawing.Point(292, 33)
        Me.NRChkBox2.Name = "NRChkBox2"
        Me.NRChkBox2.Size = New System.Drawing.Size(150, 19)
        Me.NRChkBox2.TabIndex = 218
        Me.NRChkBox2.Text = "Cash -> Btn(Converter)"
        Me.NRChkBox2.UseVisualStyleBackColor = True
        '
        'UpNRBtn3
        '
        Me.UpNRBtn3.BackColor = System.Drawing.Color.Gainsboro
        Me.UpNRBtn3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UpNRBtn3.ForeColor = System.Drawing.Color.Black
        Me.UpNRBtn3.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UpNRBtn3.Image = CType(resources.GetObject("UpNRBtn3.Image"), System.Drawing.Image)
        Me.UpNRBtn3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UpNRBtn3.Location = New System.Drawing.Point(276, 56)
        Me.UpNRBtn3.Name = "UpNRBtn3"
        Me.UpNRBtn3.Size = New System.Drawing.Size(126, 29)
        Me.UpNRBtn3.TabIndex = 220
        Me.UpNRBtn3.Text = "Refresh"
        '
        'UpNRBtn1
        '
        Me.UpNRBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.UpNRBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UpNRBtn1.ForeColor = System.Drawing.Color.Black
        Me.UpNRBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UpNRBtn1.Image = CType(resources.GetObject("UpNRBtn1.Image"), System.Drawing.Image)
        Me.UpNRBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UpNRBtn1.Location = New System.Drawing.Point(12, 56)
        Me.UpNRBtn1.Name = "UpNRBtn1"
        Me.UpNRBtn1.Size = New System.Drawing.Size(126, 29)
        Me.UpNRBtn1.TabIndex = 219
        Me.UpNRBtn1.Text = "Generate"
        '
        'UpNRBtn2
        '
        Me.UpNRBtn2.BackColor = System.Drawing.Color.Gainsboro
        Me.UpNRBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UpNRBtn2.ForeColor = System.Drawing.Color.Black
        Me.UpNRBtn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UpNRBtn2.Image = CType(resources.GetObject("UpNRBtn2.Image"), System.Drawing.Image)
        Me.UpNRBtn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UpNRBtn2.Location = New System.Drawing.Point(144, 56)
        Me.UpNRBtn2.Name = "UpNRBtn2"
        Me.UpNRBtn2.Size = New System.Drawing.Size(126, 29)
        Me.UpNRBtn2.TabIndex = 221
        Me.UpNRBtn2.Text = "Update"
        '
        'UpNRGrid
        '
        Me.UpNRGrid.AllowUserToAddRows = False
        Me.UpNRGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.UpNRGrid.Location = New System.Drawing.Point(12, 92)
        Me.UpNRGrid.Name = "UpNRGrid"
        Me.UpNRGrid.ReadOnly = True
        Me.UpNRGrid.Size = New System.Drawing.Size(426, 418)
        Me.UpNRGrid.TabIndex = 222
        '
        'NoRekBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(450, 522)
        Me.Controls.Add(Me.UpNRGrid)
        Me.Controls.Add(Me.UpNRBtn2)
        Me.Controls.Add(Me.UpNRBtn3)
        Me.Controls.Add(Me.UpNRBtn1)
        Me.Controls.Add(Me.NRChkBox2)
        Me.Controls.Add(Me.NRChkBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ViewLb2)
        Me.Controls.Add(Me.NRCmb2)
        Me.Controls.Add(Me.NRCmb1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "NoRekBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pay/No Rek Updater"
        CType(Me.UpNRGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ViewLb2 As System.Windows.Forms.Label
    Friend WithEvents NRCmb2 As System.Windows.Forms.ComboBox
    Friend WithEvents NRCmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents NRChkBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents NRChkBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents UpNRBtn3 As Glass.GlassButton
    Friend WithEvents UpNRBtn1 As Glass.GlassButton
    Friend WithEvents UpNRBtn2 As Glass.GlassButton
    Friend WithEvents UpNRGrid As System.Windows.Forms.DataGridView
End Class
