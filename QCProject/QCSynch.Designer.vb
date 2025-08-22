<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SynchBlock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SynchBlock))
        Me.SynchGrid1 = New System.Windows.Forms.DataGridView()
        Me.SynchGrid2 = New System.Windows.Forms.DataGridView()
        Me.Userlbl1 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RegLbx4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SynchRadio1 = New System.Windows.Forms.RadioButton()
        Me.SynchRadio2 = New System.Windows.Forms.RadioButton()
        Me.SynchRadio3 = New System.Windows.Forms.RadioButton()
        Me.SynchRadio4 = New System.Windows.Forms.RadioButton()
        Me.SynchRadio5 = New System.Windows.Forms.RadioButton()
        Me.SynchRadio6 = New System.Windows.Forms.RadioButton()
        Me.SynchBtn1 = New Glass.GlassButton()
        Me.SynchBtn2 = New Glass.GlassButton()
        Me.SynchBtn3 = New Glass.GlassButton()
        Me.MaskSynch1 = New System.Windows.Forms.MaskedTextBox()
        Me.MaskSynch2 = New System.Windows.Forms.MaskedTextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.RecTbx1 = New System.Windows.Forms.TextBox()
        Me.RecTbx2 = New System.Windows.Forms.TextBox()
        CType(Me.SynchGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SynchGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SynchGrid1
        '
        Me.SynchGrid1.AllowUserToAddRows = False
        Me.SynchGrid1.AllowUserToDeleteRows = False
        Me.SynchGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.SynchGrid1.Location = New System.Drawing.Point(241, 43)
        Me.SynchGrid1.Name = "SynchGrid1"
        Me.SynchGrid1.ReadOnly = True
        Me.SynchGrid1.Size = New System.Drawing.Size(492, 188)
        Me.SynchGrid1.TabIndex = 0
        '
        'SynchGrid2
        '
        Me.SynchGrid2.AllowUserToAddRows = False
        Me.SynchGrid2.AllowUserToDeleteRows = False
        Me.SynchGrid2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.SynchGrid2.Location = New System.Drawing.Point(241, 275)
        Me.SynchGrid2.Name = "SynchGrid2"
        Me.SynchGrid2.ReadOnly = True
        Me.SynchGrid2.Size = New System.Drawing.Size(492, 194)
        Me.SynchGrid2.TabIndex = 1
        '
        'Userlbl1
        '
        Me.Userlbl1.AutoSize = True
        Me.Userlbl1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Userlbl1.Location = New System.Drawing.Point(238, 25)
        Me.Userlbl1.Name = "Userlbl1"
        Me.Userlbl1.Size = New System.Drawing.Size(124, 15)
        Me.Userlbl1.TabIndex = 2
        Me.Userlbl1.Text = "Per Item Table Synch:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(238, 257)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 15)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Per Total Table Synch:"
        '
        'RegLbx4
        '
        Me.RegLbx4.AutoSize = True
        Me.RegLbx4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx4.Location = New System.Drawing.Point(12, 25)
        Me.RegLbx4.Name = "RegLbx4"
        Me.RegLbx4.Size = New System.Drawing.Size(67, 15)
        Me.RegLbx4.TabIndex = 27
        Me.RegLbx4.Text = "Date Start:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 15)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Date To:"
        '
        'SynchRadio1
        '
        Me.SynchRadio1.AutoSize = True
        Me.SynchRadio1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchRadio1.Location = New System.Drawing.Point(15, 112)
        Me.SynchRadio1.Name = "SynchRadio1"
        Me.SynchRadio1.Size = New System.Drawing.Size(84, 19)
        Me.SynchRadio1.TabIndex = 31
        Me.SynchRadio1.TabStop = True
        Me.SynchRadio1.Text = "Conveyour"
        Me.SynchRadio1.UseVisualStyleBackColor = True
        '
        'SynchRadio2
        '
        Me.SynchRadio2.AutoSize = True
        Me.SynchRadio2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchRadio2.Location = New System.Drawing.Point(15, 137)
        Me.SynchRadio2.Name = "SynchRadio2"
        Me.SynchRadio2.Size = New System.Drawing.Size(64, 19)
        Me.SynchRadio2.TabIndex = 32
        Me.SynchRadio2.TabStop = True
        Me.SynchRadio2.Text = "Mutu II"
        Me.SynchRadio2.UseVisualStyleBackColor = True
        '
        'SynchRadio3
        '
        Me.SynchRadio3.AutoSize = True
        Me.SynchRadio3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchRadio3.Location = New System.Drawing.Point(15, 162)
        Me.SynchRadio3.Name = "SynchRadio3"
        Me.SynchRadio3.Size = New System.Drawing.Size(61, 19)
        Me.SynchRadio3.TabIndex = 33
        Me.SynchRadio3.TabStop = True
        Me.SynchRadio3.Text = "Wallet"
        Me.SynchRadio3.UseVisualStyleBackColor = True
        '
        'SynchRadio4
        '
        Me.SynchRadio4.AutoSize = True
        Me.SynchRadio4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchRadio4.Location = New System.Drawing.Point(15, 187)
        Me.SynchRadio4.Name = "SynchRadio4"
        Me.SynchRadio4.Size = New System.Drawing.Size(65, 19)
        Me.SynchRadio4.TabIndex = 34
        Me.SynchRadio4.TabStop = True
        Me.SynchRadio4.Text = "Packing"
        Me.SynchRadio4.UseVisualStyleBackColor = True
        '
        'SynchRadio5
        '
        Me.SynchRadio5.AutoSize = True
        Me.SynchRadio5.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchRadio5.Location = New System.Drawing.Point(15, 212)
        Me.SynchRadio5.Name = "SynchRadio5"
        Me.SynchRadio5.Size = New System.Drawing.Size(101, 19)
        Me.SynchRadio5.TabIndex = 35
        Me.SynchRadio5.TabStop = True
        Me.SynchRadio5.Text = "Miscellaneous"
        Me.SynchRadio5.UseVisualStyleBackColor = True
        '
        'SynchRadio6
        '
        Me.SynchRadio6.AutoSize = True
        Me.SynchRadio6.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchRadio6.Location = New System.Drawing.Point(15, 237)
        Me.SynchRadio6.Name = "SynchRadio6"
        Me.SynchRadio6.Size = New System.Drawing.Size(62, 19)
        Me.SynchRadio6.TabIndex = 36
        Me.SynchRadio6.TabStop = True
        Me.SynchRadio6.Text = "Sortasi"
        Me.SynchRadio6.UseVisualStyleBackColor = True
        '
        'SynchBtn1
        '
        Me.SynchBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.SynchBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchBtn1.ForeColor = System.Drawing.Color.Black
        Me.SynchBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.SynchBtn1.Image = CType(resources.GetObject("SynchBtn1.Image"), System.Drawing.Image)
        Me.SynchBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.SynchBtn1.Location = New System.Drawing.Point(15, 275)
        Me.SynchBtn1.Name = "SynchBtn1"
        Me.SynchBtn1.Size = New System.Drawing.Size(116, 29)
        Me.SynchBtn1.TabIndex = 145
        Me.SynchBtn1.Text = "Generate"
        '
        'SynchBtn2
        '
        Me.SynchBtn2.BackColor = System.Drawing.Color.Gainsboro
        Me.SynchBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchBtn2.ForeColor = System.Drawing.Color.Black
        Me.SynchBtn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.SynchBtn2.Image = CType(resources.GetObject("SynchBtn2.Image"), System.Drawing.Image)
        Me.SynchBtn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.SynchBtn2.Location = New System.Drawing.Point(15, 310)
        Me.SynchBtn2.Name = "SynchBtn2"
        Me.SynchBtn2.Size = New System.Drawing.Size(116, 29)
        Me.SynchBtn2.TabIndex = 146
        Me.SynchBtn2.Text = "Synch"
        '
        'SynchBtn3
        '
        Me.SynchBtn3.BackColor = System.Drawing.Color.Gainsboro
        Me.SynchBtn3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SynchBtn3.ForeColor = System.Drawing.Color.Black
        Me.SynchBtn3.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.SynchBtn3.Image = CType(resources.GetObject("SynchBtn3.Image"), System.Drawing.Image)
        Me.SynchBtn3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.SynchBtn3.Location = New System.Drawing.Point(15, 345)
        Me.SynchBtn3.Name = "SynchBtn3"
        Me.SynchBtn3.Size = New System.Drawing.Size(116, 29)
        Me.SynchBtn3.TabIndex = 147
        Me.SynchBtn3.Text = "Refresh"
        '
        'MaskSynch1
        '
        Me.MaskSynch1.Location = New System.Drawing.Point(15, 45)
        Me.MaskSynch1.Name = "MaskSynch1"
        Me.MaskSynch1.Size = New System.Drawing.Size(171, 20)
        Me.MaskSynch1.TabIndex = 148
        '
        'MaskSynch2
        '
        Me.MaskSynch2.Location = New System.Drawing.Point(15, 86)
        Me.MaskSynch2.Name = "MaskSynch2"
        Me.MaskSynch2.Size = New System.Drawing.Size(171, 20)
        Me.MaskSynch2.TabIndex = 149
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(13, 391)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(129, 15)
        Me.Label3.TabIndex = 150
        Me.Label3.Text = "RecordCount Per Item"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(13, 434)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(129, 15)
        Me.Label4.TabIndex = 151
        Me.Label4.Text = "RecordCount Per Total"
        '
        'RecTbx1
        '
        Me.RecTbx1.Enabled = False
        Me.RecTbx1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecTbx1.Location = New System.Drawing.Point(16, 409)
        Me.RecTbx1.Name = "RecTbx1"
        Me.RecTbx1.Size = New System.Drawing.Size(126, 22)
        Me.RecTbx1.TabIndex = 152
        '
        'RecTbx2
        '
        Me.RecTbx2.Enabled = False
        Me.RecTbx2.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecTbx2.Location = New System.Drawing.Point(16, 452)
        Me.RecTbx2.Name = "RecTbx2"
        Me.RecTbx2.Size = New System.Drawing.Size(126, 22)
        Me.RecTbx2.TabIndex = 153
        '
        'SynchBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(745, 482)
        Me.Controls.Add(Me.RecTbx2)
        Me.Controls.Add(Me.RecTbx1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.MaskSynch2)
        Me.Controls.Add(Me.MaskSynch1)
        Me.Controls.Add(Me.SynchBtn3)
        Me.Controls.Add(Me.SynchBtn2)
        Me.Controls.Add(Me.SynchBtn1)
        Me.Controls.Add(Me.SynchRadio6)
        Me.Controls.Add(Me.SynchRadio5)
        Me.Controls.Add(Me.SynchRadio4)
        Me.Controls.Add(Me.SynchRadio3)
        Me.Controls.Add(Me.SynchRadio2)
        Me.Controls.Add(Me.SynchRadio1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.RegLbx4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Userlbl1)
        Me.Controls.Add(Me.SynchGrid2)
        Me.Controls.Add(Me.SynchGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "SynchBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "QC ~ Synch"
        CType(Me.SynchGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SynchGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SynchGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents SynchGrid2 As System.Windows.Forms.DataGridView
    Friend WithEvents Userlbl1 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RegLbx4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SynchRadio1 As System.Windows.Forms.RadioButton
    Friend WithEvents SynchRadio2 As System.Windows.Forms.RadioButton
    Friend WithEvents SynchRadio3 As System.Windows.Forms.RadioButton
    Friend WithEvents SynchRadio4 As System.Windows.Forms.RadioButton
    Friend WithEvents SynchRadio5 As System.Windows.Forms.RadioButton
    Friend WithEvents SynchRadio6 As System.Windows.Forms.RadioButton
    Friend WithEvents SynchBtn1 As Glass.GlassButton
    Friend WithEvents SynchBtn2 As Glass.GlassButton
    Friend WithEvents SynchBtn3 As Glass.GlassButton
    Friend WithEvents MaskSynch1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents MaskSynch2 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents RecTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents RecTbx2 As System.Windows.Forms.TextBox
End Class
