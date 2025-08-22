<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class QCIncentivesGEN
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
        Me.IncCmb01 = New System.Windows.Forms.ComboBox()
        Me.IncGrid1 = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.InceTbx01 = New Glass.GlassButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.IncGenTbx1 = New System.Windows.Forms.TextBox()
        Me.InceTbx02 = New Glass.GlassButton()
        CType(Me.IncGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'IncCmb01
        '
        Me.IncCmb01.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IncCmb01.FormattingEnabled = True
        Me.IncCmb01.Location = New System.Drawing.Point(15, 30)
        Me.IncCmb01.MaxDropDownItems = 5
        Me.IncCmb01.Name = "IncCmb01"
        Me.IncCmb01.Size = New System.Drawing.Size(200, 23)
        Me.IncCmb01.TabIndex = 73
        '
        'IncGrid1
        '
        Me.IncGrid1.AllowUserToAddRows = False
        Me.IncGrid1.AllowUserToDeleteRows = False
        Me.IncGrid1.BackgroundColor = System.Drawing.Color.WhiteSmoke
        Me.IncGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.IncGrid1.Location = New System.Drawing.Point(240, 3)
        Me.IncGrid1.Name = "IncGrid1"
        Me.IncGrid1.ReadOnly = True
        Me.IncGrid1.Size = New System.Drawing.Size(664, 473)
        Me.IncGrid1.TabIndex = 77
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 15)
        Me.Label1.TabIndex = 78
        Me.Label1.Text = "Periode Range"
        '
        'InceTbx01
        '
        Me.InceTbx01.BackColor = System.Drawing.Color.Gainsboro
        Me.InceTbx01.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InceTbx01.ForeColor = System.Drawing.Color.Black
        Me.InceTbx01.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.InceTbx01.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.InceTbx01.Location = New System.Drawing.Point(15, 122)
        Me.InceTbx01.Name = "InceTbx01"
        Me.InceTbx01.Size = New System.Drawing.Size(109, 29)
        Me.InceTbx01.TabIndex = 79
        Me.InceTbx01.Text = "Read"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 68)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 15)
        Me.Label4.TabIndex = 201
        Me.Label4.Text = "Days / Hari :"
        '
        'IncGenTbx1
        '
        Me.IncGenTbx1.Location = New System.Drawing.Point(15, 86)
        Me.IncGenTbx1.Name = "IncGenTbx1"
        Me.IncGenTbx1.Size = New System.Drawing.Size(109, 20)
        Me.IncGenTbx1.TabIndex = 200
        '
        'InceTbx02
        '
        Me.InceTbx02.BackColor = System.Drawing.Color.Gainsboro
        Me.InceTbx02.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InceTbx02.ForeColor = System.Drawing.Color.Black
        Me.InceTbx02.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.InceTbx02.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.InceTbx02.Location = New System.Drawing.Point(15, 153)
        Me.InceTbx02.Name = "InceTbx02"
        Me.InceTbx02.Size = New System.Drawing.Size(109, 29)
        Me.InceTbx02.TabIndex = 202
        Me.InceTbx02.Text = "Save "
        '
        'QCIncentivesGEN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(905, 478)
        Me.Controls.Add(Me.InceTbx02)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.IncGenTbx1)
        Me.Controls.Add(Me.InceTbx01)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.IncGrid1)
        Me.Controls.Add(Me.IncCmb01)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "QCIncentivesGEN"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Incentives Control"
        CType(Me.IncGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents IncCmb01 As System.Windows.Forms.ComboBox
    Friend WithEvents IncGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents InceTbx01 As Glass.GlassButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents IncGenTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents InceTbx02 As Glass.GlassButton
End Class
