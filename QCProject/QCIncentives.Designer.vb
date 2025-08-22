<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IncentivesBlock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IncentivesBlock))
        Me.InceDateCmb1 = New System.Windows.Forms.ComboBox()
        Me.InceBtn1 = New Glass.GlassButton()
        Me.InceLbl1 = New System.Windows.Forms.Label()
        Me.InceBtn2 = New Glass.GlassButton()
        Me.InceTbx1 = New System.Windows.Forms.TextBox()
        Me.IncLbl2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'InceDateCmb1
        '
        Me.InceDateCmb1.Enabled = False
        Me.InceDateCmb1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InceDateCmb1.FormattingEnabled = True
        Me.InceDateCmb1.Location = New System.Drawing.Point(15, 34)
        Me.InceDateCmb1.Name = "InceDateCmb1"
        Me.InceDateCmb1.Size = New System.Drawing.Size(254, 23)
        Me.InceDateCmb1.TabIndex = 76
        '
        'InceBtn1
        '
        Me.InceBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.InceBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InceBtn1.ForeColor = System.Drawing.Color.Black
        Me.InceBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.InceBtn1.Image = CType(resources.GetObject("InceBtn1.Image"), System.Drawing.Image)
        Me.InceBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.InceBtn1.Location = New System.Drawing.Point(15, 115)
        Me.InceBtn1.Name = "InceBtn1"
        Me.InceBtn1.Size = New System.Drawing.Size(109, 29)
        Me.InceBtn1.TabIndex = 75
        Me.InceBtn1.Text = "New"
        '
        'InceLbl1
        '
        Me.InceLbl1.AutoSize = True
        Me.InceLbl1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InceLbl1.Location = New System.Drawing.Point(12, 16)
        Me.InceLbl1.Name = "InceLbl1"
        Me.InceLbl1.Size = New System.Drawing.Size(133, 15)
        Me.InceLbl1.TabIndex = 74
        Me.InceLbl1.Text = "Month Incentive Range"
        '
        'InceBtn2
        '
        Me.InceBtn2.BackColor = System.Drawing.Color.Gainsboro
        Me.InceBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InceBtn2.ForeColor = System.Drawing.Color.Black
        Me.InceBtn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.InceBtn2.Image = CType(resources.GetObject("InceBtn2.Image"), System.Drawing.Image)
        Me.InceBtn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.InceBtn2.Location = New System.Drawing.Point(160, 115)
        Me.InceBtn2.Name = "InceBtn2"
        Me.InceBtn2.Size = New System.Drawing.Size(109, 29)
        Me.InceBtn2.TabIndex = 73
        Me.InceBtn2.Text = "Save"
        '
        'InceTbx1
        '
        Me.InceTbx1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InceTbx1.Location = New System.Drawing.Point(15, 86)
        Me.InceTbx1.Name = "InceTbx1"
        Me.InceTbx1.Size = New System.Drawing.Size(109, 23)
        Me.InceTbx1.TabIndex = 78
        '
        'IncLbl2
        '
        Me.IncLbl2.AutoSize = True
        Me.IncLbl2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IncLbl2.Location = New System.Drawing.Point(12, 68)
        Me.IncLbl2.Name = "IncLbl2"
        Me.IncLbl2.Size = New System.Drawing.Size(63, 15)
        Me.IncLbl2.TabIndex = 77
        Me.IncLbl2.Text = "Day Range"
        '
        'IncentivesBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(290, 151)
        Me.Controls.Add(Me.InceTbx1)
        Me.Controls.Add(Me.IncLbl2)
        Me.Controls.Add(Me.InceDateCmb1)
        Me.Controls.Add(Me.InceBtn1)
        Me.Controls.Add(Me.InceLbl1)
        Me.Controls.Add(Me.InceBtn2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "IncentivesBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Incentives"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents InceDateCmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents InceBtn1 As Glass.GlassButton
    Friend WithEvents InceLbl1 As System.Windows.Forms.Label
    Friend WithEvents InceBtn2 As Glass.GlassButton
    Friend WithEvents InceTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents IncLbl2 As System.Windows.Forms.Label
End Class
