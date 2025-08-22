<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StandardBlock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(StandardBlock))
        Me.StanTbx1 = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StanCmb1 = New System.Windows.Forms.ComboBox()
        Me.StanBtn1 = New Glass.GlassButton()
        Me.SuspendLayout()
        '
        'StanTbx1
        '
        Me.StanTbx1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StanTbx1.Location = New System.Drawing.Point(15, 70)
        Me.StanTbx1.Name = "StanTbx1"
        Me.StanTbx1.Size = New System.Drawing.Size(230, 23)
        Me.StanTbx1.TabIndex = 68
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.Location = New System.Drawing.Point(12, 52)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(89, 15)
        Me.Label33.TabIndex = 67
        Me.Label33.Text = "Standard Value"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 15)
        Me.Label1.TabIndex = 69
        Me.Label1.Text = "Standard Name"
        '
        'StanCmb1
        '
        Me.StanCmb1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StanCmb1.FormattingEnabled = True
        Me.StanCmb1.Items.AddRange(New Object() {"OriginalWage", "Jamsostek", "Minimum Trainee", "Subsidi", "Address1", "Address2", "Address3", "SignBy", "GajiCtrl"})
        Me.StanCmb1.Location = New System.Drawing.Point(15, 28)
        Me.StanCmb1.Name = "StanCmb1"
        Me.StanCmb1.Size = New System.Drawing.Size(230, 23)
        Me.StanCmb1.TabIndex = 70
        '
        'StanBtn1
        '
        Me.StanBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.StanBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StanBtn1.ForeColor = System.Drawing.Color.Black
        Me.StanBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.StanBtn1.Image = CType(resources.GetObject("StanBtn1.Image"), System.Drawing.Image)
        Me.StanBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.StanBtn1.Location = New System.Drawing.Point(61, 99)
        Me.StanBtn1.Name = "StanBtn1"
        Me.StanBtn1.Size = New System.Drawing.Size(129, 29)
        Me.StanBtn1.TabIndex = 71
        Me.StanBtn1.Text = "Save"
        '
        'StandardBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(250, 132)
        Me.Controls.Add(Me.StanBtn1)
        Me.Controls.Add(Me.StanCmb1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.StanTbx1)
        Me.Controls.Add(Me.Label33)
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "StandardBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "QC Codex~ Standard"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StanTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StanCmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents StanBtn1 As Glass.GlassButton
End Class
