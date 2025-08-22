<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Q
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Q))
        Me.DateCmb1 = New System.Windows.Forms.ComboBox()
        Me.DateBtn1 = New Glass.GlassButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateBtn2 = New Glass.GlassButton()
        Me.SuspendLayout()
        '
        'DateCmb1
        '
        Me.DateCmb1.Enabled = False
        Me.DateCmb1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateCmb1.FormattingEnabled = True
        Me.DateCmb1.Location = New System.Drawing.Point(32, 39)
        Me.DateCmb1.Name = "DateCmb1"
        Me.DateCmb1.Size = New System.Drawing.Size(224, 23)
        Me.DateCmb1.TabIndex = 76
        '
        'DateBtn1
        '
        Me.DateBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.DateBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateBtn1.ForeColor = System.Drawing.Color.Black
        Me.DateBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DateBtn1.Image = CType(resources.GetObject("DateBtn1.Image"), System.Drawing.Image)
        Me.DateBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DateBtn1.Location = New System.Drawing.Point(32, 103)
        Me.DateBtn1.Name = "DateBtn1"
        Me.DateBtn1.Size = New System.Drawing.Size(109, 29)
        Me.DateBtn1.TabIndex = 75
        Me.DateBtn1.Text = "New"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(144, 15)
        Me.Label1.TabIndex = 74
        Me.Label1.Text = "Incentives Periode Range"
        '
        'DateBtn2
        '
        Me.DateBtn2.BackColor = System.Drawing.Color.Gainsboro
        Me.DateBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateBtn2.ForeColor = System.Drawing.Color.Black
        Me.DateBtn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DateBtn2.Image = CType(resources.GetObject("DateBtn2.Image"), System.Drawing.Image)
        Me.DateBtn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DateBtn2.Location = New System.Drawing.Point(147, 103)
        Me.DateBtn2.Name = "DateBtn2"
        Me.DateBtn2.Size = New System.Drawing.Size(109, 29)
        Me.DateBtn2.TabIndex = 73
        Me.DateBtn2.Text = "Save"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(289, 142)
        Me.Controls.Add(Me.DateCmb1)
        Me.Controls.Add(Me.DateBtn1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateBtn2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Incentives Setup"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DateCmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents DateBtn1 As Glass.GlassButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateBtn2 As Glass.GlassButton
End Class
