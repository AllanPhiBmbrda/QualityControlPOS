<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PermitBlock
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PermitBlock))
        Me.DokTbx1 = New System.Windows.Forms.TextBox()
        Me.HolLbl1 = New System.Windows.Forms.Label()
        Me.DokTbx2 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.DokCal = New System.Windows.Forms.DateTimePicker()
        Me.DokLook = New Glass.GlassButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DokTime = New System.Windows.Forms.Label()
        Me.DokTbx3 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DokSave = New Glass.GlassButton()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.DokAdd = New Glass.GlassButton()
        Me.SuspendLayout()
        '
        'DokTbx1
        '
        Me.DokTbx1.Enabled = False
        Me.DokTbx1.Location = New System.Drawing.Point(66, 26)
        Me.DokTbx1.Name = "DokTbx1"
        Me.DokTbx1.Size = New System.Drawing.Size(223, 20)
        Me.DokTbx1.TabIndex = 83
        '
        'HolLbl1
        '
        Me.HolLbl1.AutoSize = True
        Me.HolLbl1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HolLbl1.Location = New System.Drawing.Point(6, 31)
        Me.HolLbl1.Name = "HolLbl1"
        Me.HolLbl1.Size = New System.Drawing.Size(29, 15)
        Me.HolLbl1.TabIndex = 82
        Me.HolLbl1.Text = "Nik:"
        '
        'DokTbx2
        '
        Me.DokTbx2.Enabled = False
        Me.DokTbx2.Location = New System.Drawing.Point(66, 52)
        Me.DokTbx2.Name = "DokTbx2"
        Me.DokTbx2.Size = New System.Drawing.Size(223, 20)
        Me.DokTbx2.TabIndex = 85
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 15)
        Me.Label1.TabIndex = 84
        Me.Label1.Text = "Nama:"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(6, 80)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(37, 15)
        Me.Label11.TabIndex = 190
        Me.Label11.Text = "Date:"
        '
        'DokCal
        '
        Me.DokCal.Enabled = False
        Me.DokCal.Location = New System.Drawing.Point(66, 80)
        Me.DokCal.Name = "DokCal"
        Me.DokCal.Size = New System.Drawing.Size(139, 20)
        Me.DokCal.TabIndex = 189
        '
        'DokLook
        '
        Me.DokLook.BackColor = System.Drawing.Color.Gainsboro
        Me.DokLook.Enabled = False
        Me.DokLook.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DokLook.ForeColor = System.Drawing.Color.Black
        Me.DokLook.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DokLook.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DokLook.Location = New System.Drawing.Point(295, 26)
        Me.DokLook.Name = "DokLook"
        Me.DokLook.Size = New System.Drawing.Size(31, 24)
        Me.DokLook.TabIndex = 191
        Me.DokLook.Text = ". . ."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 107)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 15)
        Me.Label2.TabIndex = 192
        Me.Label2.Text = "Time:"
        '
        'DokTime
        '
        Me.DokTime.AutoSize = True
        Me.DokTime.Enabled = False
        Me.DokTime.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DokTime.Location = New System.Drawing.Point(63, 107)
        Me.DokTime.Name = "DokTime"
        Me.DokTime.Size = New System.Drawing.Size(57, 15)
        Me.DokTime.TabIndex = 193
        Me.DokTime.Text = "00:00:00"
        '
        'DokTbx3
        '
        Me.DokTbx3.Enabled = False
        Me.DokTbx3.Location = New System.Drawing.Point(66, 132)
        Me.DokTbx3.Multiline = True
        Me.DokTbx3.Name = "DokTbx3"
        Me.DokTbx3.Size = New System.Drawing.Size(260, 99)
        Me.DokTbx3.TabIndex = 194
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 134)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 15)
        Me.Label4.TabIndex = 195
        Me.Label4.Text = "Remark:"
        '
        'DokSave
        '
        Me.DokSave.BackColor = System.Drawing.Color.Gainsboro
        Me.DokSave.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DokSave.ForeColor = System.Drawing.Color.Black
        Me.DokSave.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DokSave.Image = CType(resources.GetObject("DokSave.Image"), System.Drawing.Image)
        Me.DokSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DokSave.Location = New System.Drawing.Point(181, 237)
        Me.DokSave.Name = "DokSave"
        Me.DokSave.Size = New System.Drawing.Size(108, 29)
        Me.DokSave.TabIndex = 196
        Me.DokSave.Text = "&Save"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'DokAdd
        '
        Me.DokAdd.BackColor = System.Drawing.Color.Gainsboro
        Me.DokAdd.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DokAdd.ForeColor = System.Drawing.Color.Black
        Me.DokAdd.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.DokAdd.Image = CType(resources.GetObject("DokAdd.Image"), System.Drawing.Image)
        Me.DokAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DokAdd.Location = New System.Drawing.Point(66, 237)
        Me.DokAdd.Name = "DokAdd"
        Me.DokAdd.Size = New System.Drawing.Size(108, 29)
        Me.DokAdd.TabIndex = 197
        Me.DokAdd.Text = "&Add"
        '
        'PermitBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(354, 270)
        Me.Controls.Add(Me.DokAdd)
        Me.Controls.Add(Me.DokSave)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DokTbx3)
        Me.Controls.Add(Me.DokTime)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DokLook)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.DokCal)
        Me.Controls.Add(Me.DokTbx2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DokTbx1)
        Me.Controls.Add(Me.HolLbl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "PermitBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Surat Dokter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DokTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents HolLbl1 As System.Windows.Forms.Label
    Friend WithEvents DokTbx2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DokCal As System.Windows.Forms.DateTimePicker
    Friend WithEvents DokLook As Glass.GlassButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DokTime As System.Windows.Forms.Label
    Friend WithEvents DokTbx3 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DokSave As Glass.GlassButton
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents DokAdd As Glass.GlassButton
End Class
