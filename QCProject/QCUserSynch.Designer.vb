<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserSynchBlock
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UserSynchBlock))
        Me.UserSynchGrid1 = New System.Windows.Forms.DataGridView()
        Me.UserSynchRadio1 = New System.Windows.Forms.RadioButton()
        Me.RegLbx4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.UserSynchBtn1 = New Glass.GlassButton()
        Me.UserSynchBtn3 = New Glass.GlassButton()
        Me.UserSynchBtn2 = New Glass.GlassButton()
        Me.USDTPick01 = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.USDTPick02 = New System.Windows.Forms.DateTimePicker()
        CType(Me.UserSynchGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UserSynchGrid1
        '
        Me.UserSynchGrid1.AllowUserToAddRows = False
        Me.UserSynchGrid1.AllowUserToDeleteRows = False
        Me.UserSynchGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.UserSynchGrid1.Location = New System.Drawing.Point(236, 12)
        Me.UserSynchGrid1.Name = "UserSynchGrid1"
        Me.UserSynchGrid1.ReadOnly = True
        Me.UserSynchGrid1.Size = New System.Drawing.Size(860, 509)
        Me.UserSynchGrid1.TabIndex = 1
        '
        'UserSynchRadio1
        '
        Me.UserSynchRadio1.AutoSize = True
        Me.UserSynchRadio1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserSynchRadio1.Location = New System.Drawing.Point(12, 30)
        Me.UserSynchRadio1.Name = "UserSynchRadio1"
        Me.UserSynchRadio1.Size = New System.Drawing.Size(101, 19)
        Me.UserSynchRadio1.TabIndex = 32
        Me.UserSynchRadio1.TabStop = True
        Me.UserSynchRadio1.Text = "Employee List"
        Me.UserSynchRadio1.UseVisualStyleBackColor = True
        '
        'RegLbx4
        '
        Me.RegLbx4.AutoSize = True
        Me.RegLbx4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RegLbx4.Location = New System.Drawing.Point(12, 12)
        Me.RegLbx4.Name = "RegLbx4"
        Me.RegLbx4.Size = New System.Drawing.Size(109, 15)
        Me.RegLbx4.TabIndex = 73
        Me.RegLbx4.Text = "Select for Function"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 69)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 15)
        Me.Label1.TabIndex = 149
        Me.Label1.Text = "Date of Registry"
        '
        'UserSynchBtn1
        '
        Me.UserSynchBtn1.BackColor = System.Drawing.Color.Gainsboro
        Me.UserSynchBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserSynchBtn1.ForeColor = System.Drawing.Color.Black
        Me.UserSynchBtn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UserSynchBtn1.Image = CType(resources.GetObject("UserSynchBtn1.Image"), System.Drawing.Image)
        Me.UserSynchBtn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UserSynchBtn1.Location = New System.Drawing.Point(12, 209)
        Me.UserSynchBtn1.Name = "UserSynchBtn1"
        Me.UserSynchBtn1.Size = New System.Drawing.Size(116, 29)
        Me.UserSynchBtn1.TabIndex = 152
        Me.UserSynchBtn1.Text = "Generate"
        '
        'UserSynchBtn3
        '
        Me.UserSynchBtn3.BackColor = System.Drawing.Color.Gainsboro
        Me.UserSynchBtn3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserSynchBtn3.ForeColor = System.Drawing.Color.Black
        Me.UserSynchBtn3.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UserSynchBtn3.Image = CType(resources.GetObject("UserSynchBtn3.Image"), System.Drawing.Image)
        Me.UserSynchBtn3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UserSynchBtn3.Location = New System.Drawing.Point(12, 279)
        Me.UserSynchBtn3.Name = "UserSynchBtn3"
        Me.UserSynchBtn3.Size = New System.Drawing.Size(116, 29)
        Me.UserSynchBtn3.TabIndex = 154
        Me.UserSynchBtn3.Text = "Refresh"
        '
        'UserSynchBtn2
        '
        Me.UserSynchBtn2.BackColor = System.Drawing.Color.Gainsboro
        Me.UserSynchBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserSynchBtn2.ForeColor = System.Drawing.Color.Black
        Me.UserSynchBtn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.UserSynchBtn2.Image = CType(resources.GetObject("UserSynchBtn2.Image"), System.Drawing.Image)
        Me.UserSynchBtn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.UserSynchBtn2.Location = New System.Drawing.Point(12, 244)
        Me.UserSynchBtn2.Name = "UserSynchBtn2"
        Me.UserSynchBtn2.Size = New System.Drawing.Size(116, 29)
        Me.UserSynchBtn2.TabIndex = 153
        Me.UserSynchBtn2.Text = "Synch"
        '
        'USDTPick01
        '
        Me.USDTPick01.Location = New System.Drawing.Point(12, 115)
        Me.USDTPick01.Name = "USDTPick01"
        Me.USDTPick01.Size = New System.Drawing.Size(146, 20)
        Me.USDTPick01.TabIndex = 155
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 138)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(19, 15)
        Me.Label2.TabIndex = 156
        Me.Label2.Text = "To"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(36, 15)
        Me.Label3.TabIndex = 157
        Me.Label3.Text = "From"
        '
        'USDTPick02
        '
        Me.USDTPick02.Location = New System.Drawing.Point(12, 156)
        Me.USDTPick02.Name = "USDTPick02"
        Me.USDTPick02.Size = New System.Drawing.Size(146, 20)
        Me.USDTPick02.TabIndex = 158
        '
        'UserSynchBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(1102, 528)
        Me.Controls.Add(Me.USDTPick02)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.USDTPick01)
        Me.Controls.Add(Me.UserSynchBtn3)
        Me.Controls.Add(Me.UserSynchBtn2)
        Me.Controls.Add(Me.UserSynchBtn1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.RegLbx4)
        Me.Controls.Add(Me.UserSynchRadio1)
        Me.Controls.Add(Me.UserSynchGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "UserSynchBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Codex  ~ For User Synch"
        CType(Me.UserSynchGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UserSynchGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents UserSynchRadio1 As System.Windows.Forms.RadioButton
    Friend WithEvents RegLbx4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents UserSynchBtn1 As Glass.GlassButton
    Friend WithEvents UserSynchBtn3 As Glass.GlassButton
    Friend WithEvents UserSynchBtn2 As Glass.GlassButton
    Friend WithEvents USDTPick01 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents USDTPick02 As System.Windows.Forms.DateTimePicker
End Class
