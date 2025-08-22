<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoginDoor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LoginDoor))
        Me.LoginBox = New System.Windows.Forms.GroupBox()
        Me.LogCmd02 = New System.Windows.Forms.Button()
        Me.LogCmd01 = New System.Windows.Forms.Button()
        Me.LoginTbx2 = New System.Windows.Forms.TextBox()
        Me.Loginlbl1 = New System.Windows.Forms.Label()
        Me.LoginTbx1 = New System.Windows.Forms.TextBox()
        Me.Loginlbl2 = New System.Windows.Forms.Label()
        Me.LogTimer = New System.Windows.Forms.Timer(Me.components)
        Me.LoginBox.SuspendLayout()
        Me.SuspendLayout()
        '
        'LoginBox
        '
        Me.LoginBox.BackColor = System.Drawing.Color.PowderBlue
        Me.LoginBox.Controls.Add(Me.LogCmd02)
        Me.LoginBox.Controls.Add(Me.LogCmd01)
        Me.LoginBox.Controls.Add(Me.LoginTbx2)
        Me.LoginBox.Controls.Add(Me.Loginlbl1)
        Me.LoginBox.Controls.Add(Me.LoginTbx1)
        Me.LoginBox.Controls.Add(Me.Loginlbl2)
        Me.LoginBox.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.LoginBox.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LoginBox.Location = New System.Drawing.Point(0, 0)
        Me.LoginBox.Name = "LoginBox"
        Me.LoginBox.Size = New System.Drawing.Size(281, 133)
        Me.LoginBox.TabIndex = 0
        Me.LoginBox.TabStop = False
        '
        'LogCmd02
        '
        Me.LogCmd02.Location = New System.Drawing.Point(165, 97)
        Me.LogCmd02.Name = "LogCmd02"
        Me.LogCmd02.Size = New System.Drawing.Size(89, 23)
        Me.LogCmd02.TabIndex = 8
        Me.LogCmd02.Text = "Cancel"
        Me.LogCmd02.UseVisualStyleBackColor = True
        '
        'LogCmd01
        '
        Me.LogCmd01.Location = New System.Drawing.Point(25, 97)
        Me.LogCmd01.Name = "LogCmd01"
        Me.LogCmd01.Size = New System.Drawing.Size(89, 23)
        Me.LogCmd01.TabIndex = 7
        Me.LogCmd01.Text = "Login"
        Me.LogCmd01.UseVisualStyleBackColor = True
        '
        'LoginTbx2
        '
        Me.LoginTbx2.Location = New System.Drawing.Point(104, 52)
        Me.LoginTbx2.Name = "LoginTbx2"
        Me.LoginTbx2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(8226)
        Me.LoginTbx2.Size = New System.Drawing.Size(150, 23)
        Me.LoginTbx2.TabIndex = 6
        '
        'Loginlbl1
        '
        Me.Loginlbl1.AutoSize = True
        Me.Loginlbl1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Loginlbl1.Location = New System.Drawing.Point(22, 23)
        Me.Loginlbl1.Name = "Loginlbl1"
        Me.Loginlbl1.Size = New System.Drawing.Size(76, 18)
        Me.Loginlbl1.TabIndex = 3
        Me.Loginlbl1.Text = "Username:"
        '
        'LoginTbx1
        '
        Me.LoginTbx1.Location = New System.Drawing.Point(104, 23)
        Me.LoginTbx1.Name = "LoginTbx1"
        Me.LoginTbx1.Size = New System.Drawing.Size(150, 23)
        Me.LoginTbx1.TabIndex = 4
        '
        'Loginlbl2
        '
        Me.Loginlbl2.AutoSize = True
        Me.Loginlbl2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Loginlbl2.Location = New System.Drawing.Point(22, 52)
        Me.Loginlbl2.Name = "Loginlbl2"
        Me.Loginlbl2.Size = New System.Drawing.Size(71, 18)
        Me.Loginlbl2.TabIndex = 5
        Me.Loginlbl2.Text = "Password:"
        '
        'LogTimer
        '
        Me.LogTimer.Enabled = True
        Me.LogTimer.Interval = 1
        '
        'LoginDoor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(282, 135)
        Me.Controls.Add(Me.LoginBox)
        Me.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LoginDoor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Codex ~ Login"
        Me.LoginBox.ResumeLayout(False)
        Me.LoginBox.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LoginBox As System.Windows.Forms.GroupBox
    Friend WithEvents LoginTbx2 As System.Windows.Forms.TextBox
    Friend WithEvents Loginlbl1 As System.Windows.Forms.Label
    Friend WithEvents LoginTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents Loginlbl2 As System.Windows.Forms.Label
    Friend WithEvents LogCmd02 As System.Windows.Forms.Button
    Friend WithEvents LogCmd01 As System.Windows.Forms.Button
    Friend WithEvents LogTimer As System.Windows.Forms.Timer
End Class
