<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DateCtrlBlock
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
        Me.HolGrid1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.HolTbx1 = New System.Windows.Forms.TextBox()
        Me.HolLbl1 = New System.Windows.Forms.Label()
        Me.HolCal1 = New System.Windows.Forms.MonthCalendar()
        Me.HolBtn1 = New System.Windows.Forms.Button()
        Me.HolLbl2 = New System.Windows.Forms.Label()
        Me.HolMaskTbx1 = New System.Windows.Forms.MaskedTextBox()
        Me.HolTimer1 = New System.Windows.Forms.Timer(Me.components)
        Me.HolBtn2 = New System.Windows.Forms.Button()
        CType(Me.HolGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'HolGrid1
        '
        Me.HolGrid1.AllowUserToAddRows = False
        Me.HolGrid1.AllowUserToDeleteRows = False
        Me.HolGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.HolGrid1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2})
        Me.HolGrid1.Location = New System.Drawing.Point(370, 12)
        Me.HolGrid1.Name = "HolGrid1"
        Me.HolGrid1.ReadOnly = True
        Me.HolGrid1.Size = New System.Drawing.Size(296, 287)
        Me.HolGrid1.TabIndex = 0
        '
        'Col1
        '
        Me.Col1.HeaderText = "Date"
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        Me.Col1.Width = 125
        '
        'Col2
        '
        Me.Col2.HeaderText = "Event"
        Me.Col2.Name = "Col2"
        Me.Col2.ReadOnly = True
        Me.Col2.Width = 125
        '
        'HolTbx1
        '
        Me.HolTbx1.Location = New System.Drawing.Point(12, 250)
        Me.HolTbx1.Name = "HolTbx1"
        Me.HolTbx1.Size = New System.Drawing.Size(227, 20)
        Me.HolTbx1.TabIndex = 10
        '
        'HolLbl1
        '
        Me.HolLbl1.AutoSize = True
        Me.HolLbl1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HolLbl1.Location = New System.Drawing.Point(9, 232)
        Me.HolLbl1.Name = "HolLbl1"
        Me.HolLbl1.Size = New System.Drawing.Size(99, 15)
        Me.HolLbl1.TabIndex = 9
        Me.HolLbl1.Text = "Event/Peristiwa:"
        '
        'HolCal1
        '
        Me.HolCal1.Location = New System.Drawing.Point(12, 12)
        Me.HolCal1.MaxSelectionCount = 1
        Me.HolCal1.Name = "HolCal1"
        Me.HolCal1.TabIndex = 11
        '
        'HolBtn1
        '
        Me.HolBtn1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.HolBtn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HolBtn1.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.HolBtn1.Location = New System.Drawing.Point(12, 276)
        Me.HolBtn1.Name = "HolBtn1"
        Me.HolBtn1.Size = New System.Drawing.Size(75, 23)
        Me.HolBtn1.TabIndex = 26
        Me.HolBtn1.Text = "Save"
        Me.HolBtn1.UseVisualStyleBackColor = True
        '
        'HolLbl2
        '
        Me.HolLbl2.AutoSize = True
        Me.HolLbl2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HolLbl2.Location = New System.Drawing.Point(12, 183)
        Me.HolLbl2.Name = "HolLbl2"
        Me.HolLbl2.Size = New System.Drawing.Size(37, 15)
        Me.HolLbl2.TabIndex = 27
        Me.HolLbl2.Text = "Date:"
        '
        'HolMaskTbx1
        '
        Me.HolMaskTbx1.Location = New System.Drawing.Point(15, 201)
        Me.HolMaskTbx1.Name = "HolMaskTbx1"
        Me.HolMaskTbx1.Size = New System.Drawing.Size(162, 20)
        Me.HolMaskTbx1.TabIndex = 28
        '
        'HolTimer1
        '
        Me.HolTimer1.Interval = 1000
        '
        'HolBtn2
        '
        Me.HolBtn2.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.HolBtn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HolBtn2.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.HolBtn2.Location = New System.Drawing.Point(164, 276)
        Me.HolBtn2.Name = "HolBtn2"
        Me.HolBtn2.Size = New System.Drawing.Size(75, 23)
        Me.HolBtn2.TabIndex = 29
        Me.HolBtn2.Text = "Delete"
        Me.HolBtn2.UseVisualStyleBackColor = True
        '
        'DateCtrlBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(671, 307)
        Me.Controls.Add(Me.HolBtn2)
        Me.Controls.Add(Me.HolMaskTbx1)
        Me.Controls.Add(Me.HolLbl2)
        Me.Controls.Add(Me.HolBtn1)
        Me.Controls.Add(Me.HolCal1)
        Me.Controls.Add(Me.HolTbx1)
        Me.Controls.Add(Me.HolLbl1)
        Me.Controls.Add(Me.HolGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "DateCtrlBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "QC ~ Holiday"
        CType(Me.HolGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents HolGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HolTbx1 As System.Windows.Forms.TextBox
    Friend WithEvents HolLbl1 As System.Windows.Forms.Label
    Friend WithEvents HolCal1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents HolBtn1 As System.Windows.Forms.Button
    Friend WithEvents HolLbl2 As System.Windows.Forms.Label
    Friend WithEvents HolMaskTbx1 As System.Windows.Forms.MaskedTextBox
    Private WithEvents HolTimer1 As System.Windows.Forms.Timer
    Friend WithEvents HolBtn2 As System.Windows.Forms.Button
End Class
