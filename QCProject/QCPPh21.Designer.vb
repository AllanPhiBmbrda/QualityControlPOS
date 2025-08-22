<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PPH21Block
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
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PPH21Block))
        Me.PPhGrid1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PPh21cmb2 = New System.Windows.Forms.ComboBox()
        Me.PPh21cmb1 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ViewLb2 = New System.Windows.Forms.Label()
        Me.PPh21Btn3 = New Glass.GlassButton()
        Me.PPh21Btn2 = New Glass.GlassButton()
        Me.PPh21Btn1 = New Glass.GlassButton()
        CType(Me.PPhGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PPhGrid1
        '
        Me.PPhGrid1.AllowUserToAddRows = False
        Me.PPhGrid1.AllowUserToDeleteRows = False
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.PPhGrid1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle9
        Me.PPhGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.PPhGrid1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5, Me.Col6, Me.Col7, Me.Col8, Me.Col9, Me.Col10})
        Me.PPhGrid1.Location = New System.Drawing.Point(5, 60)
        Me.PPhGrid1.Name = "PPhGrid1"
        Me.PPhGrid1.ReadOnly = True
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle16.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.PPhGrid1.RowHeadersDefaultCellStyle = DataGridViewCellStyle16
        Me.PPhGrid1.Size = New System.Drawing.Size(1045, 430)
        Me.PPhGrid1.TabIndex = 1
        '
        'Col1
        '
        Me.Col1.HeaderText = "NIK"
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        '
        'Col2
        '
        Me.Col2.HeaderText = "NAMA"
        Me.Col2.Name = "Col2"
        Me.Col2.ReadOnly = True
        '
        'Col3
        '
        DataGridViewCellStyle10.NullValue = """-"""
        Me.Col3.DefaultCellStyle = DataGridViewCellStyle10
        Me.Col3.HeaderText = "Address"
        Me.Col3.Name = "Col3"
        Me.Col3.ReadOnly = True
        '
        'Col4
        '
        DataGridViewCellStyle11.NullValue = Nothing
        Me.Col4.DefaultCellStyle = DataGridViewCellStyle11
        Me.Col4.HeaderText = "KTP"
        Me.Col4.Name = "Col4"
        Me.Col4.ReadOnly = True
        '
        'Col5
        '
        Me.Col5.HeaderText = "NPWP"
        Me.Col5.Name = "Col5"
        Me.Col5.ReadOnly = True
        '
        'Col6
        '
        DataGridViewCellStyle12.NullValue = Nothing
        Me.Col6.DefaultCellStyle = DataGridViewCellStyle12
        Me.Col6.HeaderText = "Incentif"
        Me.Col6.Name = "Col6"
        Me.Col6.ReadOnly = True
        '
        'Col7
        '
        DataGridViewCellStyle13.NullValue = Nothing
        Me.Col7.DefaultCellStyle = DataGridViewCellStyle13
        Me.Col7.HeaderText = "Astek"
        Me.Col7.Name = "Col7"
        Me.Col7.ReadOnly = True
        '
        'Col8
        '
        DataGridViewCellStyle14.NullValue = Nothing
        Me.Col8.DefaultCellStyle = DataGridViewCellStyle14
        Me.Col8.HeaderText = "Result 1"
        Me.Col8.Name = "Col8"
        Me.Col8.ReadOnly = True
        '
        'Col9
        '
        DataGridViewCellStyle15.NullValue = Nothing
        Me.Col9.DefaultCellStyle = DataGridViewCellStyle15
        Me.Col9.HeaderText = "Result 2"
        Me.Col9.Name = "Col9"
        Me.Col9.ReadOnly = True
        '
        'Col10
        '
        Me.Col10.HeaderText = "Result 3"
        Me.Col10.Name = "Col10"
        Me.Col10.ReadOnly = True
        '
        'PPh21cmb2
        '
        Me.PPh21cmb2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPh21cmb2.FormattingEnabled = True
        Me.PPh21cmb2.Items.AddRange(New Object() {"All", "BTN", "CASH"})
        Me.PPh21cmb2.Location = New System.Drawing.Point(192, 31)
        Me.PPh21cmb2.Name = "PPh21cmb2"
        Me.PPh21cmb2.Size = New System.Drawing.Size(67, 23)
        Me.PPh21cmb2.TabIndex = 76
        Me.PPh21cmb2.Text = "All"
        '
        'PPh21cmb1
        '
        Me.PPh21cmb1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPh21cmb1.FormattingEnabled = True
        Me.PPh21cmb1.Location = New System.Drawing.Point(12, 31)
        Me.PPh21cmb1.MaxDropDownItems = 5
        Me.PPh21cmb1.Name = "PPh21cmb1"
        Me.PPh21cmb1.Size = New System.Drawing.Size(151, 23)
        Me.PPh21cmb1.TabIndex = 75
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(189, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(33, 15)
        Me.Label2.TabIndex = 78
        Me.Label2.Text = "Bank"
        '
        'ViewLb2
        '
        Me.ViewLb2.AutoSize = True
        Me.ViewLb2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ViewLb2.Location = New System.Drawing.Point(9, 9)
        Me.ViewLb2.Name = "ViewLb2"
        Me.ViewLb2.Size = New System.Drawing.Size(86, 15)
        Me.ViewLb2.TabIndex = 77
        Me.ViewLb2.Text = "Periode Range"
        '
        'PPh21Btn3
        '
        Me.PPh21Btn3.BackColor = System.Drawing.Color.Gainsboro
        Me.PPh21Btn3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPh21Btn3.ForeColor = System.Drawing.Color.Black
        Me.PPh21Btn3.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.PPh21Btn3.Image = CType(resources.GetObject("PPh21Btn3.Image"), System.Drawing.Image)
        Me.PPh21Btn3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.PPh21Btn3.Location = New System.Drawing.Point(562, 25)
        Me.PPh21Btn3.Name = "PPh21Btn3"
        Me.PPh21Btn3.Size = New System.Drawing.Size(116, 29)
        Me.PPh21Btn3.TabIndex = 150
        Me.PPh21Btn3.Text = "Refresh"
        '
        'PPh21Btn2
        '
        Me.PPh21Btn2.BackColor = System.Drawing.Color.Gainsboro
        Me.PPh21Btn2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPh21Btn2.ForeColor = System.Drawing.Color.Black
        Me.PPh21Btn2.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.PPh21Btn2.Image = CType(resources.GetObject("PPh21Btn2.Image"), System.Drawing.Image)
        Me.PPh21Btn2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.PPh21Btn2.Location = New System.Drawing.Point(440, 25)
        Me.PPh21Btn2.Name = "PPh21Btn2"
        Me.PPh21Btn2.Size = New System.Drawing.Size(116, 29)
        Me.PPh21Btn2.TabIndex = 149
        Me.PPh21Btn2.Text = "Excel"
        '
        'PPh21Btn1
        '
        Me.PPh21Btn1.BackColor = System.Drawing.Color.Gainsboro
        Me.PPh21Btn1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPh21Btn1.ForeColor = System.Drawing.Color.Black
        Me.PPh21Btn1.GlowColor = System.Drawing.Color.LightSkyBlue
        Me.PPh21Btn1.Image = CType(resources.GetObject("PPh21Btn1.Image"), System.Drawing.Image)
        Me.PPh21Btn1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.PPh21Btn1.Location = New System.Drawing.Point(318, 25)
        Me.PPh21Btn1.Name = "PPh21Btn1"
        Me.PPh21Btn1.Size = New System.Drawing.Size(116, 29)
        Me.PPh21Btn1.TabIndex = 148
        Me.PPh21Btn1.Text = "Generate"
        '
        'PPH21Block
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(1055, 494)
        Me.Controls.Add(Me.PPh21Btn3)
        Me.Controls.Add(Me.PPh21Btn2)
        Me.Controls.Add(Me.PPh21Btn1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ViewLb2)
        Me.Controls.Add(Me.PPh21cmb2)
        Me.Controls.Add(Me.PPh21cmb1)
        Me.Controls.Add(Me.PPhGrid1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "PPH21Block"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PPh21 Report Field"
        CType(Me.PPhGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PPhGrid1 As System.Windows.Forms.DataGridView
    Friend WithEvents PPh21cmb2 As System.Windows.Forms.ComboBox
    Friend WithEvents PPh21cmb1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ViewLb2 As System.Windows.Forms.Label
    Friend WithEvents PPh21Btn3 As Glass.GlassButton
    Friend WithEvents PPh21Btn2 As Glass.GlassButton
    Friend WithEvents PPh21Btn1 As Glass.GlassButton
    Friend WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col10 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
