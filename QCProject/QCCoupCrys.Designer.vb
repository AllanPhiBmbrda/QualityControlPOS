<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CouponBlock
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
        Me.CryCoup01 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'CryCoup01
        '
        Me.CryCoup01.ActiveViewIndex = -1
        Me.CryCoup01.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CryCoup01.Cursor = System.Windows.Forms.Cursors.Default
        Me.CryCoup01.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CryCoup01.Location = New System.Drawing.Point(0, 0)
        Me.CryCoup01.Name = "CryCoup01"
        Me.CryCoup01.Size = New System.Drawing.Size(934, 676)
        Me.CryCoup01.TabIndex = 0
        Me.CryCoup01.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'CouponBlock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(934, 676)
        Me.Controls.Add(Me.CryCoup01)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "CouponBlock"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Coupon Slip"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents CryCoup01 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    'Friend WithEvents QCCouDesign1 As WindowsApplication1.QCCouDesign
    'Friend WithEvents QCCouDesign2 As WindowsApplication1.QCCouDesign
End Class
