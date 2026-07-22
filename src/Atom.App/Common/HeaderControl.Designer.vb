Namespace Common
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class HeaderControl
        Inherits System.Windows.Forms.UserControl

        'UserControl overrides dispose to clean up the component list.
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

        Private components As System.ComponentModel.IContainer

        <System.Diagnostics.DebuggerStepThrough()> _
        Private Sub InitializeComponent()
            Me.lblNameCaption = New System.Windows.Forms.Label()
            Me.lblName = New System.Windows.Forms.Label()
            Me.lblTitle = New System.Windows.Forms.Label()
            Me.pnlButtons = New System.Windows.Forms.FlowLayoutPanel()
            Me.SuspendLayout()
            '
            'lblNameCaption
            '
            Me.lblNameCaption.AutoSize = True
            Me.lblNameCaption.ForeColor = System.Drawing.SystemColors.GrayText
            Me.lblNameCaption.Location = New System.Drawing.Point(11, 12)
            Me.lblNameCaption.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblNameCaption.Name = "lblNameCaption"
            Me.lblNameCaption.Size = New System.Drawing.Size(40, 15)
            Me.lblNameCaption.TabIndex = 0
            Me.lblNameCaption.Text = "氏名:"
            '
            'lblName
            '
            Me.lblName.BackColor = System.Drawing.Color.FromArgb(CType(CType(238, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(242, Byte), Integer))
            Me.lblName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblName.Location = New System.Drawing.Point(60, 9)
            Me.lblName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblName.Name = "lblName"
            Me.lblName.Size = New System.Drawing.Size(187, 25)
            Me.lblName.TabIndex = 1
            Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'lblTitle
            '
            Me.lblTitle.Font = New System.Drawing.Font("Meiryo UI", 15.0!, System.Drawing.FontStyle.Bold)
            Me.lblTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(36, Byte), Integer), CType(CType(59, Byte), Integer), CType(CType(83, Byte), Integer))
            Me.lblTitle.Location = New System.Drawing.Point(267, 9)
            Me.lblTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblTitle.Name = "lblTitle"
            Me.lblTitle.Size = New System.Drawing.Size(480, 25)
            Me.lblTitle.TabIndex = 2
            Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'pnlButtons
            '
            Me.pnlButtons.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.pnlButtons.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
            Me.pnlButtons.Location = New System.Drawing.Point(2246, 9)
            Me.pnlButtons.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.pnlButtons.Name = "pnlButtons"
            Me.pnlButtons.Size = New System.Drawing.Size(427, 38)
            Me.pnlButtons.TabIndex = 3
            Me.pnlButtons.WrapContents = False
            '
            'HeaderControl
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.SystemColors.Control
            Me.Controls.Add(Me.lblTitle)
            Me.Controls.Add(Me.lblName)
            Me.Controls.Add(Me.lblNameCaption)
            Me.Controls.Add(Me.pnlButtons)
            Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.Name = "HeaderControl"
            Me.Size = New System.Drawing.Size(1350, 60)
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Friend WithEvents lblNameCaption As System.Windows.Forms.Label
        Friend WithEvents lblName As System.Windows.Forms.Label
        Friend WithEvents lblTitle As System.Windows.Forms.Label
        Friend WithEvents pnlButtons As System.Windows.Forms.FlowLayoutPanel
    End Class
End Namespace