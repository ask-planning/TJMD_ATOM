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
            Me.lblLoginCaption = New System.Windows.Forms.Label()
            Me.lblLogin = New System.Windows.Forms.Label()
            Me.lblNameCaption = New System.Windows.Forms.Label()
            Me.lblName = New System.Windows.Forms.Label()
            Me.lblDeptCaption = New System.Windows.Forms.Label()
            Me.lblDept = New System.Windows.Forms.Label()
            Me.lblTitle = New System.Windows.Forms.Label()
            Me.SuspendLayout()
            '
            'lblLoginCaption
            '
            Me.lblLoginCaption.AutoSize = True
            Me.lblLoginCaption.ForeColor = System.Drawing.SystemColors.GrayText
            Me.lblLoginCaption.Location = New System.Drawing.Point(8, 10)
            Me.lblLoginCaption.Name = "lblLoginCaption"
            Me.lblLoginCaption.Size = New System.Drawing.Size(52, 12)
            Me.lblLoginCaption.TabIndex = 0
            Me.lblLoginCaption.Text = "ログイン:"
            '
            'lblLogin
            '
            Me.lblLogin.BackColor = System.Drawing.Color.FromArgb(CType(238, Integer), CType(240, Integer), CType(242, Integer))
            Me.lblLogin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblLogin.Location = New System.Drawing.Point(70, 7)
            Me.lblLogin.Name = "lblLogin"
            Me.lblLogin.Size = New System.Drawing.Size(70, 20)
            Me.lblLogin.TabIndex = 1
            Me.lblLogin.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'lblNameCaption
            '
            Me.lblNameCaption.AutoSize = True
            Me.lblNameCaption.ForeColor = System.Drawing.SystemColors.GrayText
            Me.lblNameCaption.Location = New System.Drawing.Point(150, 10)
            Me.lblNameCaption.Name = "lblNameCaption"
            Me.lblNameCaption.Size = New System.Drawing.Size(29, 12)
            Me.lblNameCaption.TabIndex = 2
            Me.lblNameCaption.Text = "氏名:"
            '
            'lblName
            '
            Me.lblName.BackColor = System.Drawing.Color.FromArgb(CType(238, Integer), CType(240, Integer), CType(242, Integer))
            Me.lblName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblName.Location = New System.Drawing.Point(185, 7)
            Me.lblName.Name = "lblName"
            Me.lblName.Size = New System.Drawing.Size(120, 20)
            Me.lblName.TabIndex = 3
            Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'lblDeptCaption
            '
            Me.lblDeptCaption.AutoSize = True
            Me.lblDeptCaption.ForeColor = System.Drawing.SystemColors.GrayText
            Me.lblDeptCaption.Location = New System.Drawing.Point(315, 10)
            Me.lblDeptCaption.Name = "lblDeptCaption"
            Me.lblDeptCaption.Size = New System.Drawing.Size(29, 12)
            Me.lblDeptCaption.TabIndex = 4
            Me.lblDeptCaption.Text = "部門:"
            '
            'lblDept
            '
            Me.lblDept.BackColor = System.Drawing.Color.FromArgb(CType(238, Integer), CType(240, Integer), CType(242, Integer))
            Me.lblDept.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblDept.Location = New System.Drawing.Point(350, 7)
            Me.lblDept.Name = "lblDept"
            Me.lblDept.Size = New System.Drawing.Size(90, 20)
            Me.lblDept.TabIndex = 5
            Me.lblDept.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'lblTitle
            '
            Me.lblTitle.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.lblTitle.Font = New System.Drawing.Font("Meiryo UI", 10.0!, System.Drawing.FontStyle.Bold)
            Me.lblTitle.ForeColor = System.Drawing.Color.FromArgb(CType(36, Integer), CType(59, Integer), CType(83, Integer))
            Me.lblTitle.Location = New System.Drawing.Point(692, 7)
            Me.lblTitle.Name = "lblTitle"
            Me.lblTitle.Size = New System.Drawing.Size(300, 20)
            Me.lblTitle.TabIndex = 6
            Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            '
            'HeaderControl
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.SystemColors.Control
            Me.Controls.Add(Me.lblTitle)
            Me.Controls.Add(Me.lblDept)
            Me.Controls.Add(Me.lblDeptCaption)
            Me.Controls.Add(Me.lblName)
            Me.Controls.Add(Me.lblNameCaption)
            Me.Controls.Add(Me.lblLogin)
            Me.Controls.Add(Me.lblLoginCaption)
            Me.Dock = System.Windows.Forms.DockStyle.Top
            Me.Name = "HeaderControl"
            Me.Size = New System.Drawing.Size(1000, 34)
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

        Friend WithEvents lblLoginCaption As System.Windows.Forms.Label
        Friend WithEvents lblLogin As System.Windows.Forms.Label
        Friend WithEvents lblNameCaption As System.Windows.Forms.Label
        Friend WithEvents lblName As System.Windows.Forms.Label
        Friend WithEvents lblDeptCaption As System.Windows.Forms.Label
        Friend WithEvents lblDept As System.Windows.Forms.Label
        Friend WithEvents lblTitle As System.Windows.Forms.Label
    End Class
End Namespace
