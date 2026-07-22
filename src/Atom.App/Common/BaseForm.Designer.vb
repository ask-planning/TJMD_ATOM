Namespace Common
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class BaseForm
        Inherits System.Windows.Forms.Form

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
            Me.hdrHeader = New Atom.App.Common.HeaderControl()
            Me.pnlFooter = New System.Windows.Forms.Panel()
            Me.btnReconnect = New System.Windows.Forms.Button()
            Me.pnlContent = New System.Windows.Forms.Panel()
            Me.pnlFooter.SuspendLayout()
            Me.SuspendLayout()
            '
            'hdrHeader
            '
            Me.hdrHeader.Dock = System.Windows.Forms.DockStyle.Top
            Me.hdrHeader.Name = "hdrHeader"
            Me.hdrHeader.Size = New System.Drawing.Size(984, 34)
            Me.hdrHeader.TabIndex = 0
            '
            'pnlFooter
            '
            ' 「閉じる」はヘッダーへ移した。フッターは再接続のみ。
            Me.pnlFooter.Controls.Add(Me.btnReconnect)
            Me.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom
            Me.pnlFooter.Name = "pnlFooter"
            Me.pnlFooter.Padding = New System.Windows.Forms.Padding(8)
            Me.pnlFooter.Size = New System.Drawing.Size(984, 46)
            Me.pnlFooter.TabIndex = 1
            '
            'btnReconnect
            '
            Me.btnReconnect.Dock = System.Windows.Forms.DockStyle.Right
            Me.btnReconnect.Name = "btnReconnect"
            Me.btnReconnect.Size = New System.Drawing.Size(100, 30)
            Me.btnReconnect.TabIndex = 0
            Me.btnReconnect.Text = "再接続"
            Me.btnReconnect.UseVisualStyleBackColor = True
            '
            'pnlContent
            '
            Me.pnlContent.Dock = System.Windows.Forms.DockStyle.Fill
            Me.pnlContent.Name = "pnlContent"
            Me.pnlContent.Size = New System.Drawing.Size(984, 531)
            Me.pnlContent.TabIndex = 2
            '
            'BaseForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(984, 611)
            Me.Controls.Add(Me.pnlContent)
            Me.Controls.Add(Me.pnlFooter)
            Me.Controls.Add(Me.hdrHeader)
            Me.Font = New System.Drawing.Font("Meiryo UI", 9.0!)
            Me.Name = "BaseForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "BaseForm"
            Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
            Me.pnlFooter.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub

        Friend WithEvents hdrHeader As Atom.App.Common.HeaderControl
        Friend WithEvents pnlFooter As System.Windows.Forms.Panel
        Friend WithEvents btnReconnect As System.Windows.Forms.Button
        Protected Friend WithEvents pnlContent As System.Windows.Forms.Panel
    End Class
End Namespace
