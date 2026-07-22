Namespace Forms.Expense
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class ExpenseCodeMaintenanceForm
        Inherits Atom.App.Common.BaseForm

        <System.Diagnostics.DebuggerNonUserCode()>
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

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Me.dgvList = New System.Windows.Forms.DataGridView()
            Me.bsExpense = New System.Windows.Forms.BindingSource(Me.components)
            Me.colDenpyo = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.colText = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colRef = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colKbn = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.colHoken = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.pnlContent.SuspendLayout()
            CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.bsExpense, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'pnlContent
            '
            Me.pnlContent.Controls.Add(Me.dgvList)
            '
            'dgvList
            '
            Me.dgvList.AutoGenerateColumns = False
            Me.dgvList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colDenpyo, Me.colText, Me.colRef, Me.colKbn, Me.colHoken})
            Me.dgvList.DataSource = Me.bsExpense
            Me.dgvList.Dock = System.Windows.Forms.DockStyle.Fill
            Me.dgvList.Location = New System.Drawing.Point(0, 0)
            Me.dgvList.Name = "dgvList"
            Me.dgvList.RowHeadersWidth = 24
            Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.dgvList.Size = New System.Drawing.Size(984, 531)
            Me.dgvList.TabIndex = 0
            '
            'colDenpyo
            '
            Me.colDenpyo.DataPropertyName = "pk_keihi_code"
            DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
            Me.colDenpyo.DefaultCellStyle = DataGridViewCellStyle1
            Me.colDenpyo.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
            Me.colDenpyo.HeaderText = "伝票コード"
            Me.colDenpyo.MinimumWidth = 6
            Me.colDenpyo.Name = "colDenpyo"
            Me.colDenpyo.Width = 78
            '
            'colText
            '
            Me.colText.DataPropertyName = "keihi_code_text"
            Me.colText.HeaderText = "経費名称"
            Me.colText.MinimumWidth = 6
            Me.colText.Name = "colText"
            Me.colText.Width = 98
            '
            'colRef
            '
            Me.colRef.DataPropertyName = "keihi_code_ref"
            Me.colRef.HeaderText = "REF"
            Me.colRef.MinimumWidth = 6
            Me.colRef.Name = "colRef"
            Me.colRef.Width = 66
            '
            'colKbn
            '
            Me.colKbn.DataPropertyName = "keihi_kbn"
            DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
            Me.colKbn.DefaultCellStyle = DataGridViewCellStyle2
            Me.colKbn.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
            Me.colKbn.HeaderText = "区分"
            Me.colKbn.MinimumWidth = 6
            Me.colKbn.Name = "colKbn"
            Me.colKbn.Width = 46
            '
            'colHoken
            '
            Me.colHoken.DataPropertyName = "hoken_taisyo_flg"
            DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
            Me.colHoken.DefaultCellStyle = DataGridViewCellStyle3
            Me.colHoken.HeaderText = "保険対象"
            Me.colHoken.MinimumWidth = 6
            Me.colHoken.Name = "colHoken"
            Me.colHoken.Width = 98
            '
            'ExpenseCodeMaintenanceForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
            Me.ClientSize = New System.Drawing.Size(984, 611)
            Me.Location = New System.Drawing.Point(0, 0)
            Me.Name = "ExpenseCodeMaintenanceForm"
            Me.Text = ""
            Me.pnlContent.ResumeLayout(False)
            CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.bsExpense, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

        Friend WithEvents dgvList As System.Windows.Forms.DataGridView
        Friend WithEvents bsExpense As System.Windows.Forms.BindingSource
        Friend WithEvents colDenpyo As Windows.Forms.DataGridViewComboBoxColumn
        Friend WithEvents colText As Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colRef As Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colKbn As Windows.Forms.DataGridViewComboBoxColumn
        Friend WithEvents colHoken As Windows.Forms.DataGridViewTextBoxColumn
    End Class
End Namespace