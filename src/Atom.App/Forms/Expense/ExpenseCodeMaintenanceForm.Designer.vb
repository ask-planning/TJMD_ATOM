Namespace Forms.Expense
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class ExpenseCodeMaintenanceForm
        Inherits Atom.App.Common.BaseForm

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
            Me.components = New System.ComponentModel.Container()
            Me.dgvList = New System.Windows.Forms.DataGridView()
            Me.bsExpense = New System.Windows.Forms.BindingSource(Me.components)
            Me.expenseDataSet = New ExpenseDataSet()
            Me.colDenpyo = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.colText = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colRef = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colKbn = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.colHoken = New System.Windows.Forms.DataGridViewTextBoxColumn()
            CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.bsExpense, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.expenseDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'dgvList
            '
            Me.dgvList.AllowUserToAddRows = True
            Me.dgvList.AllowUserToDeleteRows = True
            Me.dgvList.AutoGenerateColumns = False
            Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colDenpyo, Me.colText, Me.colRef, Me.colKbn, Me.colHoken})
            Me.dgvList.DataSource = Me.bsExpense
            Me.dgvList.Dock = System.Windows.Forms.DockStyle.Fill
            Me.dgvList.Name = "dgvList"
            Me.dgvList.RowHeadersWidth = 24
            Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.dgvList.Size = New System.Drawing.Size(984, 531)
            Me.dgvList.TabIndex = 0
            '
            'bsExpense
            '
            Me.bsExpense.DataMember = "ExpenseCode"
            Me.bsExpense.DataSource = Me.expenseDataSet
            '
            'expenseDataSet
            '
            Me.expenseDataSet.DataSetName = "ExpenseDataSet"
            Me.expenseDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
            '
            'colDenpyo
            '
            Me.colDenpyo.DataPropertyName = "pk_keihi_code"
            Me.colDenpyo.HeaderText = "伝票コード"
            Me.colDenpyo.Name = "colDenpyo"
            Me.colDenpyo.Width = 90
            '
            'colText
            '
            Me.colText.DataPropertyName = "keihi_code_text"
            Me.colText.HeaderText = "経費名称"
            Me.colText.Name = "colText"
            Me.colText.Width = 260
            '
            'colRef
            '
            Me.colRef.DataPropertyName = "keihi_code_ref"
            Me.colRef.HeaderText = "REF"
            Me.colRef.Name = "colRef"
            Me.colRef.Width = 200
            '
            'colKbn
            '
            Me.colKbn.DataPropertyName = "keihi_kbn"
            Me.colKbn.HeaderText = "区分"
            Me.colKbn.Name = "colKbn"
            Me.colKbn.Width = 110
            '
            'colHoken
            '
            Me.colHoken.DataPropertyName = "hoken_taisyo_flg"
            Me.colHoken.HeaderText = "保険対象"
            Me.colHoken.Name = "colHoken"
            Me.colHoken.Width = 70
            '
            'ExpenseCodeMaintenanceForm
            '
            Me.ContentPanel.Controls.Add(Me.dgvList)
            Me.Name = "ExpenseCodeMaintenanceForm"
            CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.bsExpense, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.expenseDataSet, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
        End Sub

        Friend WithEvents dgvList As System.Windows.Forms.DataGridView
        Friend WithEvents bsExpense As System.Windows.Forms.BindingSource
        Friend WithEvents expenseDataSet As ExpenseDataSet
        Friend WithEvents colDenpyo As System.Windows.Forms.DataGridViewComboBoxColumn
        Friend WithEvents colText As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colRef As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colKbn As System.Windows.Forms.DataGridViewComboBoxColumn
        Friend WithEvents colHoken As System.Windows.Forms.DataGridViewTextBoxColumn
    End Class
End Namespace
