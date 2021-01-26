<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmImportadorAutomatico
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()

	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
                fTerminateCalled = True
            End If

			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdVefivicar As System.Windows.Forms.Button
	Public WithEvents txtHoras As System.Windows.Forms.TextBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
    Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportadorAutomatico))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdVefivicar = New System.Windows.Forms.Button
        Me.txtHoras = New System.Windows.Forms.TextBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ConfigurarToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.FecharOProgramaToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportarBindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.ImportarBindingNavigatorSaveItem = New System.Windows.Forms.ToolStripButton
        Me.ImportarDataGridView = New System.Windows.Forms.DataGridView
        Me.ImportarBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ImportacaoDataSet = New ImportaçãoAutomática.ImportacaoDataSet
        Me.ImportarTableAdapter = New ImportaçãoAutomática.ImportacaoDataSetTableAdapters.ImportarTableAdapter
        Me.TableAdapterManager = New ImportaçãoAutomática.ImportacaoDataSetTableAdapters.TableAdapterManager
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UltimoCaixa = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ContextMenuStrip1.SuspendLayout()
        CType(Me.ImportarBindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ImportarBindingNavigator.SuspendLayout()
        CType(Me.ImportarDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ImportarBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ImportacaoDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdVefivicar
        '
        Me.cmdVefivicar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdVefivicar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdVefivicar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdVefivicar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdVefivicar.Location = New System.Drawing.Point(546, 259)
        Me.cmdVefivicar.Name = "cmdVefivicar"
        Me.cmdVefivicar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdVefivicar.Size = New System.Drawing.Size(87, 33)
        Me.cmdVefivicar.TabIndex = 10
        Me.cmdVefivicar.Text = "Verificar"
        Me.cmdVefivicar.UseVisualStyleBackColor = False
        '
        'txtHoras
        '
        Me.txtHoras.AcceptsReturn = True
        Me.txtHoras.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtHoras.BackColor = Global.ImportaçãoAutomática.My.MySettings.Default.TextBackColor
        Me.txtHoras.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtHoras.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.ImportaçãoAutomática.My.MySettings.Default, "TextForeColor", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.txtHoras.DataBindings.Add(New System.Windows.Forms.Binding("BackColor", Global.ImportaçãoAutomática.My.MySettings.Default, "TextBackColor", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.txtHoras.ForeColor = Global.ImportaçãoAutomática.My.MySettings.Default.TextForeColor
        Me.txtHoras.Location = New System.Drawing.Point(121, 266)
        Me.txtHoras.MaxLength = 0
        Me.txtHoras.Name = "txtHoras"
        Me.txtHoras.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtHoras.Size = New System.Drawing.Size(33, 20)
        Me.txtHoras.TabIndex = 8
        Me.txtHoras.Text = "3"
        Me.txtHoras.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 65535
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.BackColor = Global.ImportaçãoAutomática.My.MySettings.Default.LabelBackColor
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.DataBindings.Add(New System.Windows.Forms.Binding("BackColor", Global.ImportaçãoAutomática.My.MySettings.Default, "LabelBackColor", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label4.DataBindings.Add(New System.Windows.Forms.Binding("Font", Global.ImportaçãoAutomática.My.MySettings.Default, "LabelFont", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label4.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.ImportaçãoAutomática.My.MySettings.Default, "LabelForeColor", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label4.Font = Global.ImportaçãoAutomática.My.MySettings.Default.LabelFont
        Me.Label4.ForeColor = Global.ImportaçãoAutomática.My.MySettings.Default.LabelForeColor
        Me.Label4.Location = New System.Drawing.Point(160, 269)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(66, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Horas:"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.BackColor = Global.ImportaçãoAutomática.My.MySettings.Default.LabelBackColor
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.DataBindings.Add(New System.Windows.Forms.Binding("BackColor", Global.ImportaçãoAutomática.My.MySettings.Default, "LabelBackColor", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label3.DataBindings.Add(New System.Windows.Forms.Binding("Font", Global.ImportaçãoAutomática.My.MySettings.Default, "LabelFont", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label3.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.ImportaçãoAutomática.My.MySettings.Default, "LabelForeColor", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label3.Font = Global.ImportaçãoAutomática.My.MySettings.Default.LabelFont
        Me.Label3.ForeColor = Global.ImportaçãoAutomática.My.MySettings.Default.LabelForeColor
        Me.Label3.Location = New System.Drawing.Point(9, 266)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(125, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Importar a cada"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Importação Automática"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConfigurarToolStripMenuItem1, Me.FecharOProgramaToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(177, 48)
        '
        'ConfigurarToolStripMenuItem1
        '
        Me.ConfigurarToolStripMenuItem1.Name = "ConfigurarToolStripMenuItem1"
        Me.ConfigurarToolStripMenuItem1.Size = New System.Drawing.Size(176, 22)
        Me.ConfigurarToolStripMenuItem1.Text = "Configurar"
        '
        'FecharOProgramaToolStripMenuItem1
        '
        Me.FecharOProgramaToolStripMenuItem1.Name = "FecharOProgramaToolStripMenuItem1"
        Me.FecharOProgramaToolStripMenuItem1.Size = New System.Drawing.Size(176, 22)
        Me.FecharOProgramaToolStripMenuItem1.Text = "Fechar o programa"
        '
        'ImportarBindingNavigator
        '
        Me.ImportarBindingNavigator.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.ImportarBindingNavigator.BindingSource = Me.ImportarBindingSource
        Me.ImportarBindingNavigator.CountItem = Me.BindingNavigatorCountItem
        Me.ImportarBindingNavigator.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.ImportarBindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.ImportarBindingNavigatorSaveItem})
        Me.ImportarBindingNavigator.Location = New System.Drawing.Point(0, 0)
        Me.ImportarBindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.ImportarBindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.ImportarBindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.ImportarBindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.ImportarBindingNavigator.Name = "ImportarBindingNavigator"
        Me.ImportarBindingNavigator.PositionItem = Me.BindingNavigatorPositionItem
        Me.ImportarBindingNavigator.Size = New System.Drawing.Size(645, 25)
        Me.ImportarBindingNavigator.TabIndex = 11
        Me.ImportarBindingNavigator.Text = "BindingNavigator1"
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(36, 22)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 21)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ImportarBindingNavigatorSaveItem
        '
        Me.ImportarBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ImportarBindingNavigatorSaveItem.Image = CType(resources.GetObject("ImportarBindingNavigatorSaveItem.Image"), System.Drawing.Image)
        Me.ImportarBindingNavigatorSaveItem.Name = "ImportarBindingNavigatorSaveItem"
        Me.ImportarBindingNavigatorSaveItem.Size = New System.Drawing.Size(23, 22)
        Me.ImportarBindingNavigatorSaveItem.Text = "Save Data"
        '
        'ImportarDataGridView
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightSkyBlue
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.SteelBlue
        Me.ImportarDataGridView.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.ImportarDataGridView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ImportarDataGridView.AutoGenerateColumns = False
        Me.ImportarDataGridView.BackgroundColor = Global.ImportaçãoAutomática.My.MySettings.Default.TextBackColor
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.DarkBlue
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ImportarDataGridView.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.ImportarDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ImportarDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.UltimoCaixa, Me.DataGridViewTextBoxColumn4})
        Me.ImportarDataGridView.DataBindings.Add(New System.Windows.Forms.Binding("BackgroundColor", Global.ImportaçãoAutomática.My.MySettings.Default, "TextBackColor", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.ImportarDataGridView.DataSource = Me.ImportarBindingSource
        Me.ImportarDataGridView.Location = New System.Drawing.Point(0, 28)
        Me.ImportarDataGridView.Name = "ImportarDataGridView"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.DarkBlue
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ImportarDataGridView.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.LightCyan
        Me.ImportarDataGridView.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.ImportarDataGridView.Size = New System.Drawing.Size(645, 223)
        Me.ImportarDataGridView.TabIndex = 12
        '
        'ImportarBindingSource
        '
        Me.ImportarBindingSource.DataMember = "Importar"
        Me.ImportarBindingSource.DataSource = Me.ImportacaoDataSet
        '
        'ImportacaoDataSet
        '
        Me.ImportacaoDataSet.DataSetName = "ImportacaoDataSet"
        Me.ImportacaoDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ImportarTableAdapter
        '
        Me.ImportarTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.ImportarTableAdapter = Me.ImportarTableAdapter
        Me.TableAdapterManager.UpdateOrder = ImportaçãoAutomática.ImportacaoDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "codigoPosto"
        Me.DataGridViewTextBoxColumn1.HeaderText = "Codigo"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 65
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "NomePosto"
        Me.DataGridViewTextBoxColumn2.HeaderText = "Nome"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "LocalDB"
        Me.DataGridViewTextBoxColumn3.HeaderText = "Local do banco de dados"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        '
        'UltimoCaixa
        '
        Me.UltimoCaixa.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.UltimoCaixa.DataPropertyName = "UltimoCaixa"
        Me.UltimoCaixa.HeaderText = "Último Caixa"
        Me.UltimoCaixa.Name = "UltimoCaixa"
        Me.UltimoCaixa.Width = 90
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "UltimaImportacao"
        Me.DataGridViewTextBoxColumn4.HeaderText = "Ultima Importação"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 107
        '
        'frmImportadorAutomatico
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(645, 300)
        Me.Controls.Add(Me.ImportarBindingNavigator)
        Me.Controls.Add(Me.ImportarDataGridView)
        Me.Controls.Add(Me.txtHoras)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdVefivicar)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(3, 49)
        Me.Name = "frmImportadorAutomatico"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Importação automática"
        Me.ContextMenuStrip1.ResumeLayout(False)
        CType(Me.ImportarBindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ImportarBindingNavigator.ResumeLayout(False)
        Me.ImportarBindingNavigator.PerformLayout()
        CType(Me.ImportarDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ImportarBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ImportacaoDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
#Region "Upgrade Support"
	
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ConfigurarToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FecharOProgramaToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportacaoDataSet As ImportaçãoAutomática.ImportacaoDataSet
    Friend WithEvents ImportarBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ImportarTableAdapter As ImportaçãoAutomática.ImportacaoDataSetTableAdapters.ImportarTableAdapter
    Friend WithEvents TableAdapterManager As ImportaçãoAutomática.ImportacaoDataSetTableAdapters.TableAdapterManager
    Friend WithEvents ImportarBindingNavigator As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ImportarBindingNavigatorSaveItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents ImportarDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UltimoCaixa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class