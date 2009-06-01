Imports ciloci.FormulaEngine
Imports System.ComponentModel
Imports System.ComponentModel.Design

' Reference implementation showing how the formula engine can be used to build an Excel like
' application.
Public Class MainForm
	Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents edFormulaBar As System.Windows.Forms.TextBox
	Friend WithEvents tabSheets As System.Windows.Forms.TabControl
	Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
	Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents NewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents SaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents SaveAsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
	Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents CutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents CopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents PasteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
	Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents FillToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents FillDownToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents FillRightToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents NamesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents InformationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents cmsSheet As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents InsertSheetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents DeleteSheetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents menuDeleteSheetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents menuInsertSheetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
	Friend WithEvents ChartToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents EvaluateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ClearToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Me.edFormulaBar = New System.Windows.Forms.TextBox
		Me.tabSheets = New System.Windows.Forms.TabControl
		Me.cmsSheet = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.InsertSheetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.DeleteSheetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
		Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
		Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.CutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
		Me.FillToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.FillDownToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.FillRightToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.ClearToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.menuDeleteSheetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
		Me.menuInsertSheetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.ChartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.NamesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.InformationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.EvaluateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.Panel1 = New System.Windows.Forms.Panel
		Me.PictureBox1 = New System.Windows.Forms.PictureBox
		Me.cmsSheet.SuspendLayout()
		Me.MenuStrip1.SuspendLayout()
		Me.Panel1.SuspendLayout()
		CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'edFormulaBar
		'
		Me.edFormulaBar.BackColor = System.Drawing.SystemColors.Window
		Me.edFormulaBar.Dock = System.Windows.Forms.DockStyle.Fill
		Me.edFormulaBar.Location = New System.Drawing.Point(16, 0)
		Me.edFormulaBar.Name = "edFormulaBar"
		Me.edFormulaBar.Size = New System.Drawing.Size(608, 20)
		Me.edFormulaBar.TabIndex = 4
		Me.edFormulaBar.Text = "TextBox1"
		'
		'tabSheets
		'
		Me.tabSheets.Alignment = System.Windows.Forms.TabAlignment.Bottom
		Me.tabSheets.Dock = System.Windows.Forms.DockStyle.Fill
		Me.tabSheets.Location = New System.Drawing.Point(0, 44)
		Me.tabSheets.Name = "tabSheets"
		Me.tabSheets.SelectedIndex = 0
		Me.tabSheets.Size = New System.Drawing.Size(624, 393)
		Me.tabSheets.TabIndex = 5
		'
		'cmsSheet
		'
		Me.cmsSheet.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InsertSheetToolStripMenuItem, Me.DeleteSheetToolStripMenuItem})
		Me.cmsSheet.Name = "cmsSheet"
		Me.cmsSheet.Size = New System.Drawing.Size(117, 48)
		'
		'InsertSheetToolStripMenuItem
		'
		Me.InsertSheetToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.InsertSheet
		Me.InsertSheetToolStripMenuItem.Name = "InsertSheetToolStripMenuItem"
		Me.InsertSheetToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
		Me.InsertSheetToolStripMenuItem.Text = "Insert"
		'
		'DeleteSheetToolStripMenuItem
		'
		Me.DeleteSheetToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.DeleteHS
		Me.DeleteSheetToolStripMenuItem.Name = "DeleteSheetToolStripMenuItem"
		Me.DeleteSheetToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
		Me.DeleteSheetToolStripMenuItem.Text = "Delete"
		'
		'MenuStrip1
		'
		Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.ToolStripMenuItem1, Me.ToolsToolStripMenuItem})
		Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
		Me.MenuStrip1.Name = "MenuStrip1"
		Me.MenuStrip1.Size = New System.Drawing.Size(624, 24)
		Me.MenuStrip1.TabIndex = 6
		Me.MenuStrip1.Text = "MenuStrip1"
		'
		'FileToolStripMenuItem
		'
		Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.OpenToolStripMenuItem, Me.SaveToolStripMenuItem, Me.SaveAsToolStripMenuItem, Me.ToolStripSeparator1, Me.ExitToolStripMenuItem})
		Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
		Me.FileToolStripMenuItem.Size = New System.Drawing.Size(35, 20)
		Me.FileToolStripMenuItem.Text = "File"
		'
		'NewToolStripMenuItem
		'
		Me.NewToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.NewDocumentHS
		Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
		Me.NewToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
		Me.NewToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
		Me.NewToolStripMenuItem.Text = "New"
		'
		'OpenToolStripMenuItem
		'
		Me.OpenToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.Open
		Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
		Me.OpenToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
		Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
		Me.OpenToolStripMenuItem.Text = "Open..."
		'
		'SaveToolStripMenuItem
		'
		Me.SaveToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.Save
		Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
		Me.SaveToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
		Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
		Me.SaveToolStripMenuItem.Text = "Save"
		'
		'SaveAsToolStripMenuItem
		'
		Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
		Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
		Me.SaveAsToolStripMenuItem.Text = "Save As..."
		'
		'ToolStripSeparator1
		'
		Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
		Me.ToolStripSeparator1.Size = New System.Drawing.Size(160, 6)
		'
		'ExitToolStripMenuItem
		'
		Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
		Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
		Me.ExitToolStripMenuItem.Text = "Exit"
		'
		'EditToolStripMenuItem
		'
		Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CutToolStripMenuItem, Me.CopyToolStripMenuItem, Me.PasteToolStripMenuItem, Me.ToolStripSeparator2, Me.FillToolStripMenuItem, Me.ClearToolStripMenuItem, Me.menuDeleteSheetToolStripMenuItem})
		Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
		Me.EditToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
		Me.EditToolStripMenuItem.Text = "Edit"
		'
		'CutToolStripMenuItem
		'
		Me.CutToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.Cut
		Me.CutToolStripMenuItem.Name = "CutToolStripMenuItem"
		Me.CutToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
		Me.CutToolStripMenuItem.Size = New System.Drawing.Size(150, 22)
		Me.CutToolStripMenuItem.Text = "Cut"
		'
		'CopyToolStripMenuItem
		'
		Me.CopyToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.Copy
		Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
		Me.CopyToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
		Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(150, 22)
		Me.CopyToolStripMenuItem.Text = "Copy"
		'
		'PasteToolStripMenuItem
		'
		Me.PasteToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.Paste
		Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
		Me.PasteToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
		Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(150, 22)
		Me.PasteToolStripMenuItem.Text = "Paste"
		'
		'ToolStripSeparator2
		'
		Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
		Me.ToolStripSeparator2.Size = New System.Drawing.Size(147, 6)
		'
		'FillToolStripMenuItem
		'
		Me.FillToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FillDownToolStripMenuItem, Me.FillRightToolStripMenuItem})
		Me.FillToolStripMenuItem.Name = "FillToolStripMenuItem"
		Me.FillToolStripMenuItem.Size = New System.Drawing.Size(150, 22)
		Me.FillToolStripMenuItem.Text = "Fill"
		'
		'FillDownToolStripMenuItem
		'
		Me.FillDownToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.FillDownHS
		Me.FillDownToolStripMenuItem.Name = "FillDownToolStripMenuItem"
		Me.FillDownToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
		Me.FillDownToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
		Me.FillDownToolStripMenuItem.Text = "Down"
		'
		'FillRightToolStripMenuItem
		'
		Me.FillRightToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.FillRightHS
		Me.FillRightToolStripMenuItem.Name = "FillRightToolStripMenuItem"
		Me.FillRightToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
		Me.FillRightToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
		Me.FillRightToolStripMenuItem.Text = "Right"
		'
		'ClearToolStripMenuItem
		'
		Me.ClearToolStripMenuItem.Name = "ClearToolStripMenuItem"
		Me.ClearToolStripMenuItem.Size = New System.Drawing.Size(150, 22)
		Me.ClearToolStripMenuItem.Text = "Clear"
		'
		'menuDeleteSheetToolStripMenuItem
		'
		Me.menuDeleteSheetToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.DeleteHS
		Me.menuDeleteSheetToolStripMenuItem.Name = "menuDeleteSheetToolStripMenuItem"
		Me.menuDeleteSheetToolStripMenuItem.Size = New System.Drawing.Size(150, 22)
		Me.menuDeleteSheetToolStripMenuItem.Text = "Delete sheet"
		'
		'ToolStripMenuItem1
		'
		Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuInsertSheetToolStripMenuItem, Me.ChartToolStripMenuItem})
		Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
		Me.ToolStripMenuItem1.Size = New System.Drawing.Size(48, 20)
		Me.ToolStripMenuItem1.Text = "Insert"
		'
		'menuInsertSheetToolStripMenuItem
		'
		Me.menuInsertSheetToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.InsertSheet
		Me.menuInsertSheetToolStripMenuItem.Name = "menuInsertSheetToolStripMenuItem"
		Me.menuInsertSheetToolStripMenuItem.Size = New System.Drawing.Size(113, 22)
		Me.menuInsertSheetToolStripMenuItem.Text = "Sheet"
		'
		'ChartToolStripMenuItem
		'
		Me.ChartToolStripMenuItem.Name = "ChartToolStripMenuItem"
		Me.ChartToolStripMenuItem.Size = New System.Drawing.Size(113, 22)
		Me.ChartToolStripMenuItem.Text = "Chart"
		'
		'ToolsToolStripMenuItem
		'
		Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NamesToolStripMenuItem, Me.InformationToolStripMenuItem, Me.EvaluateToolStripMenuItem})
		Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
		Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
		Me.ToolsToolStripMenuItem.Text = "Tools"
		'
		'NamesToolStripMenuItem
		'
		Me.NamesToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.Names
		Me.NamesToolStripMenuItem.Name = "NamesToolStripMenuItem"
		Me.NamesToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
		Me.NamesToolStripMenuItem.Text = "Names..."
		'
		'InformationToolStripMenuItem
		'
		Me.InformationToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.Information
		Me.InformationToolStripMenuItem.Name = "InformationToolStripMenuItem"
		Me.InformationToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
		Me.InformationToolStripMenuItem.Text = "Information..."
		'
		'EvaluateToolStripMenuItem
		'
		Me.EvaluateToolStripMenuItem.Image = Global.GridSample.My.Resources.Images.CalculatorHS
		Me.EvaluateToolStripMenuItem.Name = "EvaluateToolStripMenuItem"
		Me.EvaluateToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
		Me.EvaluateToolStripMenuItem.Text = "Evaluate"
		'
		'Panel1
		'
		Me.Panel1.Controls.Add(Me.edFormulaBar)
		Me.Panel1.Controls.Add(Me.PictureBox1)
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
		Me.Panel1.Location = New System.Drawing.Point(0, 24)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(624, 20)
		Me.Panel1.TabIndex = 7
		'
		'PictureBox1
		'
		Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Left
		Me.PictureBox1.Image = Global.GridSample.My.Resources.Images.FunctionHS
		Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
		Me.PictureBox1.Name = "PictureBox1"
		Me.PictureBox1.Size = New System.Drawing.Size(16, 20)
		Me.PictureBox1.TabIndex = 5
		Me.PictureBox1.TabStop = False
		'
		'MainForm
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(624, 437)
		Me.Controls.Add(Me.tabSheets)
		Me.Controls.Add(Me.Panel1)
		Me.Controls.Add(Me.MenuStrip1)
		Me.MainMenuStrip = Me.MenuStrip1
		Me.Name = "MainForm"
		Me.Text = "Grid Sample"
		Me.cmsSheet.ResumeLayout(False)
		Me.MenuStrip1.ResumeLayout(False)
		Me.MenuStrip1.PerformLayout()
		Me.Panel1.ResumeLayout(False)
		Me.Panel1.PerformLayout()
		CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

#End Region

	Private MyServices As IServiceContainer
	Private MyActiveSheet As Worksheet
	Private Const MAX_SHEETS As Integer = 5
	Private MySheets As IList
	Private MyActiveStream As System.IO.FileStream		' The file stream to the current document we are working with
	Public Const APP_NAME As String = "Grid Sample"
	Public Event ActiveSheetChanged As EventHandler
	Public Event RefreshRequired As EventHandler

	Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
		MyBase.OnLoad(e)

		MySheets = New ArrayList

		Me.SetUpComponentsAndServices()

		Me.AddSheet()
		Me.SetActiveStream(Nothing)
	End Sub

	Private Sub SetUpComponentsAndServices()
		MyServices = New ServiceContainer
		Me.components = New CustomContainer(MyServices)

		Dim ms As New MessageService()
		ms.SetActiveDialog(Me)
		MyServices.AddService(GetType(MessageService), ms)
		MyServices.AddService(GetType(DictionaryService), New DictionaryService())
		MyServices.AddService(GetType(DocumentStatusService), New DocumentStatusService())

		MyServices.AddService(GetType(MainForm), Me)
		Dim engine As New FormulaEngine()
		Me.SetFormulaEngine(engine)

		Dim elc As New ExcelLikeCell(Me)
		MyServices.AddService(GetType(ExcelLikeCell), elc)
		Me.AddComponentService(GetType(FormulaBar))
		Me.AddComponentService(GetType(ClipboardManager))
	End Sub

	Private Sub AddComponentService(ByVal componentType As Type)
		Dim instance As IComponent = Activator.CreateInstance(componentType)
		Me.components.Add(instance)
		Dim init As IInitialize = TryCast(instance, IInitialize)
		If Not init Is Nothing Then
			init.Initialize()
		End If
		MyServices.AddService(componentType, instance)
	End Sub

	Private Sub SetFormulaEngine(ByVal engine As FormulaEngine)
		Dim oldEngine As FormulaEngine = MyServices.GetService(GetType(FormulaEngine))
		If Not oldEngine Is Nothing Then
			RemoveHandler engine.CircularReferenceDetected, AddressOf OnCircularReferenceDetected
		End If

		MyServices.RemoveService(GetType(FormulaEngine))
		MyServices.AddService(GetType(FormulaEngine), engine)

		AddHandler engine.CircularReferenceDetected, AddressOf OnCircularReferenceDetected
	End Sub

	Protected Overrides Sub OnClosed(ByVal e As System.EventArgs)
		MyBase.OnClosed(e)
		Me.SetActiveStream(Nothing)
	End Sub

	Friend Sub ShowMessage(ByVal message As String, ByVal mt As MessageType)
		Dim ms As MessageService = MyServices.GetService(GetType(MessageService))
		ms.ShowMessage(message, mt)
	End Sub

	Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
		Me.Close()
	End Sub

	Private Sub miShowNamesForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NamesToolStripMenuItem.Click
		Dim frm As New NamesForm
		Me.components.Add(frm)
		frm.Initialize()
		frm.ShowDialog(Me)
		Me.components.Remove(frm)
		frm.Dispose()
	End Sub

	Private Sub miInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InformationToolStripMenuItem.Click
		Dim frm As New InfoForm
		frm.SetEngine(Me.Engine)
		frm.ShowDialog(Me)
		frm.Dispose()
	End Sub

	Private Sub EvaluateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EvaluateToolStripMenuItem.Click
		Dim dialog As New EvaluateForm()
		Me.EvaluateToolStripMenuItem.Enabled = False
		Me.components.Add(dialog)
		AddHandler dialog.Closed, AddressOf OnEvaluateFormClosed
		dialog.Show(Me)
	End Sub

	Private Sub OnEvaluateFormClosed(ByVal sender As Object, ByVal e As EventArgs)
		Dim dialog As EvaluateForm = sender
		RemoveHandler dialog.Closed, AddressOf OnEvaluateFormClosed
		Me.components.Remove(dialog)
		Me.EvaluateToolStripMenuItem.Enabled = True
	End Sub

	Private Sub InsertSheetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertSheetToolStripMenuItem.Click, menuInsertSheetToolStripMenuItem.Click
		If MySheets.Count < MAX_SHEETS Then
			Me.AddSheet()
		Else
			Me.ShowMessage(String.Format("Cannot have more than {0} sheets", MAX_SHEETS), MessageType.Error)
		End If
	End Sub

	Private Sub DeleteSheetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteSheetToolStripMenuItem.Click, menuDeleteSheetToolStripMenuItem.Click
		If Me.tabSheets.TabCount > 1 Then
			Me.RemoveActiveSheet()
		Else
			Me.ShowMessage("Cannot delete all sheets", MessageType.Error)
		End If
	End Sub

	' Add a worksheet to the application
	Private Sub AddSheet()
		Dim page As New TabPage

		' Get a name
		Dim name As String = Me.CreateSheetName()
		' Create a worksheet
		Dim sheet As New Worksheet(name)

		page.Text = name
		Me.components.Add(sheet)
		sheet.SizeToDefault()
		sheet.SelectFirstCell()
		' Register the sheet with the formula engine
		Me.Engine.Sheets.Add(sheet)
		MySheets.Add(sheet)

		sheet.SetGridParent(page)

		Me.tabSheets.SelectedTab = page
		Me.tabSheets.TabPages.Add(page)
		Me.SetActiveSheet(sheet)
	End Sub

	Private Function CreateSheetName() As String
		Dim ordinal As Integer = MySheets.Count + 1

		Do
			Dim name As String = String.Concat("Sheet", ordinal)
			If Me.Engine.Sheets.GetSheetByName(name) Is Nothing Then
				Return name
			Else
				ordinal += 1
			End If
		Loop
	End Function

	Private Sub TabSheet_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tabSheets.MouseClick
		If e.Button <> Windows.Forms.MouseButtons.Right Then
			Return
		End If
		Dim index As Integer = Me.GetTabIndexForPoint(e.Location)
		Me.tabSheets.SelectTab(index)
		Me.cmsSheet.Show(sender, e.Location)
	End Sub

	Private Function GetTabIndexForPoint(ByVal p As Point) As Integer
		For i As Integer = 0 To Me.tabSheets.TabPages.Count - 1
			Dim rect As Rectangle = Me.tabSheets.GetTabRect(i)
			If rect.Contains(p) = True Then
				Return i
			End If
		Next

		Return -1
	End Function

	Private Sub tabSheets_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabSheets.SelectedIndexChanged
		If Me.tabSheets.SelectedIndex <> -1 Then
			Dim selectedIndex As Integer = Me.tabSheets.SelectedIndex
			Me.SetActiveSheet(MySheets.Item(selectedIndex))
		End If
	End Sub

	' Set our active worksheet
	Private Sub SetActiveSheet(ByVal sheet As Worksheet)
		MyActiveSheet = sheet
		' Let the formula engine know that this is the active sheet
		Me.Engine.Sheets.ActiveSheet = sheet
		Me.ActiveControl = MyActiveSheet.Grid
		MyServices.RemoveService(GetType(Worksheet))
		MyServices.AddService(GetType(Worksheet), sheet)
		RaiseEvent ActiveSheetChanged(Me, EventArgs.Empty)
	End Sub

	Private Sub RemoveActiveSheet()
		Dim index As Integer = MySheets.IndexOf(MyActiveSheet)
		Me.RemoveSheet(index)
	End Sub

	' Remove a worksheet
	Private Sub RemoveSheet(ByVal sheetIndex As Integer)
		Dim sheet As Worksheet = MySheets.Item(sheetIndex)
		' Unregister the sheet with the formula engine
		Me.Engine.Sheets.Remove(sheet)
		MySheets.RemoveAt(sheetIndex)

		Dim tp As TabPage = Me.tabSheets.TabPages.Item(sheetIndex)
		Me.tabSheets.TabPages.RemoveAt(sheetIndex)

		sheet.Dispose()
		Me.components.Remove(sheet)
		tp.Dispose()
	End Sub

	Private Sub miFillDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FillDownToolStripMenuItem.Click
		MyActiveSheet.DoFillDown()
	End Sub

	Private Sub miFillRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FillRightToolStripMenuItem.Click
		MyActiveSheet.DoFillRight()
	End Sub

	Private Sub miClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
		MyActiveSheet.ClearSelectedCells()
	End Sub

	Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
		MyActiveSheet.DoCopy()
	End Sub

	Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
		MyActiveSheet.DoCut()
	End Sub

	Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
		MyActiveSheet.DoPaste()
	End Sub

	Private Sub ChartToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChartToolStripMenuItem.Click
		Me.SetupChartSampleData()
		Dim dialog As New ChartForm()
		AddHandler dialog.Closed, AddressOf OnChartFormClosed
		Me.components.Add(dialog)
		dialog.Show(Me)
	End Sub

	Private Sub OnChartFormClosed(ByVal sender As Object, ByVal e As EventArgs)
		Dim dialog As ChartForm = sender
		RemoveHandler dialog.Closed, AddressOf OnChartFormClosed
		Me.components.Remove(dialog)
	End Sub

	' Write some sample data to the active sheet for the chart demo
	Private Sub SetupChartSampleData()
		MyActiveSheet.ClearRange(MyActiveSheet.GridRange)
		Dim g As DataGridView = MyActiveSheet.Grid

		g.Item(0, 1).Value = "Canada"
		g.Item(0, 2).Value = "USA"
		g.Item(0, 3).Value = "China"

		g.Item(1, 0).Value = "Jan"
		g.Item(2, 0).Value = "Feb"

		g.Item(1, 1).Value = 500
		g.Item(1, 2).Value = 90
		g.Item(1, 3).Value = 300

		g.Item(2, 1).Value = 220
		g.Item(2, 2).Value = 180
		g.Item(2, 3).Value = 390
	End Sub

	Private Sub SetSheetCount(ByVal count As Integer)
		Dim currentCount As Integer = MySheets.Count

		If currentCount < count Then
			For i As Integer = 0 To count - currentCount - 1
				Me.AddSheet()
			Next
		ElseIf count < currentCount Then
			For i As Integer = currentCount - 1 To count Step -1
				Me.RemoveSheet(i)
			Next
		End If
	End Sub

	' Called by the engine when circular references are detected
	Private Sub OnCircularReferenceDetected(ByVal sender As Object, ByVal e As CircularReferenceDetectedEventArgs)
		' Simply inform the user of the circular references
		Dim arr(2 - 1) As String
		arr(0) = "The following references have circular references:"
		arr(1) = Me.ReferenceList2String(e.Roots)
		Dim msg As String = String.Join(System.Environment.NewLine, arr)
		Me.ShowMessage(msg, MessageType.Error)
	End Sub

	Private Function ReferenceList2String(ByVal refs As IReference()) As String
		Dim arr(refs.Length - 1) As String

		For i As Integer = 0 To refs.Length - 1
			arr(i) = refs(i).ToString()
		Next

		Return String.Join(", ", arr)
	End Function

	' Creates a formula with error handling
	Public Function CreateFormula(ByVal expression As String) As ciloci.FormulaEngine.Formula
		Try
			Return Me.Engine.CreateFormula(expression)
		Catch ex As ciloci.FormulaEngine.InvalidFormulaException
			' This is the only exception that the CreateFormula method should throw
			Me.ShowFormulaCreateError(ex)
			Return Nothing
		End Try
	End Function

	Private Sub ShowFormulaCreateError(ByVal ex As ciloci.FormulaEngine.InvalidFormulaException)
		Dim arr(3 - 1) As String
		arr(0) = "The entered formula is not valid"
		arr(1) = "Details:"
		' The InvalidFormulaException will always have an inner exception which will hold the specifics
		arr(2) = ex.InnerException.Message
		Dim msg As String = String.Join(System.Environment.NewLine, arr)
		Me.ShowMessage(msg, MessageType.Error)
	End Sub

	' Reset our state to one empty sheet
	Private Sub mnuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
		Dim dcs As DocumentStatusService = MyServices.GetService(GetType(DocumentStatusService))
		' Let everyone know we are resetting
		dcs.RaiseNewDocument()
		' We only want one sheet
		Me.SetSheetCount(1)
		Me.SetActiveSheet(MySheets.Item(0))
		MyActiveSheet.SizeToDefault()
		MyActiveSheet.ClearRange(MyActiveSheet.GridRange)
		' Remove everything from the formula engine (except functions)
		Me.Engine.Clear()
		' We have to re-add our first sheet
		Me.Engine.Sheets.Add(MyActiveSheet)
		Me.SetActiveStream(Nothing)
	End Sub

	Private Sub mnuSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
		Me.SaveAs()
	End Sub

	Private Sub mnuSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
		If MyActiveStream Is Nothing Then
			Me.SaveAs()
		Else
			Me.Save()
		End If
	End Sub

	' Load a previously saved workbook
	Private Sub mnuOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
		Dim ofd As New OpenFileDialog()
		ofd.Filter = "Workbook files (.wb)|*.wb"

		If ofd.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
			Dim dcs As DocumentStatusService = MyServices.GetService(GetType(DocumentStatusService))
			dcs.RaiseOpenDocument()
			' We have to create the stream ourselves since we need read/write access
			Dim instream As New System.IO.FileStream(ofd.FileName, IO.FileMode.Open, IO.FileAccess.ReadWrite, IO.FileShare.None)
			Me.LoadFile(instream)
			Me.SetActiveStream(instream)
		End If

		ofd.Dispose()
	End Sub

	' Set the active document we are working with
	Private Sub SetActiveStream(ByVal s As System.IO.FileStream)
		If Not MyActiveStream Is Nothing Then
			' Make sure to close the old one
			MyActiveStream.Close()
		End If

		MyActiveStream = s

		Dim activeFileName As String

		If MyActiveStream Is Nothing Then
			activeFileName = "Untitled"
		Else
			activeFileName = System.IO.Path.GetFileName(MyActiveStream.Name)
		End If

		Me.Text = String.Format("{0} - {1}", APP_NAME, activeFileName)
	End Sub

	Private Sub SaveAs()
		Dim sfd As New SaveFileDialog()
		sfd.Filter = "Workbook files (.wb)|*.wb"

		If sfd.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
			Dim outstream As System.IO.Stream = sfd.OpenFile()
			Me.SetActiveStream(outstream)
			Me.Save()
		End If

		sfd.Dispose()
	End Sub

	' Save our state to a file
	Private Sub Save()
		' Create our memento class
		Dim memento As New WorkbookMemento()
		' Populate it with data
		memento.Load(Me.GetSheets(), Me.Engine)
		Dim formatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
		MyActiveStream.Seek(0, IO.SeekOrigin.Begin)
		' ...and save
		formatter.Serialize(MyActiveStream, memento)
		MyActiveStream.SetLength(MyActiveStream.Position)

		Me.ShowMessage("Save successful", MessageType.Information)
	End Sub

	' Load a saved workbook
	Private Sub LoadFile(ByVal instream As System.IO.Stream)
		Dim formatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
		' Get our saved class
		Dim memento As WorkbookMemento = formatter.Deserialize(instream)

		Dim savedSheets As ISheet() = memento.Sheets
		' We want the same number of sheets as what was saved.  We have to be careful because
		' the deserialized engine already has all the sheets added so we have to add/remove
		' sheets on the OLD formula engine.
		Me.SetSheetCount(savedSheets.Length)
		Dim currentSheets As Worksheet() = Me.GetSheets()

		' Replace all our sheets with the saved ones
		For i As Integer = 0 To savedSheets.Length - 1
			Me.ReplaceSheet(i, currentSheets(i), savedSheets(i))
		Next

		' Write the saved data into our sheets
		memento.Store(Me.GetSheets())

		' Store the new formula engine instance
		Me.SetFormulaEngine(memento.Engine)

		Me.SetActiveSheet(MySheets.Item(0))

		' Get rid of old sheets
		For Each sheet As Worksheet In currentSheets
			sheet.Detach()
		Next
	End Sub

	Private Sub ReplaceSheet(ByVal index As Integer, ByVal currentSheet As Worksheet, ByVal savedSheet As Worksheet)
		Me.components.Remove(currentSheet)
		Me.components.Add(savedSheet)
		MySheets.Item(index) = savedSheet
		Me.tabSheets.TabPages.Item(index).Text = savedSheet.Name
		currentSheet.CopyTo(savedSheet)
	End Sub

	Private Function GetSheets() As Worksheet()
		Dim arr(MySheets.Count - 1) As Worksheet
		MySheets.CopyTo(arr, 0)
		Return arr
	End Function

	Friend Sub RaiseRefreshRequired()
		RaiseEvent RefreshRequired(Me, EventArgs.Empty)
	End Sub

	Friend ReadOnly Property ActiveSheet() As Worksheet
		Get
			Return MyActiveSheet
		End Get
	End Property

	Public ReadOnly Property Engine() As FormulaEngine
		Get
			Return MyServices.GetService(GetType(FormulaEngine))
		End Get
	End Property
End Class