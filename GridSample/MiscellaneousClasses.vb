' Various supporting classes

Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports ciloci.FormulaEngine

Friend Class CustomContainer
	Inherits Container

	Private MyServices As IServiceProvider

	Public Sub New(ByVal services As IServiceProvider)
		MyServices = services
	End Sub

	Protected Overrides Function GetService(ByVal service As System.Type) As Object
		Return MyServices.GetService(service)
	End Function
End Class

' Manages clipboard operations
Friend Class ClipboardManager
	Inherits System.ComponentModel.Component
	Implements IInitialize

	Private MyData As SheetData
	Private MySourceRange As Rectangle
	Private MyIsCut As Boolean
	Private MySourceSheet As Worksheet
	Private WithEvents MySourceGrid As DataGridView
	Private WithEvents MyDocumentStatusService As DocumentStatusService

	Public Sub New()

	End Sub

	Public Sub Initialize() Implements IInitialize.Initialize
		MyDocumentStatusService = Me.Site.GetService(GetType(DocumentStatusService))
	End Sub

	Private Sub MyDocumentStatusService_NewOrOpenDocument(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyDocumentStatusService.NewDocument, MyDocumentStatusService.OpenDocument
		Me.DoCancel()
	End Sub

	Private Sub MySourceGrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MySourceGrid.KeyDown
		If e.KeyCode = Keys.Escape Then
			e.Handled = True
			Me.DoCancel()
		End If
	End Sub

	' Draws an indicator of the source range of a clipboard operation
	Private Sub MySourceGrid_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MySourceGrid.Paint
		Dim rect As Rectangle = MySourceSheet.GetRangeRectangle(Me.ClippedSourceRange)
		Dim b As New SolidBrush(Color.FromArgb(128, 100, 220, 100))
		e.Graphics.FillRectangle(b, rect)
		b.Dispose()
	End Sub

	Public Sub DoCopy(ByVal sheet As Worksheet, ByVal range As Rectangle)
		' Make sure we cancel if we are already in an operation
		If Not MyData Is Nothing Then
			Me.DoCancel()
		End If

		' Load the copy data
		MyData = New SheetData()
		MyData.LoadData(sheet, range)
		' Note the source range, sheet, and grid
		MySourceRange = range
		MySourceSheet = sheet
		MySourceGrid = sheet.Grid
		' Repaint the copied range
		MySourceSheet.RefreshRange(MySourceRange)
		' Remember that we did a copy
		MyIsCut = False
	End Sub

	Public Sub DoCut(ByVal sheet As Worksheet, ByVal range As Rectangle)
		' Do the same thing as a copy but not that we did a cut
		Me.DoCopy(sheet, range)
		MyIsCut = True
	End Sub

	Public Sub DoPaste(ByVal sheet As Worksheet, ByVal pos As Point)
		If MyData Is Nothing Then
			' If we have no internal data then try to do a paste of data on the clipboard
			Me.PasteCommonFormat(sheet, pos)
			Return
		End If

		Dim success As Boolean
		Dim destRange As Rectangle = Me.GetDestRange(sheet, pos)

		' We don't allow pastes that won't fit on the sheet
		If sheet.GridContainsRange(destRange) = False Then
			Dim ms As MessageService = Me.Site.GetService(GetType(MessageService))
			ms.ShowMessage("Destination range is not contained in sheet", MessageType.Error)
			Return
		End If

		If MyIsCut = True Then
			success = Me.DoPasteFromCut(sheet, destRange)
		Else
			success = Me.DoPasteFromCopy(sheet, destRange)
		End If

		If success = True Then
			Me.DoCancel()
		End If
	End Sub

	Private Function GetDestRange(ByVal sheet As Worksheet, ByVal pos As Point) As Rectangle
		Dim rt As RangeType = sheet.GetRangeType(MySourceRange)
		If rt = RangeType.Cells Then
			Return New Rectangle(pos.X, pos.Y, MySourceRange.Width, MySourceRange.Height)
		ElseIf rt = RangeType.Columns Then
			Return New Rectangle(pos.X, 0, MySourceRange.Width, MySourceRange.Height)
		ElseIf rt = RangeType.Rows Then
			Return New Rectangle(0, pos.Y, MySourceRange.Width, MySourceRange.Height)
		End If
	End Function

	Private Function DoPasteFromCut(ByVal sheet As Worksheet, ByVal range As Rectangle) As Boolean
		sheet.DoCutPaste(MyData, range)
		Return True
	End Function

	Private Function DoPasteFromCopy(ByVal sheet As Worksheet, ByVal range As Rectangle) As Boolean
		sheet.DoCopyPaste(MyData, range)
		Return True
	End Function

	' Paste a common clipboard format
	Private Sub PasteCommonFormat(ByVal sheet As Worksheet, ByVal pos As Point)
		Dim dobj As IDataObject = Clipboard.GetDataObject()
		If dobj.GetDataPresent(DataFormats.CommaSeparatedValue) = True Then
			Me.DoCSVPaste(sheet, pos, dobj)
		ElseIf Clipboard.ContainsText() = True Then
			Me.DoPlainTextPaste(sheet, pos)
		End If
	End Sub

	Private Sub DoPlainTextPaste(ByVal sheet As Worksheet, ByVal pos As Point)
		Dim text As String = Clipboard.GetText()
		Dim cell As ExcelLikeCell = sheet.Grid.Item(pos.X, pos.Y)
		cell.SetTextValue(text)
	End Sub

	Private Sub DoCSVPaste(ByVal sheet As Worksheet, ByVal pos As Point, ByVal dobj As IDataObject)
		Dim csvData As String()() = Me.GetCSVData(dobj)
		Me.PasteCSVData(sheet, pos, csvData)
	End Sub

	' Get an array of strings from the raw clipboard CSV data
	Private Function GetCSVData(ByVal dobj As IDataObject) As String()()
		Dim data As System.IO.Stream = dobj.GetData(DataFormats.CommaSeparatedValue, False)
		Dim sr As New System.IO.StreamReader(data, System.Text.Encoding.ASCII)

		' We will let the regex engine do the real work
		Const FIELD_PATTERN As String = "(?<field>""(""""|[^""])*""|[^,]*)"
		Dim csvRegex As New System.Text.RegularExpressions.Regex(String.Format("^{0}(,{0})*$", FIELD_PATTERN), System.Text.RegularExpressions.RegexOptions.ExplicitCapture)

		Dim lines As IList = New ArrayList
		Dim line As String = sr.ReadLine()

		While Not line Is Nothing
			Dim values As String() = Me.ParseCSVLine(line, csvRegex)
			lines.Add(values)
			line = sr.ReadLine()
		End While

		sr.Close()

		' Get rid of final blank line
		lines.RemoveAt(lines.Count - 1)
		Dim lineArr(lines.Count - 1)() As String
		lines.CopyTo(lineArr, 0)

		Return lineArr
	End Function

	' Gets the individual fields of a CSV line
	Private Function ParseCSVLine(ByVal line As String, ByVal regEx As System.Text.RegularExpressions.Regex) As String()
		Dim match As System.Text.RegularExpressions.Match = regEx.Match(line)
		Dim group As System.Text.RegularExpressions.Group = match.Groups.Item("field")

		Dim arr(group.Captures.Count - 1) As String

		For i As Integer = 0 To group.Captures.Count - 1
			Dim value As String = group.Captures.Item(i).Value
			If value.IndexOf(","c) <> -1 Then
				' Values with commas in them will be in double quotes
				' Remove outer double quotes
				value = value.Substring(1, value.Length - 2)
				' Unescape double quotes
				value = value.Replace("""""", """")
			End If
			arr(i) = value
		Next

		Return arr
	End Function

	' Go through each field in the parsed CSV data and try to put it into a cell
	Private Sub PasteCSVData(ByVal sheet As Worksheet, ByVal pos As Point, ByVal data As String()())
		For row As Integer = 0 To data.Length - 1
			Dim line As String() = data(row)
			For col As Integer = 0 To line.Length - 1
				Dim text As String = line(col)
				'We treat an empty string as an empty cell
				If text.Length = 0 Then
					text = Nothing
				End If

				' Compute the actual co-ordinates
				Dim realRow As Integer = row + pos.Y
				Dim realCol As Integer = col + pos.X

				' The real point may not be on the sheet
				If sheet.GridContainsPoint(realRow, realCol) = True Then
					' Let the cell parse the text into the final value
					Dim cell As ExcelLikeCell = sheet.Grid.Item(realCol, realRow)
					cell.SetTextValue(text)
				End If
			Next
		Next
	End Sub

	' Cancel any clipboard operation
	Private Sub DoCancel()
		If Not MyData Is Nothing Then
			MyData = Nothing
			MySourceGrid = Nothing
			' Stop drawing our marker
			If MySourceSheet.Grid.Visible = True Then
				MySourceSheet.RefreshRange(Me.ClippedSourceRange)
			End If
			MySourceRange = Rectangle.Empty
			MySourceSheet = Nothing
		End If
	End Sub

	Private ReadOnly Property ClippedSourceRange() As Rectangle
		Get
			Return MySourceSheet.ClipToGrid(MySourceRange)
		End Get
	End Property

	Private ReadOnly Property Engine() As ciloci.FormulaEngine.FormulaEngine
		Get
			Return Me.Site.GetService(GetType(ciloci.FormulaEngine.FormulaEngine))
		End Get
	End Property
End Class

' A base class for operations that change cells on a sheet
Friend MustInherit Class CellTransform
	Public MustOverride Sub TransformCell(ByVal target As ExcelLikeCell, ByVal row As Integer, ByVal col As Integer)
End Class

' Clears a range on a worksheet
Friend Class ClearCellTransform
	Inherits CellTransform

	Public Overrides Sub TransformCell(ByVal target As ExcelLikeCell, ByVal row As Integer, ByVal col As Integer)
		target.Value = Nothing
	End Sub
End Class

' Responsible for showing messages to the user
Friend Class MessageService

	Private MyActiveDialog As Form

	Public Sub New()

	End Sub

	Public Sub SetActiveDialog(ByVal dialog As Form)
		MyActiveDialog = dialog
	End Sub

	Public Sub ShowMessage(ByVal msg As String, ByVal mt As MessageType)
		MessageBox.Show(MyActiveDialog, msg, MainForm.APP_NAME, MessageBoxButtons.OK, Me.GetMessageBoxIcon(mt))
	End Sub

	Private Function GetMessageBoxIcon(ByVal type As MessageType) As MessageBoxIcon
		Select Case type
			Case MessageType.Information
				Return MessageBoxIcon.Information
			Case MessageType.Error
				Return MessageBoxIcon.Warning
			Case Else
				Debug.Assert(False, "unknown type")
		End Select
	End Function
End Class

' Stores all required data when we save or load a workbook
<Serializable()> _
Friend Class WorkbookMemento
	Implements System.Runtime.Serialization.ISerializable

	Private MySheets As ISheet()				' The set of sheets in the workbook
	Private MyData As SheetData()				' The contents of each sheet
	Private MyFormulaEngine As FormulaEngine	' The formula engine
	Private Const VERSION As Integer = 1

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MySheets = info.GetValue("Sheets", GetType(ISheet()))
		MyData = info.GetValue("SheetsData", GetType(SheetData()))
		MyFormulaEngine = info.GetValue("FormulaEngine", GetType(FormulaEngine))
	End Sub

	Public Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
		info.AddValue("Sheets", MySheets)
		info.AddValue("SheetsData", MyData)
		info.AddValue("FormulaEngine", MyFormulaEngine)
	End Sub

	' Loads all the data we need for a save
	Public Sub Load(ByVal sheets As Worksheet(), ByVal engine As FormulaEngine)
		MySheets = New ISheet(sheets.Length - 1) {}
		MyData = New SheetData(sheets.Length - 1) {}

		For i As Integer = 0 To MySheets.Length - 1
			Dim sheet As Worksheet = sheets(i)
			MySheets(i) = sheet
			Dim gd As New SheetData
			gd.LoadData(sheet, sheet.GridRange)
			MyData(i) = gd
		Next

		MyFormulaEngine = engine
	End Sub

	' Transfers saved data into a set of worksheets
	Public Sub Store(ByVal sheets As Worksheet())
		For i As Integer = 0 To sheets.Length - 1
			Dim sheet As Worksheet = sheets(i)
			Dim gd As SheetData = MyData(i)
			' Make the sheet the same size as the saved one
			sheet.SetSize(gd.DataSize)
			' Write the saved data into the sheet
			gd.WriteData(sheet, sheet.GridRange)
		Next
	End Sub

	Public ReadOnly Property Engine() As FormulaEngine
		Get
			Return MyFormulaEngine
		End Get
	End Property

	Public ReadOnly Property Sheets() As ISheet()
		Get
			Return MySheets
		End Get
	End Property
End Class

' Implements a simple formula bar
Friend Class FormulaBar
	Inherits System.ComponentModel.Component
	Implements IInitialize

	Private WithEvents MyGrid As DataGridView
	Private WithEvents MyOwner As MainForm
	Private WithEvents MyTextbox As TextBox

	Public Sub Initialize() Implements IInitialize.Initialize
		MyOwner = Me.Site.GetService(GetType(MainForm))
		MyTextbox = MyOwner.edFormulaBar
	End Sub

	' Update our text when the active sheet changes
	Private Sub MyOwner_ActiveSheetChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyOwner.ActiveSheetChanged
		Dim sheet As Worksheet = Me.Site.GetService(GetType(Worksheet))
		MyGrid = sheet.Grid
		Me.DisplayCurrentCell()
	End Sub

	Private Sub MyOwner_RefreshRequired(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyOwner.RefreshRequired
		Me.DisplayCurrentCell()
	End Sub

	Private Sub MyGrid_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles MyGrid.CellValueChanged
		Dim curCellPos As Point = MyGrid.CurrentCellAddress
		Dim changedPos As New Point(e.ColumnIndex, e.RowIndex)
		If curCellPos = changedPos Then
			Me.DisplayCurrentCell()
		End If
	End Sub

	Private Sub MyGrid_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyGrid.CurrentCellChanged
		Me.DisplayCurrentCell()
	End Sub

	Private Sub DisplayCurrentCell()
		Dim p As Point = MyGrid.CurrentCellAddress
		If p.X = -1 Or p.Y = -1 Then
			Return
		End If
		Me.Display(p.Y + 1, p.X + 1)
	End Sub

	Private Sub Display(ByVal row As Integer, ByVal col As Integer)
		Dim engine As FormulaEngine = Me.Site.GetService(GetType(FormulaEngine))
		' Get a reference to the row and column
		Dim ref As ISheetReference = engine.ReferenceFactory.Cell(row, col)
		' Query the engine if there is a formula at that reference
		Dim f As Formula = engine.GetFormulaAt(ref)

		Dim text As String

		If Not f Is Nothing Then
			' There is a formula there; display its text
			text = f.ToString()
		Else
			' No formula exists at that reference so get the text of the cell's value
			Dim value As Object = MyGrid.CurrentCell.Value
			If value Is Nothing Then
				value = String.Empty
			End If
			text = value.ToString()
		End If

		MyOwner.edFormulaBar.Text = text
	End Sub

	Private Sub HandleEscape(ByVal e As KeyEventArgs)
		Me.DisplayCurrentCell()
		MyGrid.Focus()
	End Sub

	Private Sub HandleEnter(ByVal e As KeyEventArgs)
		Dim text As String = MyTextbox.Text
		Dim cell As ExcelLikeCell = MyGrid.CurrentCell

		Dim success As Boolean = cell.SetTextValue(text)

		If success = True Then
			MyGrid.Focus()
			Me.DisplayCurrentCell()
		End If
	End Sub

	Private Sub MyTextbox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyTextbox.KeyDown
		If e.KeyCode = Keys.Escape Then
			Me.HandleEscape(e)
		ElseIf e.KeyCode = Keys.Enter Then
			Me.HandleEnter(e)
		End If
	End Sub
End Class

' Manages the selection of entire rows since the datagridview can't do it for us.
Friend Class RowSelectionManager

	Private WithEvents MyGrid As DataGridView
	Private MySheet As Worksheet
	Private MyMouseDown As Boolean
	Private MyMouseDownRow As Integer

	Public Sub New(ByVal sheet As Worksheet)
		MySheet = sheet
		MyGrid = sheet.Grid
	End Sub

	Private Sub MyGrid_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyGrid.MouseDown
		If e.Button <> MouseButtons.Left Then
			Return
		End If
		MyMouseDown = True

		Dim rowIndex As Integer = Me.GetRowIndex(e.Location)

		If rowIndex = -1 Then
			Return
		End If

		MyMouseDownRow = rowIndex
		Me.SelectRows(rowIndex)
	End Sub

	Private Sub MyGrid_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyGrid.MouseMove
		If MyMouseDown = False Then
			Return
		End If

		Dim rowIndex As Integer = Me.GetRowIndex(e.Location)

		If rowIndex = -1 Then
			Return
		End If

		Me.SelectRows(rowIndex)
	End Sub

	Private Function GetRowIndex(ByVal pos As Point) As Integer
		Dim hti As DataGridView.HitTestInfo = MyGrid.HitTest(pos.X, pos.Y)
		If hti.Type = DataGridViewHitTestType.RowHeader Then
			Return hti.RowIndex
		Else
			Return -1
		End If
	End Function

	Private Sub MyGrid_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyGrid.MouseUp
		MyMouseDown = False
	End Sub

	Private Sub SelectRows(ByVal endRow As Integer)
		MyGrid.ClearSelection()
		If endRow < MyMouseDownRow Then
			Dim i As Integer = 0
		End If
		Dim startRow As Integer = System.Math.Min(MyMouseDownRow, endRow)
		endRow = System.Math.Max(endRow, MyMouseDownRow)
		For row As Integer = startRow To endRow
			MySheet.SelectRow(row)
		Next
	End Sub

	Public Sub Dispose()
		MyGrid = Nothing
	End Sub
End Class

' Interface for components that need to initialize after being added to a container
Friend Interface IInitialize
	Sub Initialize()
End Interface

Friend Enum RangeType
	Cells
	Rows
	Columns
End Enum

' Base class for performing fills on a worksheet
Friend MustInherit Class FillExecutor

	Protected MyPrimaryOffset As Integer
	Protected MyTargetRange As Rectangle
	Protected MyEngine As FormulaEngine
	Protected MySourceFormula As Formula

	Public Sub Initialize(ByVal targetRange As Rectangle, ByVal engine As FormulaEngine)
		MyTargetRange = targetRange
		MyEngine = engine
	End Sub

	Protected MustOverride Function GetSourcePoint() As Point
	Protected MustOverride Function GetDestPoint(ByVal offset As Integer) As Point

	Public Sub SetPrimaryOffset(ByVal offset As Integer)
		MyPrimaryOffset = offset
		Dim sourcePoint As Point = Me.GetSourcePoint()
		Dim sourceRef As ISheetReference = Me.GetPointReference(sourcePoint)
		MySourceFormula = MyEngine.GetFormulaAt(sourceRef)
	End Sub

	Public MustOverride Function GetFillCount() As Integer
	Public MustOverride Function GetFillRange() As Rectangle
	Public MustOverride Function IsValidFill() As Boolean

	Private Function GetPointReference(ByVal p As Point) As ISheetReference
		Return MyEngine.ReferenceFactory.Cell(p.Y + 1, p.X + 1)
	End Function

	' Do a fill on a cell
	Public Sub DoFill(ByVal offset As Integer, ByVal grid As DataGridView)
		Dim destPoint As Point = Me.GetDestPoint(offset)
		Dim sourcePoint As Point = Me.GetSourcePoint()

		If MySourceFormula Is Nothing Then
			' We are not working with a formula so just copy the source cell's value into the destination cell
			grid.Item(destPoint.X, destPoint.Y).Value = grid.Item(sourcePoint.X, sourcePoint.Y).Value
		Else
			' We are working with a formula; ask the engine to copy the source formula into the destination cell
			Dim destRef As ISheetReference = Me.GetPointReference(destPoint)
			MyEngine.CopySheetFormula(MySourceFormula, destRef)
		End If
	End Sub
End Class

' Implements a fill down
Friend Class FillDownExecutor
	Inherits FillExecutor

	Public Overrides Function GetFillCount() As Integer
		Return MyTargetRange.Height - 1
	End Function

	Public Overrides Function GetFillRange() As System.Drawing.Rectangle
		Return Rectangle.FromLTRB(MyTargetRange.Left, MyTargetRange.Top + 1, MyTargetRange.Right, MyTargetRange.Bottom)
	End Function

	Protected Overrides Function GetDestPoint(ByVal offset As Integer) As System.Drawing.Point
		Return New Point(MyTargetRange.Left + MyPrimaryOffset, MyTargetRange.Top + offset + 1)
	End Function

	Protected Overrides Function GetSourcePoint() As System.Drawing.Point
		Return New Point(MyTargetRange.Left + MyPrimaryOffset, MyTargetRange.Top)
	End Function

	Public Overrides Function IsValidFill() As Boolean
		Return MyTargetRange.Height > 1
	End Function
End Class

' Implements a fill right
Friend Class FillRightExecutor
	Inherits FillExecutor

	Public Overrides Function GetFillCount() As Integer
		Return MyTargetRange.Width - 1
	End Function

	Public Overrides Function GetFillRange() As System.Drawing.Rectangle
		Return Rectangle.FromLTRB(MyTargetRange.Left + 1, MyTargetRange.Top, MyTargetRange.Right, MyTargetRange.Bottom)
	End Function

	Protected Overrides Function GetSourcePoint() As System.Drawing.Point
		Return New Point(MyTargetRange.Left, MyTargetRange.Top + MyPrimaryOffset)
	End Function

	Protected Overrides Function GetDestPoint(ByVal offset As Integer) As System.Drawing.Point
		Return New Point(MyTargetRange.Left + offset + 1, MyTargetRange.Top + MyPrimaryOffset)
	End Function

	Public Overrides Function IsValidFill() As Boolean
		Return MyTargetRange.Width > 1
	End Function
End Class

' Simple class for storing state
Friend Class DictionaryService
	Private MyMap As IDictionary

	Public Sub New()
		MyMap = New Hashtable
	End Sub

	Public Sub SetValue(ByVal key As String, ByVal value As Object)
		MyMap.Add(key, value)
	End Sub

	Public Function GetValue(ByVal key As String) As Object
		Return MyMap.Item(key)
	End Function

	Public Sub RemoveValue(ByVal key As String)
		MyMap.Remove(key)
	End Sub
End Class

Friend Enum MessageType
	Information
	[Error]
End Enum

' Notifies any listeners about events related to document status
Friend Class DocumentStatusService
	Public Event NewDocument As EventHandler
	Public Event OpenDocument As EventHandler

	Public Sub RaiseNewDocument()
		RaiseEvent NewDocument(Me, EventArgs.Empty)
	End Sub

	Public Sub RaiseOpenDocument()
		RaiseEvent OpenDocument(Me, EventArgs.Empty)
	End Sub
End Class