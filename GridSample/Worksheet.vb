Imports ciloci.FormulaEngine

' This class represents a worksheet.  It uses a datagridview as its grid and implements the formulaengine ISheet interface
' so that it can be used in formulas.
<Serializable()> _
Friend Class Worksheet
	Inherits System.ComponentModel.Component
	Implements ciloci.FormulaEngine.ISheet
	Implements System.Runtime.Serialization.ISerializable

	Private WithEvents MyGrid As DataGridView
	Private MyCellMenu As ContextMenuStrip
	Private MyColumnMenu As ContextMenuStrip
	Private MyRowMenu As ContextMenuStrip
	Private MyRowSelectionManager As RowSelectionManager
	Private MyName As String

	Public Sub New(ByVal name As String)
		MyName = name
		MyGrid = New DataGridView()
		Me.CustomizeGrid()
		Me.Initialize()
	End Sub

	Public Sub CopyTo(ByVal dest As Worksheet)
		dest.MyGrid = MyGrid
		dest.Initialize()
	End Sub

	Public Sub Detach()
		MyGrid = Nothing
		MyRowSelectionManager.Dispose()
	End Sub

	Private Sub Initialize()
		MyCellMenu = Me.CreateCellMenu()
		MyColumnMenu = Me.CreateColumnMenu()
		MyRowMenu = Me.CreateRowMenu()
		MyRowSelectionManager = New RowSelectionManager(Me)
	End Sub

	Public Sub SelectFirstCell()
		MyGrid.Item(0, 0).Selected = True
	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyName = info.GetString("Name")
	End Sub

	Public Overridable Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Name", MyName)
	End Sub

	Private Function CreateCellMenu() As ContextMenuStrip
		Dim cms As New ContextMenuStrip()

		cms.Items.Add(New ToolStripMenuItem("Cut", My.Resources.Images.Cut, AddressOf OnCutMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Copy", My.Resources.Images.Copy, AddressOf OnCopyMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Paste", My.Resources.Images.Paste, AddressOf OnPasteMenuItemClick))
		cms.Items.Add(New ToolStripSeparator())
		cms.Items.Add(New ToolStripMenuItem("Fill Down", My.Resources.Images.FillDownHS, AddressOf OnFillDownMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Fill Right", My.Resources.Images.FillRightHS, AddressOf OnFillRightMenuItemClick))
		cms.Items.Add(New ToolStripSeparator())
		cms.Items.Add(New ToolStripMenuItem("Clear", DirectCast(Nothing, Image), AddressOf OnClearMenuItemClick))
		Return cms
	End Function

	Private Function CreateColumnMenu() As ContextMenuStrip
		Dim cms As New ContextMenuStrip()

		cms.Items.Add(New ToolStripMenuItem("Cut", My.Resources.Images.Cut, AddressOf OnCutMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Copy", My.Resources.Images.Copy, AddressOf OnCopyMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Paste", My.Resources.Images.Paste, AddressOf OnPasteMenuItemClick))
		cms.Items.Add(New ToolStripSeparator())
		cms.Items.Add(New ToolStripMenuItem("Insert", DirectCast(Nothing, Image), AddressOf OnInsertColumnMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Delete", DirectCast(Nothing, Image), AddressOf OnRemoveColumnMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Clear", DirectCast(Nothing, Image), AddressOf OnClearMenuItemClick))
		Return cms
	End Function

	Private Function CreateRowMenu() As ContextMenuStrip
		Dim cms As New ContextMenuStrip()

		cms.Items.Add(New ToolStripMenuItem("Cut", My.Resources.Images.Cut, AddressOf OnCutMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Copy", My.Resources.Images.Copy, AddressOf OnCopyMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Paste", My.Resources.Images.Paste, AddressOf OnPasteMenuItemClick))
		cms.Items.Add(New ToolStripSeparator())
		cms.Items.Add(New ToolStripMenuItem("Insert", DirectCast(Nothing, Image), AddressOf OnInsertRowMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Delete", DirectCast(Nothing, Image), AddressOf OnRemoveRowMenuItemClick))
		cms.Items.Add(New ToolStripMenuItem("Clear", DirectCast(Nothing, Image), AddressOf OnClearMenuItemClick))
		Return cms
	End Function

	Private Sub CustomizeGrid()
		MyGrid.AllowUserToAddRows = False
		MyGrid.AllowUserToDeleteRows = False
		MyGrid.AllowUserToOrderColumns = False
		MyGrid.Dock = DockStyle.Fill
		MyGrid.ColumnHeadersHeight = 17
		MyGrid.RowHeadersWidth = 26
		MyGrid.ShowEditingIcon = False
		MyGrid.ShowCellToolTips = False
	End Sub

	Public Sub SetSize(ByVal s As Size)
		MyGrid.SelectionMode = DataGridViewSelectionMode.CellSelect
		MyGrid.ColumnCount = s.Width
		Me.ProcessColumns()
		MyGrid.Rows.Clear()
		MyGrid.RowCount = s.Height
		Me.ProcessAllRows()
		MyGrid.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect
	End Sub

	Private Sub ProcessColumns()
		For i As Integer = 0 To MyGrid.ColumnCount - 1
			Dim col As DataGridViewColumn = MyGrid.Columns.Item(i)
			Me.ProcessColumn(col, i)
		Next
	End Sub

	Private Sub ProcessColumn(ByVal col As DataGridViewColumn, ByVal ordinal As Integer)
		Dim cellTemplate As ExcelLikeCell = Me.Site.GetService(GetType(ExcelLikeCell))
		col.CellTemplate = cellTemplate
		col.Name = Chr(ordinal + 65)
		col.Width = 64
		col.SortMode = DataGridViewColumnSortMode.NotSortable
	End Sub

	Private Sub ProcessAllRows()
		Me.ProcessRows(0, MyGrid.RowCount)
	End Sub

	Private Sub ProcessRows(ByVal start As Integer, ByVal count As Integer)
		For i As Integer = start To start + count - 1
			Dim row As DataGridViewRow = MyGrid.Rows.Item(i)
			row.Height = 17
		Next
	End Sub

	Public Sub SetGridParent(ByVal parent As Control)
		MyGrid.Parent = parent
	End Sub

	' Clear all cells in the current selection
	Public Sub ClearSelectedCells()
		Dim selected As Rectangle = Me.GetSelectedRange()
		If selected.IsEmpty = True Then
			Return
		End If

		' Clear the values of the cells
		Me.ClearRange(selected)
		' Get a reference to the selected range
		Dim ref As ISheetReference = Me.GetRectangleReference(selected)
		' Clear all formulas in the range
		Me.Engine.RemoveFormulasInRange(ref)
		' Recalculate any dependents of the range
		Me.Engine.Recalculate(ref)
		Me.RaiseRefreshRequired()
	End Sub

	' Clear the values of all cells in a range
	Public Sub ClearRange(ByVal range As Rectangle)
		Me.Transform(range, New ClearCellTransform())
	End Sub

	' Perform an Excel-like "Fill"
	Private Sub DoFill(ByVal range As Rectangle, ByVal primarySize As Integer, ByVal executor As FillExecutor)
		executor.Initialize(range, Me.Engine)

		If executor.IsValidFill() = False Then
			Return
		End If

		' Get the range that we will fill
		Dim fillRange As Rectangle = executor.GetFillRange()
		' Get an equivalent sheet reference to the range
		Dim fillRef As ISheetReference = Me.GetRectangleReference(fillRange)
		' Get rid of any formulas in the range we will fill
		Me.Engine.RemoveFormulasInRange(fillRef)

		' Do the fill
		For primaryIndex As Integer = 0 To primarySize - 1
			executor.SetPrimaryOffset(primaryIndex)
			For secondaryIndex As Integer = 0 To executor.GetFillCount() - 1
				executor.DoFill(secondaryIndex, MyGrid)
			Next
		Next

		' Recalculate the filled range
		Me.Engine.Recalculate(fillRef)
	End Sub

	Public Sub DoFillDown()
		Dim targetRange As Rectangle = Me.GetSelectedRange()
		Me.DoFill(targetRange, targetRange.Width, New FillDownExecutor())
	End Sub

	Public Sub DoFillRight()
		Dim targetRange As Rectangle = Me.GetSelectedRange()
		Me.DoFill(targetRange, targetRange.Height, New FillRightExecutor())
	End Sub

	' Get a rectangle representing the currently selected range.  We have to do this manually since
	' grid doesn't provide such a method
	Private Function GetSelectedRange() As Rectangle
		Dim rect As Rectangle

		If MyGrid.SelectedCells.Count > 0 Then
			' At least once cell is selected; get its rectangle
			Dim first As DataGridViewCell = MyGrid.SelectedCells.Item(0)
			rect = New Rectangle(first.ColumnIndex, first.RowIndex, 1, 1)
		Else
			Return Rectangle.Empty
		End If

		' Go through all selected cells and build up a rectangle
		For i As Integer = 1 To MyGrid.SelectedCells.Count - 1
			Dim cell As DataGridViewCell = MyGrid.SelectedCells.Item(i)
			rect = Rectangle.Union(rect, New Rectangle(cell.ColumnIndex, cell.RowIndex, 1, 1))
		Next

		Return rect
	End Function

	' Get a sheet reference from a rectangle range
	Private Function GetRectangleReference(ByVal rect As Rectangle) As ISheetReference
		Return Me.GetRectangleReference(Me, rect)
	End Function

	' Get a sheet reference from a sheet and rectangle range
	Private Function GetRectangleReference(ByVal sheet As Worksheet, ByVal rect As Rectangle) As ISheetReference
		' Remember, the engine's rectangles start at (1,1)
		rect.Offset(1, 1)
		Return Me.Engine.ReferenceFactory.FromRectangle(sheet, rect)
	End Function

	' Get a sheet reference to a cell on a worksheet
	Private Function GetCellReference(ByVal sheet As Worksheet, ByVal row As Integer, ByVal col As Integer) As ISheetReference
		' Remember, the top-left cell is (1,1)
		Return Me.Engine.ReferenceFactory.Cell(sheet, row + 1, col + 1)
	End Function

	' Get a sheet reference to a cell on the current work sheet
	Private Function GetCellReference(ByVal row As Integer, ByVal col As Integer) As ISheetReference
		Return Me.GetCellReference(Me, row, col)
	End Function

	Private Sub HandleRightClick(ByVal rowIndex As Integer, ByVal colIndex As Integer)
		If rowIndex = -1 Then
			Me.HandleColumnRightClick(colIndex)
		ElseIf colIndex = -1 Then
			Me.ShowContextMenu(MyRowMenu)
		Else
			Me.HandleCellRightClick(rowIndex, colIndex)
		End If
	End Sub

	Private Sub HandleColumnRightClick(ByVal colIndex As Integer)
		Dim col As DataGridViewColumn = MyGrid.Columns.Item(colIndex)
		If col.Selected = False Then
			MyGrid.ClearSelection()
			col.Selected = True
		End If

		Me.ShowContextMenu(MyColumnMenu)
	End Sub

	Private Sub HandleCellRightClick(ByVal rowIndex As Integer, ByVal colIndex As Integer)
		Dim cell As DataGridViewCell = MyGrid.Item(colIndex, rowIndex)
		If cell.Selected = False Then
			MyGrid.ClearSelection()
			cell.Selected = True
			MyGrid.CurrentCell = cell
		End If

		Me.ShowContextMenu(MyCellMenu)
	End Sub

	Private Sub ShowContextMenu(ByVal mnu As ContextMenuStrip)
		Dim pos As Point = Control.MousePosition
		pos = MyGrid.PointToClient(pos)
		mnu.Show(MyGrid, pos)
	End Sub

	' Take over cell parsing so that we store real values instead of strings
	Private Sub MyGrid_CellParsing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellParsingEventArgs) Handles MyGrid.CellParsing
		Dim text As String = DirectCast(e.Value, String)
		If text.StartsWith("=") = False Then
			' If the string value is not a formula then try to parse it into a real value
			e.Value = ciloci.FormulaEngine.Utility.Parse(text)
			e.ParsingApplied = True
		End If
	End Sub

	' Validate a formula entered into a cell
	Private Sub MyGrid_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles MyGrid.CellValidating
		If MyGrid.IsCurrentCellInEditMode = False Then
			Return
		End If

		Dim text As String = DirectCast(e.FormattedValue, String)

		' Is the text a formula?
		If text.StartsWith("=") = True Then
			' Try to create a formula from it
			Dim f As Formula = Me.Owner.CreateFormula(text)
			If f Is Nothing Then
				' Could not create a formula so fail validation
				e.Cancel = True
			Else
				' Store it for later use in validated event
				Dim ds As DictionaryService = Me.Site.GetService(GetType(DictionaryService))
				ds.SetValue("Formula", f)
			End If
		End If
	End Sub

	' Perfom the task of storing a valid formula
	Private Sub MyGrid_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles MyGrid.CellValidated
		If MyGrid.IsCurrentCellInEditMode = False Then
			Return
		End If

		Dim ds As DictionaryService = Me.Site.GetService(GetType(DictionaryService))
		' Get a sheet reference to the cell
		Dim ref As ISheetReference = Me.GetCellReference(e.RowIndex, e.ColumnIndex)

		' Remove any formulas at this cell
		Me.Engine.RemoveFormulasInRange(ref)

		' Get our stored formula
		Dim f As Formula = ds.GetValue("Formula")

		If Not f Is Nothing Then
			' Clear the cell's value and add a formula to it
			MyGrid.Item(e.ColumnIndex, e.RowIndex).Value = Nothing
			Me.Engine.AddFormula(f, ref)
			' Remove our saved formula
			ds.RemoveValue("Formula")
		End If

		' Finally, tell the engine to recalculate the cell
		Me.Engine.Recalculate(ref)
	End Sub

	Private Sub MyGrid_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles MyGrid.CellMouseClick
		If e.Button = MouseButtons.Right Then
			Me.HandleRightClick(e.RowIndex, e.ColumnIndex)
		ElseIf e.Button = MouseButtons.Left Then
			' Disable selection of individual cells using click+CTRL since we can't handle them
			If Control.ModifierKeys = Keys.Control Then
				MyGrid.ClearSelection()
				Me.ShowMessage("Selecting individual cells using CTRL is not supported", MessageType.Information)
			End If
		End If
	End Sub

	Private Sub MyGrid_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles MyGrid.DataError
		e.ThrowException = False
	End Sub

	Private Sub MyGrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyGrid.KeyDown
		e.Handled = True

		If e.KeyCode = Keys.Delete Then
			Me.ClearSelectedCells()
		Else
			e.Handled = False
		End If
	End Sub

	Public Sub DoCopy()
		Dim selected As Rectangle = Me.GetSelectedRange()
		If selected.IsEmpty = True Then
			Return
		End If

		Dim manager As ClipboardManager = Me.Site.GetService(GetType(ClipboardManager))
		manager.DoCopy(Me, selected)
	End Sub

	Public Sub DoCut()
		Dim selected As Rectangle = Me.GetSelectedRange()
		If selected.IsEmpty = True Then
			Return
		End If

		Dim manager As ClipboardManager = Me.Site.GetService(GetType(ClipboardManager))
		manager.DoCut(Me, selected)
	End Sub

	Public Sub DoPaste()
		Dim curPos As Point = MyGrid.CurrentCellAddress
		Dim manager As ClipboardManager = Me.Site.GetService(GetType(ClipboardManager))
		manager.DoPaste(Me, curPos)
	End Sub

	' Insert columns into our grid
	Private Sub DoColumnInsert()
		Dim insertAt As Integer = Me.GetFirstSelectedColumnIndex()
		Dim count As Integer = MyGrid.SelectedColumns.Count

		For i As Integer = 0 To count - 1
			Dim col As New DataGridViewColumn
			Me.ProcessColumn(col, i)
			MyGrid.Columns.Insert(insertAt, col)
		Next

		Me.RenumberColumns(insertAt)
		' Notify the engine of the insertion
		' Remember, first column is 1
		Me.Engine.OnColumnsInserted(insertAt + 1, count)
	End Sub

	' Remove columns from our grid
	Private Sub DoColumnRemove()
		Dim removeAt As Integer = Me.GetFirstSelectedColumnIndex()
		Dim count As Integer = MyGrid.SelectedColumns.Count

		If count = MyGrid.ColumnCount Then
			Me.ShowMessage("Cannot remove all columns", MessageType.Error)
			Return
		End If

		For i As Integer = 0 To count - 1
			MyGrid.Columns.RemoveAt(removeAt)
		Next

		Me.RenumberColumns(removeAt)
		' Notify the engine that columns were removed
		' Remember, first column is 1
		Me.Engine.OnColumnsRemoved(removeAt + 1, count)
	End Sub

	Private Sub DoRowInsert()
		Dim selected As Rectangle = Me.GetSelectedRange()
		Dim insertAt As Integer = selected.Top
		Dim count As Integer = selected.Height

		MyGrid.Rows.Insert(insertAt, count)
		Me.ProcessRows(insertAt, count)
		' Notify the engine that rows were inserted
		' Remember, first row is 1
		Me.Engine.OnRowsInserted(insertAt + 1, count)
	End Sub

	Private Sub DoRowRemove()
		Dim selected As Rectangle = Me.GetSelectedRange()
		Dim removeAt As Integer = selected.Top
		Dim count As Integer = selected.Height

		If count = MyGrid.RowCount Then
			Me.ShowMessage("Cannot remove all rows", MessageType.Error)
			Return
		End If

		For i As Integer = 0 To count - 1
			MyGrid.Rows.RemoveAt(removeAt)
		Next

		' Notity the engine that rows were removed
		' Remember, first row is 1
		Me.Engine.OnRowsRemoved(removeAt + 1, count)
	End Sub

	Private Sub RenumberColumns(ByVal startAt As Integer)
		For i As Integer = startAt To MyGrid.Columns.Count - 1
			Dim col As DataGridViewColumn = MyGrid.Columns.Item(i)
			col.Name = Chr(i + 65)
		Next
	End Sub

	Private Function GetFirstSelectedColumnIndex() As Integer
		Dim index As Integer = Integer.MaxValue
		For Each col As DataGridViewColumn In MyGrid.SelectedColumns
			If col.Index < index Then
				index = col.Index
			End If
		Next

		Return index
	End Function

	Private Sub ShowMessage(ByVal message As String, ByVal mt As MessageType)
		Dim ms As MessageService = Me.Site.GetService(GetType(MessageService))
		ms.ShowMessage(message, mt)
	End Sub

	' Handle a cut-paste which is essentially a move of a range
	Friend Sub DoCutPaste(ByVal data As SheetData, ByVal destRange As Rectangle)
		Dim sourceRange As Rectangle = data.SourceRange
		Dim rowOffset, colOffset As Integer
		' Compute the row,col offsets of the move
		colOffset = destRange.Left - sourceRange.Left
		rowOffset = destRange.Top - sourceRange.Top

		' Save the formulas in the source range.  We do this so we can handle overlapping source and destination ranges
		Dim sourceFormulas As Formula(,) = Me.GetRangeFormulas(data.SourceSheet, sourceRange)
		' Get a reference to the source range
		Dim sourceRef As ciloci.FormulaEngine.ISheetReference = Me.GetRectangleReference(data.SourceSheet, sourceRange)
		' Clear all formulas in the destination range
		Me.Engine.RemoveFormulasInRange(Me.GetRectangleReference(destRange))
		' Clear all formulas in the source range
		Me.Engine.RemoveFormulasInRange(sourceRef)

		' Add our stored formulas 
		Me.AddFormulas(sourceFormulas, data.SourceSheet, sourceRange.Top, sourceRange.Left)
		data.SourceSheet.ClearRange(sourceRange)
		data.WriteData(Me, destRange)

		' Notify the engine of the range move so that it can adjust references as required
		Me.Engine.OnRangeMoved(sourceRef, rowOffset, colOffset)
		Me.RaiseRefreshRequired()
	End Sub

	' Add formulas in a table to the engine
	Private Sub AddFormulas(ByVal formulas As Formula(,), ByVal sheet As Worksheet, ByVal startRow As Integer, ByVal startCol As Integer)
		For row As Integer = 0 To formulas.GetLength(0) - 1
			For col As Integer = 0 To formulas.GetLength(1) - 1
				Dim f As Formula = formulas(row, col)
				If Not f Is Nothing Then
					Dim ref As ISheetReference = Me.GetCellReference(sheet, row + startRow, col + startCol)
					Me.Engine.AddFormula(f, ref)
				End If
			Next
		Next
	End Sub

	' Copy and paste formulas and data
	Friend Sub DoCopyPaste(ByVal data As SheetData, ByVal destRange As Rectangle)
		Dim sourceRange As Rectangle = data.SourceRange

		' Save the source formulas
		Dim sourceFormulas As Formula(,) = Me.GetRangeFormulas(data.SourceSheet, sourceRange)
		' Clear formulas in the destination range
		Me.Engine.RemoveFormulasInRange(Me.GetRectangleReference(destRange))
		' Write the actual data
		data.WriteData(Me, destRange)

		' Copy and adjust all the formulas we saved
		For row As Integer = destRange.Top To destRange.Bottom - 1
			For col As Integer = destRange.Left To destRange.Right - 1
				Dim sourceFormula As Formula = sourceFormulas(row - destRange.Top, col - destRange.Left)
				If Not sourceFormula Is Nothing Then
					' Get a reference to the destination cell
					Dim destRef As ISheetReference = Me.GetCellReference(row, col)
					' Let the engine copy and adjust the formula
					Me.Engine.CopySheetFormula(sourceFormula, destRef)
				End If
			Next
		Next

		Me.RaiseRefreshRequired()
	End Sub

	' Get all formulas in a range
	Private Function GetRangeFormulas(ByVal sheet As Worksheet, ByVal range As Rectangle) As Formula(,)
		Dim arr(range.Height - 1, range.Width - 1) As Formula

		For row As Integer = range.Top To range.Bottom - 1
			For col As Integer = range.Left To range.Right - 1
				Dim ref As ISheetReference = Me.GetCellReference(sheet, row, col)
				arr(row - range.Top, col - range.Left) = Me.Engine.GetFormulaAt(ref)
			Next
		Next

		Return arr
	End Function

	Public Function GetRangeRectangle(ByVal range As Rectangle) As Rectangle
		Dim rt As RangeType = Me.GetRangeType(range)
		Dim startCell, endCell As Point

		If rt = RangeType.Columns Then
			startCell = New Point(range.Left, 0)
			endCell = New Point(range.Right - 1, MyGrid.RowCount - 1)
		ElseIf rt = RangeType.Rows Then
			startCell = New Point(0, range.Top)
			endCell = New Point(MyGrid.ColumnCount - 1, range.Bottom - 1)
		Else
			startCell = New Point(range.Left, range.Top)
			endCell = New Point(range.Right - 1, range.Bottom - 1)
		End If

		Dim lastCol As DataGridViewColumn = MyGrid.Columns.GetLastColumn(DataGridViewElementStates.Displayed, DataGridViewElementStates.None)
		Dim lastColIndex As Integer = lastCol.Index
		Dim lastRowIndex As Integer = MyGrid.Rows.GetLastRow(DataGridViewElementStates.Displayed)

		endCell.X = System.Math.Min(endCell.X, lastColIndex)
		endCell.Y = System.Math.Min(endCell.Y, lastRowIndex)

		Dim r1, r2 As Rectangle
		r1 = MyGrid.GetCellDisplayRectangle(startCell.X, startCell.Y, True)
		r2 = MyGrid.GetCellDisplayRectangle(endCell.X, endCell.Y, True)

		Return Rectangle.Union(r1, r2)
	End Function

	Public Sub RefreshRange(ByVal range As Rectangle)
		Dim rect As Rectangle = Me.GetRangeRectangle(range)
		MyGrid.Invalidate(rect)
		MyGrid.Update()
	End Sub

	Public Function GetRangeType(ByVal range As Rectangle) As RangeType
		If range.Height = MyGrid.RowCount Then
			Return RangeType.Columns
		ElseIf range.Width = MyGrid.ColumnCount Then
			Return RangeType.Rows
		Else
			Return RangeType.Cells
		End If
	End Function

	' Draw the row number for each row
	Private Sub MyGrid_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles MyGrid.RowPostPaint
		e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, Brushes.Black, e.RowBounds.Left, e.RowBounds.Top)
	End Sub

	Private Sub OnClearMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.ClearSelectedCells()
	End Sub

	Private Sub OnCopyMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoCopy()
	End Sub

	Private Sub OnCutMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoCut()
	End Sub

	Private Sub OnPasteMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoPaste()
	End Sub

	Private Sub OnFillDownMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoFillDown()
	End Sub

	Private Sub OnFillRightMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoFillRight()
	End Sub

	Private Sub OnInsertColumnMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoColumnInsert()
	End Sub

	Private Sub OnRemoveColumnMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoColumnRemove()
	End Sub

	Private Sub OnInsertRowMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoRowInsert()
	End Sub

	Private Sub OnRemoveRowMenuItemClick(ByVal sender As Object, ByVal e As EventArgs)
		Me.DoRowRemove()
	End Sub

	Public Sub SizeToDefault()
		Me.SetSize(New Size(9, 22))
	End Sub

	' Our implementation of ISheet: Get a cell value
	Public Function GetCellValue(ByVal row As Integer, ByVal column As Integer) As Object Implements ciloci.FormulaEngine.ISheet.GetCellValue
		' Return the cell at row,col in our grid.
		' Offset by -1,-1 since our top-left is (0,0) and the engine's is (1,1)
		Return MyGrid.Item(column - 1, row - 1).Value
	End Function

	' Our implementation of ISheet: Save the result of a formula
	Public Sub SetFormulaResult(ByVal result As Object, ByVal row As Integer, ByVal column As Integer) Implements ciloci.FormulaEngine.ISheet.SetFormulaResult
		' Look up the cell and set its value.
		' Offset by -1,-1 since our top-left is (0,0) and the engine's is (1,1)
		Dim cell As ExcelLikeCell = MyGrid.Item(column - 1, row - 1)
		cell.Value = result
	End Sub

	Public Function GridContainsPoint(ByVal row As Integer, ByVal col As Integer) As Boolean
		Dim gridRect As New Rectangle(0, 0, MyGrid.ColumnCount, MyGrid.RowCount)
		Return gridRect.Contains(col, row)
	End Function

	Public Function GridContainsRange(ByVal range As Rectangle) As Boolean
		Dim gridRect As New Rectangle(0, 0, MyGrid.ColumnCount, MyGrid.RowCount)
		Return gridRect.Contains(range)
	End Function

	Public Function ClipToGrid(ByVal range As Rectangle) As Rectangle
		Return Rectangle.Intersect(range, Me.GridRange)
	End Function

	Public Sub Transform(ByVal target As Rectangle, ByVal transform As CellTransform)
		For row As Integer = target.Top To target.Bottom - 1
			For col As Integer = target.Left To target.Right - 1
				transform.TransformCell(MyGrid.Item(col, row), row, col)
			Next
		Next
	End Sub

	Public Sub SelectRow(ByVal rowIndex As Integer)
		Dim row As DataGridViewRow = MyGrid.Rows.Item(rowIndex)
		For Each cell As DataGridViewCell In row.Cells
			cell.Selected = True
		Next
	End Sub

	Private Sub RaiseRefreshRequired()
		Dim owner As MainForm = Me.Site.GetService(GetType(MainForm))
		owner.RaiseRefreshRequired()
	End Sub

	Public ReadOnly Property GridRange() As Rectangle
		Get
			Return New Rectangle(0, 0, MyGrid.ColumnCount, MyGrid.RowCount)
		End Get
	End Property

	Private ReadOnly Property Owner() As MainForm
		Get
			Return Me.Site.GetService(GetType(MainForm))
		End Get
	End Property

	Private ReadOnly Property Engine() As ciloci.FormulaEngine.FormulaEngine
		Get
			Return Me.Site.GetService(GetType(ciloci.FormulaEngine.FormulaEngine))
		End Get
	End Property

	' Our implementation if ISheet: return our given name
	Public ReadOnly Property Name() As String Implements ciloci.FormulaEngine.ISheet.Name
		Get
			Return MyName
		End Get
	End Property

	' Our implementation if ISheet: return our grid's row count
	Public ReadOnly Property RowCount() As Integer Implements ciloci.FormulaEngine.ISheet.RowCount
		Get
			Return MyGrid.RowCount
		End Get
	End Property

	' Our implementation if ISheet: return our grid's column count
	Public ReadOnly Property ColumnCount() As Integer Implements ciloci.FormulaEngine.ISheet.ColumnCount
		Get
			Return MyGrid.ColumnCount
		End Get
	End Property

	Public ReadOnly Property Grid() As DataGridView
		Get
			Return MyGrid
		End Get
	End Property
End Class