Imports ciloci.FormulaEngine

Friend Class FormulaEvaluateGrid
	Implements ISheet

	Private Const ROW_COUNT As Integer = 7
	Private Const COL_COUNT As Integer = 5
	Private MyTable As Object(,)

	Public Sub New()
		MyTable = New Object(ROW_COUNT - 1, COL_COUNT - 1) {}
		Me.FillTable()
	End Sub

	Private Sub FillTable()
		' Numbers
		MyTable(0, 0) = 123.34
		MyTable(0, 1) = CDbl(13)
		MyTable(0, 2) = CDbl(-5)
		MyTable(0, 3) = 1000.23
		MyTable(0, 4) = 1300		' Integer

		' More numbers
		MyTable(1, 0) = 0.56
		MyTable(1, 1) = CDbl(100)
		MyTable(1, 2) = CDbl(0)
		MyTable(1, 3) = 3.45
		MyTable(1, 4) = 155		'Integer

		' Numbers as strings
		MyTable(2, 0) = "123"
		MyTable(2, 1) = "45.23"
		MyTable(2, 2) = "0"
		MyTable(2, 3) = "-11"
		MyTable(2, 4) = String.Empty

		' Random strings
		MyTable(3, 0) = "eugene"
		MyTable(3, 1) = "not a number"
		MyTable(3, 2) = "****"
		MyTable(3, 3) = "()()^^^"
		MyTable(3, 4) = "S"

		' Errors
		MyTable(4, 0) = FormulaEngine.CreateError(ErrorValueType.Div0)
		MyTable(4, 1) = FormulaEngine.CreateError(ErrorValueType.NA)
		MyTable(4, 2) = FormulaEngine.CreateError(ErrorValueType.Name)
		MyTable(4, 3) = FormulaEngine.CreateError(ErrorValueType.Null)
		MyTable(4, 4) = FormulaEngine.CreateError(ErrorValueType.Num)

		' Empty cells
		MyTable(5, 0) = Nothing
		MyTable(5, 1) = Nothing
		MyTable(5, 2) = Nothing
		MyTable(5, 3) = Nothing
		MyTable(5, 4) = Nothing

		' Booleans
		MyTable(6, 0) = True
		MyTable(6, 1) = False
		MyTable(6, 2) = False
		MyTable(6, 3) = True
		MyTable(6, 4) = True
	End Sub

	Public Function GetCellValue(ByVal row As Integer, ByVal column As Integer) As Object Implements ciloci.FormulaEngine.ISheet.GetCellValue
		Return MyTable(row - 1, column - 1)
	End Function

	Public Sub SetFormulaResult(ByVal result As Object, ByVal row As Integer, ByVal column As Integer) Implements ciloci.FormulaEngine.ISheet.SetFormulaResult
		MyTable(row - 1, column - 1) = result
	End Sub

	Public ReadOnly Property Name() As String Implements ciloci.FormulaEngine.ISheet.Name
		Get
			Return "Sheet1"
		End Get
	End Property

	Public ReadOnly Property RowCount() As Integer Implements ciloci.FormulaEngine.ISheet.RowCount
		Get
			Return ROW_COUNT
		End Get
	End Property

	Public ReadOnly Property ColumnCount() As Integer Implements ciloci.FormulaEngine.ISheet.ColumnCount
		Get
			Return COL_COUNT
		End Get
	End Property
End Class

Friend Class FormulaEngineTestGrid
	Implements ISheet

	Private Const ROW_COUNT As Integer = 15
	Private Const COL_COUNT As Integer = 15
	Private MyCells As Object(,)
	Private MyName As String
	Private MyRowCount, MyColCount As Integer

	Public Sub New(ByVal name As String)
		MyCells = New Object(ROW_COUNT - 1, COL_COUNT - 1) {}
		MyName = name
		Me.ResetSize()
	End Sub

	Public Sub ResetSize()
		MyRowCount = ROW_COUNT
		MyColCount = COL_COUNT
	End Sub

	Public Sub SetSize(ByVal rowCount As Integer, ByVal colCount As Integer)
		MyRowCount = rowCount
		MyColCount = colCount
	End Sub

	Public Sub ClearCell(ByVal row As Integer, ByVal col As Integer)
		MyCells(row - 1, col - 1) = Nothing
	End Sub

	Public Function GetCellValue(ByVal row As Integer, ByVal column As Integer) As Object Implements ciloci.FormulaEngine.ISheet.GetCellValue
		Return MyCells(row - 1, column - 1)
	End Function

	Public Sub SetFormulaResult(ByVal result As Object, ByVal row As Integer, ByVal column As Integer) Implements ciloci.FormulaEngine.ISheet.SetFormulaResult
		MyCells(row - 1, column - 1) = result
	End Sub

	Public ReadOnly Property Name() As String Implements ciloci.FormulaEngine.ISheet.Name
		Get
			Return MyName
		End Get
	End Property

	Public ReadOnly Property RowCount() As Integer Implements ciloci.FormulaEngine.ISheet.RowCount
		Get
			Return MyRowCount
		End Get
	End Property

	Public ReadOnly Property ColumnCount() As Integer Implements ciloci.FormulaEngine.ISheet.ColumnCount
		Get
			Return MyColCount
		End Get
	End Property
End Class