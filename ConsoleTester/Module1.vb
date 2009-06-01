Imports ciloci.FormulaEngine
Imports System.Drawing

Module Module1

	Private MyEngine As FormulaEngine

	Sub Main()
		Dim sheet1 As New DumbSheet("Sheet1")
		sheet1.SetCellValue(1, 1, 14.56)
		sheet1.SetCellValue(2, 1, "")
		sheet1.SetCellValue(3, 1, 156)
		sheet1.SetCellValue(1, 2, 100)
		sheet1.SetCellValue(2, 2, 200)
		sheet1.SetCellValue(3, 2, 300)

		Dim sheet2 As New DumbSheet("Sheet2")

		MyEngine = New FormulaEngine()

		Dim result As Object = MyEngine.Evaluate("Substitute(""bbb"", """", ""!!"")")

		'MyEngine.Sheets.Add(sheet1)
		'MyEngine.Sheets.Add(sheet2)

		'Dim f As Formula = CreateFormula("b1+1", "A1")

		Dim ref As IExternalReference = MyEngine.ReferenceFactory.External()
		MyEngine.AddFormula("cos(45)", ref)

		Dim m As New Memento
		m.MyRef = ref
		m.MyEngine = MyEngine

		Dim ms As New System.IO.MemoryStream()
		Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
		bf.Serialize(ms, m)

		ms.Seek(0, IO.SeekOrigin.Begin)

		Dim m2 As Memento = bf.Deserialize(ms)
	End Sub

	<Serializable()> _
	Private Class Memento
		Public MyRef As IExternalReference
		Public MyEngine As FormulaEngine
	End Class

	Private Function CreateFormula(ByVal expression As String, ByVal ref As String) As Formula
		Dim gridRef As ISheetReference = MyEngine.ReferenceFactory.Parse(ref)
		Dim f As Formula = MyEngine.AddFormula(expression, gridRef)
		Return f
	End Function
End Module

<Serializable()> _
Friend Class DumbSheet
	Implements ciloci.FormulaEngine.ISheet

	Private MyCells As Object(,)
	Private Const ROW_COUNT As Integer = 16
	Private Const COL_COUNT As Integer = 10
	Private MyName As String

	Public Sub New(ByVal name As String)
		MyCells = New Object(ROW_COUNT - 1, COL_COUNT - 1) {}
		MyName = name
	End Sub

	Public Sub SetCellValue(ByVal row As Integer, ByVal col As Integer, ByVal value As Object)
		MyCells(row - 1, col - 1) = value
	End Sub

	Public Function GetCellValue(ByVal row As Integer, ByVal column As Integer) As Object Implements ciloci.FormulaEngine.ISheet.GetCellValue
		Return MyCells(row - 1, column - 1)
	End Function

	Public Sub SetFormulaResult(ByVal result As Object, ByVal row As Integer, ByVal column As Integer) Implements ciloci.FormulaEngine.ISheet.SetFormulaResult

	End Sub

	Public ReadOnly Property Name() As String Implements ciloci.FormulaEngine.ISheet.Name
		Get
			Return MyName
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