Imports NUnit.Framework
Imports Microsoft.Office.Interop

' Tests specific to formulas
<TestFixture(), CLSCompliant(False)> _
Public Class FormulaEvaluateTestFixture
	Inherits TestFixtureBase

    Protected MyExcelApplication As Excel.Application
	Protected MyWorksheet As Excel.Worksheet
	Private MyGrid As FormulaEvaluateGrid

	Protected Overrides Sub DoFixtureSetup()
		MyGrid = New FormulaEvaluateGrid
		MyExcelApplication = New Excel.ApplicationClass
		Dim wb As Excel.Workbook = MyExcelApplication.Workbooks.Add()
		MyWorksheet = wb.Worksheets.Item(1)
		Me.CopyTableToWorkSheet()
	End Sub

	Protected Overrides Sub DoFixtureTearDown()
		MyExcelApplication.Workbooks.Item(1).Saved = True
		MyExcelApplication.Workbooks.Close()
		MyExcelApplication.Quit()
	End Sub

	Private Sub ClearFormulaEngine()
		MyFormulaEngine.Clear()
		MyFormulaEngine.Sheets.Add(MyGrid)
	End Sub

	Private Sub CopyTableToWorkSheet()
		Dim grid As ciloci.FormulaEngine.ISheet = MyGrid

		For row As Integer = 1 To grid.RowCount
			For col As Integer = 1 To grid.ColumnCount
				Dim value As Object = grid.GetCellValue(row, col)

				If value Is Nothing Then
					value = value
				ElseIf value.GetType() Is GetType(ciloci.FormulaEngine.ErrorValueWrapper) Then
					value = value.ToString()
				ElseIf value.GetType() Is GetType(String) Then
					value = "'" & value
				End If

				MyWorksheet.Cells.Item(row, col) = value
			Next
		Next
	End Sub

	<Test()> _
	Public Sub TestFormulaEvaluateAgainstExcel()
		Me.ProcessScriptTests("ValidTestFormulas.txt", AddressOf ProcessFormulaEvaluateAgainstExcel)
	End Sub

	<Test()> _
	Public Sub TestCultureSensitiveParse()
		' Test that we set the decimal point and argument separator based on the current culture
		Dim oldCi As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
		Dim ci As New System.Globalization.CultureInfo("et-EE")
		System.Threading.Thread.CurrentThread.CurrentCulture = ci
		MyFormulaEngine.RecreateParsers()
		Try
			Dim f As ciloci.FormulaEngine.Formula = MyFormulaEngine.CreateFormula("13,45 + sum(1;2;3)")
			Dim result As Double = DirectCast(f.Evaluate(), Double)
			Assert.AreEqual(19.45, result)
		Catch ex As ciloci.FormulaEngine.InvalidFormulaException
			Assert.IsTrue(False)
		Finally
			' Reset this or excel will give us problems
			System.Threading.Thread.CurrentThread.CurrentCulture = oldCi
			' Recreate parser in default culture for other tests
			MyFormulaEngine.RecreateParsers()
		End Try
	End Sub

	Private Sub ProcessFormulaEvaluateAgainstExcel(ByVal formula As String)
		Me.ClearFormulaEngine()
		'Console.WriteLine(formula)
		Try
			Dim formulaResult As Object = MyFormulaEngine.Evaluate(formula)
			Dim excelResult As Object = Me.EvaluateFormulaInExcel(formula)

			Me.CompareResults(formulaResult, excelResult, formula)
		Catch ex As Exception
			Console.WriteLine("Failed formula: {0}", formula)
			Throw ex
		End Try
	End Sub

	Private Function EvaluateFormulaInExcel(ByVal formula As String) As Object
		' Do it this way so that we get exactly the same results as a user would get
		Dim r As Excel.Range = MyWorksheet.Range("F9")

		If formula.StartsWith("=") = False Then
			formula = "=" & formula
		End If

		r.Formula = formula
		r.NumberFormat = "General"

		Return r.Value
	End Function

	Private Sub CompareResults(ByVal formulaResult As Object, ByVal excelResult As Object, ByVal formula As String)
		If formulaResult.GetType() Is GetType(ciloci.FormulaEngine.ErrorValueWrapper) Then
			If Not excelResult.GetType() Is GetType(Integer) Then
				Throw New ArgumentException("Formula returned an error but excel did not")
			End If
			' Error value
			Dim formulaError As ciloci.FormulaEngine.ErrorValueWrapper = formulaResult
			Dim excelError As ciloci.FormulaEngine.ErrorValueWrapper = Me.ExcelError2FormulaError(DirectCast(excelResult, Integer))

			Assert.AreEqual(excelError, formulaError, formula)
		ElseIf TypeOf (formulaResult) Is ciloci.FormulaEngine.ISheetReference Then
			Me.CompareRanges(formulaResult, excelResult)
		ElseIf formulaResult.GetType() Is GetType(DateTime) Then
			' Sometimes excel gives us a date and sometimes a double
			Me.CompareDate(formulaResult, excelResult)
		Else
			formulaResult = Me.NormalizeValue(formulaResult)
			excelResult = Me.NormalizeValue(excelResult)
			Assert.AreEqual(excelResult, formulaResult, formula)
		End If
	End Sub

	Private Sub CompareDate(ByVal formulaResult As Object, ByVal excelResult As Object)
		Dim formulaDate As DateTime = DirectCast(formulaResult, DateTime)
		Dim excelDate As DateTime

		If excelResult.GetType() Is GetType(DateTime) Then
			excelDate = DirectCast(excelResult, DateTime)
		ElseIf excelResult.GetType() Is GetType(Double) Then
			excelDate = DateTime.FromOADate(DirectCast(excelResult, Double))
		Else
			Throw New ArgumentException("Formula returned a date but excel did not")
		End If

		Assert.AreEqual(excelDate, formulaDate)
	End Sub

	Private Function ExcelError2FormulaError(ByVal value As Integer) As ciloci.FormulaEngine.ErrorValueWrapper
		Dim excelErrorValue As Excel.XlCVError = value And &HFFFF
		Dim errorValue As ciloci.FormulaEngine.ErrorValueType

		Select Case excelErrorValue
			Case Excel.XlCVError.xlErrDiv0
				errorValue = ciloci.FormulaEngine.ErrorValueType.Div0
			Case Excel.XlCVError.xlErrNA
				errorValue = ciloci.FormulaEngine.ErrorValueType.NA
			Case Excel.XlCVError.xlErrName
				errorValue = ciloci.FormulaEngine.ErrorValueType.Name
			Case Excel.XlCVError.xlErrNull
				errorValue = ciloci.FormulaEngine.ErrorValueType.Null
			Case Excel.XlCVError.xlErrNum
				errorValue = ciloci.FormulaEngine.ErrorValueType.Num
			Case Excel.XlCVError.xlErrRef
				errorValue = ciloci.FormulaEngine.ErrorValueType.Ref
			Case Excel.XlCVError.xlErrValue
				errorValue = ciloci.FormulaEngine.ErrorValueType.Value
			Case Else
				Throw New InvalidOperationException("Unknown error code")
		End Select

		Return ciloci.FormulaEngine.FormulaEngine.CreateError(errorValue)
	End Function

	Private Function NormalizeValue(ByVal value As Object) As Object
		If value.GetType() Is GetType(Double) Then
			Dim d As Double = DirectCast(value, Double)
			d = System.Math.Round(d, 8)
			Return d
		Else
			Return value
		End If
	End Function

	Protected Sub CompareRanges(ByVal ref As ciloci.FormulaEngine.ISheetReference, ByVal range As Excel.Range)
		Dim excelRect As System.Drawing.Rectangle = Me.ExcelRangeToRectangle(range)
		Dim refRect As System.Drawing.Rectangle = ref.Area
		Assert.AreEqual(excelRect, refRect)
	End Sub

	Protected Function ExcelRangeToRectangle(ByVal r As Excel.Range) As System.Drawing.Rectangle
		Dim rect As New System.Drawing.Rectangle(r.Column, r.Row, r.Columns.Count, r.Rows.Count)
		Return rect
	End Function
End Class