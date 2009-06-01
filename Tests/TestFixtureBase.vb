Imports NUnit.Framework

Public MustInherit Class TestFixtureBase

	Protected Const COMMENT_CHAR As Char = "'"c
	Protected MyFormulaEngine As ciloci.FormulaEngine.FormulaEngine

	Protected Const ARG_DELIMITER As Char = ";"c
	Protected Const TEST_COMPONENT_DELIMITER As Char = "•"c
	Protected Const ELEMENT_DELIMITER As Char = "†"c

	<TestFixtureSetUp()> _
	Public Sub FixtureSetUp()
		MyFormulaEngine = New ciloci.FormulaEngine.FormulaEngine()
		Me.DoFixtureSetup()
	End Sub

	<TestFixtureTearDown()> _
	Public Sub FixtureTearDown()
		Me.DoFixtureTearDown()
	End Sub

	Protected Overridable Sub DoFixtureSetup()

	End Sub

	Protected Overridable Sub DoFixtureTearDown()

	End Sub

	<SetUp()> _
	Public Sub TestSetup()
		Me.DoTestSetup()
	End Sub

	Protected Overridable Sub DoTestSetup()

	End Sub

	Protected Function CreateFormula(ByVal expression As String, ByVal row As Integer, ByVal col As Integer) As ciloci.FormulaEngine.Formula
		Dim ref As ciloci.FormulaEngine.ISheetReference = MyFormulaEngine.ReferenceFactory.Cell(row, col)
		Dim f As ciloci.FormulaEngine.Formula = MyFormulaEngine.AddFormula(expression, ref)
		Return f
	End Function

	Protected Sub ProcessScriptTests(ByVal scriptFileName As String, ByVal processor As LineProcessor)
		Dim scriptPath As String = System.IO.Path.Combine("../TestScripts", scriptFileName)
		Dim instream As New System.IO.FileStream(scriptPath, IO.FileMode.Open, IO.FileAccess.Read)

		Dim sr As New System.IO.StreamReader(instream)
		Dim line As String = sr.ReadLine()

		While Not line Is Nothing
			If line.StartsWith(COMMENT_CHAR) = False Then
				processor(line)
			End If

			line = sr.ReadLine()
		End While

		instream.Close()
	End Sub

	Protected Function CreateType(ByVal name As String) As Object
		Return Activator.CreateInstance(Type.GetType("Tests." & name, True))
	End Function
End Class
