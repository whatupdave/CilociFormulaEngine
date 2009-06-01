Imports NUnit.Framework
Imports System.Drawing
Imports ciloci.FormulaEngine

<TestFixture()> _
Public Class FormulaEngineTestFixture
	Inherits TestFixtureBase

	Private MySheet1, MySheet2 As FormulaEngineTestGrid
	Public MyCircularReferenceFlag As Boolean

	Public Sub New()

	End Sub

	Protected Overrides Sub DoFixtureSetup()
		MySheet1 = New FormulaEngineTestGrid("Sheet1")
		MySheet2 = New FormulaEngineTestGrid("Sheet2")
		AddHandler MyFormulaEngine.CircularReferenceDetected, AddressOf OnCircularReferenceDetected
	End Sub

	Protected Overrides Sub DoTestSetup()
		Me.Clear()
	End Sub

	Private Sub Clear()
		MyFormulaEngine.Clear()
		MyFormulaEngine.Sheets.Add(MySheet1)
		MyFormulaEngine.Sheets.Add(MySheet2)
		MyCircularReferenceFlag = False
	End Sub

	Public Sub OnCircularReferenceDetected(ByVal sender As Object, ByVal e As CircularReferenceDetectedEventArgs)
		MyCircularReferenceFlag = True
	End Sub

	<Test()> _
	Public Sub TestMultiThreadedParsing()
		' Test that we can parse formulas from multiple threads
		Dim fe As New FormulaEngine()
		Dim t1 As New System.Threading.Thread(AddressOf ThreadStart)
		Dim t2 As New System.Threading.Thread(AddressOf ThreadStart)

		t1.Start(fe)
		t2.Start(fe)

		t1.Join()
		t2.Join()
	End Sub

	Private Sub ThreadStart(ByVal o As Object)
		Dim fe As FormulaEngine = o

		For i As Integer = 0 To 20 - 1
			Dim result As Object = fe.Evaluate("=cos(pi())")
		Next
	End Sub

	<Test()> _
	Public Sub TestGeneralGrammar()
		Dim v As Variable = MyFormulaEngine.DefineVariable("a")
		v.Value = 10
		Dim result As Double = MyFormulaEngine.Evaluate("(1+100-2.2) * (a ^ 3)", GrammarType.General)
		Assert.AreEqual(98800, result)
	End Sub

	<Test()> _
	Public Sub TestSerialization()
		Dim engine As New FormulaEngine()
		engine.AddFormula("=1+1", engine.ReferenceFactory.Named("test"))
		Dim ms As New System.IO.MemoryStream
		Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
		bf.Serialize(ms, engine)

		ms.Seek(0, IO.SeekOrigin.Begin)

		Dim engine2 As FormulaEngine = bf.Deserialize(ms)

		Assert.AreEqual(1, engine2.FormulaCount)
		Assert.IsTrue(engine2.HasFormulaAt(engine2.ReferenceFactory.Named("test")))

		Dim result As Double = DirectCast(engine2.Evaluate("=1+2+3"), Double)
		Assert.AreEqual(6.0, result)
	End Sub

	<Test(), ExpectedException(GetType(ArgumentException))> _
	Public Sub TestAddOverlappingFormula()
		' Add a formula at the same location as an existing one
		Me.CreateFormula("1+1", 1, 1)
		Me.CreateFormula("2*56", 1, 1)
	End Sub

	<Test()> _
	Public Sub TestAddFormula()
		' Null formula
		Me.DoAddFormula(Nothing, MyFormulaEngine.ReferenceFactory.Parse("a1"))
		Dim f As Formula = MyFormulaEngine.CreateFormula("1+1")
		' Null ref
		Me.DoAddFormula(f, Nothing)
		' Invalid references
		Me.DoAddFormula(f, MyFormulaEngine.ReferenceFactory.Cells(2, 2, 4, 4))
		Me.DoAddFormula(f, MyFormulaEngine.ReferenceFactory.Columns(3, 4))
		Me.DoAddFormula(f, MyFormulaEngine.ReferenceFactory.Rows(5, 6))

		' Add with duplicate reference
		MyFormulaEngine.AddFormula(f, MyFormulaEngine.ReferenceFactory.Parse("A1"))
		Me.DoAddFormula(f, MyFormulaEngine.ReferenceFactory.Parse("A1"))
	End Sub

	Private Sub DoAddFormula(ByVal f As Formula, ByVal ref As IReference)
		Dim sawException As Boolean = False
		Try
			MyFormulaEngine.AddFormula(f, ref)
		Catch ex As Exception
			sawException = True
		End Try

		Assert.IsTrue(sawException)
	End Sub

	<Test()> _
	Public Sub TestReferenceParsing()
		' Test our column index parse/format
		Me.TestColumnIndexParseFormat(1)
		Me.TestColumnIndexParseFormat(15)
		Me.TestColumnIndexParseFormat(27)
		Me.TestColumnIndexParseFormat(256)
		Me.TestColumnIndexParseFormat(71)

		MySheet1.SetSize(65536, 256)

		' Test that we are properly interpreting the row and column of string cell references
		Me.DoTestReferenceParsing("A1", 1, 1)
		Me.DoTestReferenceParsing("b2", 2, 2)
		Me.DoTestReferenceParsing("o15", 15, 15)
		Me.DoTestReferenceParsing("z133", 133, 26)
		Me.DoTestReferenceParsing("AA155", 155, 27)
		Me.DoTestReferenceParsing("l12", 12, 12)
		Me.DoTestReferenceParsing("O15", 15, 15)
		Me.DoTestReferenceParsing("He13", 13, 213)
		Me.DoTestReferenceParsing("IV100", 100, 256)
		Me.DoTestReferenceParsing("BS6", 6, 71)
		Me.DoTestReferenceParsing("A65535", 65535, 1)
		Me.DoTestReferenceParsing("CH2000", 2000, 86)
		Me.DoTestReferenceParsing("IV35536", 35536, 256)
		Me.DoTestReferenceParsing("IV65535", 65535, 256)

		MySheet1.ResetSize()
	End Sub

	Private Sub TestColumnIndexParseFormat(ByVal columnIndex As Integer)
		Dim s As String = ciloci.FormulaEngine.Utility.ColumnIndex2Label(columnIndex)
		Dim index As Integer = ciloci.FormulaEngine.Utility.ColumnLabel2Index(s)
		Assert.AreEqual(columnIndex, index)
	End Sub

	Private Sub DoTestReferenceParsing(ByVal image As String, ByVal expectedRow As Integer, ByVal expectedColumn As Integer)
		Dim ref As ISheetReference = MyFormulaEngine.ReferenceFactory.Parse(image)
		Assert.AreEqual(expectedRow, ref.Row)
		Assert.AreEqual(expectedColumn, ref.Column)
	End Sub

	<Test()> _
	Public Sub TestReferenceValidation()
		Me.DoReferenceValidation("aaa90")
		Me.DoReferenceValidation("zz90")
		Me.DoReferenceValidation("a0")
		Me.DoReferenceValidation("a100000")
		Me.DoReferenceValidation("zzz80000")
		Me.DoReferenceValidation("aa")
		Me.DoReferenceValidation("10")
		Me.DoReferenceValidation("a10:zz45")
		Me.DoReferenceValidation("a99999:b4")
		Me.DoReferenceValidation("Sh  et!a1:b4")
		Me.DoReferenceValidation("9sheet!a1:b4")
		Me.DoReferenceValidation("Sheet1 a1:b4")
	End Sub

	Private Sub DoReferenceValidation(ByVal image As String)
		Dim invalid As Boolean

		Try
			MyFormulaEngine.ReferenceFactory.Parse(image)
			invalid = False
		Catch ex As Exception
			invalid = True
		End Try

		Assert.IsTrue(invalid, image)
	End Sub

	Private Sub ProcessTestScriptLine(ByVal line As String)
		'Console.WriteLine(line)
		Dim infos As TestComponentInfo() = Me.CreateTestComponentInfos(line)
		Me.RunTest(infos, line)
	End Sub

	Private Function CreateTestComponentInfos(ByVal line As String) As TestComponentInfo()
		Dim strings As String() = line.Split(TestFixtureBase.TEST_COMPONENT_DELIMITER)
		Dim infos(strings.Length - 1) As TestComponentInfo

		For i As Integer = 0 To strings.Length - 1
			Dim strings2 As String() = strings(i).Split(TestFixtureBase.ELEMENT_DELIMITER)
			Dim info As TestComponentInfo
			info.Component = Me.CreateType(strings2(0))
			info.Args = strings2(1).Split(TestFixtureBase.ARG_DELIMITER)
			infos(i) = info
		Next

		Return infos
	End Function

	Private Sub RunTest(ByVal infos As TestComponentInfo(), ByVal line As String)
		Dim state As IDictionary = New Hashtable
		state.Add(GetType(FormulaEngineTestFixture), Me)

		Me.Clear()

		For i As Integer = 0 To infos.Length - 1
			Dim info As TestComponentInfo
			info = infos(i)
			Dim component As TestComponent = info.Component
			component.AcceptParameters(info.Args, MyFormulaEngine)
			Me.ExecuteTestComponent(component, state, line)
		Next
	End Sub

	Private Sub ExecuteTestComponent(ByVal component As TestComponent, ByVal state As IDictionary, ByVal line As String)
		Try
			component.Execute(MyFormulaEngine, state)
		Catch ex As Exception
			Dim arr(3 - 1) As String
			arr(0) = line
			arr(1) = ex.ToString()
			arr(2) = New String("-", 80)
			Dim msg As String = String.Join(System.Environment.NewLine, arr)
			Assert.Fail(msg)
		End Try
	End Sub

	<Test()> _
	Public Sub TestMiscellaneous()
		Me.ProcessScriptTests("MiscellaneousTests.txt", AddressOf ProcessTestScriptLine)
	End Sub

	<Test()> _
	Public Sub TestReferenceAdjusts()
		Me.ProcessScriptTests("ReferenceTests.txt", AddressOf ProcessTestScriptLine)
	End Sub

	<Test()> _
	Public Sub TestDependencies()
		Me.ProcessScriptTests("DependencyTests.txt", AddressOf ProcessTestScriptLine)
	End Sub

	<Test()> _
	Public Sub TestCalculationOrder()
		Me.ProcessScriptTests("CalculationOrderTests.txt", AddressOf ProcessTestScriptLine)
	End Sub

	<Test()> _
	Public Sub TestCircularReferenceHandling()
		Me.ProcessScriptTests("CircularReferenceTests.txt", AddressOf ProcessTestScriptLine)
	End Sub
End Class