Imports ciloci.FormulaEngine
Imports NUnit.Framework

Public Delegate Sub LineProcessor(ByVal line As String)

Friend MustInherit Class TestComponent
	Public MustOverride Sub AcceptParameters(ByVal params As String(), ByVal engine As FormulaEngine)
	Public MustOverride Sub Execute(ByVal engine As FormulaEngine, ByVal state As IDictionary)
End Class

Friend Class FirstReferenceEqualValidator
	Inherits TestComponent

	Private MyExpected As IReference

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyExpected = engine.ReferenceFactory.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim formulas As Formula() = state.Item("Formulas")
		Dim f As Formula = formulas(0)
		Dim refs As IReference() = f.References
		Dim firstRef As IReference = refs(0)
		Assert.IsTrue(MyExpected.Equals(firstRef))
	End Sub
End Class

Friend Class SetActiveFormulaExecutor
	Inherits TestComponent

	Private MyLocation As IReference

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyLocation = engine.ReferenceFactory.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim f As Formula = engine.GetFormulaAt(MyLocation)
		If f Is Nothing Then
			Throw New InvalidOperationException("No formula exists at given reference")
		Else
			state.Item("ActiveFormula") = f
		End If
	End Sub
End Class

Friend Class FormulaReferencesEqualValidator
	Inherits TestComponent

	Private MyExpected As IReference()

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		Dim arr(params.Length - 1) As IReference
		For i As Integer = 0 To params.Length - 1
			arr(i) = engine.ReferenceFactory.Parse(params(i))
		Next
		MyExpected = arr
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim f As Formula = state.Item("ActiveFormula")
		Dim refs As IReference() = f.References

		If MyExpected.Length <> refs.Length Then
			Throw New InvalidOperationException("Formula reference count does not match expected reference count")
		End If

		For i As Integer = 0 To MyExpected.Length - 1
			Assert.IsTrue(MyExpected(i).Equals(refs(i)))
		Next
	End Sub
End Class

Friend Class FormulaResultEqualsErrorValidator
	Inherits TestComponent

	Private MyExpectedError As ErrorValueType

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyExpectedError = System.Enum.Parse(GetType(ErrorValueType), params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim formulas As Formula() = state.Item("Formulas")
		Dim f As Formula = formulas(0)
		Dim result As ErrorValueWrapper = f.Evaluate()
		Assert.AreEqual(MyExpectedError, result.ErrorValue)
	End Sub
End Class

Friend Class FormulaFormatValidator
	Inherits TestComponent

	Private MyExpected As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyExpected = params(0)
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim formulas As Formula() = state.Item("Formulas")
		Dim f As Formula = formulas(0)
		Dim s As String = f.ToString()
		Assert.IsTrue(String.Equals(MyExpected, s, StringComparison.OrdinalIgnoreCase), s)
	End Sub
End Class

Friend Class PooledReferenceCountValidator
	Inherits TestComponent

	Private MyExpected As Integer

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyExpected = Integer.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Assert.AreEqual(MyExpected, engine.Info.ReferenceCount)
	End Sub
End Class

Friend Class FormulaEngineEmptyValidator
	Inherits TestComponent

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)

	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Assert.AreEqual(0, engine.Info.ReferenceCount)
		Assert.AreEqual(0, engine.Info.DependentsCount)
		Assert.AreEqual(0, engine.Info.PrecedentsCount)
		Assert.AreEqual(0, engine.FormulaCount)
	End Sub
End Class

Friend Class NoDependenciesValidator
	Inherits TestComponent

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)

	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Assert.AreEqual(0, engine.Info.DependentsCount)
	End Sub
End Class

Friend Class NumDependenciesValidator
	Inherits TestComponent

	Private MyNumDependencies As Integer

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyNumDependencies = Integer.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Assert.AreEqual(MyNumDependencies, engine.Info.DependentsCount)
	End Sub
End Class

Friend Class NumReferencesValidator
	Inherits TestComponent

	Private MyExpectedCount As Integer

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyExpectedCount = Integer.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Assert.AreEqual(MyExpectedCount, engine.Info.ReferenceCount)
	End Sub
End Class

Friend Class CalculationOrderValidator
	Inherits TestComponent

	Private MyOrder As Integer()

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		Dim arr(params.Length - 1) As Integer

		For i As Integer = 0 To arr.Length - 1
			arr(i) = Integer.Parse(params(i))
		Next

		MyOrder = arr
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim formulas As Formula() = state.Item("Formulas")
		Dim calculationList As Formula() = state.Item("CalculationList")

		If formulas.Length <> calculationList.Length Then
			Throw New InvalidOperationException("Formula and calculation list have different sizes")
		End If

		If formulas.Length <> MyOrder.Length Then
			Throw New InvalidOperationException("Wrong sizes")
		End If

		For i As Integer = 0 To formulas.Length - 1
			Dim f As Formula = formulas(i)
			Dim index As Integer = System.Array.IndexOf(calculationList, f)
			Dim expected As Integer = MyOrder(i)
			Assert.AreEqual(expected, index)
		Next
	End Sub
End Class

Friend Class EmptyCalculationOrderValidator
	Inherits TestComponent

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)

	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim calculationList As Formula() = state.Item("CalculationList")
		Assert.AreEqual(0, calculationList.Length)
	End Sub
End Class

Friend Class NonEmptyCalculationOrderValidator
	Inherits TestComponent

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)

	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim calculationList As Formula() = state.Item("CalculationList")
		Assert.AreNotEqual(0, calculationList.Length)
	End Sub
End Class

Friend Class CircularReferenceDetectedValidator
	Inherits TestComponent

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)

	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim fixture As FormulaEngineTestFixture = state.Item(GetType(FormulaEngineTestFixture))
		Assert.IsTrue(fixture.MyCircularReferenceFlag)
	End Sub
End Class

Friend Class NumDirectPrecedentsValidator
	Inherits TestComponent

	Private MyRef As IReference
	Private MyExpectedCount As Integer

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyRef = engine.ReferenceFactory.Parse(params(0))
		MyExpectedCount = Integer.Parse(params(1))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Assert.AreEqual(MyExpectedCount, engine.Info.GetDirectPrecedentsCount(MyRef))
	End Sub
End Class

Friend Class CellValueEqualsErrorValidator
	Inherits TestComponent

	Private MyError As ErrorValueType
	Private MyCellRef As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyCellRef = params(0)
		MyError = System.Enum.Parse(GetType(ErrorValueType), params(1))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim ref As ISheetReference = engine.ReferenceFactory.Parse(MyCellRef)
		Dim cellValue As Object = engine.Sheets.ActiveSheet.GetCellValue(ref.Row, ref.Column)
		Assert.AreEqual(FormulaEngine.CreateError(MyError), cellValue)
	End Sub
End Class

Friend Class FirstReferenceInvalidValidator
	Inherits TestComponent

	Private MyTargetRef As ISheetReference

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		If params.Length = 1 And params(0).Length > 0 Then
			MyTargetRef = engine.ReferenceFactory.Parse(params(0))
		End If
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim f As Formula

		If MyTargetRef Is Nothing Then
			Dim formulas As Formula() = state.Item("Formulas")
			f = formulas(0)
		Else
			f = engine.GetFormulaAt(MyTargetRef)
		End If

		Dim refs As IReference() = f.References
		Dim firstRef As IReference = refs(0)
		Assert.IsFalse(engine.Info.IsReferenceValid(firstRef))
	End Sub
End Class

Friend Class ReferenceValidValidator
	Inherits TestComponent

	Private MyReferences() As String
	Private MyValids() As Boolean

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		Dim len As Integer = (params.Length \ 2) - 1
		MyReferences = New String(len) {}
		MyValids = New Boolean(len) {}

		For i As Integer = 0 To len
			MyReferences(i) = params(i * 2)
			MyValids(i) = Boolean.Parse(params(i * 2 + 1))
		Next
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		For i As Integer = 0 To MyReferences.Length - 1
			Dim ref As IReference = engine.ReferenceFactory.Parse(MyReferences(i))
			Assert.AreEqual(MyValids(i), engine.Info.IsReferenceValid(ref))
		Next
	End Sub
End Class

Friend MustInherit Class HeaderManipulator
	Inherits TestComponent

	Protected MyManipulateAt As Integer
	Protected MyCount As Integer

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyManipulateAt = Integer.Parse(params(0))
		MyCount = Integer.Parse(params(1))
	End Sub
End Class

Friend Class RemoveColumnsExecutor
	Inherits HeaderManipulator

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		engine.OnColumnsRemoved(MyManipulateAt, MyCount)
	End Sub
End Class

Friend Class RemoveRowsExecutor
	Inherits HeaderManipulator

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		engine.OnRowsRemoved(MyManipulateAt, MyCount)
	End Sub
End Class

Friend Class InsertColumnsExecutor
	Inherits HeaderManipulator

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		engine.OnColumnsInserted(MyManipulateAt, MyCount)
	End Sub
End Class

Friend Class InsertRowsExecutor
	Inherits HeaderManipulator

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		engine.OnRowsInserted(MyManipulateAt, MyCount)
	End Sub
End Class

Friend Class MoveRangeExecutor
	Inherits TestComponent

	Private MySourceRange As IReference
	Private MyRowOffset, MyColOffset As Integer

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MySourceRange = engine.ReferenceFactory.Parse(params(0))
		MyRowOffset = Integer.Parse(params(1))
		MyColOffset = Integer.Parse(params(2))
	End Sub

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		engine.OnRangeMoved(MySourceRange, MyRowOffset, MyColOffset)
	End Sub
End Class

Friend Class CreateFormulaExecutor
	Inherits TestComponent

	Private Structure FormulaInfo
		Public Ref As ISheetReference
		Public Formula As String

		Public Sub New(ByVal ref As ISheetReference, ByVal formula As String)
			Me.Ref = ref
			Me.Formula = formula
		End Sub
	End Structure

	Private MyInfos As FormulaInfo()

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		Dim count As Integer = params.Length \ 2
		Dim infos(count - 1) As FormulaInfo

		For i As Integer = 0 To params.Length - 1 Step 2
			Dim ref As ISheetReference = engine.ReferenceFactory.Parse(params(i))
			Dim formula As String = params(i + 1)

			Dim info As New FormulaInfo(ref, formula)
			infos(i \ 2) = info
		Next

		MyInfos = infos
	End Sub

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim formulas(MyInfos.Length - 1) As Formula

		For i As Integer = 0 To MyInfos.Length - 1
			Dim info As FormulaInfo = MyInfos(i)
			Dim f As Formula = engine.AddFormula(info.Formula, info.Ref)
			formulas(i) = f
		Next

		state.Item("Formulas") = formulas
	End Sub
End Class

Friend Class DestroyFormulaExecutor
	Inherits TestComponent

	Private MyIndices As Integer()

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		Dim arr(params.Length - 1) As Integer

		For i As Integer = 0 To arr.Length - 1
			arr(i) = Integer.Parse(params(i))
		Next

		MyIndices = arr
	End Sub

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim formulas As Formula() = state.Item("Formulas")

		For Each index As Integer In MyIndices
			Dim f As Formula = formulas(index)
			engine.RemoveFormula(f)
		Next
	End Sub
End Class

Friend Class GetCalculationListExecutor
	Inherits TestComponent

	Private MyRange As IReference

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyRange = engine.ReferenceFactory.Parse(params(0))
	End Sub

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim calcList As Formula() = engine.Info.GetCalculationList(MyRange)
		state.Add("CalculationList", calcList)
	End Sub
End Class

Friend Class GetNamedCalculationListExecutor
	Inherits TestComponent

	Private MyName As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)
		MyName = params(0)
	End Sub

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim ref As INamedReference = engine.ReferenceFactory.Named(MyName)
		Dim calcList As Formula() = engine.Info.GetCalculationList(ref)
		state.Add("CalculationList", calcList)
	End Sub
End Class

Friend Class NullExecutor
	Inherits TestComponent

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As FormulaEngine)

	End Sub

	Public Overloads Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)

	End Sub
End Class

Friend Class ClearCellValueExecutor
	Inherits TestComponent

	Private MyCellRef As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyCellRef = params(0)
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim ref As ISheetReference = engine.ReferenceFactory.Parse(MyCellRef)
		DirectCast(engine.Sheets.ActiveSheet, FormulaEngineTestGrid).ClearCell(ref.Row, ref.Column)
	End Sub
End Class

Friend Class SetActiveSheetExecutor
	Inherits TestComponent

	Private MyTargetSheetName As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyTargetSheetName = params(0)
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim sheet As ISheet = engine.Sheets.GetSheetByName(MyTargetSheetName)
		engine.Sheets.ActiveSheet = sheet
	End Sub
End Class

Friend Class CopyFormulaExecutor
	Inherits TestComponent

	Private MyDestRef As ISheetReference

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyDestRef = engine.ReferenceFactory.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim formulas As Formula() = state.Item("Formulas")
		Dim target As Formula = formulas(0)
		engine.CopySheetFormula(target, MyDestRef)
	End Sub
End Class

Friend Class DestroyFormulasInRangeExecutor
	Inherits TestComponent

	Private MyRange As ISheetReference

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyRange = engine.ReferenceFactory.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		engine.RemoveFormulasInRange(MyRange)
	End Sub
End Class

Friend Class FormulaInvalidValidator
	Inherits TestComponent

	Private MyExpression As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyExpression = params(0)
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim invalid As Boolean = False

		Try
			engine.CreateFormula(MyExpression)
		Catch ex As InvalidFormulaException
			invalid = True
		End Try

		Assert.IsTrue(invalid)
	End Sub
End Class

Friend Class RemoveSheetExecutor
	Inherits TestComponent

	Private MySheetIndex As Integer

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MySheetIndex = Integer.Parse(params(0))
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim sheet As ISheet = engine.Sheets.Item(MySheetIndex)
		engine.Sheets.Remove(sheet)
	End Sub
End Class

Friend Class CreateNamedFormulaExecutor
	Inherits TestComponent

	Private MyName As String
	Private MyExpression As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyName = params(0)
		MyExpression = params(1)
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim ref As INamedReference = engine.ReferenceFactory.Named(MyName)
		Dim f As Formula = engine.CreateFormula(MyExpression)
		engine.AddFormula(f, ref)
	End Sub
End Class

Friend Class DestroyNamedFormulaExecutor
	Inherits TestComponent

	Private MyName As String

	Public Overrides Sub AcceptParameters(ByVal params() As String, ByVal engine As ciloci.FormulaEngine.FormulaEngine)
		MyName = params(0)
	End Sub

	Public Overrides Sub Execute(ByVal engine As ciloci.FormulaEngine.FormulaEngine, ByVal state As System.Collections.IDictionary)
		Dim ref As INamedReference = engine.ReferenceFactory.Named(MyName)
		engine.RemoveFormulaAt(ref)
	End Sub
End Class

Friend Structure TestComponentInfo
	Public Component As TestComponent
	Public Args As String()
End Structure