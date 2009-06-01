' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public License
' as published by the Free Software Foundation; either version 2.1
' of the License, or (at your option) any later version.
' 
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free
' Software Foundation, Inc., 59 Temple Place, Suite 330, Boston,
' MA 02111-1307, USA.
' 
' FormulaEngine - A library for parsing and managing formulas
' Copyright © 2007 Eugene Ciloci
'

''' <summary>
''' Implements all the functions that come standard with the formula engine
''' </summary>
<Serializable()> _
Friend Class BuiltinFunctions

	Private Const MAX_FACTORIAL As Integer = 170
	Private MyFactorialTable As Integer()
	Private MyRandom As Random

	Public Sub New()
		MyFactorialTable = Me.CreateFactorialTable()
		MyRandom = New Random
	End Sub

	Private Function CreateFactorialTable() As Integer()
		Dim arr(MAX_FACTORIAL - 1) As Integer

		For i As Integer = 0 To MAX_FACTORIAL - 1
			arr(i) = i + 1
		Next

		Return arr
	End Function

#Region "Statistical functions"
	<VariableArgumentFormulaFunction()> _
	Public Sub Stdev(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeStandardDeviation(processor.Values, True, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub StdevA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New AverageAProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeStandardDeviation(processor.Values, True, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub StdevP(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeStandardDeviation(processor.Values, False, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub StdevPA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New AverageAProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeStandardDeviation(processor.Values, False, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Var(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeVariance(processor.Values, True, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub VarA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New AverageAProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeVariance(processor.Values, True, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub VarP(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeVariance(processor.Values, False, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub VarPA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New AverageAProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeVariance(processor.Values, False, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Max(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Dim max As Double = Me.ComputeMax(processor.Values)
			result.SetValue(max)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub MaxA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New AverageAProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Dim max As Double = Me.ComputeMax(processor.Values)
			result.SetValue(max)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Min(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Dim min As Double = Me.ComputeMin(processor.Values)
			result.SetValue(min)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub MinA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New AverageAProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Dim min As Double = Me.ComputeMin(processor.Values)
			result.SetValue(min)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Average(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Me.ComputeAverageFunction(args, result, New SumProcessor())
	End Sub

	Private Sub ComputeAverageFunction(ByVal args As Argument(), ByVal result As FunctionResult, ByVal processor As DoubleBasedReferenceValueProcessor)
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			If processor.Values.Count = 0 Then
				result.SetError(ErrorValueType.Div0)
			Else
				result.SetValue(Me.ComputeAverage(processor.Values))
			End If
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub AverageA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Me.ComputeAverageFunction(args, result, New AverageAProcessor())
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Count(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New CountProcessor()
		processor.ProcessArguments(args)
		result.SetValue(processor.Count)
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub CountA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New CountAProcessor()
		processor.ProcessArguments(args)
		result.SetValue(processor.Count)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Reference})> _
	Public Sub CountBlank(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New CountBlankProcessor()
		processor.ProcessArguments(args)
		result.SetValue(processor.Count)
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.SheetReference, OperandType.Primitive})> _
	Public Sub CountIf(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim range As ISheetReference = args(0).ValueAsReference
		Dim criteria As Object = args(1).ValueAsPrimitive

		Dim pred As SheetValuePredicate = SheetValuePredicate.Create(criteria)
		Dim values As Object(,) = range.GetValuesTable()
		Dim processor As New CountIfConditionalSheetProcessor()
		Me.DoConditionalTableOp(values, processor, pred)
		result.SetValue(processor.Count)
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Mode(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeMode(processor.Values, result)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Median(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Me.ComputeMedian(processor.Values, result)
		End If
	End Sub
#End Region

#Region "Math functions"
	<VariableArgumentFormulaFunction()> _
	Public Sub Sum(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Dim sum As Double = Me.ComputeSum(processor.Values)
			result.SetValue(sum)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub SumSq(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Dim sum As Double = Me.ComputeSumOfSquares(processor.Values)
			result.SetValue(sum)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Product(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New SumProcessor()
		Dim success As Boolean = processor.ProcessArguments(args)

		If success = False Then
			result.SetError(processor.ErrorValue)
		Else
			Dim product As Double = Me.ComputeProduct(processor.Values)
			result.SetValue(product)
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Double, OperandType.Integer})> _
	Public Sub Round(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim decimals As Integer = args(1).ValueAsInteger

		number = Me.ComputeRound(number, decimals)

		result.SetValue(number)
	End Sub

	Private Function ComputeRound(ByVal number As Double, ByVal decimals As Integer) As Double
		If decimals >= 0 Then
			Return System.Math.Round(number, decimals)
		Else
			Dim scale As Double = System.Math.Pow(10, System.Math.Abs(decimals))
			Return System.Math.Round(number / scale) * scale
		End If
	End Function

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Sin(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(System.Math.Sin(args(0).ValueAsDouble))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Sinh(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(System.Math.Sinh(args(0).ValueAsDouble))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub ASin(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim d As Double = args(0).ValueAsDouble

		If d < -1 Or d > 1 Then
			result.SetError(ErrorValueType.Num)
		Else
			result.SetValue(System.Math.Asin(args(0).ValueAsDouble))
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Cos(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(System.Math.Cos(args(0).ValueAsDouble))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Cosh(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(System.Math.Cosh(args(0).ValueAsDouble))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub ACos(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim d As Double = args(0).ValueAsDouble

		If d < -1 Or d > 1 Then
			result.SetError(ErrorValueType.Num)
		Else
			result.SetValue(System.Math.Acos(args(0).ValueAsDouble))
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Tan(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(System.Math.Tan(args(0).ValueAsDouble))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Tanh(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(System.Math.Tanh(args(0).ValueAsDouble))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub ATan(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim d As Double = args(0).ValueAsDouble
		result.SetValue(System.Math.Atan(args(0).ValueAsDouble))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Abs(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim d As Double = args(0).ValueAsDouble
		result.SetValue(System.Math.Abs(d))
	End Sub

	<FixedArgumentFormulaFunction(0, New OperandType() {})> _
	Public Sub PI(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(System.Math.PI)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Degrees(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim rads As Double = args(0).ValueAsDouble
		Dim degs As Double = rads * (180 / System.Math.PI)
		result.SetValue(degs)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Radians(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim degs As Double = args(0).ValueAsDouble
		Dim rads As Double = degs * (System.Math.PI / 180)
		result.SetValue(rads)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Exp(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim d As Double = args(0).ValueAsDouble
		result.SetValue(System.Math.Exp(d))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Even(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim sign As Integer = System.Math.Sign(number)

		number = System.Math.Abs(number)
		number = System.Math.Ceiling(number)

		If number Mod 2 <> 0 Then
			number += 1
		End If

		result.SetValue(number * sign)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Odd(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim sign As Integer = System.Math.Sign(number)

		number = System.Math.Abs(number)
		number = System.Math.Ceiling(number)

		If number Mod 2 = 0 Then
			number += 1
		End If

		result.SetValue(number * sign)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Fact(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim n As Double = args(0).ValueAsDouble

		If n < 0 Or n > MAX_FACTORIAL Then
			result.SetError(ErrorValueType.Num)
		ElseIf n = 0 Then
			result.SetValue(1.0)
		Else
			result.SetValue(Me.ComputeFactorial(n))
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Double, OperandType.Double})> _
	Public Sub Ceiling(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double
		Dim significance As Double

		number = args(0).ValueAsDouble
		significance = args(1).ValueAsDouble

		If number = 0 Or significance = 0 Then
			result.SetValue(0)
			Return
		End If

		If System.Math.Sign(number) <> System.Math.Sign(significance) Then
			result.SetError(ErrorValueType.Num)
		Else
			Dim ceil As Double = System.Math.Ceiling(number / significance) * significance
			result.SetValue(ceil)
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Double, OperandType.Double})> _
	Public Sub Floor(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double
		Dim significance As Double

		number = args(0).ValueAsDouble
		significance = args(1).ValueAsDouble

		If number = 0 Or significance = 0 Then
			result.SetValue(0)
			Return
		End If

		If System.Math.Sign(number) <> System.Math.Sign(significance) Then
			result.SetError(ErrorValueType.Num)
		Else
			Dim floor As Double = System.Math.Floor(number / significance) * significance
			result.SetValue(floor)
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Int(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		result.SetValue(System.Math.Floor(number))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Ln(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble

		If number <= 0 Then
			result.SetError(ErrorValueType.Num)
		Else
			result.SetValue(System.Math.Log(number))
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Log10(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble

		If number <= 0 Then
			result.SetError(ErrorValueType.Num)
		Else
			result.SetValue(System.Math.Log10(number))
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.Double, OperandType.Double})> _
	Public Sub Log(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim base As Double = 10

		If args.Length > 1 Then
			base = args(1).ValueAsDouble
		End If

		If number <= 0 Or base <= 0 Then
			result.SetError(ErrorValueType.Num)
		ElseIf base = 1 Then
			result.SetError(ErrorValueType.Div0)
		Else
			result.SetValue(System.Math.Log(number) / System.Math.Log(base))
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Double, OperandType.Double})> _
	Public Sub [Mod](ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim divisor As Double = args(1).ValueAsDouble

		If divisor = 0 Then
			result.SetError(ErrorValueType.Div0)
		Else
			Dim value As Double = number - divisor * System.Math.Floor(number / divisor)
			result.SetValue(value)
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Double, OperandType.Double})> _
	Public Sub Power(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim power As Double = args(1).ValueAsDouble

		' Relying on result to handle NAN and infinity
		result.SetValue(System.Math.Pow(number, power))
	End Sub

	<FixedArgumentFormulaFunction(0, New OperandType() {}), VolatileFunction()> _
	Public Sub Rand(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(MyRandom.NextDouble())
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Integer, OperandType.Integer}), VolatileFunction()> _
	Public Sub Randbetween(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(MyRandom.Next(args(0).ValueAsInteger, args(1).ValueAsInteger))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Sign(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		result.SetValue(System.Math.Sign(number))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})> _
	Public Sub Sqrt(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble

		If number < 0 Then
			result.SetError(ErrorValueType.Num)
		Else
			result.SetValue(System.Math.Sqrt(number))
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.Double, OperandType.Integer})> _
	Public Sub Trunc(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim decimals As Integer = 0

		If args.Length > 1 Then
			decimals = args(1).ValueAsInteger
		End If

		Dim scale As Double = System.Math.Pow(10, decimals)

		number = number * scale
		number = System.Math.Truncate(number)
		number = number / scale

		result.SetValue(number)
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Double, OperandType.Integer})> _
	Public Sub RoundDown(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim decimals As Integer = args(1).ValueAsInteger

		Dim half As Double = 0.5 * System.Math.Pow(10, -decimals) * System.Math.Sign(number)
		number = number - half
		number = Me.ComputeRound(number, decimals)

		result.SetValue(number)
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Double, OperandType.Integer})> _
	Public Sub RoundUp(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim decimals As Integer = args(1).ValueAsInteger

		Dim half As Double = 0.5 * System.Math.Pow(10, -decimals) * System.Math.Sign(number)
		number = number + half
		number = Me.ComputeRound(number, decimals)

		result.SetValue(number)
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Integer, OperandType.Integer})> _
	Public Sub Combin(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim n As Integer = args(0).ValueAsInteger
		Dim k As Integer = args(1).ValueAsInteger

		If n < 0 Or k < 0 Or n < k Then
			result.SetError(ErrorValueType.Num)
		Else
			result.SetValue(Me.ComputeCombinations(n, k))
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, 3, New OperandType() {OperandType.SheetReference, OperandType.Primitive, OperandType.SheetReference})> _
	Public Sub SumIf(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim sourceRange As ISheetReference = args(0).ValueAsReference
		Dim criteria As Object = args(1).ValueAsPrimitive
		Dim sumRange As ISheetReference = sourceRange

		If args.Length = 3 Then
			sumRange = args(2).ValueAsReference
			Dim rect As System.Drawing.Rectangle = sourceRange.Area
			rect.Offset(sumRange.Area.Left - sourceRange.Area.Left, sumRange.Area.Top - sourceRange.Area.Top)

			If Utility.IsRectangleInSheet(rect, sumRange.Sheet) = False Then
				result.SetError(ErrorValueType.Ref)
				Return
			End If
			sumRange = engine.ReferenceFactory.FromRectangle(rect)
		End If

		Dim pred As SheetValuePredicate = SheetValuePredicate.Create(criteria)
		Dim sourceValues As Object(,) = sourceRange.GetValuesTable()
		Dim sumValues As Object(,) = sourceValues

		If Not sourceRange Is sumRange Then
			sumValues = sumRange.GetValuesTable()
		End If

		Dim processor As New SumIfConditionalSheetProcessor()
		processor.Initialize(sumValues)

		Me.DoConditionalTableOp(sourceValues, processor, pred)

		Dim sum As Double = Me.ComputeSum(processor.Values)
		result.SetValue(sum)
	End Sub
#End Region

#Region "Information functions"
	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsLogical(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		Dim isValue As Boolean = (arg.IsPrimitive = True) AndAlso (TypeOf (arg.ValueAsPrimitive) Is Boolean)
		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsNumber(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		Dim isValue As Boolean = (arg.IsPrimitive = True) AndAlso (TypeOf (arg.ValueAsPrimitive) Is Integer Or TypeOf (arg.ValueAsPrimitive) Is Double)
		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsText(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		Dim isValue As Boolean = (arg.IsPrimitive = True) AndAlso (TypeOf (arg.ValueAsPrimitive) Is String)
		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsBlank(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		Dim isValue As Boolean = (arg.IsPrimitive = True) AndAlso (arg.ValueAsPrimitive Is Nothing)
		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsError(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		' For whatever reason, IsError called with a non-cell range returns true....
		Dim isValue As Boolean = (arg.IsPrimitive = False) OrElse (TypeOf (arg.ValueAsPrimitive) Is ErrorValueType)
		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsErr(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		Dim isValue As Boolean

		If arg.IsPrimitive = False Then
			isValue = True
		ElseIf arg.IsError = False Then
			isValue = False
		Else
			isValue = arg.ValueAsError <> ErrorValueType.NA
		End If

		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsNA(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		Dim isValue As Boolean

		If arg.IsPrimitive = False Then
			isValue = False
		ElseIf arg.IsError = False Then
			isValue = False
		Else
			isValue = arg.ValueAsError = ErrorValueType.NA
		End If

		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsRef(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(args(0).IsReference)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub IsNonText(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)
		Dim isValue As Boolean = (arg.IsPrimitive = False) OrElse (Not TypeOf (arg.ValueAsPrimitive) Is String)
		result.SetValue(isValue)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub Type(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)

		If arg.IsPrimitive = False Then
			result.SetValue(16)
			Return
		End If

		Dim value As Object = arg.ValueAsPrimitive

		If value Is Nothing Then
			result.SetValue(1)
			Return
		End If

		Dim t As Type = value.GetType()
		Dim typeCode As Integer

		If Utility.IsNumericType(t) = True Then
			typeCode = 1
		ElseIf t Is GetType(String) Then
			typeCode = 2
		ElseIf t Is GetType(Boolean) Then
			typeCode = 4
		ElseIf t Is GetType(ErrorValueType) Then
			typeCode = 16
		Else
			Debug.Assert(False, "unknown type")
		End If

		result.SetValue(typeCode)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Self})> _
	Public Sub Error_Type(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim arg As Argument = args(0)

		Dim ev As ErrorValueType

		If arg.IsPrimitive = False Then
			ev = ErrorValueType.Value
		ElseIf arg.IsError = True Then
			ev = arg.ValueAsError
		Else
			result.SetError(ErrorValueType.NA)
			Return
		End If

		Dim values As ErrorValueType() = {ErrorValueType.Null, ErrorValueType.Div0, ErrorValueType.Value, ErrorValueType.Ref, ErrorValueType.Name, ErrorValueType.Num, ErrorValueType.NA}
		Dim index As Integer = System.Array.IndexOf(values, ev) + 1
		Debug.Assert(index <> -1, "unknown error value")
		result.SetValue(index)
	End Sub
#End Region

#Region "Text functions"
	<FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.String, OperandType.Integer})> _
	Public Sub Left(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String
		Dim count As Integer

		text = args(0).ValueAsString

		If args.Length = 1 Then
			count = 1
		Else
			count = args(1).ValueAsInteger
		End If

		If count < 0 Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		count = System.Math.Min(count, text.Length)

		text = text.Substring(0, count)
		result.SetValue(text)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Integer})> _
	Public Sub [Char](ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim code As Integer = args(0).ValueAsInteger

		If code <= 0 Or code > 65535 Then
			result.SetError(ErrorValueType.Value)
		Else
			Dim c As Char = System.Convert.ToChar(code)
			result.SetValue(c.ToString())
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub Clean(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s As String = args(0).ValueAsString
		s = System.Text.RegularExpressions.Regex.Replace(s, "\p{C}+", String.Empty)
		result.SetValue(s)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub Code(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s As String = args(0).ValueAsString

		If s.Length < 1 Then
			result.SetError(ErrorValueType.Value)
		Else
			Dim c As Char = args(0).ValueAsString.Chars(0)
			result.SetValue(System.Convert.ToInt32(c))
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.Double, OperandType.Integer})> _
	Public Sub Dollar(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim decimals As Integer

		If args.Length = 2 Then
			decimals = args(1).ValueAsInteger
		Else
			decimals = 2
		End If

		number = Me.ComputeRound(number, decimals)

		Dim format As String = "C"

		decimals = System.Math.Max(0, decimals)
		format = format & decimals

		result.SetValue(number.ToString(format))
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.String, OperandType.String})> _
	Public Sub Exact(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s1, s2 As String
		s1 = args(0).ValueAsString
		s2 = args(1).ValueAsString

		Dim equal As Boolean = String.Equals(s1, s2, StringComparison.Ordinal)
		result.SetValue(equal)
	End Sub

	<FixedArgumentFormulaFunction(2, 3, New OperandType() {OperandType.String, OperandType.String, OperandType.Integer})> _
	Public Sub Find(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim findText, withinText As String
		Dim startIndex As Integer

		findText = args(0).ValueAsString
		withinText = args(1).ValueAsString

		If args.Length = 3 Then
			startIndex = args(2).ValueAsInteger
		Else
			startIndex = 1
		End If

		If startIndex <= 0 Or startIndex > withinText.Length Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		startIndex -= 1

		Dim index As Integer = withinText.IndexOf(findText, startIndex, StringComparison.CurrentCulture)

		If index = -1 Then
			result.SetError(ErrorValueType.Value)
		Else
			result.SetValue(index + 1)
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, 3, New OperandType() {OperandType.Double, OperandType.Integer, OperandType.Boolean})> _
	Public Sub Fixed(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim number As Double = args(0).ValueAsDouble
		Dim decimals As Integer
		Dim noCommas As Boolean

		If args.Length > 1 Then
			decimals = args(1).ValueAsInteger
		Else
			decimals = 2
		End If

		If args.Length > 2 Then
			noCommas = args(2).ValueAsBoolean
		Else
			noCommas = False
		End If

		number = Me.ComputeRound(number, decimals)

		Dim format As String

		If noCommas = False Then
			format = "N"
		Else
			format = "F"
		End If

		decimals = System.Math.Max(0, decimals)
		format = format & decimals

		result.SetValue(number.ToString(format))
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub Len(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s As String = args(0).ValueAsString
		result.SetValue(s.Length)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub Lower(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s As String = args(0).ValueAsString
		result.SetValue(s.ToLower())
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub Upper(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s As String = args(0).ValueAsString
		result.SetValue(s.ToUpper())
	End Sub

	<FixedArgumentFormulaFunction(3, New OperandType() {OperandType.String, OperandType.Integer, OperandType.Integer})> _
	Public Sub Mid(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString
		Dim startIndex As Integer = args(1).ValueAsInteger
		Dim length As Integer = args(2).ValueAsInteger

		If startIndex < 1 Or length < 0 Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		startIndex -= 1

		Dim value As String

		If startIndex >= text.Length Then
			value = String.Empty
		Else
			length = System.Math.Min(length, text.Length - startIndex)
			value = text.Substring(startIndex, length)
		End If

		result.SetValue(value)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub Proper(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString
		Dim sb As New System.Text.StringBuilder(text)
		Dim firstLetter As Boolean = False

		For i As Integer = 0 To sb.Length - 1
			Dim c As Char = sb.Chars(i)
			If Char.IsLetter(c) = True Then
				If firstLetter = False Then
					firstLetter = True
					c = Char.ToUpper(c)
				ElseIf Char.IsLetter(sb.Chars(i - 1)) = False Then
					c = Char.ToUpper(c)
				Else
					c = Char.ToLower(c)
				End If

				sb.Chars(i) = c
			End If
		Next

		text = sb.ToString()
		result.SetValue(text)
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.String, OperandType.Integer})> _
	Public Sub Rept(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString
		Dim count As Integer = args(1).ValueAsInteger

		If count < 0 Or text.Length * count > Short.MaxValue Then
			result.SetError(ErrorValueType.Value)
		Else
			Dim sb As New System.Text.StringBuilder(text.Length * count)
			sb.Insert(0, text, count)
			result.SetValue(sb.ToString())
		End If
	End Sub

	<FixedArgumentFormulaFunction(4, New OperandType() {OperandType.String, OperandType.Integer, OperandType.Integer, OperandType.String})> _
	Public Sub Replace(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString
		Dim startIndex As Integer = args(1).ValueAsInteger
		Dim count As Integer = args(2).ValueAsInteger
		Dim newText As String = args(3).ValueAsString

		If startIndex < 1 Or count < 0 Then
			result.SetError(ErrorValueType.Value)
		Else
			startIndex -= 1
			count = System.Math.Min(count, text.Length)
			If startIndex > text.Length Then
				startIndex = text.Length
				count = 0
			End If

			Dim sb As New System.Text.StringBuilder(text)
			sb.Remove(startIndex, count)
			sb.Insert(startIndex, newText)
			result.SetValue(sb.ToString())
		End If
	End Sub

	'SUBSTITUTE(text,old_text,new_text,instance_num)
	<FixedArgumentFormulaFunction(3, 4, New OperandType() {OperandType.String, OperandType.String, OperandType.String, OperandType.Integer})> _
	Public Sub Substitute(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString
		Dim oldText As String = args(1).ValueAsString
		Dim newText As String = args(2).ValueAsString
		Dim instanceNum As Integer = -1

		If args.Length = 4 Then
			instanceNum = args(3).ValueAsInteger

			If instanceNum < 1 Then
				result.SetError(ErrorValueType.Value)
				Return
			Else
				instanceNum -= 1
			End If
		End If

		If oldText.Length = 0 Then
			result.SetValue(text)
			Return
		End If

		If instanceNum = -1 Then
			text = text.Replace(oldText, newText)
		Else
			Dim indices As Integer() = Me.GetIndicesOf(text, oldText)

			If instanceNum < indices.Length Then
				Dim index As Integer = indices(instanceNum)
				text = text.Remove(index, oldText.Length)
				text = text.Insert(index, newText)
			End If
		End If

		result.SetValue(text)
	End Sub

	Private Function GetIndicesOf(ByVal text As String, ByVal s As String) As Integer()
		Dim list As New Generic.List(Of Integer)

		Dim index As Integer = 0

		While index < text.Length
			Dim tempIndex As Integer = text.IndexOf(s, index)
			If tempIndex <> -1 Then
				list.Add(tempIndex)
				index = tempIndex + s.Length
			Else
				Exit While
			End If
		End While

		Dim arr(list.Count - 1) As Integer
		list.CopyTo(arr, 0)
		Return arr
	End Function

	<FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.String, OperandType.Integer})> _
	Public Sub Right(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString
		Dim count As Integer = 1

		If args.Length > 1 Then
			count = args(1).ValueAsInteger
		End If

		If count < 0 Then
			result.SetError(ErrorValueType.Value)
		Else
			count = System.Math.Min(count, text.Length)
			Dim startIndex As Integer = text.Length - count
			text = text.Substring(startIndex, count)
			result.SetValue(text)
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, New OperandType() {OperandType.Primitive, OperandType.String})> _
	Public Sub Text(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim value As Object = args(0).ValueAsPrimitive
		Dim format As String = args(1).ValueAsString
		Dim formattable As IFormattable = TryCast(value, IFormattable)

		If formattable Is Nothing Then
			result.SetError(ErrorValueType.Value)
			Return
		End If
		
		Try
			Dim text As String = formattable.ToString(format, Nothing)
			result.SetValue(text)
		Catch ex As FormatException
			result.SetError(ErrorValueType.Value)
		End Try
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub Trim(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString
		' This is an inefficient but quick way to do this
		text = System.Text.RegularExpressions.Regex.Replace(text, "^ +| +$", String.Empty)
		text = System.Text.RegularExpressions.Regex.Replace(text, " +", " ")
		result.SetValue(text)
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub Concatenate(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim sb As New System.Text.StringBuilder()
		For i As Integer = 0 To args.Length - 1
			Dim arg As Argument = args(i)
			If arg.IsError = True Then
				result.SetError(arg.ValueAsError)
				Return
			ElseIf arg.IsString = True Then
				sb.Append(arg.ValueAsString)
			Else
				result.SetError(ErrorValueType.Value)
				Return
			End If
		Next

		result.SetValue(sb.ToString())
	End Sub

	<FixedArgumentFormulaFunction(2, 3, New OperandType() {OperandType.String, OperandType.String, OperandType.Integer})> _
	Public Sub Search(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim findText, withinText As String
		Dim startIndex As Integer

		findText = args(0).ValueAsString
		withinText = args(1).ValueAsString

		If args.Length = 3 Then
			startIndex = args(2).ValueAsInteger
		Else
			startIndex = 1
		End If

		If startIndex <= 0 Or startIndex > withinText.Length Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		startIndex -= 1

		findText = Utility.Wildcard2Regex(findText)
		Dim re As New System.Text.RegularExpressions.Regex(findText, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
		Dim m As System.Text.RegularExpressions.Match = re.Match(withinText, startIndex)

		If m Is System.Text.RegularExpressions.Match.Empty Then
			result.SetError(ErrorValueType.Value)
		Else
			result.SetValue(m.Index + 1)
		End If
	End Sub
#End Region

#Region "Logical functions"
	<FixedArgumentFormulaFunction(2, 3, New OperandType() {OperandType.Boolean, OperandType.Self, OperandType.Self})> _
	Public Sub [If](ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim condition As Boolean = args(0).ValueAsBoolean

		If condition = True Then
			result.SetValue(args(1))
		Else
			If args.Length = 2 Then
				result.SetValue(False)
			Else
				result.SetValue(args(2))
			End If
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Boolean})> _
	Public Sub [Not](ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim b As Boolean = args(0).ValueAsBoolean
		result.SetValue(Not b)
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub [And](ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New LogicalFunctionProcessor()
		If processor.ProcessArguments(args) = False Then
			result.SetError(processor.ErrorValue)
			Return
		End If

		If processor.Values.Count = 0 Then
			result.SetError(ErrorValueType.Value)
		Else
			Dim b As Boolean = Me.ComputeAnd(processor.Values)
			result.SetValue(b)
		End If
	End Sub

	<VariableArgumentFormulaFunction()> _
	Public Sub [Or](ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim processor As New LogicalFunctionProcessor()
		If processor.ProcessArguments(args) = False Then
			result.SetError(processor.ErrorValue)
			Return
		End If

		If processor.Values.Count = 0 Then
			result.SetError(ErrorValueType.Value)
		Else
			Dim b As Boolean = Me.ComputeOr(processor.Values)
			result.SetValue(b)
		End If
	End Sub
#End Region

#Region "Lookup functions"
	<VariableArgumentFormulaFunction()> _
	 Public Sub Choose(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		If args(0).IsInteger = False Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		Dim index As Integer = args(0).ValueAsInteger

		If index < 1 Or index > args.Length - 1 Then
			result.SetError(ErrorValueType.Value)
		Else
			result.SetValue(args(index))
		End If
	End Sub

	<FixedArgumentFormulaFunction(2, 5, New OperandType() {OperandType.Integer, OperandType.Integer, OperandType.Integer, OperandType.Boolean, OperandType.String})> _
	Public Sub Address(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim rowIndex As Integer = args(0).ValueAsInteger
		Dim columnIndex As Integer = args(1).ValueAsInteger

		Dim refType As Integer = 1
		Dim a1Style As Boolean = True

		If args.Length > 2 Then
			refType = args(2).ValueAsInteger
		End If

		If refType < 1 Or refType > 4 Or SheetReference.IsValidColumnIndex(columnIndex) = False Or SheetReference.IsValidRowIndex(rowIndex) = False Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		Dim sheetText As String = Me.CreateSheetString(args)

		Dim rowDollar As String = String.Empty
		Dim colDollar As String = String.Empty

		If refType = 1 Or refType = 2 Then
			rowDollar = "$"
		End If

		If refType = 1 Or refType = 3 Then
			colDollar = "$"
		End If

		Dim address As String = String.Format("{0}{1}{2}{3}{4}", sheetText, colDollar, SheetReference.ColumnIndex2Label(columnIndex), rowDollar, rowIndex)
		result.SetValue(address)
	End Sub

	Private Function CreateSheetString(ByVal args As Argument()) As String
		If args.Length < 5 Then
			Return String.Empty
		End If

		Dim s As String = args(4).ValueAsString

		If s.IndexOf(" "c) <> -1 Then
			Return String.Format("'{0}'!", s)
		Else
			Return String.Concat(s, "!")
		End If
	End Function

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.SheetReference})> _
	Public Sub Column(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim ref As ISheetReference = args(0).ValueAsReference
		result.SetValue(ref.Column)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.SheetReference})> _
	Public Sub Row(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim ref As ISheetReference = args(0).ValueAsReference
		result.SetValue(ref.Row)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.SheetReference})> _
	Public Sub Columns(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim ref As ISheetReference = args(0).ValueAsReference
		result.SetValue(ref.Width)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.SheetReference})> _
	Public Sub Rows(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim ref As ISheetReference = args(0).ValueAsReference
		result.SetValue(ref.Height)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.SheetReference})> _
	Public Sub Areas(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(1)
	End Sub

	<FixedArgumentFormulaFunction(3, 4, New OperandType() {OperandType.SheetReference, OperandType.Integer, OperandType.Integer, OperandType.Integer})> _
	Public Sub Index(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim ref As ISheetReference = args(0).ValueAsReference
		Dim rowOffset As Integer = args(1).ValueAsInteger
		Dim colOffset As Integer = args(2).ValueAsInteger
		Dim refRect As System.Drawing.Rectangle = ref.Area

		If rowOffset < 0 Or colOffset < 0 Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		Dim height, width As Integer
		Dim x, y As Integer

		If rowOffset = 0 Then
			height = refRect.Height
			y = 0
		Else
			height = 1
			y = rowOffset - 1
		End If

		If colOffset = 0 Then
			width = refRect.Width
			x = 0
		Else
			width = 1
			x = colOffset - 1
		End If

		Dim rect As New System.Drawing.Rectangle(refRect.Left + x, refRect.Top + y, width, height)

		If refRect.Contains(rect) = False Then
			result.SetError(ErrorValueType.Ref)
		Else
			result.SetValue(engine.ReferenceFactory.FromRectangle(rect))
		End If
	End Sub

	<FixedArgumentFormulaFunction(3, 5, New OperandType() {OperandType.SheetReference, OperandType.Integer, OperandType.Integer, OperandType.Integer, OperandType.Integer}), VolatileFunction()> _
	Public Sub Offset(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim ref As ISheetReference = args(0).ValueAsReference
		Dim rowOffset As Integer = args(1).ValueAsInteger
		Dim colOffset As Integer = args(2).ValueAsInteger
		Dim refRect As System.Drawing.Rectangle = ref.Area
		Dim height As Integer = refRect.Height
		Dim width As Integer = refRect.Width

		If args.Length > 3 Then
			height = args(3).ValueAsInteger
		End If

		If args.Length > 4 Then
			width = args(4).ValueAsInteger
		End If

		If width = 0 Or height = 0 Then
			result.SetError(ErrorValueType.Ref)
			Return
		ElseIf width < 0 Or height < 0 Then
			result.SetError(ErrorValueType.Value)
			Return
		End If

		refRect.Offset(colOffset, rowOffset)
		Dim newRect As New System.Drawing.Rectangle(refRect.Left, refRect.Top, width, height)
		If SheetReference.IsRectangleInSheet(newRect, ref.Sheet) = False Then
			result.SetError(ErrorValueType.Ref)
		Else
			result.SetValue(engine.ReferenceFactory.FromRectangle(newRect))
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.String, OperandType.Boolean}), VolatileFunction()> _
	Public Sub Indirect(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim text As String = args(0).ValueAsString

		Try
			Dim ref As ISheetReference = engine.ReferenceFactory.Parse(text)
			result.SetValue(DirectCast(ref, IReference))
		Catch ex As Exception
			result.SetError(ErrorValueType.Ref)
		End Try
	End Sub

	<FixedArgumentFormulaFunction(3, 4, New OperandType() {OperandType.Primitive, OperandType.SheetReference, OperandType.Integer, OperandType.Boolean})> _
	Public Sub HLookup(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Me.DoHVLookup(args, result, New HLookupProcessor())
	End Sub

	<FixedArgumentFormulaFunction(3, 4, New OperandType() {OperandType.Primitive, OperandType.SheetReference, OperandType.Integer, OperandType.Boolean})> _
	Public Sub VLookup(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Me.DoHVLookup(args, result, New VLookupProcessor())
	End Sub

	<FixedArgumentFormulaFunction(3, New OperandType() {OperandType.Primitive, OperandType.SheetReference, OperandType.SheetReference})> _
	Public Sub Lookup(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim lookupValue As Object = args(0).ValueAsPrimitive
		Dim lookupRef As ISheetReference = args(1).ValueAsReference
		Dim resultRef As ISheetReference = args(2).ValueAsReference

		If lookupRef.Area.Size <> resultRef.Area.Size Then
			result.SetError(ErrorValueType.Ref)
			Return
		End If

		Dim lookupTable As Object(,) = lookupRef.GetValuesTable()
		Dim resultTable As Object(,) = resultRef.GetValuesTable()

		Dim lookupVector As Object()
		Dim resultVector As Object()

		If lookupTable.GetLength(0) = 1 Then
			lookupVector = Utility.GetTableRow(lookupTable, 0)
			resultVector = Utility.GetTableRow(resultTable, 0)
		Else
			lookupVector = Utility.GetTableColumn(lookupTable, 0)
			resultVector = Utility.GetTableColumn(resultTable, 0)
		End If

		Dim processor As New PlainLookupProcessor()
		processor.Initialize(lookupVector, resultVector)

		Me.DoGenericLookup(processor, lookupValue, False, result)
	End Sub

	<FixedArgumentFormulaFunction(2, 3, New OperandType() {OperandType.Primitive, OperandType.SheetReference, OperandType.Integer})> _
	Public Sub Match(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim lookupValue As Object = args(0).ValueAsPrimitive
		Dim lookupRef As ISheetReference = args(1).ValueAsReference
		Dim matchType As Integer = 1

		If args.Length = 3 Then
			matchType = args(2).ValueAsInteger
		End If

		Dim lookupTable As Object(,) = lookupRef.GetValuesTable()
		Dim lookupVector As Object() = Utility.GetTableColumn(lookupTable, 0)

		Dim processor As New PlainLookupProcessor()
		processor.Initialize(lookupVector, lookupVector)

		Dim matchIndex As Integer
		Dim predicate As SheetValuePredicate = Me.CreateLookupPredicate(lookupValue)

		If matchType > 0 Then
			predicate.SetCompareType(CompareType.LessThanOrEqual)
			matchIndex = Me.IndexOfMatch(lookupVector, predicate, True)
		ElseIf matchType = 0 Then
			predicate.SetCompareType(CompareType.Equal)
			matchIndex = Me.IndexOfMatch(lookupVector, predicate, True)
		Else
			predicate.SetCompareType(CompareType.GreaterThanOrEqual)
			matchIndex = Me.IndexOfMatch(lookupVector, predicate, False)
		End If

		If matchIndex = -1 Then
			result.SetError(ErrorValueType.NA)
		Else
			result.SetValue(matchIndex + 1)
		End If
	End Sub

	Private Sub DoHVLookup(ByVal args As Argument(), ByVal result As FunctionResult, ByVal processor As HVLookupProcessor)
		Dim lookupValue As Object = args(0).ValueAsPrimitive
		Dim tableRef As ISheetReference = args(1).ValueAsReference
		Dim index As Integer = args(2).ValueAsInteger
		Dim exactMatch As Boolean = False

		If lookupValue Is Nothing Then
			result.SetError(ErrorValueType.NA)
			Return
		ElseIf args(0).IsError = True Then
			result.SetValue(args(0))
			Return
		End If

		If args.Length = 4 Then
			exactMatch = Not args(3).ValueAsBoolean
		End If

		Dim refRect As System.Drawing.Rectangle = tableRef.Area

		If index < 1 Then
			result.SetError(ErrorValueType.Value)
			Return
		ElseIf processor.IsValidIndex(refRect, index) = False Then
			result.SetError(ErrorValueType.Ref)
			Return
		End If

		Dim table As Object(,) = tableRef.GetValuesTable()
		processor.Initialize(table, index)

		Me.DoGenericLookup(processor, lookupValue, exactMatch, result)
	End Sub

	Private Sub DoGenericLookup(ByVal processor As LookupProcessor, ByVal lookupValue As Object, ByVal exactMatch As Boolean, ByVal result As FunctionResult)
		Dim lookupVector As Object() = processor.GetLookupVector()
		Dim predicate As SheetValuePredicate = Me.CreateLookupPredicate(lookupValue)
		Dim matchIndex As Integer

		predicate.SetCompareType(CompareType.Equal)
		matchIndex = Me.IndexOfMatch(lookupVector, predicate, True)

		If matchIndex = -1 And exactMatch = False Then
			predicate.SetCompareType(CompareType.LessThan)
			matchIndex = Me.IndexOfMatch(lookupVector, predicate, True)
		End If

		If matchIndex = -1 Then
			result.SetError(ErrorValueType.NA)
		Else
			Dim resultVector As Object() = processor.GetResultVector()
			result.SetValueFromSheet(resultVector(matchIndex))
		End If
	End Sub

	Private Function IndexOfMatch(ByVal line As Object(), ByVal predicate As SheetValuePredicate, ByVal useGreatest As Boolean) As Integer
		Dim matches As IList = New ArrayList

		For i As Integer = 0 To line.Length - 1
			Dim value As Object = line(i)
			value = Utility.NormalizeIfNumericValue(value)
			line(i) = value
			If predicate.IsMatch(value) = True Then
				matches.Add(value)
			End If
		Next

		If matches.Count = 0 Then
			Return -1
		End If

		Dim arr(matches.Count - 1) As Object
		matches.CopyTo(arr, 0)
		System.Array.Sort(arr)

		Dim nextIndex As Integer

		If useGreatest = True Then
			nextIndex = arr.Length - 1
		Else
			nextIndex = 0
		End If

		Dim nextValue As Object = arr(nextIndex)
		Return System.Array.IndexOf(line, nextValue)
	End Function

	Private Function CreateLookupPredicate(ByVal targetValue As Object) As SheetValuePredicate
		Dim s As String = TryCast(targetValue, String)
		Dim pred As SheetValuePredicate

		If s Is Nothing Then
			pred = New ComparerBasedPredicate(targetValue, New Comparer(System.Globalization.CultureInfo.CurrentCulture))
		ElseIf s.IndexOfAny(New Char() {"?", "*"}) <> -1 Then
			pred = New WildcardPredicate(s)
		Else
			pred = New ComparerBasedPredicate(s, StringComparer.OrdinalIgnoreCase)
		End If

		pred.SetCompareType(CompareType.Equal)
		Return pred
	End Function
#End Region

#Region "Date and time functions"
	<FixedArgumentFormulaFunction(3, New OperandType() {OperandType.Integer, OperandType.Integer, OperandType.Integer})> _
	Public Sub [Date](ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim day, month, year As Integer
		year = args(0).ValueAsInteger
		month = args(1).ValueAsInteger
		day = args(2).ValueAsInteger

		If year < 1 Or year > 9999 Then
			result.SetError(ErrorValueType.Value)
		Else
			Dim dt As Nullable(Of DateTime) = Me.CreateDate(year, month, day)
			If dt.HasValue = True Then
				result.SetValue(dt.Value)
			Else
				result.SetError(ErrorValueType.Value)
			End If
		End If
	End Sub

	Private Function CreateDate(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer) As Nullable(Of DateTime)
		Try
			Dim dt As New DateTime(year, 1, 1)
			dt = dt.AddMonths(month - 1)
			dt = dt.AddDays(day - 1)
			Return dt
		Catch ex As ArgumentOutOfRangeException
			Return Nothing
		End Try
	End Function

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub DateValue(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s As String = args(0).ValueAsString
		Dim dt As DateTime

		If DateTime.TryParseExact(s, New String() {"d", "D"}, Nothing, Globalization.DateTimeStyles.AllowWhiteSpaces = Globalization.DateTimeStyles.None, dt) = True Then
			result.SetValue(dt)
		Else
			result.SetError(ErrorValueType.Value)
		End If
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.DateTime})> _
	Public Sub Day(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dt As DateTime = args(0).ValueAsDateTime
		result.SetValue(dt.Day)
	End Sub

	<FixedArgumentFormulaFunction(2, 3, New OperandType() {OperandType.DateTime, OperandType.DateTime, OperandType.Boolean})> _
	Public Sub Days360(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dateStart As DateTime = args(0).ValueAsDateTime
		Dim dateEnd As DateTime = args(1).ValueAsDateTime
		Dim useEuropeanMethod As Boolean = False

		If args.Length = 3 Then
			useEuropeanMethod = args(2).ValueAsBoolean
		End If

		Dim daysStart, daysEnd As Integer
		Dim daysInStartMonth As Integer = DateTime.DaysInMonth(dateStart.Year, dateStart.Month)
		Dim daysInEndMonth As Integer = DateTime.DaysInMonth(dateEnd.Year, dateEnd.Month)
		daysStart = dateStart.Day
		daysEnd = dateEnd.Day

		If useEuropeanMethod = True Then
			daysEnd = System.Math.Min(30, dateEnd.Day)
		Else
			If dateStart.Month = 2 And dateStart.Day = daysInStartMonth Then
				daysStart = 30
				If dateEnd.Month = 2 And dateEnd.Day = daysInEndMonth Then
					daysEnd = 30
				End If
			End If

			If daysEnd = 31 And daysStart >= 30 Then
				daysEnd = 30
			End If
		End If

		daysStart = System.Math.Min(30, daysStart)

		result.SetValue(Me.ComputeDays360(dateStart, dateEnd, daysStart, daysEnd))
	End Sub

	Private Function ComputeDays360(ByVal dateStart As DateTime, ByVal dateEnd As DateTime, ByVal daysStart As Integer, ByVal daysEnd As Integer) As Integer
		Dim sign As Integer = 1

		If DateTime.Compare(dateStart, dateEnd) = 1 Then
			Dim dt As DateTime = dateStart
			dateStart = dateEnd
			dateEnd = dt
			Dim days As Integer = daysStart
			daysStart = daysEnd
			daysEnd = days
			sign = -1
		End If

		daysStart = dateStart.Year * 360 + dateStart.Month * 30 + daysStart
		daysEnd = dateEnd.Year * 360 + dateEnd.Month * 30 + daysEnd

		Dim result As Integer = (daysEnd - daysStart) * sign
		Return result
	End Function

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.DateTime})> _
	Public Sub Hour(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dt As DateTime = args(0).ValueAsDateTime
		result.SetValue(dt.Hour)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.DateTime})> _
	Public Sub Minute(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dt As DateTime = args(0).ValueAsDateTime
		result.SetValue(dt.Minute)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.DateTime})> _
	Public Sub Month(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dt As DateTime = args(0).ValueAsDateTime
		result.SetValue(dt.Month)
	End Sub

	<FixedArgumentFormulaFunction(0, New OperandType() {}), VolatileFunction()> _
	Public Sub Now(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(DateTime.Now)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.DateTime})> _
	Public Sub Second(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dt As DateTime = args(0).ValueAsDateTime
		result.SetValue(dt.Second)
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.DateTime})> _
	Public Sub Year(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dt As DateTime = args(0).ValueAsDateTime
		result.SetValue(dt.Year)
	End Sub

	<FixedArgumentFormulaFunction(3, New OperandType() {OperandType.Integer, OperandType.Integer, OperandType.Integer})> _
	Public Sub Time(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim hours As Integer = args(0).ValueAsInteger
		Dim minutes As Integer = args(1).ValueAsInteger
		Dim seconds As Integer = args(2).ValueAsInteger

		Try
			Dim ts As New TimeSpan(hours, minutes, seconds)
			Dim dt As DateTime = DateTime.Today
			dt = dt.Add(ts)
			result.SetValue(dt)
		Catch ex As ArgumentOutOfRangeException
			result.SetError(ErrorValueType.Value)
		End Try
	End Sub

	<FixedArgumentFormulaFunction(1, New OperandType() {OperandType.String})> _
	Public Sub TimeValue(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim s As String = args(0).ValueAsString
		Dim dt As DateTime

		If DateTime.TryParseExact(s, New String() {"t", "T"}, Nothing, Globalization.DateTimeStyles.AllowWhiteSpaces = Globalization.DateTimeStyles.None, dt) = True Then
			result.SetValue(dt)
		Else
			result.SetError(ErrorValueType.Value)
		End If
	End Sub

	<FixedArgumentFormulaFunction(0, New OperandType() {}), VolatileFunction()> _
	Public Sub Today(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		result.SetValue(DateTime.Today)
	End Sub

	<FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.DateTime, OperandType.Integer})> _
	Public Sub Weekday(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
		Dim dt As DateTime = args(0).ValueAsDateTime
		Dim returnType As Integer = 1

		If args.Length = 2 Then
			returnType = args(1).ValueAsInteger
		End If

		If returnType < 1 Or returnType > 3 Then
			result.SetError(ErrorValueType.Num)
			Return
		End If

		Dim arr1 As DayOfWeek() = {DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday}
		Dim arr2 As DayOfWeek() = {DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday}
		Dim index As Integer

		If returnType = 1 Then
			index = System.Array.IndexOf(Of DayOfWeek)(arr1, dt.DayOfWeek)
			index += 1
		ElseIf returnType = 2 Then
			index = System.Array.IndexOf(Of DayOfWeek)(arr2, dt.DayOfWeek)
			index += 1
		ElseIf returnType = 3 Then
			index = System.Array.IndexOf(Of DayOfWeek)(arr2, dt.DayOfWeek)
		End If

		result.SetValue(index)
	End Sub
#End Region

#Region "Utility"
	Private Function ComputeCombinations(ByVal n As Integer, ByVal k As Integer) As Double
		Dim C As Double = 1

		For i As Integer = 0 To k - 1
			C *= (n - i) / (k - i)
		Next

		Return C
	End Function

	Private Function ComputeSum(ByVal values As System.Collections.Generic.IList(Of Double)) As Double
		Dim sum As Double = 0.0

		For i As Integer = 0 To values.Count - 1
			Dim d As Double = values.Item(i)
			sum += d
		Next

		Return sum
	End Function

	Private Function ComputeSumOfSquares(ByVal values As System.Collections.Generic.IList(Of Double)) As Double
		Dim sum As Double = 0.0

		For i As Integer = 0 To values.Count - 1
			Dim d As Double = values.Item(i)
			sum += d ^ 2
		Next

		Return sum
	End Function

	Private Function ComputeVariance(ByVal values As System.Collections.Generic.IList(Of Double), ByVal isSample As Boolean, ByVal result As FunctionResult) As Nullable(Of Double)
		Dim count As Integer = values.Count

		If count = 0 Or (count = 1 And isSample = True) Then
			result.SetError(ErrorValueType.Div0)
			Return Nothing
		End If

		Dim mean As Double = Me.ComputeAverage(values)
		Dim sum As Double

		For i As Integer = 0 To count - 1
			Dim x As Double = values.Item(i)
			sum += (x - mean) ^ 2
		Next

		If isSample = True Then
			count -= 1
		End If
		Dim var As Double = sum / count
		result.SetValue(var)
		Return var
	End Function

	Private Sub ComputeStandardDeviation(ByVal values As System.Collections.Generic.IList(Of Double), ByVal isSample As Boolean, ByVal result As FunctionResult)
		Dim variance As Nullable(Of Double) = Me.ComputeVariance(values, isSample, result)
		If variance.HasValue = True Then
			result.SetValue(System.Math.Sqrt(variance.Value))
		End If
	End Sub

	Private Sub ComputeMode(ByVal values As System.Collections.Generic.IList(Of Double), ByVal result As FunctionResult)
		Dim frequencies As New Generic.Dictionary(Of Double, Integer)

		For i As Integer = 0 To values.Count - 1
			Dim value As Double = values.Item(i)
			If frequencies.ContainsKey(value) = False Then
				frequencies.Add(value, 1)
			Else
				frequencies.Item(value) += 1
			End If
		Next

		Dim modePair As Generic.KeyValuePair(Of Double, Integer)
		Dim max As Integer = 1

		For Each pair As Generic.KeyValuePair(Of Double, Integer) In frequencies
			If pair.Value > max Then
				modePair = pair
				max = pair.Value
			End If
		Next

		If max = 1 Then
			result.SetError(ErrorValueType.NA)
		Else
			result.SetValue(modePair.Key)
		End If
	End Sub

	Private Sub ComputeMedian(ByVal values As System.Collections.Generic.IList(Of Double), ByVal result As FunctionResult)
		If values.Count = 0 Then
			result.SetError(ErrorValueType.Num)
			Return
		End If

		Dim arr(values.Count - 1) As Double
		values.CopyTo(arr, 0)
		System.Array.Sort(Of Double)(arr)

		Dim median As Double

		If (arr.Length And 1) = 1 Then
			median = arr((arr.Length - 1) \ 2)
		Else
			Dim d1 As Double = arr(arr.Length \ 2)
			Dim d2 As Double = arr((arr.Length \ 2) - 1)
			median = (d1 + d2) / 2
		End If

		result.SetValue(median)
	End Sub

	Private Function ComputeProduct(ByVal values As System.Collections.Generic.IList(Of Double)) As Double
		If values.Count = 0 Then
			Return 0.0
		End If

		Dim product As Double = 1

		For i As Integer = 0 To values.Count - 1
			Dim d As Double = values.Item(i)
			product *= d
		Next

		Return product
	End Function

	Private Function ComputeAverage(ByVal values As System.Collections.Generic.IList(Of Double)) As Double
		Dim sum As Double = Me.ComputeSum(values)
		Return sum / values.Count
	End Function

	Private Function ComputeMax(ByVal values As System.Collections.Generic.IList(Of Double)) As Double
		If values.Count = 0 Then
			Return 0.0
		End If

		Dim max As Double = Double.MinValue

		For i As Integer = 0 To values.Count - 1
			Dim d As Double = values.Item(i)
			If d > max Then
				max = d
			End If
		Next

		Return max
	End Function

	Private Function ComputeMin(ByVal values As System.Collections.Generic.IList(Of Double)) As Double
		If values.Count = 0 Then
			Return 0.0
		End If

		Dim min As Double = Double.MaxValue

		For i As Integer = 0 To values.Count - 1
			Dim d As Double = values.Item(i)
			If d < min Then
				min = d
			End If
		Next

		Return min
	End Function

	Private Function ComputeFactorial(ByVal n As Double) As Double
		Dim product As Double = 1
		n = System.Math.Truncate(n)

		For i As Integer = 0 To n - 1
			product = product * MyFactorialTable(i)
		Next

		Return product
	End Function

	Private Function ComputeAnd(ByVal values As System.Collections.Generic.IList(Of Boolean)) As Boolean
		For i As Integer = 0 To values.Count - 1
			Dim b As Boolean = values.Item(i)
			If b = False Then
				Return False
			End If
		Next
		Return True
	End Function

	Private Function ComputeOr(ByVal values As System.Collections.Generic.IList(Of Boolean)) As Boolean
		For i As Integer = 0 To values.Count - 1
			Dim b As Boolean = values.Item(i)
			If b = True Then
				Return True
			End If
		Next
		Return False
	End Function

	Private Sub DoConditionalTableOp(ByVal compareValues As Object(,), ByVal processor As ConditionalSheetProcessor, ByVal predicate As SheetValuePredicate)
		For row As Integer = 0 To compareValues.GetUpperBound(0)
			For col As Integer = 0 To compareValues.GetUpperBound(1)
				Dim compareValue As Object = compareValues(row, col)
				If predicate.IsMatch(compareValue) = True Then
					processor.OnMatch(row, col)
				End If
			Next
		Next
	End Sub
#End Region
End Class