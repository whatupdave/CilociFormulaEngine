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
''' Processes the parse tree and generates the information necessary to build a formula
''' </summary>
Friend Class CustomExcelAnalyzer
	Inherits ExcelAnalyzer
	Implements IAnalyzer

	Private MyReferenceInfos As IList

	Public Sub New()
		MyReferenceInfos = New ArrayList
	End Sub

	Public Sub Reset() Implements IAnalyzer.Reset
		MyReferenceInfos.Clear()
	End Sub

	Public Overrides Function ExitFormula(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim childValues As IList = Me.GetChildValues(node)
		Debug.Assert(childValues.Count = 1 Or childValues.Count = 2, "Should have 1 or 2 values")
		If childValues.Count = 2 Then
			' Remove leading EQ expression
			childValues.RemoveAt(0)
		End If
		node.AddValues(childValues)
		Return node
	End Function

	Public Overrides Function ExitScalarformula(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitPrimaryExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitLogicalExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Me.AddBinaryOperation(node)
		Return node
	End Function

	Public Overrides Function ExitLogicalOp(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitConcatExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Me.AddBinaryOperation(node)
		Return node
	End Function

	Public Overrides Function ExitAdditiveExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Me.AddBinaryOperation(node)
		Return node
	End Function

	Public Overrides Function ExitAdditiveOp(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitMultiplicativeExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Me.AddBinaryOperation(node)
		Return node
	End Function

	Public Overrides Function ExitMultiplicativeOp(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitExponentiationExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Me.AddBinaryOperation(node)
		Return node
	End Function

	Public Overrides Function ExitPercentExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim values As IList = Me.GetChildValues(node)
		Dim first As ParseTreeElement = values.Item(0)

		If values.Count > 1 Then
			Dim pe As New PercentOperator(values.Count - 1)
			node.AddValue(New UnaryOperatorElement(pe, first))
		Else
			node.AddValue(first)
		End If

		Return node
	End Function

	Public Overrides Function ExitUnaryExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim values As IList = Me.GetChildValues(node)
		Dim last As ParseTreeElement = values.Item(values.Count - 1)

		Dim negCount As Integer

		For i As Integer = 0 To values.Count - 2
			Dim op As Object = values.Item(i)
			If op.GetType() Is GetType(BinarySubOperator) Then
				negCount += 1
			End If
		Next

		If (negCount And 1) <> 0 Then
			Dim ne As New UnaryNegateOperator
			node.AddValue(New UnaryOperatorElement(ne, last))
		Else
			node.AddValue(last)
		End If

		Return node
	End Function

	Public Overrides Function ExitBasicExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitReference(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim values As IList = MyBase.GetChildValues(node)
		Dim oe As New ContainerElement(values.Item(0))
		node.AddValue(oe)
		Return node
	End Function

	Public Overrides Function ExitExpressionGroup(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitFunctionCall(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim fe As New FunctionCallElement
		fe.AcceptValues(Me.GetChildValues(node))
		node.AddValue(fe)
		Return node
	End Function

	Public Overrides Function ExitArgumentList(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitPrimitive(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim values As IList = MyBase.GetChildValues(node)
		Dim oe As New ContainerElement(values.Item(0))
		node.AddValue(oe)
		Return node
	End Function

	Public Overrides Function ExitAdd(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New BinaryAddOperator)
		Return node
	End Function

	Public Overrides Function ExitSub(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New BinarySubOperator)
		Return node
	End Function

	Public Overrides Function ExitMul(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New BinaryMultiplyOperator)
		Return node
	End Function

	Public Overrides Function ExitDiv(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New BinaryDivisionOperator)
		Return node
	End Function

	Public Overrides Function ExitConcat(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ConcatenationOperator)
		Return node
	End Function

	Public Overrides Function ExitExp(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New BinaryPowerOperator)
		Return node
	End Function

	Public Overrides Function ExitEQ(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New LogicalOperator(CompareType.Equal))
		Return node
	End Function

	Public Overrides Function ExitNE(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New LogicalOperator(CompareType.NotEqual))
		Return node
	End Function

	Public Overrides Function ExitLT(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New LogicalOperator(CompareType.LessThan))
		Return node
	End Function

	Public Overrides Function ExitGT(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New LogicalOperator(CompareType.GreaterThan))
		Return node
	End Function

	Public Overrides Function ExitLTE(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New LogicalOperator(CompareType.LessThanOrEqual))
		Return node
	End Function

	Public Overrides Function ExitGTE(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New LogicalOperator(CompareType.GreaterThanOrEqual))
		Return node
	End Function

	Public Overrides Function ExitPercent(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(node.Image)
		Return node
	End Function

	Public Overrides Function ExitNumber(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		Dim value As Double
		Dim op As PrimitiveOperand

		' Try to store the number as an integer if possible
		Dim success As Boolean = Double.TryParse(node.Image, Globalization.NumberStyles.Integer, Nothing, value)

		If success = True Then
			If value >= Int32.MinValue And value <= Int32.MaxValue Then
				op = New IntegerOperand(Convert.ToInt32(value))
			Else
				op = New DoubleOperand(value)
			End If
		Else
			value = Double.Parse(node.Image)
			op = New DoubleOperand(value)
		End If

		node.AddValue(op)
		Return node
	End Function

	Public Overrides Function ExitStringLiteral(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		Dim s As String = node.Image
		' Remove first and last characters
		s = s.Substring(1, s.Length - 2)
		' Replace any double quotes with a single quote
		s = s.Replace("""""", """")
		node.AddValue(New StringOperand(s))
		Return node
	End Function

	Public Overrides Function ExitBoolean(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim b As Boolean = node.GetChildAt(0).Values.Item(0)
		node.AddValue(New BooleanOperand(b))
		Return node
	End Function

	Public Overrides Function ExitTrue(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(True)
		Return node
	End Function

	Public Overrides Function ExitFalse(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(False)
		Return node
	End Function

	Public Overrides Function ExitErrorExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitDivError(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ErrorValueOperand(ErrorValueType.Div0))
		Return node
	End Function

	Public Overrides Function ExitNAError(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ErrorValueOperand(ErrorValueType.NA))
		Return node
	End Function

	Public Overrides Function ExitNameError(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ErrorValueOperand(ErrorValueType.Name))
		Return node
	End Function

	Public Overrides Function ExitNullError(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ErrorValueOperand(ErrorValueType.Null))
		Return node
	End Function

	Public Overrides Function ExitRefError(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ErrorValueOperand(ErrorValueType.Ref))
		Return node
	End Function

	Public Overrides Function ExitValueError(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ErrorValueOperand(ErrorValueType.Value))
		Return node
	End Function

	Public Overrides Function ExitNumError(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New ErrorValueOperand(ErrorValueType.Num))
		Return node
	End Function

	Public Overrides Function ExitDefinedName(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		node.AddValue(New NamedFormulaOperator(node.Image))
		Return node
	End Function

	Public Overrides Function ExitFunctionName(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		Dim functionName As String = node.Image
		' Remove trailing '('
		functionName = functionName.Remove(functionName.Length - 1, 1)
		node.AddValue(functionName)
		Return node
	End Function

	Private Sub AddBinaryOperation(ByVal node As PerCederberg.Grammatica.Runtime.Production)
		If node.GetChildCount() > 1 Then
			Dim element As New BinaryOperatorElement
			element.AcceptValues(MyBase.GetChildValues(node))
			node.AddValue(element)
		Else
			node.AddValues(MyBase.GetChildValues(node))
		End If
	End Sub

	Public Overrides Function ExitCell(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		Dim creator As New CellReferenceCreator
		creator.Initialize(node.Image)
		node.AddValue(creator)
		Return node
	End Function

	Public Overrides Function ExitCellRange(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		Dim creator As New CellRangeReferenceCreator
		creator.Initialize(node.Image)
		node.AddValue(creator)
		Return node
	End Function

	Public Overrides Function ExitRowRange(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		Dim creator As New RowReferenceCreator
		creator.Initialize(node.Image)
		node.AddValue(creator)
		Return node
	End Function

	Public Overrides Function ExitColumnRange(ByVal node As PerCederberg.Grammatica.Runtime.Token) As PerCederberg.Grammatica.Runtime.Node
		Dim creator As New ColumnReferenceCreator
		creator.Initialize(node.Image)
		node.AddValue(creator)
		Return node
	End Function

	Public Overrides Function ExitGridReference(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		node.AddValues(Me.GetChildValues(node))
		Return node
	End Function

	Public Overrides Function ExitGridReferenceExpression(ByVal node As PerCederberg.Grammatica.Runtime.Production) As PerCederberg.Grammatica.Runtime.Node
		Dim cr As System.Drawing.CharacterRange
		Dim sheetName As String = Nothing

		Dim refNode As PerCederberg.Grammatica.Runtime.Node = node.GetChildAt(node.GetChildCount() - 1)
		Dim creator As ReferenceCreator = refNode.GetValue(0)
		cr.First = refNode.StartColumn - 1
		cr.Length = refNode.EndColumn - refNode.StartColumn + 1

		If node.GetChildCount() = 2 Then
			Dim sheetToken As PerCederberg.Grammatica.Runtime.Token = node.GetChildAt(0)
			sheetName = sheetToken.Image.Substring(0, sheetToken.Image.Length - 1)
			cr.First = sheetToken.StartColumn - 1
			cr.Length = refNode.EndColumn - sheetToken.StartColumn + 1
		End If

		Dim ref As Reference = creator.CreateReference()
		Dim info As New ReferenceParseInfo
		info.Target = ref
		info.Location = cr
		info.Properties = creator.CreateProperties(node.GetChildCount() = 1)
		info.ParseProperties = New ReferenceParseProperties()
		info.ParseProperties.SheetName = sheetName
		MyReferenceInfos.Add(info)

		node.AddValue(ref)
		Return node
	End Function

	Public ReadOnly Property ReferenceInfos() As ReferenceParseInfo() Implements IAnalyzer.ReferenceInfos
		Get
			Dim arr(MyReferenceInfos.Count - 1) As ReferenceParseInfo
			MyReferenceInfos.CopyTo(arr, 0)
			Return arr
		End Get
	End Property
End Class

Friend MustInherit Class ReferenceCreator

	Protected MyImage As String

	Public Sub Initialize(ByVal image As String)
		MyImage = image
	End Sub

	Public MustOverride Function CreateReference() As Reference
	Public MustOverride Function CreateProperties(ByVal implicitSheet As Boolean) As ReferenceProperties
End Class

Friend Class CellReferenceCreator
	Inherits ReferenceCreator

	Public Overrides Function CreateProperties(ByVal implicitSheet As Boolean) As ReferenceProperties
		Return CellReference.CreateProperties(implicitSheet, MyImage)
	End Function

	Public Overrides Function CreateReference() As Reference
		Return CellReference.FromString(MyImage)
	End Function
End Class

Friend Class CellRangeReferenceCreator
	Inherits ReferenceCreator

	Public Overrides Function CreateProperties(ByVal implicitSheet As Boolean) As ReferenceProperties
		Return CellRangeReference.CreateProperties(implicitSheet, MyImage)
	End Function

	Public Overrides Function CreateReference() As Reference
		Return CellRangeReference.FromString(MyImage)
	End Function
End Class

Friend Class RowReferenceCreator
	Inherits ReferenceCreator

	Public Overrides Function CreateProperties(ByVal implicitSheet As Boolean) As ReferenceProperties
		Return RowColumnReference.CreateProperties(implicitSheet, MyImage)
	End Function

	Public Overrides Function CreateReference() As Reference
		Return RowReference.FromString(MyImage)
	End Function
End Class

Friend Class ColumnReferenceCreator
	Inherits ReferenceCreator

	Public Overrides Function CreateProperties(ByVal implicitSheet As Boolean) As ReferenceProperties
		Return RowColumnReference.CreateProperties(implicitSheet, MyImage)
	End Function

	Public Overrides Function CreateReference() As Reference
		Return ColumnReference.FromString(MyImage)
	End Function
End Class