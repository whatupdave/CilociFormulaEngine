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

' Implementation of all operators in formulas

''' <summary>
''' Base class for all operators
''' </summary>
<Serializable()> _
Friend MustInherit Class OperatorBase
	Implements IFormulaComponent
	Implements System.Runtime.Serialization.ISerializable

	Private Const VERSION As Integer = 1

	Protected Sub New()

	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)

	End Sub

	Public Overridable Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
	End Sub

	Public MustOverride Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine) Implements IFormulaComponent.Evaluate

	Public Function Clone() As Object Implements System.ICloneable.Clone
		Return Me.MemberwiseClone()
	End Function

	Public Overridable Sub EvaluateForDependencyReference(ByVal references As System.Collections.IList, ByVal engine As FormulaEngine) Implements IFormulaComponent.EvaluateForDependencyReference

	End Sub

	Public Overridable Sub Validate(ByVal engine As FormulaEngine) Implements IFormulaComponent.Validate

	End Sub

	Public Shared Function IsInvalidDouble(ByVal d As Double) As Boolean
		Return Double.IsInfinity(d) Or Double.IsNaN(d)
	End Function
End Class

''' <summary>
''' Represents an operator on one operand
''' </summary>
<Serializable()> _
Friend MustInherit Class UnaryOperator
	Inherits OperatorBase

	Protected Sub New()

	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Public Overloads Overrides Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine)
		Dim op As IOperand = state.Pop()

		Dim result As IOperand
		result = Me.ComputeValueInternal(op)
		state.Push(result)
	End Sub

	Private Function ComputeValueInternal(ByVal op As IOperand) As IOperand
		Dim knownType As OperandType = Me.GetKnownArgumentType()

		Dim convertedOp As IOperand = op.Convert(knownType)

		If convertedOp Is Nothing Then
			Return Me.GetConvertError(op)
		Else
			Return Me.ComputeValue(convertedOp)
		End If
	End Function

	Private Function GetConvertError(ByVal op As IOperand) As IOperand
		Dim errorOp As IOperand = op.Convert(OperandType.Error)

		If Not errorOp Is Nothing Then
			Return errorOp
		Else
			Return New ErrorValueOperand(ErrorValueType.Value)
		End If
	End Function

	Protected MustOverride Function GetKnownArgumentType() As OperandType
	Protected MustOverride Function ComputeValue(ByVal value As IOperand) As IOperand
End Class

<Serializable()> _
Friend Class UnaryNegateOperator
	Inherits UnaryOperator

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function GetKnownArgumentType() As OperandType
		Return OperandType.Double
	End Function

	Protected Overloads Overrides Function ComputeValue(ByVal value As IOperand) As IOperand
		Return New DoubleOperand(-DirectCast(value, DoubleOperand).ValueAsDouble)
	End Function
End Class

<Serializable()> _
Friend Class PercentOperator
	Inherits UnaryOperator

	Private MyScale As Integer

	Public Sub New(ByVal factor As Integer)
		MyScale = 100 ^ factor
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyScale = info.GetInt32("Scale")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Scale", MyScale)
	End Sub

	Protected Overrides Function GetKnownArgumentType() As OperandType
		Return OperandType.Double
	End Function

	Protected Overloads Overrides Function ComputeValue(ByVal value As IOperand) As IOperand
		Return New DoubleOperand(DirectCast(value, DoubleOperand).ValueAsDouble / MyScale)
	End Function
End Class

''' <summary>
''' An operator on two operands
''' </summary>
<Serializable()> _
Friend MustInherit Class BinaryOperator
	Inherits OperatorBase

	Protected Delegate Function DoubleOperator(ByVal lhs As Double, ByVal rhs As Double) As Double

	Protected Sub New()

	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Public Overrides Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine)
		Dim rhs, lhs As IOperand
		rhs = state.Pop()
		lhs = state.Pop()

		Dim result As IOperand

		Dim lhsConverted, rhsConverted As IOperand
		lhsConverted = Me.ConvertOperand(lhs)
		rhsConverted = Me.ConvertOperand(rhs)

		If lhsConverted Is Nothing Or rhsConverted Is Nothing Then
			result = Me.GetConvertError(lhs, rhs)
		Else
			result = Me.ComputeValue(lhsConverted, rhsConverted, engine)
		End If

		state.Push(result)
	End Sub

	Protected Overridable Function ConvertOperand(ByVal op As IOperand) As IOperand
		Return op.Convert(Me.ArgumentType)
	End Function

	Private Function GetConvertError(ByVal lhs As IOperand, ByVal rhs As IOperand) As IOperand
		Dim errorOp As IOperand

		errorOp = Me.GetErrorOperand(lhs, rhs)

		If Not errorOp Is Nothing Then
			Return errorOp
		Else
			Return New ErrorValueOperand(ErrorValueType.Value)
		End If
	End Function

	Protected Function GetErrorOperand(ByVal lhs As IOperand, ByVal rhs As IOperand) As IOperand
		Dim errorOp As IOperand

		errorOp = lhs.Convert(OperandType.Error)

		If Not errorOp Is Nothing Then
			Return errorOp
		End If

		errorOp = rhs.Convert(OperandType.Error)

		If Not errorOp Is Nothing Then
			Return errorOp
		End If

		Return Nothing
	End Function

	Protected Function DoDoubleOperation(ByVal lhs As DoubleOperand, ByVal rhs As DoubleOperand, ByVal [operator] As DoubleOperator) As IOperand
		Dim result As Double = [operator](lhs.ValueAsDouble, rhs.ValueAsDouble)

		If OperatorBase.IsInvalidDouble(result) = True Then
			Return New ErrorValueOperand(ErrorValueType.Num)
		Else
			Return New DoubleOperand(result)
		End If
	End Function

	Protected Function DoDateAndDoubleOperation(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal op As DoubleOperator) As IOperand
		If Me.IsDateAndDouble(lhs, rhs) = True Then
			Return Me.ComputeDateAndDouble(lhs, rhs, op)
		ElseIf Me.IsDateAndDouble(rhs, lhs) = True Then
			Return Me.ComputeDateAndDouble(rhs, lhs, op)
		Else
			Return Me.DoDoubleOperation(lhs, rhs, op)
		End If
	End Function

	Private Function IsDateAndDouble(ByVal op1 As IOperand, ByVal op2 As IOperand) As Boolean
		Return (TypeOf (op1) Is DateTimeOperand) AndAlso (TypeOf (op2) Is DoubleOperand)
	End Function

	Private Function ComputeDateAndDouble(ByVal dateOp As DateTimeOperand, ByVal doubleOp As DoubleOperand, ByVal op As DoubleOperator) As IOperand
		Try
			Dim dt As DateTime = dateOp.ValueAsDateTime
			Dim result As Double = op(dt.ToOADate(), doubleOp.ValueAsDouble)
			dt = DateTime.FromOADate(result)
			Return New DateTimeOperand(dt)
		Catch ex As ArgumentException
			Return New ErrorValueOperand(ErrorValueType.Value)
		End Try
	End Function

	Protected MustOverride Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
	Protected MustOverride ReadOnly Property ArgumentType() As OperandType
End Class

<Serializable()> _
Friend Class BinaryAddOperator
	Inherits BinaryOperator

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
		Return MyBase.DoDateAndDoubleOperation(lhs, rhs, AddressOf DoAdd)
	End Function

	Private Function DoAdd(ByVal lhs As Double, ByVal rhs As Double) As Double
		Return lhs + rhs
	End Function

	Protected Overrides Function ConvertOperand(ByVal op As IOperand) As IOperand
		Dim dateOp As DateTimeOperand = op.Convert(OperandType.DateTime)
		If dateOp Is Nothing Then
			Return MyBase.ConvertOperand(op)
		Else
			Return dateOp
		End If
	End Function

	Protected Overrides ReadOnly Property ArgumentType() As OperandType
		Get
			Return OperandType.Double
		End Get
	End Property
End Class

<Serializable()> _
Friend Class BinarySubOperator
	Inherits BinaryOperator

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
		Return MyBase.DoDateAndDoubleOperation(lhs, rhs, AddressOf DoSub)
	End Function

	Private Function DoSub(ByVal lhs As Double, ByVal rhs As Double) As Double
		Return lhs - rhs
	End Function

	Protected Overrides Function ConvertOperand(ByVal op As IOperand) As IOperand
		Dim dateOp As DateTimeOperand = op.Convert(OperandType.DateTime)
		If dateOp Is Nothing Then
			Return MyBase.ConvertOperand(op)
		Else
			Return dateOp
		End If
	End Function

	Protected Overrides ReadOnly Property ArgumentType() As OperandType
		Get
			Return OperandType.Double
		End Get
	End Property
End Class

<Serializable()> _
Friend Class BinaryMultiplyOperator
	Inherits BinaryOperator

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
		Return MyBase.DoDoubleOperation(lhs, rhs, AddressOf DoMul)
	End Function

	Private Function DoMul(ByVal lhs As Double, ByVal rhs As Double) As Double
		Return lhs * rhs
	End Function

	Protected Overrides ReadOnly Property ArgumentType() As OperandType
		Get
			Return OperandType.Double
		End Get
	End Property
End Class

<Serializable()> _
Friend Class BinaryDivisionOperator
	Inherits BinaryOperator

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
		Dim rhsDouble As Double = DirectCast(rhs, DoubleOperand).ValueAsDouble
		If rhsDouble = 0 Then
			Return New ErrorValueOperand(ErrorValueType.Div0)
		End If

		Return MyBase.DoDoubleOperation(lhs, rhs, AddressOf DoDiv)
	End Function

	Private Function DoDiv(ByVal lhs As Double, ByVal rhs As Double) As Double
		Return lhs / rhs
	End Function

	Protected Overrides ReadOnly Property ArgumentType() As OperandType
		Get
			Return OperandType.Double
		End Get
	End Property
End Class

<Serializable()> _
Friend Class ConcatenationOperator
	Inherits BinaryOperator

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
		Return New StringOperand(DirectCast(lhs, StringOperand).ValueAsString & DirectCast(rhs, StringOperand).ValueAsString)
	End Function

	Protected Overrides ReadOnly Property ArgumentType() As OperandType
		Get
			Return OperandType.String
		End Get
	End Property
End Class

<Serializable()> _
Friend Class BinaryPowerOperator
	Inherits BinaryOperator

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
		Return MyBase.DoDoubleOperation(lhs, rhs, AddressOf DoPower)
	End Function

	Private Function DoPower(ByVal lhs As Double, ByVal rhs As Double) As Double
		Return System.Math.Pow(lhs, rhs)
	End Function

	Protected Overrides ReadOnly Property ArgumentType() As OperandType
		Get
			Return OperandType.Double
		End Get
	End Property
End Class

<Serializable()> _
Friend Class LogicalOperator
	Inherits BinaryOperator

	Private MyCompareType As CompareType

	Public Sub New(ByVal ct As CompareType)
		MyCompareType = ct
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyCompareType = info.GetInt32("CompareType")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("CompareType", MyCompareType)
	End Sub

	Protected Overrides Function ComputeValue(ByVal lhs As IOperand, ByVal rhs As IOperand, ByVal engine As FormulaEngine) As IOperand
		Dim errorOp As IOperand = MyBase.GetErrorOperand(lhs, rhs)

		If Not errorOp Is Nothing Then
			Return errorOp
		End If

		Dim cr As CompareResult = Me.Compare(lhs, rhs)

		Dim result As Boolean = GetBooleanResult(MyCompareType, cr)
		Return New BooleanOperand(result)
	End Function

	Private Function Compare(ByVal lhs As IOperand, ByVal rhs As IOperand) As CompareResult
		Dim convertTypes As OperandType() = {OperandType.Integer, OperandType.Double, OperandType.String, OperandType.Boolean, OperandType.DateTime}

		Dim commonTypeIndex As Integer = Me.IndexOfCommonType(lhs, rhs, convertTypes)

		If commonTypeIndex = -1 Then
			Return Me.CompareDifferentTypes(lhs, rhs)
		Else
			Dim commonType As OperandType = convertTypes(commonTypeIndex)
			Dim lhsPrim, rhsPrim As PrimitiveOperand
			lhsPrim = lhs.Convert(commonType)
			rhsPrim = rhs.Convert(commonType)
			Dim result As Integer = lhsPrim.Compare(rhsPrim)
			Return Compare2CompareResult(result)
		End If
	End Function

	Private Function IndexOfCommonType(ByVal lhs As PrimitiveOperand, ByVal rhs As PrimitiveOperand, ByVal types As OperandType()) As Integer
		For i As Integer = 0 To types.Length - 1
			Dim ot As OperandType = types(i)
			If lhs.CanConvertForCompare(ot) = True And rhs.CanConvertForCompare(ot) = True Then
				Return i
			End If
		Next
		Return -1
	End Function

	Private Function CompareDifferentTypes(ByVal lhs As PrimitiveOperand, ByVal rhs As PrimitiveOperand) As CompareResult
		Dim lhsValue, rhsValue As Object
		lhsValue = lhs.Value
		rhsValue = rhs.Value

		' Excel has weird rules when comparing values of different types.  Basically it assigns each type a rank and 
		' compares the rank of both values.
		Dim lhsRank, rhsRank As Integer
		lhsRank = GetTypeRank(lhsValue.GetType())
		rhsRank = GetTypeRank(rhsValue.GetType())
		Return LogicalOperator.Compare2CompareResult(lhsRank.CompareTo(rhsRank))
	End Function

	Private Shared Function GetTypeRank(ByVal t As Type) As Integer
		If t Is GetType(Double) Or t Is GetType(Integer) Then
			Return 1
		ElseIf t Is GetType(String) Then
			Return 2
		ElseIf t Is GetType(Boolean) Then
			Return 3
		ElseIf t Is GetType(DateTime) Then
			Return 4
		Else
			Debug.Assert(False, "unknown type")
			Return 0
		End If
	End Function

	Public Shared Function GetBooleanResult(ByVal ct As CompareType, ByVal cr As CompareResult) As Boolean
		Dim result As Boolean

		Select Case ct
			Case CompareType.Equal
				result = cr = CompareResult.Equal
			Case CompareType.NotEqual
				result = cr <> CompareResult.Equal
			Case CompareType.LessThan
				result = cr = CompareResult.LessThan
			Case CompareType.GreaterThan
				result = cr = CompareResult.GreaterThan
			Case CompareType.LessThanOrEqual
				result = cr = CompareResult.LessThan Or cr = CompareResult.Equal
			Case CompareType.GreaterThanOrEqual
				result = cr = CompareResult.GreaterThan Or cr = CompareResult.Equal
			Case Else
				Debug.Assert(False, "unknown compare type")
		End Select

		Return result
	End Function

	Public Shared Function Compare2CompareResult(ByVal compare As Integer) As CompareResult
		If compare < 0 Then
			Return CompareResult.LessThan
		ElseIf compare = 0 Then
			Return CompareResult.Equal
		Else
			Return CompareResult.GreaterThan
		End If
	End Function

	Protected Overrides ReadOnly Property ArgumentType() As OperandType
		Get
			Return OperandType.Primitive
		End Get
	End Property
End Class

<Serializable()> _
Friend Class FunctionCallOperator
	Inherits OperatorBase

	Private MyFunctionName As String
	Private MyArgumentCount As Integer

	Public Sub New(ByVal functionName As String, ByVal argumentCount As Integer)
		MyFunctionName = functionName
		MyArgumentCount = argumentCount
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyFunctionName = info.GetString("FunctionName")
		MyArgumentCount = info.GetInt32("ArgumentCount")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("FunctionName", MyFunctionName)
		info.AddValue("ArgumentCount", MyArgumentCount)
	End Sub

	Public Overrides Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine)
		Dim fl As FunctionLibrary = engine.FunctionLibrary
		fl.InvokeFunction(MyFunctionName, state, MyArgumentCount)
	End Sub

	Public Overrides Sub EvaluateForDependencyReference(ByVal references As System.Collections.IList, ByVal engine As FormulaEngine)
		MyBase.EvaluateForDependencyReference(references, engine)
		If engine.FunctionLibrary.IsFunctionVolatile(MyFunctionName) = True Then
			Dim ref As New VolatileFunctionReference()
			references.Add(ref)
		End If
	End Sub

	Public Overrides Sub Validate(ByVal engine As FormulaEngine)
		If engine.FunctionLibrary.IsFunctionDefined(MyFunctionName) = False Then
			Dim inner As New InvalidOperationException(String.Format("The function {0} is not defined", MyFunctionName))
			Throw New InvalidFormulaException(inner)
		ElseIf engine.FunctionLibrary.IsValidArgumentCount(MyFunctionName, MyArgumentCount) = False Then
			Dim inner As New ArgumentException(String.Format("Invalid number of arguments provided for function {0}", MyFunctionName))
			Throw New InvalidFormulaException(inner)
		End If
	End Sub
End Class

<Serializable()> _
Friend Class NamedFormulaOperator
	Inherits OperatorBase

	Private MyName As String

	Public Sub New(ByVal name As String)
		MyName = name
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyName = info.GetString("Name")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Name", MyName)
	End Sub

	Public Overrides Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine)
		Dim namedRef As Reference = engine.ReferenceFactory.Named(MyName)
		Dim target As Formula = engine.GetFormulaAt(namedRef)

		If Not target Is Nothing Then
			Dim selfRef As NamedReference = target.SelfReferenceInternal
			Dim op As IOperand = selfRef.ValueOperand
			state.Push(op)
		Else
			state.Push(New ErrorValueOperand(ErrorValueType.Name))
		End If
	End Sub

	Public Overrides Sub EvaluateForDependencyReference(ByVal references As System.Collections.IList, ByVal engine As FormulaEngine)
		Dim namedRef As Reference = engine.ReferenceFactory.Named(MyName)
		references.Add(namedRef)
	End Sub
End Class