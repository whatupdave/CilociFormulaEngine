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
''' Base class for a constant, primitive value that can be converted to other values.
''' </summary>
<Serializable()> _
Friend MustInherit Class PrimitiveOperand
	Implements IFormulaComponent
	Implements IOperand
	Implements System.Runtime.Serialization.ISerializable

	Private Const VERSION As Integer = 1

	Protected Sub New()

	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)

	End Sub

	Public Overridable Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
	End Sub

	Protected MustOverride Sub SetDeserializedValue(ByVal value As Object)

	Public Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine) Implements IFormulaComponent.Evaluate
		state.Push(Me)
	End Sub

	Public Overridable Sub EvaluateForDependencyReference(ByVal references As System.Collections.IList, ByVal engine As FormulaEngine) Implements IFormulaComponent.EvaluateForDependencyReference

	End Sub

	Public MustOverride Function Compare(ByVal rhs As PrimitiveOperand) As Integer

	Protected MustOverride Function ConvertToDouble() As IOperand
	Protected MustOverride Function ConvertToInteger() As IOperand
	Protected MustOverride Function ConvertToString() As IOperand
	Protected MustOverride Function ConvertToBoolean() As IOperand

	Protected Overridable Function ConvertToError() As IOperand
		Return Nothing
	End Function

	Protected Overridable Function ConvertToBlank() As IOperand
		Return Nothing
	End Function

	Protected Overridable Function ConvertToDateTime() As IOperand
		Return Nothing
	End Function

	Public Function Convert(ByVal convertType As OperandType) As IOperand Implements IOperand.Convert
		Select Case convertType
			Case OperandType.Double
				Return Me.ConvertToDouble()
			Case OperandType.String
				Return Me.ConvertToString()
			Case OperandType.Boolean
				Return Me.ConvertToBoolean()
			Case OperandType.Integer
				Return Me.ConvertToInteger()
			Case OperandType.Reference, OperandType.SheetReference
				Return Nothing
			Case OperandType.Self
				Return Me
			Case OperandType.Primitive
				Return Me
			Case OperandType.Error
				Return Me.ConvertToError()
			Case OperandType.Blank
				Return Me.ConvertToBlank()
			Case OperandType.DateTime
				Return Me.ConvertToDateTime()
			Case Else
				Throw New ArgumentException("Unknown convert type")
		End Select
	End Function

	Public MustOverride ReadOnly Property Value() As Object Implements IOperand.Value

	Public Function Clone() As Object Implements System.ICloneable.Clone
		Return Me.MemberwiseClone()
	End Function

	Public MustOverride Function CanConvertForCompare(ByVal ot As OperandType) As Boolean

	Public Sub Validate(ByVal engine As FormulaEngine) Implements IFormulaComponent.Validate

	End Sub

	Public MustOverride ReadOnly Property NativeType() As OperandType Implements IOperand.NativeType
End Class

<Serializable()> _
Friend Class DoubleOperand
	Inherits PrimitiveOperand

	Private MyValue As Double

	Public Sub New(ByVal value As Double)
		MyValue = value
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyValue = info.GetDouble("Value")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Value", MyValue)
	End Sub

	Protected Overrides Sub SetDeserializedValue(ByVal value As Object)
		MyValue = DirectCast(value, Double)
	End Sub

	Public Overrides Function Compare(ByVal rhs As PrimitiveOperand) As Integer
		Dim dRhs As Double
		dRhs = DirectCast(rhs, DoubleOperand).MyValue
		Return MyValue.CompareTo(dRhs)
	End Function

	Protected Overrides Function ConvertToBoolean() As IOperand
		Return New BooleanOperand(System.Convert.ToBoolean(MyValue))
	End Function

	Protected Overrides Function ConvertToDouble() As IOperand
		Return Me
	End Function

	Protected Overrides Function ConvertToInteger() As IOperand
		If MyValue < Int32.MinValue Or MyValue > Int32.MaxValue Then
			Return Nothing
		Else
			' Excel does a hard truncate ....system.convert will do a round
			Dim d As Double = System.Math.Truncate(MyValue)
			Return New IntegerOperand(System.Convert.ToInt32(d))
		End If
	End Function

	Protected Overrides Function ConvertToString() As IOperand
		Return New StringOperand(MyValue.ToString())
	End Function

	Public Overrides Function CanConvertForCompare(ByVal ot As OperandType) As Boolean
		Return ot = OperandType.Double
	End Function

	Public Overrides ReadOnly Property Value() As Object
		Get
			Return MyValue
		End Get
	End Property

	Public ReadOnly Property ValueAsDouble() As Double
		Get
			Return MyValue
		End Get
	End Property

	Public Overrides ReadOnly Property NativeType() As OperandType
		Get
			Return OperandType.Double
		End Get
	End Property
End Class

<Serializable()> _
Friend Class IntegerOperand
	Inherits PrimitiveOperand

	Private MyValue As Integer

	Public Sub New(ByVal value As Integer)
		MyValue = value
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyValue = info.GetInt32("Value")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Value", MyValue)
	End Sub

	Protected Overrides Sub SetDeserializedValue(ByVal value As Object)
		MyValue = DirectCast(value, Integer)
	End Sub

	Public Overrides Function Compare(ByVal rhs As PrimitiveOperand) As Integer
		Dim iRhs As Integer = DirectCast(rhs, IntegerOperand).MyValue
		Return MyValue.CompareTo(iRhs)
	End Function

	Public Overrides Function CanConvertForCompare(ByVal ot As OperandType) As Boolean
		Return ot = OperandType.Integer Or ot = OperandType.Double
	End Function

	Protected Overrides Function ConvertToBoolean() As IOperand
		Return New BooleanOperand(System.Convert.ToBoolean(MyValue))
	End Function

	Protected Overrides Function ConvertToDouble() As IOperand
		Return New DoubleOperand(MyValue)
	End Function

	Protected Overrides Function ConvertToInteger() As IOperand
		Return Me
	End Function

	Protected Overrides Function ConvertToString() As IOperand
		Return New StringOperand(MyValue.ToString())
	End Function

	Public ReadOnly Property ValueAsInteger() As Integer
		Get
			Return MyValue
		End Get
	End Property

	Public Overrides ReadOnly Property Value() As Object
		Get
			Return MyValue
		End Get
	End Property

	Public Overrides ReadOnly Property NativeType() As OperandType
		Get
			Return OperandType.Integer
		End Get
	End Property
End Class

<Serializable()> _
Friend Class StringOperand
	Inherits PrimitiveOperand

	Private MyValue As String

	Public Sub New(ByVal value As String)
		MyValue = value
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyValue = info.GetString("Value")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Value", MyValue)
	End Sub

	Protected Overrides Sub SetDeserializedValue(ByVal value As Object)
		MyValue = DirectCast(value, String)
	End Sub

	Public Overrides Function Compare(ByVal rhs As PrimitiveOperand) As Integer
		Dim s1, s2 As String
		s1 = Me.Value
		s2 = DirectCast(rhs, StringOperand).MyValue
		Return String.Compare(s1, s2, StringComparison.OrdinalIgnoreCase)
	End Function

	Public Overrides Function CanConvertForCompare(ByVal ot As OperandType) As Boolean
		Return ot = OperandType.String
	End Function

	Protected Overrides Function ConvertToBoolean() As IOperand
		If String.Equals(MyValue, Boolean.TrueString, StringComparison.OrdinalIgnoreCase) = True Then
			Return New BooleanOperand(True)
		ElseIf String.Equals(MyValue, Boolean.FalseString, StringComparison.OrdinalIgnoreCase) = True Then
			Return New BooleanOperand(False)
		Else
			Return Nothing
		End If
	End Function

	Protected Overrides Function ConvertToDouble() As IOperand
		Dim result As Double
		Dim success As Boolean

		success = Double.TryParse(MyValue, Globalization.NumberStyles.Float, Nothing, result)

		If success = True Then
			Return New DoubleOperand(result)
		Else
			Return Nothing
		End If
	End Function

	Protected Overrides Function ConvertToInteger() As IOperand
		Dim op As DoubleOperand = Me.ConvertToDouble()

		If op Is Nothing Then
			Return Nothing
		Else
			Return op.Convert(OperandType.Integer)
		End If
	End Function

	Protected Overrides Function ConvertToString() As IOperand
		Return Me
	End Function

	Public ReadOnly Property ValueAsString() As String
		Get
			Return MyValue
		End Get
	End Property

	Public Overrides ReadOnly Property Value() As Object
		Get
			Return MyValue
		End Get
	End Property

	Public Overrides ReadOnly Property NativeType() As OperandType
		Get
			Return OperandType.String
		End Get
	End Property
End Class

<Serializable()> _
Friend Class BooleanOperand
	Inherits PrimitiveOperand

	Private MyValue As Boolean

	Public Sub New(ByVal value As Boolean)
		MyValue = value
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyValue = info.GetBoolean("Value")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Value", MyValue)
	End Sub

	Protected Overrides Sub SetDeserializedValue(ByVal value As Object)
		MyValue = DirectCast(value, Boolean)
	End Sub

	Public Overrides Function Compare(ByVal rhs As PrimitiveOperand) As Integer
		Dim i1, i2 As IntegerOperand
		i1 = Me.ConvertToInteger()
		i2 = rhs.Convert(OperandType.Integer)
		Return i1.Compare(i2)
	End Function

	Public Overrides Function CanConvertForCompare(ByVal ot As OperandType) As Boolean
		Return ot = OperandType.Boolean
	End Function

	Protected Overrides Function ConvertToBoolean() As IOperand
		Return Me
	End Function

	Protected Overrides Function ConvertToDouble() As IOperand
		Return New DoubleOperand(System.Convert.ToDouble(MyValue))
	End Function

	Protected Overrides Function ConvertToInteger() As IOperand
		Return New IntegerOperand(System.Convert.ToInt32(MyValue))
	End Function

	Protected Overrides Function ConvertToString() As IOperand
		Return New StringOperand(MyValue.ToString().ToUpper())
	End Function

	Public ReadOnly Property ValueAsBoolean() As Boolean
		Get
			Return MyValue
		End Get
	End Property

	Public Overrides ReadOnly Property Value() As Object
		Get
			Return MyValue
		End Get
	End Property

	Public Overrides ReadOnly Property NativeType() As OperandType
		Get
			Return OperandType.Boolean
		End Get
	End Property
End Class

<Serializable()> _
Friend Class DateTimeOperand
	Inherits PrimitiveOperand

	Private MyValue As DateTime

	Public Sub New(ByVal value As DateTime)
		MyValue = value
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyValue = info.GetDateTime("Value")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Value", MyValue)
	End Sub

	Protected Overrides Sub SetDeserializedValue(ByVal value As Object)
		MyValue = DirectCast(value, DateTime)
	End Sub

	Public Overrides Function CanConvertForCompare(ByVal ot As OperandType) As Boolean
		Return ot = OperandType.DateTime
	End Function

	Public Overrides Function Compare(ByVal rhs As PrimitiveOperand) As Integer
		Dim rhsValue As DateTime = DirectCast(rhs, DateTimeOperand).MyValue
		Return MyValue.CompareTo(rhsValue)
	End Function

	Protected Overrides Function ConvertToBoolean() As IOperand
		Dim b As Boolean = MyValue.Equals(DateTime.MinValue) = False
		Return New BooleanOperand(b)
	End Function

	Protected Overrides Function ConvertToDouble() As IOperand
		Return Nothing
	End Function

	Protected Overrides Function ConvertToInteger() As IOperand
		Return Nothing
	End Function

	Protected Overrides Function ConvertToString() As IOperand
		Return MyValue.ToString()
	End Function

	Protected Overrides Function ConvertToDateTime() As IOperand
		Return Me
	End Function

	Public Overrides ReadOnly Property NativeType() As OperandType
		Get
			Return OperandType.DateTime
		End Get
	End Property

	Public Overrides ReadOnly Property Value() As Object
		Get
			Return MyValue
		End Get
	End Property

	Public ReadOnly Property ValueAsDateTime() As DateTime
		Get
			Return MyValue
		End Get
	End Property
End Class

<Serializable()> _
Friend Class ErrorValueOperand
	Inherits PrimitiveOperand

	Private MyValue As ErrorValueType

	Public Sub New(ByVal value As ErrorValueType)
		MyValue = value
	End Sub

	Public Sub New(ByVal wrapper As ErrorValueWrapper)
		MyValue = wrapper.ErrorValue
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyValue = info.GetInt32("Value")
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Value", MyValue)
	End Sub

	Protected Overrides Sub SetDeserializedValue(ByVal value As Object)
		MyValue = DirectCast(value, ErrorValueType)
	End Sub

	Public Overrides Function Compare(ByVal rhs As PrimitiveOperand) As Integer
		Debug.Assert(False, "should not be called")
	End Function

	Public Overrides Function CanConvertForCompare(ByVal ot As OperandType) As Boolean
		Debug.Assert(False, "should not be called")
	End Function

	Protected Overrides Function ConvertToBoolean() As IOperand
		Return Nothing
	End Function

	Protected Overrides Function ConvertToDouble() As IOperand
		Return Nothing
	End Function

	Protected Overrides Function ConvertToInteger() As IOperand
		Return Nothing
	End Function

	Protected Overrides Function ConvertToString() As IOperand
		Return Nothing
	End Function

	Protected Overrides Function ConvertToError() As IOperand
		Return Me
	End Function

	Public Overrides ReadOnly Property Value() As Object
		Get
			Return MyValue
		End Get
	End Property

	Public ReadOnly Property ValueAsErrorType() As ErrorValueType
		Get
			Return MyValue
		End Get
	End Property

	Public ReadOnly Property ValueAsErrorWrapper() As ErrorValueWrapper
		Get
			Return New ErrorValueWrapper(MyValue)
		End Get
	End Property

	Public Overrides ReadOnly Property NativeType() As OperandType
		Get
			Return OperandType.Error
		End Get
	End Property
End Class

<Serializable()> _
Friend Class NullValueOperand
	Inherits PrimitiveOperand

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Sub SetDeserializedValue(ByVal value As Object)

	End Sub

	Public Overrides Function Compare(ByVal rhs As PrimitiveOperand) As Integer
		Debug.Assert(False, "should not be called")
	End Function

	Public Overrides Function CanConvertForCompare(ByVal ot As OperandType) As Boolean
		Return True
	End Function

	Protected Overrides Function ConvertToBoolean() As IOperand
		Return New BooleanOperand(False)
	End Function

	Protected Overrides Function ConvertToDouble() As IOperand
		Return New DoubleOperand(0.0)
	End Function

	Protected Overrides Function ConvertToInteger() As IOperand
		Return New IntegerOperand(0)
	End Function

	Protected Overrides Function ConvertToString() As IOperand
		Return New StringOperand(String.Empty)
	End Function

	Protected Overrides Function ConvertToBlank() As IOperand
		Return Me
	End Function

	Public Overrides ReadOnly Property Value() As Object
		Get
			Return Nothing
		End Get
	End Property

	Public Overrides ReadOnly Property NativeType() As OperandType
		Get
			Return OperandType.Blank
		End Get
	End Property
End Class