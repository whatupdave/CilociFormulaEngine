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

' Various utilities that are required when implementing the builtin functions

''' <summary>
''' Provides a framework for processing the arguments of a variable argument formula function
''' </summary>
''' <remarks>This class provides a more reusable approach to processing the arguments of formula functions that 
''' take a variable number of arguments.  Processing the arguments to such functions requires differentiating between
''' reference and primitive arguments, deciding how to handle error and null values, and even transforming values.  This class
''' handles the core processing of these tasks and lets derived classes handle the details specific to each function.</remarks>
Public MustInherit Class VariableArgumentFunctionProcessor
	Implements IReferenceValueProcessor

	Private MyErrorValue As ErrorValueType
	Private MyProcessArgumentsFlag As Boolean

	Protected Sub New()
		MyProcessArgumentsFlag = True
	End Sub

	''' <summary>
	''' Processes all arguments to a variable argument formula function
	''' </summary>
	''' <param name="args">The arguments to process</param>
	''' <returns>True if processing was sucessful; False otherwise</returns>
	''' <remarks>This is the main method responsible for processing the function's arguments.  It handles the processing of 
	''' primitive arguments and the processing of each value of a reference argument.</remarks>
	Public Function ProcessArguments(ByVal args As Argument()) As Boolean
		Me.SortArguments(args)

		For Each arg As Argument In args
			Me.ProcessArgument(arg)
			If MyProcessArgumentsFlag = False Then
				Return False
			End If
		Next
		Return True
	End Function

	Private Sub SortArguments(ByVal args As Argument())
		' Excel seems to process primitive arguments before reference arguments
		' Sort our arguments so that primitives come first
		Dim original As Argument() = args.Clone()
		System.Array.Sort(Of Argument)(args, New ArgumentComparer(original))
	End Sub

	Private Sub ProcessArgument(ByVal arg As Argument)
		If arg.IsReference = False Then
			Dim keepGoing As Boolean = Me.ProcessPrimitiveArgumentInternal(arg)
			Me.KeepProcessingArguments(keepGoing)
		Else
			Dim ref As IReference = arg.ValueAsReference
			ref.GetReferenceValues(Me)
		End If
	End Sub

	Private Function ProcessPrimitiveArgumentInternal(ByVal arg As Argument) As Boolean
		If Me.StopOnError = True And arg.IsError = True Then
			Me.SetError(arg.ValueAsError)
			Return False
		Else
			Return Me.ProcessPrimitiveArgument(arg)
		End If
	End Function

	''' <summary>
	''' Implemented by a derived class to handle processing of a primitive argument
	''' </summary>
	''' <param name="arg">The primitive argument to process</param>
	''' <returns>True if processing was successful; False otherwise</returns>
	''' <remarks>This method will get called for each argument that is not a reference.  It is up to the derived class to decide what
	''' to do with each such argument.</remarks>
	Protected MustOverride Function ProcessPrimitiveArgument(ByVal arg As Argument) As Boolean

	Private Function ProcessValue(ByVal value As Object) As Boolean Implements IReferenceValueProcessor.ProcessValue
		Dim keepGoing As Boolean

		If value Is Nothing Then
			Me.ProcessEmptyValue()
			keepGoing = True
		Else
			keepGoing = Me.ProcessNonEmptyValue(value)
		End If

		Me.KeepProcessingArguments(keepGoing)
		Return keepGoing
	End Function

	''' <summary>
	''' Indicates how an empty reference value should be processed
	''' </summary>
	''' <remarks>This method will get called for each value of a reference that is null.  Derived classes would override this method
	''' if they need to handle null values in a special way.</remarks>
	Protected Overridable Sub ProcessEmptyValue()

	End Sub

	Private Function ProcessNonEmptyValue(ByVal value As Object) As Boolean
		Dim t As Type = value.GetType()
		If Me.IsError(t) Then
			Return Me.ProcessErrorValue(value)
		Else
			Me.ProcessReferenceValue(value, t)
			Return True
		End If
	End Function

	Private Function ProcessErrorValue(ByVal value As ErrorValueWrapper) As Boolean
		Me.OnErrorReferenceValue(value)

		If Me.StopOnError = True Then
			Me.SetError(value.ErrorValue)
			Return False
		Else
			' Ignore the error
			Return True
		End If
	End Function

	''' <summary>
	''' Determines how a reference value that is an error should be handled
	''' </summary>
	''' <param name="value">The error value</param>
	''' <remarks>This method will get called for each value of a reference that is an error value.  Derived classes can override
	''' this method to provide custom handling for such values.</remarks>
	Protected Overridable Sub OnErrorReferenceValue(ByVal value As ErrorValueWrapper)

	End Sub

	''' <summary>
	''' Determines how a non-empty reference value will be processed
	''' </summary>
	''' <param name="value">The value to process</param>
	''' <param name="valueType">The value's type</param>
	''' <remarks>Derived classes must override this method to provide customized handling of a reference's values.  This method only deals
	''' with non-empty values and thus the value parameter will never be null.</remarks>
	Protected MustOverride Sub ProcessReferenceValue(ByVal value As Object, ByVal valueType As Type)

	''' <summary>
	''' Determines if a type represents an error
	''' </summary>
	''' <param name="t">The type to test</param>
	''' <returns>True is the type represents an error; False otherwise</returns>
	''' <remarks>This function is handy when you need to determine if the type of a value is an error.</remarks>
	Protected Function IsError(ByVal t As Type) As Boolean
		Return t Is GetType(ErrorValueWrapper)
	End Function

	Private Function TransformErrorValue(ByVal value As Object) As Object
		Return DirectCast(value, ErrorValueWrapper).ErrorValue
	End Function

	''' <summary>
	''' Sets the error that will be reported at the end of processing
	''' </summary>
	''' <param name="errorType">The type of error to set</param>
	''' <remarks>Derived classes can use this method when they encounter an error during argument processing.  When processing is finished,
	''' this value will be returned by the <see cref="P:ciloci.FormulaEngine.VariableArgumentFunctionProcessor.ErrorValue"/> property to callers who
	''' need to know why processing failed.</remarks>
	Protected Sub SetError(ByVal errorType As ErrorValueType)
		MyErrorValue = errorType
	End Sub

	Private Sub KeepProcessingArguments(ByVal process As Boolean)
		MyProcessArgumentsFlag = process
	End Sub

	''' <summary>
	''' Determines if processing stops upon encountering an error value
	''' </summary>
	''' <value>True if processing should stop when an error value is encountered; False to keep going</value>
	''' <remarks>Some functions, like Sum, do not handle error values and need to stop at the first one they encounter.  Other functions
	''' like Count, simply ignore the value and keep going.  Derived classes must override this property so as to specify their way
	''' of handling error values.</remarks>
	Protected MustOverride ReadOnly Property StopOnError() As Boolean

	''' <summary>
	''' Gets the error that caused processing to fail
	''' </summary>
	''' <value>An error value</value>
	''' <remarks>When processing of arguments fails, this property will indicate the specific error that is the cause.  The caller can check
	''' this property and set the result of the function accordingly.</remarks>
	Public ReadOnly Property ErrorValue() As ErrorValueType
		Get
			Return MyErrorValue
		End Get
	End Property
End Class

''' <summary>
''' Processor that works with a list of doubles
''' </summary>
Friend MustInherit Class DoubleBasedReferenceValueProcessor
	Inherits VariableArgumentFunctionProcessor

	Private MyValues As System.Collections.Generic.IList(Of Double)

	Protected Sub New()
		MyValues = New System.Collections.Generic.List(Of Double)
	End Sub

	Protected Overrides Function ProcessPrimitiveArgument(ByVal arg As Argument) As Boolean
		If arg.IsDouble = True Then
			Me.Values.Add(arg.ValueAsDouble)
			Return True
		Else
			MyBase.SetError(ErrorValueType.Value)
			Return False
		End If
	End Function

	Public ReadOnly Property Values() As System.Collections.Generic.IList(Of Double)
		Get
			Return MyValues
		End Get
	End Property

	Protected Overrides ReadOnly Property StopOnError() As Boolean
		Get
			Return True
		End Get
	End Property
End Class

Friend Class SumProcessor
	Inherits DoubleBasedReferenceValueProcessor

	Public Sub New()

	End Sub

	Protected Overloads Overrides Sub ProcessReferenceValue(ByVal value As Object, ByVal valueType As System.Type)
		If Utility.IsNumericType(valueType) = True Then
			Dim d As Double = Utility.NormalizeNumericValue(value)
			Me.Values.Add(d)
		End If
	End Sub
End Class

Friend Class AverageAProcessor
	Inherits DoubleBasedReferenceValueProcessor

	Protected Overrides Sub ProcessReferenceValue(ByVal value As Object, ByVal valueType As System.Type)
		Dim d As Double

		If value.GetType() Is GetType(Boolean) Then
			Dim b As Boolean = DirectCast(value, Boolean)
			d = System.Convert.ToDouble(b)
		ElseIf value.GetType() Is GetType(String) Then
			d = 0.0
		ElseIf Utility.IsNumericType(valueType) = True Then
			d = Utility.NormalizeNumericValue(value)
		Else
			Throw New InvalidOperationException("Unknown type")
		End If

		Me.Values.Add(d)
	End Sub
End Class

''' <summary>
''' Processor that keeps a count of values
''' </summary>
Friend MustInherit Class CountBasedReferenceValueProcessor
	Inherits VariableArgumentFunctionProcessor

	Private MyCount As Integer

	Protected Sub IncrementCount()
		MyCount += 1
	End Sub

	Public ReadOnly Property Count() As Integer
		Get
			Return MyCount
		End Get
	End Property

	Protected Overrides ReadOnly Property StopOnError() As Boolean
		Get
			Return False
		End Get
	End Property
End Class

Friend Class CountProcessor
	Inherits CountBasedReferenceValueProcessor

	Protected Overrides Function ProcessPrimitiveArgument(ByVal arg As Argument) As Boolean
		If arg.IsDouble = True Then
			Me.IncrementCount()
		End If
		Return True
	End Function

	Protected Overloads Overrides Sub ProcessReferenceValue(ByVal value As Object, ByVal valueType As System.Type)
		If Utility.IsNumericType(valueType) = True Then
			Me.IncrementCount()
		End If
	End Sub
End Class

Friend Class CountAProcessor
	Inherits CountBasedReferenceValueProcessor

	Protected Overrides Function ProcessPrimitiveArgument(ByVal arg As Argument) As Boolean
		Me.IncrementCount()
		Return True
	End Function

	Protected Overrides Sub ProcessReferenceValue(ByVal value As Object, ByVal valueType As System.Type)
		Me.IncrementCount()
	End Sub

	Protected Overrides Sub OnErrorReferenceValue(ByVal value As ErrorValueWrapper)
		Me.IncrementCount()
	End Sub
End Class

Friend Class CountBlankProcessor
	Inherits CountBasedReferenceValueProcessor

	Protected Overrides Function ProcessPrimitiveArgument(ByVal arg As Argument) As Boolean

	End Function

	Protected Overrides Sub ProcessEmptyValue()
		Me.IncrementCount()
	End Sub

	Protected Overrides Sub ProcessReferenceValue(ByVal value As Object, ByVal valueType As System.Type)
		Dim stringValue As String = TryCast(value, String)

		If stringValue Is Nothing Then
			Return
		End If

		If stringValue.Length = 0 Then
			Me.IncrementCount()
		End If
	End Sub
End Class

Friend Class LogicalFunctionProcessor
	Inherits VariableArgumentFunctionProcessor

	Private MyValues As System.Collections.Generic.IList(Of Boolean)

	Public Sub New()
		MyValues = New System.Collections.Generic.List(Of Boolean)
	End Sub

	Protected Overrides Function ProcessPrimitiveArgument(ByVal arg As Argument) As Boolean
		If arg.IsBoolean = True Then
			MyValues.Add(arg.ValueAsBoolean)
			Return True
		Else
			MyBase.SetError(ErrorValueType.Value)
			Return False
		End If
	End Function

	Protected Overrides Sub ProcessReferenceValue(ByVal value As Object, ByVal valueType As System.Type)
		If Utility.IsNumericType(valueType) = True Then
			Dim b As Boolean = DirectCast(value, IConvertible).ToBoolean(Nothing)
			MyValues.Add(b)
		ElseIf valueType Is GetType(Boolean) Then
			MyValues.Add(DirectCast(value, Boolean))
		End If
	End Sub

	Public ReadOnly Property Values() As System.Collections.Generic.IList(Of Boolean)
		Get
			Return MyValues
		End Get
	End Property

	Protected Overrides ReadOnly Property StopOnError() As Boolean
		Get
			Return True
		End Get
	End Property
End Class

''' <summary>
''' The base class for all attributes that mark formula function methods
''' </summary>
''' <remarks>This attribute is the base class for the <see cref="T:ciloci.FormulaEngine.FixedArgumentFormulaFunctionAttribute"/> and <see cref="T:ciloci.FormulaEngine.VariableArgumentFormulaFunctionAttribute"/> classes.
''' All methods that you wish to be able to be used in formulas must be marked with one of those attributes.  By
''' doing so, you give the formula engine information about the number and type of arguments that your function requires.
''' This allows the engine to only call your function with the correct number and type of arguments and
''' eliminates the need for each function author to write manual validation code.</remarks>
''' <example>This example shows a method tagged with this attribute.  The formula engine will only call this method
''' if it is called with exactly one argument and that argument can be converted into a double.
''' <code>
''' &lt;FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})&gt; _
''' Public Sub PlusOne(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
'''		Dim value as Double = args(0).ValueAsDouble
'''		result.SetValue(value + 1)
''' End Sub
''' </code>
''' </example>
<Serializable(), AttributeUsage(AttributeTargets.Method)> _
Public MustInherit Class FormulaFunctionAttribute
	Inherits Attribute

	Protected Sub New()

	End Sub

	Friend MustOverride Function IsValidMinArgCount(ByVal count As Integer) As Boolean
	Friend MustOverride Function IsValidMaxArgCount(ByVal count As Integer) As Boolean
	Friend MustOverride Function MarshalArgument(ByVal position As Integer, ByVal op As IOperand) As ArgumentMarshalResult
End Class

''' <summary>
''' Marks a method as a formula function that takes a fixed number of arguments
''' </summary>
''' <remarks>By tagging a method with this attribute, you are informing the formula engine that your method expects a fixed
''' number of arguments and lets you specify their type.  All calls to a function marked with this attribute will only happen
''' if the number and type of arguments match the ones specified.</remarks>
''' <example>The following example declares the method Tan as a formula function taking one argument of type Double
''' <code>
''' &lt;FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})&gt; _
''' Public Sub Tan(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
''' End Sub
''' </code>
''' </example>
<Serializable()> _
Public Class FixedArgumentFormulaFunctionAttribute
	Inherits FormulaFunctionAttribute

	Private MyMinArgumentCount, MyMaxArgumentCount As Integer
	Private MyArgumentTypes As OperandType()

	''' <summary>
	''' Declares a formula function with an optional number of arguments
	''' </summary>
	''' <param name="minArgumentCount">The minimum number of arguments your function expects</param>
	''' <param name="maxArgumentCount">The maximum number of arguments your function expects</param>
	''' <param name="argumentTypes">An array of <see cref="T:ciloci.FormulaEngine.OperandType"/> that specifies the type of each argument your function expects</param>
	''' <remarks>Use this constructor when your function can have optional arguments.  The formula engine will allow calls to your function
	''' as long as at least minArgumentCount arguments and no more than maxArgumentCount arguments are specified.  Calls with a number of arguments
	''' between the two values will be allowed and it is up to you to interpret the values of the unspecified arguments.  You must specify the
	''' type of all arguments including optional ones.</remarks>
	''' <exception cref="T:System.ArgumentOutOfRangeException">
	''' <para>minArgumentCount is negative</para>
	''' <para>maxArgumentCount exceeds the <see cref="F:ciloci.FormulaEngine.FunctionLibrary.MAX_ARGUMENT_COUNT">maximum</see> number of allowed arguments</para>
	''' <para>maxArgumentCount is less than minArgumentCount</para>
	''' <para>The number of argument types is not equal to maxArgumentCount</para>
	''' </exception>
	''' <exception cref="T:System.ArgumentNullException">
	''' <para>argumentTypes is null</para>
	''' </exception>
	''' <example>The following example declares the method Ceiling as a formula function that be called with one or two Double arguments
	''' <code>
	''' &lt;FixedArgumentFormulaFunction(1, 2, New OperandType() {OperandType.Double, OperandType.Double})&gt; _
	''' Public Sub Ceiling(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
	''' End Sub
	''' </code>
	''' </example>
	Public Sub New(ByVal minArgumentCount As Integer, ByVal maxArgumentCount As Integer, ByVal argumentTypes As OperandType())
		MyMinArgumentCount = minArgumentCount
		MyMaxArgumentCount = maxArgumentCount
		MyArgumentTypes = argumentTypes

		If minArgumentCount < 0 Or maxArgumentCount > FunctionLibrary.MAX_ARGUMENT_COUNT Or maxArgumentCount < minArgumentCount Then
			Throw New ArgumentOutOfRangeException("argumentCount", "Invalid argument count value")
		End If

		If argumentTypes Is Nothing Then
			Throw New ArgumentNullException("argumentTypes")
		End If

		If argumentTypes.Length <> maxArgumentCount Then
			Throw New ArgumentOutOfRangeException("argumentTypes", "Invalid number of argument types")
		End If
	End Sub

	''' <summary>
	''' Declares a formula function with a fixed number of arguments
	''' </summary>
	''' <param name="argumentCount">The number of arguments your function expects</param>
	''' <param name="argumentTypes">The type of each argument</param>
	''' <remarks>Use this constructor when your function requires an exact number of arguments.  You must specify
	''' the count you want and the type of each argument.</remarks>
	''' <example>The following example declares a method as taking one argument of type Double
	''' <code>
	''' &lt;FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})&gt; _
	''' Public Sub Tan(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
	''' End Sub
	''' </code>
	''' </example>
	Public Sub New(ByVal argumentCount As Integer, ByVal argumentTypes As OperandType())
		Me.New(argumentCount, argumentCount, argumentTypes)
	End Sub

	Friend Overrides Function IsValidMaxArgCount(ByVal count As Integer) As Boolean
		Return count <= MyMaxArgumentCount
	End Function

	Friend Overrides Function IsValidMinArgCount(ByVal count As Integer) As Boolean
		Return count >= MyMinArgumentCount
	End Function

	Friend Overrides Function MarshalArgument(ByVal position As Integer, ByVal op As IOperand) As ArgumentMarshalResult
		' Get the operand type we expect
		Dim opType As OperandType = MyArgumentTypes(position)
		' Try to convert
		Dim result As IOperand = op.Convert(opType)

		Dim success As Boolean = Not result Is Nothing

		If result Is Nothing Then
			' Conversion failed; get the error
			result = Me.GetErrorOperand(op)
		End If

		' Return the result of marshaling
		Return New ArgumentMarshalResult(success, result)
	End Function

	Private Function GetErrorOperand(ByVal op As IOperand) As IOperand
		Dim errorOp As IOperand = op.Convert(OperandType.Error)

		If Not errorOp Is Nothing Then
			Return errorOp
		Else
			Return New ErrorValueOperand(ErrorValueType.Value)
		End If
	End Function
End Class

''' <summary>
''' Marks a formula function as taking a variable number of arguments
''' </summary>
''' <remarks>Use this attribute on a method that you wish to act as a formula function with a variable number of arguments.
''' The formula engine will allow calls to your method as long as it is called with at least one and no more than <see cref="F:ciloci.FormulaEngine.FunctionLibrary.MAX_ARGUMENT_COUNT"/>
''' arguments.  No validation is done on the type of each argument; it is up to you to examine each one and act on it as you see fit</remarks>
''' <example>The following example declares the method sum as a formula function taking a variable number of arguments
''' <code>
''' &lt;VariableArgumentFormulaFunction()&gt; _
''' Public Sub Sum(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
''' End Sub
''' </code>
''' </example>
<Serializable()> _
Public Class VariableArgumentFormulaFunctionAttribute
	Inherits FormulaFunctionAttribute

	Public Sub New()

	End Sub

	Friend Overrides Function IsValidMaxArgCount(ByVal count As Integer) As Boolean
		Return count <= FunctionLibrary.MAX_ARGUMENT_COUNT
	End Function

	Friend Overrides Function IsValidMinArgCount(ByVal count As Integer) As Boolean
		' Do like excel and require at least 1 argument
		Return count >= 1
	End Function

	Friend Overrides Function MarshalArgument(ByVal position As Integer, ByVal op As IOperand) As ArgumentMarshalResult
		' Marshaling always succeeds in these types of functions
		Return New ArgumentMarshalResult(True, op)
	End Function
End Class

''' <summary>
''' Marks a formula function as being volatile
''' </summary>
''' <remarks>Apply this attribute to a formula function to make the formula engine treat it as volatile.
''' Formulas with volatile functions are always included in any recalculations even if they do not depend on any of the references
''' being recalculated.  The typical reason why a function would be marked as volatile is that it creates or uses dynamic references
''' and thus the formula engine cannot determine any dependencies.  Without this attribute, such a function would never be recalculated.</remarks>
<AttributeUsage(AttributeTargets.Method)> _
Public Class VolatileFunctionAttribute
	Inherits Attribute
End Class

Friend Class ArgumentComparer
	Implements System.Collections.Generic.IComparer(Of Argument)

	Private MyOriginalArgs As Argument()

	Public Sub New(ByVal originalArgs As Argument())
		MyOriginalArgs = originalArgs
	End Sub

	Public Function Compare(ByVal x As Argument, ByVal y As Argument) As Integer Implements System.Collections.Generic.IComparer(Of Argument).Compare
		Dim xIsRef, yIsRef As Boolean
		xIsRef = x.IsReference
		yIsRef = y.IsReference

		If xIsRef = True And yIsRef = False Then
			Return 1
		ElseIf xIsRef = False And yIsRef = False Then
			Return Me.CompareIndexes(x, y)
		ElseIf xIsRef = True And yIsRef = True Then
			Return Me.CompareIndexes(x, y)
		Else
			Return -1
		End If
	End Function

	Private Function CompareIndexes(ByVal x As Argument, ByVal y As Argument) As Integer
		Return Me.GetIndex(x).CompareTo(Me.GetIndex(y))
	End Function

	Private Function GetIndex(ByVal a As Argument) As Integer
		Return System.Array.IndexOf(MyOriginalArgs, a)
	End Function
End Class

''' <summary>
''' Base class for predicates used in conditional functions like SumIf
''' </summary>
Friend MustInherit Class SheetValuePredicate

	Private MyCompareType As CompareType

	Protected Sub New()

	End Sub

	Public Shared Function Create(ByVal criteria As Object) As SheetValuePredicate
		Dim pred As SheetValuePredicate

		Dim s As String = TryCast(criteria, String)

		If Not s Is Nothing Then
			Dim info As New StringCriteriaInfo(s)
			pred = info.CreatePredicate()
		Else
			pred = ComparerBasedPredicate.CreateDefault(criteria)
			pred.SetCompareType(ciloci.FormulaEngine.CompareType.Equal)
		End If

		Return pred
	End Function

	Public Sub SetCompareType(ByVal ct As CompareType)
		MyCompareType = ct
	End Sub

	Public MustOverride Function IsMatch(ByVal value As Object) As Boolean

	Protected ReadOnly Property CompareType() As CompareType
		Get
			Return MyCompareType
		End Get
	End Property
End Class

Friend MustInherit Class NonNullPredicate
	Inherits SheetValuePredicate

	Private MyTarget As Object

	Protected Sub New(ByVal target As Object)
		MyTarget = Utility.NormalizeIfNumericValue(target)
	End Sub

	Protected MustOverride Function Compare(ByVal value As Object, ByVal target As Object) As Integer

	Public Overrides Function IsMatch(ByVal value As Object) As Boolean
		If value Is Nothing Then
			Return Me.CompareType = ciloci.FormulaEngine.CompareType.NotEqual
		End If

		value = Utility.NormalizeIfNumericValue(value)

		If Me.IsValidComparison(value, MyTarget) = False Then
			Return Me.CompareType = ciloci.FormulaEngine.CompareType.NotEqual
		Else
			Dim cr As CompareResult
			cr = LogicalOperator.Compare2CompareResult(Me.Compare(value, MyTarget))
			Return LogicalOperator.GetBooleanResult(Me.CompareType, cr)
		End If
	End Function

	Private Function IsValidComparison(ByVal value As Object, ByVal target As Object) As Boolean
		Return value.GetType() Is target.GetType()
	End Function
End Class

Friend Class ComparerBasedPredicate
	Inherits NonNullPredicate

	Private MyComparer As IComparer

	Public Sub New(ByVal target As Object, ByVal comparer As IComparer)
		MyBase.New(target)
		MyComparer = comparer
	End Sub

	Public Shared Function CreateDefault(ByVal target As Object) As ComparerBasedPredicate
		Return New ComparerBasedPredicate(target, New Comparer(System.Globalization.CultureInfo.CurrentCulture))
	End Function

	Protected Overrides Function Compare(ByVal value As Object, ByVal target As Object) As Integer
		Return MyComparer.Compare(value, target)
	End Function
End Class

Friend Class MatchNullPredicate
	Inherits SheetValuePredicate

	Public Overrides Function IsMatch(ByVal value As Object) As Boolean
		If value Is Nothing Then
			Return Me.CompareType = ciloci.FormulaEngine.CompareType.Equal
		Else
			Return Me.CompareType = ciloci.FormulaEngine.CompareType.NotEqual
		End If
	End Function
End Class

Friend Class MatchNullOrEmptyStringPredicate
	Inherits SheetValuePredicate

	Public Overrides Function IsMatch(ByVal value As Object) As Boolean
		If value Is Nothing Then
			Return True
		End If

		Dim s As String = TryCast(value, String)
		Return Not s Is Nothing AndAlso s.Length = 0
	End Function
End Class

Friend Class WildcardPredicate
	Inherits SheetValuePredicate

	Private MyRegex As System.Text.RegularExpressions.Regex

	Public Sub New(ByVal criteria As String)
		criteria = String.Concat("^", Utility.Wildcard2Regex(criteria), "$")
		MyRegex = New System.Text.RegularExpressions.Regex(criteria, Text.RegularExpressions.RegexOptions.IgnoreCase)
	End Sub

	Public Overrides Function IsMatch(ByVal value As Object) As Boolean
		If value Is Nothing OrElse Not value.GetType() Is GetType(String) Then
			Return False
		Else
			If MyRegex.IsMatch(DirectCast(value, String)) = True Then
				Return Me.CompareType = ciloci.FormulaEngine.CompareType.Equal
			Else
				Return Me.CompareType = ciloci.FormulaEngine.CompareType.NotEqual
			End If
		End If
	End Function
End Class

Friend MustInherit Class LookupProcessor

	Public MustOverride Function GetLookupVector() As Object()
	Public MustOverride Function GetResultVector() As Object()
End Class

Friend MustInherit Class HVLookupProcessor
	Inherits LookupProcessor

	Protected MyIndex As Integer
	Protected MyTable As Object(,)

	Public Sub Initialize(ByVal table As Object(,), ByVal index As Integer)
		MyTable = table
		MyIndex = index
	End Sub

	Public MustOverride Function IsValidIndex(ByVal refRect As System.Drawing.Rectangle, ByVal index As Integer) As Boolean
End Class

Friend Class HLookupProcessor
	Inherits HVLookupProcessor

	Public Overrides Function IsValidIndex(ByVal refRect As System.Drawing.Rectangle, ByVal index As Integer) As Boolean
		Return refRect.Top + index <= refRect.Bottom
	End Function

	Public Overrides Function GetLookupVector() As Object()
		Return Utility.GetTableRow(MyTable, 0)
	End Function

	Public Overrides Function GetResultVector() As Object()
		Return Utility.GetTableRow(MyTable, MyIndex - 1)
	End Function
End Class

Friend Class VLookupProcessor
	Inherits HVLookupProcessor

	Public Overrides Function IsValidIndex(ByVal refRect As System.Drawing.Rectangle, ByVal index As Integer) As Boolean
		Return refRect.Left + index <= refRect.Right
	End Function

	Public Overrides Function GetLookupVector() As Object()
		Return Utility.GetTableColumn(MyTable, 0)
	End Function

	Public Overrides Function GetResultVector() As Object()
		Return Utility.GetTableColumn(MyTable, MyIndex - 1)
	End Function
End Class

Friend Class PlainLookupProcessor
	Inherits LookupProcessor

	Private MyLookupVector As Object()
	Private MyResultVector As Object()

	Public Sub Initialize(ByVal lookupVector As Object(), ByVal resultVector As Object())
		MyLookupVector = lookupVector
		MyResultVector = resultVector
	End Sub

	Public Overrides Function GetLookupVector() As Object()
		Return MyLookupVector
	End Function

	Public Overrides Function GetResultVector() As Object()
		Return MyResultVector
	End Function
End Class

Friend MustInherit Class ConditionalSheetProcessor
	Public MustOverride Sub OnMatch(ByVal row As Integer, ByVal col As Integer)
End Class

Friend Class SumIfConditionalSheetProcessor
	Inherits ConditionalSheetProcessor

	Private MySumValues As Object(,)
	Private MyValues As Generic.IList(Of Double)

	Public Sub New()
		MyValues = New Generic.List(Of Double)
	End Sub

	Public Sub Initialize(ByVal sumValues As Object(,))
		MySumValues = sumValues
	End Sub

	Public Overrides Sub OnMatch(ByVal row As Integer, ByVal col As Integer)
		Dim value As Object = MySumValues(row, col)
		value = Utility.NormalizeIfNumericValue(value)

		If Not value Is Nothing AndAlso value.GetType() Is GetType(Double) Then
			MyValues.Add(DirectCast(value, Double))
		End If
	End Sub

	Public ReadOnly Property Values() As Generic.IList(Of Double)
		Get
			Return MyValues
		End Get
	End Property
End Class

Friend Class CountIfConditionalSheetProcessor
	Inherits ConditionalSheetProcessor

	Private MyCount As Integer

	Public Overrides Sub OnMatch(ByVal row As Integer, ByVal col As Integer)
		MyCount += 1
	End Sub

	Public ReadOnly Property Count() As Integer
		Get
			Return MyCount
		End Get
	End Property
End Class

Friend Structure StringCriteriaInfo
	Public [Operator] As String
	Public Value As Object

	Public Sub New(ByVal criteria As String)
		Me.GetInfo(criteria)
	End Sub

	Private Sub GetInfo(ByVal criteria As String)
		Dim stringValue As String
		If criteria.StartsWith("<>") Or criteria.StartsWith(">=") Or criteria.StartsWith("<=") Then
			Me.Operator = criteria.Substring(0, 2)
			stringValue = criteria.Substring(2)
		ElseIf criteria.StartsWith("=") Or criteria.StartsWith(">") Or criteria.StartsWith("<") Then
			Me.Operator = criteria.Substring(0, 1)
			stringValue = criteria.Substring(1)
		Else
			Me.Operator = Nothing
			stringValue = criteria
		End If

		Me.Value = Utility.Parse(stringValue)
	End Sub

	Public Function CreatePredicate() As SheetValuePredicate
		Dim pred As SheetValuePredicate

		Dim s As String = TryCast(Me.Value, String)

		If s Is Nothing Then
			pred = ComparerBasedPredicate.CreateDefault(Me.Value)
		Else
			pred = Me.CreateStringPredicate(s)
		End If

		If Me.Operator Is Nothing Then
			pred.SetCompareType(CompareType.Equal)
		Else
			pred.SetCompareType(Me.GetCompareType())
		End If

		Return pred
	End Function

	Private Function CreateStringPredicate(ByVal value As String) As SheetValuePredicate
		If value.Length = 0 Then
			If Me.Operator Is Nothing Then
				Return New MatchNullOrEmptyStringPredicate()
			Else
				Return New MatchNullPredicate()
			End If
		ElseIf value.IndexOfAny(New Char() {"?", "*"}) <> -1 Then
			Return New WildcardPredicate(value)
		Else
			Return New ComparerBasedPredicate(value, StringComparer.OrdinalIgnoreCase)
		End If
	End Function

	Private Function GetCompareType() As CompareType
		Select Case Me.Operator
			Case "="
				Return ciloci.FormulaEngine.CompareType.Equal
			Case "<>"
				Return ciloci.FormulaEngine.CompareType.NotEqual
			Case ">="
				Return ciloci.FormulaEngine.CompareType.GreaterThanOrEqual
			Case "<="
				Return ciloci.FormulaEngine.CompareType.LessThanOrEqual
			Case ">"
				Return ciloci.FormulaEngine.CompareType.GreaterThan
			Case "<"
				Return ciloci.FormulaEngine.CompareType.LessThan
			Case Else
				Debug.Assert(False, "unknown operator")
		End Select
	End Function
End Structure