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

Imports System.Reflection

''' <summary>
''' Manages all functions that can be used in formulas
''' </summary>
''' <remarks>This class is responsible for managing all functions used in formulas and for marshalling arguments during function calls.
''' It has methods for defining and undefining formula functions in bulk or individually.  The library comes with many of the most
''' common Excel functions already defined.  If a function is not defined in this class then an error will be generated when
''' a formula tries to use it.
''' <para>Adding your own function to the library requires the following steps:
''' <list type="bullet">
''' <item>Define a method that has the same signature as the <see cref="T:ciloci.FormulaEngine.FormulaFunctionCall">delegate</see></item>
''' <item>Tag the method with either the <see cref="T:ciloci.FormulaEngine.FixedArgumentFormulaFunctionAttribute"/> or <see cref="T:ciloci.FormulaEngine.VariableArgumentFormulaFunctionAttribute"/> attributes</item>
''' <item>Add the function to the function library using one of the appropriate methods</item>
''' </list>
''' </para>
''' <note>Function names are treated case-insensitively
''' <para>You cannot define/undefine functions while formulas are defined in the formula engine</para>
''' </note>
''' </remarks>
<Serializable()> _
Public Class FunctionLibrary
	Implements System.Runtime.Serialization.ISerializable

	Private Interface IDelegateCreator
		Function CreateDelegate(ByVal methodName As String) As FormulaFunctionCall
		ReadOnly Property Flags() As BindingFlags
	End Interface

	Private Class StaticDelegateCreator
		Implements IDelegateCreator

		Private MyTarget As Type

		Public Sub New(ByVal target As Type)
			MyTarget = target
		End Sub

		Public Function CreateDelegate(ByVal methodName As String) As FormulaFunctionCall Implements IDelegateCreator.CreateDelegate
			Return System.Delegate.CreateDelegate(GetType(FormulaFunctionCall), MyTarget, methodName, True, False)
		End Function

		Public ReadOnly Property Flags() As System.Reflection.BindingFlags Implements IDelegateCreator.Flags
			Get
				Return BindingFlags.Static
			End Get
		End Property
	End Class

	Private Class InstanceDelegateCreator
		Implements IDelegateCreator

		Private MyTarget As Object

		Public Sub New(ByVal target As Object)
			MyTarget = target
		End Sub

		Public Function CreateDelegate(ByVal methodName As String) As FormulaFunctionCall Implements IDelegateCreator.CreateDelegate
			Return System.Delegate.CreateDelegate(GetType(FormulaFunctionCall), MyTarget, methodName, True, False)
		End Function

		Public ReadOnly Property Flags() As System.Reflection.BindingFlags Implements IDelegateCreator.Flags
			Get
				Return BindingFlags.Instance
			End Get
		End Property
	End Class

	''' <summary>
	''' Stores all information about a function
	''' </summary>
	<Serializable()> _
 Private Class FunctionInfo

		Public FunctionTarget As FormulaFunctionCall
		Public FunctionAttribute As FormulaFunctionAttribute
		Public Volatile As Boolean

		Public Sub New(ByVal target As FormulaFunctionCall, ByVal attr As FormulaFunctionAttribute)
			Me.FunctionTarget = target
			Me.FunctionAttribute = attr
			Me.Volatile = Attribute.IsDefined(target.Method, GetType(VolatileFunctionAttribute))
		End Sub
	End Class

	Private MyFunctions As IDictionary
	Private MyBuiltinFunctions As BuiltinFunctions
	Private MyOwner As FormulaEngine
	''' <summary>The maximum number of arguments that any function can be called with</summary>
	''' <remarks>This limit is arbitrary but is implemented to prevent functions being called with an unreasonable number of arguments</remarks>
	Public Const MAX_ARGUMENT_COUNT As Integer = 30
	Private Const VERSION As Integer = 1

	Friend Sub New(ByVal owner As FormulaEngine)
		MyOwner = owner
		MyBuiltinFunctions = New BuiltinFunctions
		MyFunctions = New Hashtable(StringComparer.OrdinalIgnoreCase)
		Me.AddBuiltinFunctions()
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyOwner = info.GetValue("Owner", GetType(FormulaEngine))
		MyBuiltinFunctions = info.GetValue("BuiltinFunctions", GetType(BuiltinFunctions))
		MyFunctions = info.GetValue("Functions", GetType(IDictionary))
	End Sub

	Public Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
		info.AddValue("Owner", MyOwner)
		info.AddValue("BuiltinFunctions", MyBuiltinFunctions)
		info.AddValue("Functions", MyFunctions)
	End Sub

	''' <summary>
	''' Adds all builtin functions to the library
	''' </summary>
	''' <remarks>Use this method when you want to add all the builtin functions to the library.  Builtin functions are added by
	''' default when the function library is created.</remarks>
	Public Sub AddBuiltinFunctions()
		Me.AddInstanceFunctions(MyBuiltinFunctions)
	End Sub

	''' <summary>
	''' Adds all instance methods of the given object that can be formula functions
	''' </summary>
	''' <param name="instance">The object whose methods you wish to add</param>
	''' <remarks>Use this function when you want to add a large number of functions in bulk.  The method will search all instance methods
	''' of the type.  Methods that are tagged with a formula function attribute and have the correct signature will be added to the library.</remarks>
	''' <exception cref="T:System.ArgumentNullException">instance is null</exception>
	''' <exception cref="T:System.ArgumentException">A method is tagged with a formula function attribute but does not have the correct signature</exception>
	''' <exception cref="T:System.InvalidOperationException">
	''' The function was called while formulas are defined in the formula engine
	''' <para>A function with the same name is already defined</para>
	''' </exception>
	Public Sub AddInstanceFunctions(ByVal instance As Object)
		FormulaEngine.ValidateNonNull(instance, "instance")
		Dim creator As New InstanceDelegateCreator(instance)
		Me.AddFormulaMethods(instance.GetType(), creator)
	End Sub

	''' <summary>
	''' Adds all static methods of the given type that can be formula functions
	''' </summary>
	''' <param name="target">The type to examine</param>
	''' <remarks>This method works similarly to <see cref="M:ciloci.FormulaEngine.FunctionLibrary.AddInstanceFunctions(System.Object)"/> except that it looks at all static methods instead.</remarks>
	''' <exception cref="T:System.ArgumentNullException">target is null</exception>
	''' <exception cref="T:System.ArgumentException">A method is tagged with a formula function attribute but does not have the correct signature</exception>
	''' <exception cref="T:System.InvalidOperationException">
	''' The function was called while formulas are defined in the formula engine
	''' <para>A function with the same name is already defined</para>
	''' </exception>
	Public Sub AddStaticFunctions(ByVal target As Type)
		FormulaEngine.ValidateNonNull(target, "target")
		Dim creator As New StaticDelegateCreator(target)
		Me.AddFormulaMethods(target, creator)
	End Sub

	''' <summary>
	''' Go through all methods of a type and try to add them as formula functions
	''' </summary>
	Private Sub AddFormulaMethods(ByVal targetType As Type, ByVal creator As IDelegateCreator)
		Me.ValidateEngineStateForChangingFunctions()
		Dim flags As BindingFlags = BindingFlags.NonPublic Or BindingFlags.Public Or creator.Flags

		For Each mi As MethodInfo In targetType.GetMethods(flags)
			Dim attr As FormulaFunctionAttribute = Attribute.GetCustomAttribute(mi, GetType(FormulaFunctionAttribute))
			If Not attr Is Nothing Then
				Dim d As FormulaFunctionCall = creator.CreateDelegate(mi.Name)
				If d Is Nothing Then
					Throw New ArgumentException(String.Format("The method {0} is marked as a formula function but does not have the correct signature", mi.Name))
				End If
				Dim info As New FunctionInfo(d, attr)
                Me.AddFunctionInternal(info, d.Method.Name)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Adds an individual formula function
    ''' </summary>
    ''' <param name="functionCall">A delegate pointing to the method you wish to add</param>
    ''' <remarks>This function lets you add an individual formula function by specifying a delegate pointing to it.  The method that
    ''' the delegate refers to must be tagged with the appropriate <see cref="T:ciloci.FormulaEngine.FormulaFunctionAttribute">attribute</see>.</remarks>
    ''' <exception cref="T:System.ArgumentException">The method that the delegate points to is not tagged with the required attribute</exception>
    ''' <exception cref="T:System.InvalidOperationException">
    ''' The function was called while formulas are defined in the formula engine
    ''' <para>A function with the same name is already defined</para>
    ''' </exception>
    Public Sub AddFunction(ByVal functionCall As FormulaFunctionCall)
        Me.AddFunction(functionCall, functionCall.Method.Name)
    End Sub

    ''' <summary>
    ''' Adds an individual formula function
    ''' </summary>
    ''' <param name="functionCall">A delegate pointing to the method you wish to add</param>
    ''' <param name="functionName">The name of the excel method you wish to handle</param>
    ''' <remarks>This function lets you add an individual formula function by specifying a delegate pointing to it.  The method that
    ''' the delegate refers to must be tagged with the appropriate <see cref="T:ciloci.FormulaEngine.FormulaFunctionAttribute">attribute</see>.</remarks>
    ''' <exception cref="T:System.ArgumentException">The method that the delegate points to is not tagged with the required attribute</exception>
    ''' <exception cref="T:System.InvalidOperationException">
    ''' The function was called while formulas are defined in the formula engine
    ''' <para>A function with the same name is already defined</para>
    ''' </exception>
    Public Sub AddFunction(ByVal functionCall As FormulaFunctionCall, ByVal functionName As String)
        FormulaEngine.ValidateNonNull(functionCall, "functionCall")
        Me.ValidateEngineStateForChangingFunctions()
        Dim attr As FormulaFunctionAttribute = Attribute.GetCustomAttribute(functionCall.Method, GetType(FormulaFunctionAttribute))

        If attr Is Nothing Then
            Throw New ArgumentException("The function does not have a FormulaFunctionAttribute defined on it")
        End If

        Dim info As New FunctionInfo(functionCall, attr)
        Me.AddFunctionInternal(info, functionName)
    End Sub

    Private Sub ValidateEngineStateForChangingFunctions()
        If MyOwner.FormulaCount > 0 Then
            Throw New InvalidOperationException("Cannot add or remove functions while formulas are defined")
        End If
    End Sub

    Private Sub AddFunctionInternal(ByVal info As FunctionInfo, ByVal functionName As String)
        Dim name As String = functionName
        If MyFunctions.Contains(name) = True Then
            Throw New InvalidOperationException(String.Format("A function with the name {0} is already defined", name))
        Else
            MyFunctions.Add(name, info)
        End If
    End Sub

	''' <summary>
	''' Undefines an individual function
	''' </summary>
	''' <param name="functionName">The name of the function you wish to undefine</param>
	''' <remarks>This method removes a function from the library</remarks>
	''' <exception cref="System.ArgumentException">The given function name is not defined</exception>
	''' <exception cref="T:System.InvalidOperationException">The function was called while formulas are defined in the formula engine</exception>
	Public Sub RemoveFunction(ByVal functionName As String)
		FormulaEngine.ValidateNonNull(functionName, "functionName")
		Me.ValidateEngineStateForChangingFunctions()
		If MyFunctions.Contains(functionName) = False Then
			Throw New ArgumentException("That function is not defined")
		Else
			MyFunctions.Remove(functionName)
		End If
	End Sub

	Friend Function IsValidArgumentCount(ByVal functionName As String, ByVal argCount As Integer) As Boolean
		Dim info As FunctionInfo = Me.GetFunctionInfo(functionName)
		Dim attr As FormulaFunctionAttribute = info.FunctionAttribute
		Return attr.IsValidMinArgCount(argCount) And attr.IsValidMaxArgCount(argCount)
	End Function

	Friend Function IsFunctionDefined(ByVal functionName As String) As Boolean
		Return MyFunctions.Contains(functionName)
	End Function

	''' <summary>
	''' Perform a function call
	''' </summary>
	Friend Sub InvokeFunction(ByVal functionName As String, ByVal state As System.Collections.Stack, ByVal argumentCount As Integer)
		Dim result As New FunctionResult
		' Find the info for the function
		Dim info As FunctionInfo = Me.GetFunctionInfo(functionName)
		' Get all the operands from teh stack
		Dim ops As IOperand() = Me.GetArguments(state, argumentCount)
		Dim resultOperand As IOperand

		' Marshal the arguments
		resultOperand = Me.MarshalArguments(info, ops)

		If resultOperand Is Nothing Then
			' Marshaling succeeded so we can call the function
			Dim args As Argument() = Me.CreateArguments(ops)
			info.FunctionTarget(args, result, MyOwner)

			resultOperand = result.Operand
			Me.ValidateResultOperand(functionName, resultOperand)
		End If

		state.Push(resultOperand)
	End Sub

	Private Sub ValidateResultOperand(ByVal functionName As String, ByVal resultOperand As IOperand)
		If resultOperand Is Nothing Then
			Throw New InvalidOperationException(String.Format("Function {0} did not return a result", functionName))
		End If
	End Sub

	''' <summary>
	''' Perform validation on operands passed to a function
	''' </summary>
	Private Function MarshalArguments(ByVal info As FunctionInfo, ByVal ops As IOperand()) As ErrorValueOperand
		Dim attr As FormulaFunctionAttribute = info.FunctionAttribute

		For i As Integer = 0 To ops.Length - 1
			Dim op As IOperand = ops(i)

			' We don't want invalid references
			If Me.IsInvalidReference(op) = True Then
				Return New ErrorValueOperand(ErrorValueType.Ref)
			End If

			Dim result As ArgumentMarshalResult = attr.MarshalArgument(i, op)

			If result.Success = False Then
				Return result.Result
			Else
				ops(i) = result.Result
			End If
		Next

		Return Nothing
	End Function

	Private Function IsInvalidReference(ByVal op As IOperand) As Boolean
		Dim ref As Reference = op.Convert(OperandType.Reference)

		If ref Is Nothing Then
			Return False
		Else
			Return ref.Valid = False
		End If
	End Function

	Private Function CreateArguments(ByVal ops As IOperand()) As Argument()
		Dim args(ops.Length - 1) As Argument

		For i As Integer = 0 To ops.Length - 1
			args(i) = New Argument(ops(i))
		Next

		Return args
	End Function

	Private Function GetArguments(ByVal state As System.Collections.Stack, ByVal argumentCount As Integer) As IOperand()
		Dim arr(argumentCount - 1) As IOperand

		For i As Integer = 0 To argumentCount - 1
			Dim op As IOperand = state.Pop()
			arr(i) = op
		Next

		Return arr
	End Function

	Private Function GetFunctionInfo(ByVal functionName As String) As FunctionInfo
		Dim info As FunctionInfo = MyFunctions.Item(functionName)
		Debug.Assert(Not info Is Nothing, "expected to find function")
		Return info
	End Function

	Private Function GetDelegate(ByVal functionName As String) As FormulaFunctionCall
		Dim info As FunctionInfo = Me.GetFunctionInfo(functionName)
		Return info.FunctionTarget
	End Function

	''' <summary>
	''' Gets the names of all defined functions
	''' </summary>
	''' <returns>An array consisting of the names of all defined functions</returns>
	''' <remarks>Use this method when you need the names of all defined functions</remarks>
	Public Function GetFunctionNames() As String()
		Dim arr(MyFunctions.Keys.Count - 1) As String
		MyFunctions.Keys.CopyTo(arr, 0)
		Return arr
	End Function

	Friend Function IsFunctionVolatile(ByVal functionName As String) As Boolean
		Dim info As FunctionInfo = Me.GetFunctionInfo(functionName)
		Return info.Volatile
	End Function

	''' <summary>
	''' Undefines all functions
	''' </summary>
	''' <remarks>This method undefines all functions in the library</remarks>
	''' <exception cref="T:System.InvalidOperationException">The function was called while formulas are defined in the formula engine</exception>
	Public Sub Clear()
		Me.ValidateEngineStateForChangingFunctions()
		MyFunctions.Clear()
	End Sub

	''' <summary>
	''' Gets the number of functions defined in the library
	''' </summary>
	''' <value>A count of the number of defined functions</value>
	''' <remarks>Use this property when you need to know the number of defined functions in the library</remarks>
	Public ReadOnly Property FunctionCount() As Integer
		Get
			Return MyFunctions.Keys.Count
		End Get
	End Property
End Class

''' <summary>
''' Represents an argument to a formula function
''' </summary>
''' <remarks>This class represents an argument passed to a formula function.  Every such function will
''' receive an array of instances of this class; one for each argument the function was called with.  The class has properties
''' for determining the type of the argument passed and getting its value.</remarks>
Public NotInheritable Class Argument

	Private MyOperand As IOperand

	Friend Sub New(ByVal op As IOperand)
		MyOperand = op
	End Sub

	''' <summary>
	''' Determines if this argument can be converted to a particular type
	''' </summary>
	''' <param name="opType">The operand type you wish to test</param>
	''' <returns>True if the argument can be converted to opType; False otherwise</returns>
	''' <remarks>This method is a more generic version of the IsXXX properties</remarks>
	Public Function IsType(ByVal opType As OperandType) As Boolean
		Return Not MyOperand.Convert(opType) Is Nothing
	End Function

	''' <summary>
	''' Gets the value of an argument as a <see cref="System.Double"/>
	''' </summary>
	''' <value>The value of the argument as a <see cref="System.Double"/></value>
	''' <remarks>This property will try to convert the value of the argument to a <see cref="System.Double"/>.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to a <see cref="System.Double"/></exception>
	Public ReadOnly Property ValueAsDouble() As Double
		Get
			Dim op As DoubleOperand = MyOperand.Convert(OperandType.Double)
			If op Is Nothing Then
				Throw New InvalidOperationException("Conversion failed")
			Else
				Return op.ValueAsDouble
			End If
		End Get
	End Property

	''' <summary>
	''' Gets the value of an argument as an <see cref="System.Int32"/>
	''' </summary>
	''' <value>The value of the argument as an <see cref="System.Int32"/></value>
	''' <remarks>This property will try to convert the value of the argument to an <see cref="System.Int32"/>.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to an <see cref="System.Int32"/></exception>
	Public ReadOnly Property ValueAsInteger() As Integer
		Get
			Dim op As IntegerOperand = MyOperand.Convert(OperandType.Integer)
			If op Is Nothing Then
				Throw New InvalidOperationException("Conversion failed")
			Else
				Return op.ValueAsInteger
			End If
		End Get
	End Property

	''' <summary>
	''' Gets the value of an argument as a <see cref="System.String"/>
	''' </summary>
	''' <value>The value of the argument as a <see cref="System.String"/></value>
	''' <remarks>This property will try to convert the value of the argument to a <see cref="System.String"/>.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to a <see cref="System.String"/></exception>
	Public ReadOnly Property ValueAsString() As String
		Get
			Dim op As StringOperand = MyOperand.Convert(OperandType.String)
			If op Is Nothing Then
				Return Nothing
			Else
				Return op.ValueAsString
			End If
		End Get
	End Property

	''' <summary>
	''' Gets the value of an argument as a <see cref="System.Boolean"/>
	''' </summary>
	''' <value>The value of the argument as a <see cref="System.Boolean"/></value>
	''' <remarks>This property will try to convert the value of the argument to a <see cref="System.Boolean"/>.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to a <see cref="System.Boolean"/></exception>
	Public ReadOnly Property ValueAsBoolean() As Boolean
		Get
			Dim op As BooleanOperand = MyOperand.Convert(OperandType.Boolean)
			If op Is Nothing Then
				Throw New InvalidOperationException("Conversion failed")
			Else
				Return op.ValueAsBoolean
			End If
		End Get
	End Property

	''' <summary>
	''' Gets the value of an argument as a <see cref="System.DateTime"/>
	''' </summary>
	''' <value>The value of the argument as a <see cref="System.DateTime"/></value>
	''' <remarks>This property will try to convert the value of the argument to a <see cref="System.DateTime"/>.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to a <see cref="System.DateTime"/></exception>
	Public ReadOnly Property ValueAsDateTime() As DateTime
		Get
			Dim op As DateTimeOperand = MyOperand.Convert(OperandType.DateTime)
			If op Is Nothing Then
				Throw New InvalidOperationException("Conversion failed")
			Else
				Return op.ValueAsDateTime
			End If
		End Get
	End Property

	''' <summary>
	''' Gets the value of an argument as an <see cref="T:System.ciloci.FormulaEngine.ErrorValueType"/>
	''' </summary>
	''' <value>The value of the argument as an <see cref="T:System.ciloci.FormulaEngine.ErrorValueType"/></value>
	''' <remarks>This property will try to convert the value of the argument to an <see cref="T:System.ciloci.FormulaEngine.ErrorValueType"/>.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to an <see cref="T:System.ciloci.FormulaEngine.ErrorValueType"/></exception>
	Public ReadOnly Property ValueAsError() As ErrorValueType
		Get
			Dim op As ErrorValueOperand = MyOperand.Convert(OperandType.Error)

			If op Is Nothing Then
				Throw New InvalidOperationException("Conversion failed")
			Else
				Return DirectCast(op, ErrorValueOperand).ValueAsErrorType
			End If
		End Get
	End Property

	''' <summary>
	''' Gets the value of an argument as a reference
	''' </summary>
	''' <value>The value of the argument as a reference</value>
	''' <remarks>This property will try to convert the value of the argument to a reference.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to a reference</exception>
	Public ReadOnly Property ValueAsReference() As IReference
		Get
			Dim ref As IReference = MyOperand.Convert(OperandType.Reference)

			If ref Is Nothing Then
				Throw New InvalidOperationException("Conversion failed")
			Else
				Return ref
			End If
		End Get
	End Property

	''' <summary>
	''' Gets the value of an argument as a primitive
	''' </summary>
	''' <value>The value of the argument as a primitive</value>
	''' <remarks>This property will try to convert the value of the argument to a primitive.  A primitive is any datatype except a reference.</remarks>
	''' <exception cref="T:System.InvalidOperationException">The value could not be converted to a primitive</exception>
	Public ReadOnly Property ValueAsPrimitive() As Object
		Get
			Dim prim As IOperand = MyOperand.Convert(OperandType.Primitive)

			If prim Is Nothing Then
				Throw New InvalidOperationException("Conversion failed")
			Else
				Return prim.Value
			End If
		End Get
	End Property

	Friend ReadOnly Property ValueAsOperand() As IOperand
		Get
			Return MyOperand
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is a double
	''' </summary>
	''' <value>True if the argument can be converted to a double; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.</remarks>
	Public ReadOnly Property IsDouble() As Boolean
		Get
			Return Me.IsType(OperandType.Double)
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is an integer
	''' </summary>
	''' <value>True if the argument can be converted to an integer; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.</remarks>
	Public ReadOnly Property IsInteger() As Boolean
		Get
			Return Me.IsType(OperandType.Integer)
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is a string
	''' </summary>
	''' <value>True if the argument can be converted to a string; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.</remarks>
	Public ReadOnly Property IsString() As Boolean
		Get
			Return Me.IsType(OperandType.String)
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is a boolean
	''' </summary>
	''' <value>True if the argument can be converted to a boolean; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.</remarks>
	Public ReadOnly Property IsBoolean() As Boolean
		Get
			Return Me.IsType(OperandType.Boolean)
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is a reference
	''' </summary>
	''' <value>True if the argument can be converted to a reference; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.</remarks>
	Public ReadOnly Property IsReference() As Boolean
		Get
			Return Me.IsType(OperandType.Reference)
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is an error value
	''' </summary>
	''' <value>True if the argument can be converted to an error value; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.</remarks>
	Public ReadOnly Property IsError() As Boolean
		Get
			Return Me.IsType(OperandType.Error)
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is a DateTime
	''' </summary>
	''' <value>True if the argument can be converted to a DateTime; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.</remarks>
	Public ReadOnly Property IsDateTime() As Boolean
		Get
			Return Me.IsType(OperandType.DateTime)
		End Get
	End Property

	''' <summary>
	''' Indicates whether this argument is a primitive
	''' </summary>
	''' <value>True if the argument can be converted to a primitive; False otherwise</value>
	''' <remarks>Use this property to test if the argument can be converted to a particular data type before trying to get its value.
	''' A primitive is any data type except a reference.</remarks>
	Public ReadOnly Property IsPrimitive() As Boolean
		Get
			Return Me.IsType(OperandType.Primitive)
		End Get
	End Property
End Class

''' <summary>
''' Represents the result of a formula function
''' </summary>
''' <remarks>This class is responsible for storing the result of a formula function.  It has methods for storing a result of
''' various data types.  An instance of it is passed to all methods acting as formula functions and each such method must produce
''' a result and store it in the passed instance or an exception will be raised.</remarks>
''' <example>This example shows a formula function that expects one argument of type double and sets its result as that value incremented by one:
''' <code>
''' &lt;FixedArgumentFormulaFunction(1, New OperandType() {OperandType.Double})&gt; _
''' Public Sub PlusOne(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)
'''		Dim value as Double = args(0).ValueAsDouble
'''		result.SetValue(value + 1)
''' End Sub
''' </code>
''' </example>
Public NotInheritable Class FunctionResult

	Private MyOperand As IOperand

	Friend Sub New()

	End Sub

	''' <summary>
	''' Sets the formula function result to a double
	''' </summary>
	''' <param name="value">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when the result of your function is a double</remarks>
	Public Sub SetValue(ByVal value As Double)
		If OperatorBase.IsInvalidDouble(value) = True Then
			Me.SetError(ErrorValueType.Num)
		Else
			MyOperand = New DoubleOperand(value)
		End If
	End Sub

	''' <summary>
	''' Sets the formula function result to an integer
	''' </summary>
	''' <param name="value">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when the result of your function is an integer</remarks>
	Public Sub SetValue(ByVal value As Integer)
		MyOperand = New IntegerOperand(value)
	End Sub

	''' <summary>
	''' Sets the formula function result to a boolean
	''' </summary>
	''' <param name="value">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when the result of your function is a boolean</remarks>
	Public Sub SetValue(ByVal value As Boolean)
		MyOperand = New BooleanOperand(value)
	End Sub

	''' <summary>
	''' Sets the formula function result to a string
	''' </summary>
	''' <param name="value">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when the result of your function is a string</remarks>
	Public Sub SetValue(ByVal value As String)
		FormulaEngine.ValidateNonNull(value, "value")
		MyOperand = New StringOperand(value)
	End Sub

	''' <summary>
	''' Sets the formula function result to a DateTime
	''' </summary>
	''' <param name="value">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when the result of your function is a DateTime</remarks>
	Public Sub SetValue(ByVal value As DateTime)
		MyOperand = New DateTimeOperand(value)
	End Sub

	''' <summary>
	''' Sets the formula function result to an error
	''' </summary>
	''' <param name="value">The type of the error you wish to be the result of the function</param>
	''' <remarks>Use this method when you need to return an error as the result of your function</remarks>
	Public Sub SetError(ByVal value As ErrorValueType)
		MyOperand = New ErrorValueOperand(value)
	End Sub

	''' <summary>
	''' Sets the formula function result to a reference
	''' </summary>
	''' <param name="value">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when the result of your function is a reference.  Usually used when wishing to implement dynamic references</remarks>
	Public Sub SetValue(ByVal value As IReference)
		FormulaEngine.ValidateNonNull(value, "value")
		MyOperand = value
	End Sub

	''' <summary>
	''' Sets the formula function result to a given argument
	''' </summary>
	''' <param name="arg">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when you wish to use one of the arguments supplied to your function as its result without altering the value.
	''' The if function, for example, uses this method.</remarks>
	Public Sub SetValue(ByVal arg As Argument)
		FormulaEngine.ValidateNonNull(arg, "arg")
		MyOperand = arg.ValueAsOperand
	End Sub

	''' <summary>
	''' Sets the formula function result to a sheet value
	''' </summary>
	''' <param name="value">The value you wish to be the result of the function</param>
	''' <remarks>Use this method when you have a value on a sheet that you wish to use as the result of your function</remarks>
	Public Sub SetValueFromSheet(ByVal value As Object)
		MyOperand = OperandFactory.CreateDynamic(value)
	End Sub

	Friend ReadOnly Property Operand() As IOperand
		Get
			Return MyOperand
		End Get
	End Property
End Class