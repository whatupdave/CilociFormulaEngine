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
''' Defines constants for all errors that a formula can generate during evaluation
''' </summary>
''' <remarks>Formulas that produce an error during calculation will produce an <see cref="T:ciloci.FormulaEngine.ErrorValueWrapper"/> around
''' one of these values.  The values map directly to the values used by Excel.</remarks>
Public Enum ErrorValueType
	''' <summary>A division by zero was encountered</summary>
	Div0
	''' <summary>A formula referenced a name that is not defined</summary>
	Name
	''' <summary>A result is not available</summary>
	NA
	''' <summary>The two given sheet references do not intersect</summary>
	Null
	''' <summary>A calculation is invalid for numerical reasons</summary>
	Num
	''' <summary>A reference is not valid</summary>
	Ref
	''' <summary>The given value is not valid</summary>
	Value
End Enum

''' <summary>
''' Defines constants to represent the data types of a formula function's arguments
''' </summary>
''' <remarks>These values are used when working with functions to specify the desired data type of an argument and to get its value</remarks>
Public Enum OperandType
	''' <summary>A standard .NET Double</summary>
	[Double]
	''' <summary>A standard .NET String</summary>
	[String]
	''' <summary>A standard .NET Boolean</summary>
	[Boolean]
	''' <summary>Any type of reference</summary>
	[Reference]
	''' <summary>A reference that points to cells on a sheet</summary>
	SheetReference
	''' <summary>A standard .NET Integer</summary>
	[Integer]
	''' <summary>A conversion of an operand to itself</summary>
	Self
	''' <summary>Any data type except a reference</summary>
	Primitive
	''' <summary>An error value</summary>
	[Error]
	''' <summary>A blank value</summary>
	Blank
	''' <summary>A standard .NET DateTime</summary>
	DateTime
End Enum

''' <summary>
''' The exception that is thrown when attempting to create an invalid formula
''' </summary>
''' <remarks>
''' This exception will be thrown when attempting to create a formula that is invalid.  The most common (though not the only) reason is that
''' the syntax of the formula does not conform to the parser's grammar.  The inner exception will always be initialized and will
''' contain the specifics as to why the formula could not be created.  
''' </remarks>
Public Class InvalidFormulaException
	Inherits Exception

	Friend Sub New(ByVal innerException As Exception)
		MyBase.New("The formula is not valid", innerException)
	End Sub
End Class

''' <summary>
''' Represents a worksheet as seen by the formula engine
''' </summary>
''' <remarks>This interface defines the contract that any class wishing to act as a worksheet must implement.  All interaction
''' with cell values is done through this interface.</remarks>
Public Interface ISheet
	''' <summary>
	''' Gets the value of a particular cell
	''' </summary>
	''' <param name="row">The row of the required cell.  First row is 1</param>
	''' <param name="column">The column of the required cell.  First column is 1</param>
	''' <returns>The value of the cell at row,col</returns>
	''' <remarks>The formula engine will call this method when the value of a cell is required.  Your implementation should
	''' lookup the cell at row,col and return its value</remarks>
	Function GetCellValue(ByVal row As Integer, ByVal column As Integer) As Object
	''' <summary>
	''' Stores the value of a formula into a cell
	''' </summary>
	''' <param name="result">The formula result that must be stored</param>
	''' <param name="row">The row of the required cell.  First row is 1</param>
	''' <param name="column">The column of the required cell.  First column is 1</param>
	''' <remarks>The formula engine will call this method when the result of a grid formula should be stored into the sheet.
	''' Your implementation should lookup the cell at row,col and store the result into it.</remarks>
	Sub SetFormulaResult(ByVal result As Object, ByVal row As Integer, ByVal column As Integer)
	''' <summary>
	''' Gets the name of the worksheet
	''' </summary>
	''' <value>The name of the worksheet</value>
	''' <remarks>The name of a worksheet is used by the formula engine to find a sheet when its name is used in a reference.
	''' For example: When evaluating the formula "=Sheet3!A1 * 2", the formula engine will look through all sheets until
	''' it finds the one with the name "Sheet3".</remarks>
	''' <note>Sheet names are treated without regard to case</note>
	ReadOnly Property Name() As String
	''' <summary>
	''' Gets number of rows in the sheet
	''' </summary>
	''' <value>The number of rows in the sheet</value>
	''' <remarks>This property is used by the engine to determine sheet bounds</remarks>
	ReadOnly Property RowCount() As Integer
	''' <summary>
	''' Gets number of columns in the sheet
	''' </summary>
	''' <value>The number of columns in the sheet</value>
	''' <remarks>This property is used by the engine to determine sheet bounds</remarks>
	ReadOnly Property ColumnCount() As Integer
End Interface

''' <summary>
''' Represents a reference to a value or set of values
''' </summary>
''' <remarks>This is the base interface for all references.  A reference is simply a pointer to a value or formula.  
''' All formulas that are managed by the formula engine must be associated with a reference.  
''' You can change the context of a formula by binding it to different types of references.</remarks>
Public Interface IReference
	''' <summary>Signals that the reference's formula has been recalculated</summary>
	''' <remarks>You can listen to this event to be notified when the formula that is bound to this reference has been recalculated.</remarks>
	Event Recalculated As EventHandler
	''' <summary>Gets the values that the reference points to</summary>
	''' <param name="processor">A class responsible from processing the reference's values</param>
	''' <remarks>You should call this method when you need reference's values.  The reference will pass each of its values to the
	''' processor and it is up to that processor to store or use them to compute a result.</remarks>
	Sub GetReferenceValues(ByVal processor As IReferenceValueProcessor)
	''' <summary>Determines whether this reference equals another</summary>
	''' <param name="ref">The reference to test against</param>
	''' <returns>True is the current reference is equal to ref.  False otherwise</returns>
	''' <remarks>This method exists mostly for testing purposes.</remarks>
	Function Equals(ByVal ref As IReference) As Boolean
	''' <summary>Returns a formatted representation of the reference</summary>
	''' <remarks>This method allows you to get a string representation of the reference</remarks>
	Function ToString() As String
End Interface

''' <summary>
''' Represents a reference to cells on a sheet
''' </summary>
''' <remarks>Sheet references are the most common type of reference and consist of a sheet and an area on that sheet.  Any formula 
''' that needs values from a sheet will use sheet references and all formulas that are on a sheet must be bound to them.</remarks>
Public Interface ISheetReference
	Inherits IReference

	''' <summary>Returns a table of the reference's values</summary>
	''' <remarks>This function returns a table that represents the reference's values from its sheet.  The first dimension is the rows
	''' and the second dimension is columns.  This method is useful when you wish to do lookups on a sheet reference's values.</remarks>
	Function GetValuesTable() As Object(,)
	''' <summary>Gets row of the reference</summary>
	''' <remarks>The row of the reference on its sheet</remarks>
	ReadOnly Property Row() As Integer
	''' <summary>Gets the column of the reference</summary>
	''' <remarks>The column of the reference on its sheet</remarks>
	ReadOnly Property Column() As Integer
	''' <summary>Gets number of rows in the reference</summary>
	''' <remarks>The number of rows the reference spans on its sheet</remarks>
	ReadOnly Property Height() As Integer
	''' <summary>Gets the number of columns in the reference</summary>
	''' <remarks>The number of columns the reference spans on its sheet</remarks>
	ReadOnly Property Width() As Integer
	''' <summary>The reference's area as a rectangle</summary>
	''' <remarks>A convenience property for getting a reference's area as a rectangle</remarks>
	ReadOnly Property Area() As System.Drawing.Rectangle
	''' <summary>Gets the sheet that the reference is on</summary>
	''' <remarks>This property lets you access the sheet that the reference is on</remarks>
	ReadOnly Property Sheet() As ISheet
End Interface

''' <summary>
''' Represents a reference to a name
''' </summary>
''' <remarks>Named references allow you to associate a formula with a name.  By binding a formula to a named reference,
''' you make it possible to use that formula's result in other formulas by simply typing the name.  This can make formulas cleaner and less complex since you can
''' reuse a particular result in many formulas rather than duplicating the same expression in each one.  The formula engine will recalculate
''' all formulas that depend on a name when the value of the formula bound to the name changes.</remarks>
''' <example>This example shows how you can define a constant and use it in another formula:
''' <code>
''' Dim engine As New FormulaEngine
''' ' Add a constant named InterestRate with a value of 0.15
''' engine.AddFormula("=0.15", engine.ReferenceFactory.Named("InterestRate"))
''' ' Use the constant in a formula
''' Dim result As Object = engine.Evaluate("=1000 * InterestRate")
''' </code>
''' </example>
Public Interface INamedReference
	Inherits IReference
	''' <summary>
	''' Gets the name of the reference
	''' </summary>
	''' <value>The name of the reference</value>
	''' <remarks>This property lets you obtain the name of the reference</remarks>
	ReadOnly Property Name() As String
	''' <summary>Gets result of evaluating the reference's formula</summary>
	''' <value>The result of the reference's formula</value>
	''' <remarks>Use this property to get the result of evaluating this reference's formula</remarks>
	ReadOnly Property Result() As Object
End Interface

''' <summary>
''' Represents a reference outside of a sheet
''' </summary>
''' <remarks>External references are similar to named references in that you can use both without any sheets.  A limitation of named
''' references is that each name must be unique.  This can make it difficult to define many non-sheet formulas because you have to
''' generate a unique name for each one.  External references do not have this limitation and you can create as many as you need.  
''' Their only drawback is that they cannot be referenced in formulas like named references.  This basically makes them only useful as
''' bind targets for formulas.</remarks>
''' <note>Since each instance of an external reference is unique, you must keep track of the instance you create because you will
''' need to supply it when you wish to get the formula bound to it from the engine.</note>
Public Interface IExternalReference
	Inherits IReference
	''' <summary>Gets result of evaluating the reference's formula</summary>
	''' <value>The result of the reference's formula</value>
	''' <remarks>You typically will listen to the <see cref="E:ciloci.FormulaEngine.IExternalReference.Recalculated"/> event and when it fires you use this property in your handler to get the latest
	''' value of the reference's formula.</remarks>
	ReadOnly Property Result() As Object
End Interface

''' <summary>
''' Implemented by classes that process a reference's values
''' </summary>
''' <remarks>An implementation of this interface can be passed to a reference's <see cref="M:ciloci.FormulaEngine.IReference.GetReferenceValues(ciloci.FormulaEngine.IReferenceValueProcessor)"/> method.  The 
''' reference will call the ProcessValue method on each of the values it represents.</remarks>
Public Interface IReferenceValueProcessor
	''' <summary>Processes a reference value</summary>
	''' <param name="value">The value from the reference that is to be processed</param>
	''' <returns>True to keep processing values; False to stop</returns>
	''' <remarks>This method will be called once for each value that a reference represents.  Classes that implement this interface
	''' must decide what to do with the value.</remarks>
	Function ProcessValue(ByVal value As Object) As Boolean
End Interface

Friend Enum CompareType
	LessThan
	LessThanOrEqual
	Equal
	NotEqual
	GreaterThan
	GreaterThanOrEqual
End Enum

Friend Enum CompareResult
	LessThan
	Equal
	GreaterThan
	InvalidComparison
End Enum

''' <summary>
''' Implemented by objects that compute a formula's result
''' </summary>
Friend Interface IFormulaComponent
	Inherits ICloneable
	Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine)
	Sub EvaluateForDependencyReference(ByVal references As IList, ByVal engine As FormulaEngine)
	Sub Validate(ByVal engine As FormulaEngine)
End Interface

''' <summary>
''' Implemented by objects that act as operands to operators and functions in a formula
''' </summary>
Friend Interface IOperand
	Function Convert(ByVal convertType As OperandType) As IOperand
	ReadOnly Property Value() As Object
	ReadOnly Property NativeType() As OperandType
End Interface

''' <summary>
''' Represents the method that will handle a formula function call
''' </summary>
''' <param name="args">All the arguments that the function was called with</param>
''' <param name="result">The object where the function's result will be stored</param>
''' <param name="engine">A reference to the formula engine</param>
''' <remarks>All methods that you wish to be able to be called from within a formula must have the signature of this delegate.</remarks>
Public Delegate Sub FormulaFunctionCall(ByVal args As Argument(), ByVal result As FunctionResult, ByVal engine As FormulaEngine)

''' <summary>
''' Base class for a predicate on a reference
''' </summary>
Friend MustInherit Class ReferencePredicateBase
	Public MustOverride Function IsMatch(ByVal ref As Reference) As Boolean
End Class

''' <summary>
''' Result from doing an operation on a reference
''' </summary>
Friend Enum ReferenceOperationResultType
	NotAffected = 0
	Affected = 1
	Invalidated = 2
End Enum

''' <summary>
''' A convenient wrapper around an error value
''' </summary>
''' <remarks>This class encapsulates the parsing and formatting of an <see cref="T:ciloci.FormulaEngine.ErrorValueType"/>.
''' It exists so that the error returned by a formula will be nicely formatted without any additional work on the person working
''' with the formula engine.</remarks>
<Serializable()> _
Public Class ErrorValueWrapper

	Private MyErrorValue As ErrorValueType
	Private Const DIV0_STRING As String = "#DIV/0!"
	Private Const NA_STRING As String = "#N/A"
	Private Const NAME_STRING As String = "#NAME?"
	Private Const NULL_STRING As String = "#NULL!"
	Private Const REF_STRING As String = "#REF!"
	Private Const VALUE_STRING As String = "#VALUE!"
	Private Const NUM_STRING As String = "#NUM!"

	Friend Sub New(ByVal value As ErrorValueType)
		MyErrorValue = value
	End Sub

	''' <summary>
	''' Convenience function for equality
	''' </summary>
	''' <param name="obj">The value to test against</param>
	''' <returns>True if the wrapper equals obj</returns>
	''' <remarks>Compares the current wrapper against another value for equality.</remarks>
	Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
		If TypeOf (obj) Is ErrorValueType Then
			Return MyErrorValue = DirectCast(obj, ErrorValueType)
		ElseIf TypeOf (obj) Is ErrorValueWrapper Then
			Return MyErrorValue = DirectCast(obj, ErrorValueWrapper).MyErrorValue
		Else
			Return False
		End If
	End Function

	''' <summary>
	''' Tries to parse a string in to an error value wrapper
	''' </summary>
	''' <param name="s">The string to parse</param>
	''' <returns>A wrapper instance if the string was sucessfully parsed; a null reference otherwise</returns>
	''' <remarks>Use this function when you wish to parse a string into an error value wrapper.
	''' The function recognizes the following strings:
	''' <list type="bullet">
	''' <item>"#DIV/0!"</item>
	''' <item>"#N/A"</item>
	''' <item>"#NAME?"</item>
	''' <item>"#NULL!"</item>
	''' <item>"#REF!"</item>
	''' <item>"#VALUE!"</item>
	''' <item>"#NUM!"</item>
	''' </list>
	''' </remarks>
	Public Shared Function TryParse(ByVal s As String) As ErrorValueWrapper
		Dim ev As ErrorValueType

		If s.Equals(DIV0_STRING, StringComparison.OrdinalIgnoreCase) = True Then
			ev = ErrorValueType.Div0
		ElseIf s.Equals(NA_STRING, StringComparison.OrdinalIgnoreCase) = True Then
			ev = ErrorValueType.NA
		ElseIf s.Equals(NAME_STRING, StringComparison.OrdinalIgnoreCase) = True Then
			ev = ErrorValueType.Name
		ElseIf s.Equals(NULL_STRING, StringComparison.OrdinalIgnoreCase) = True Then
			ev = ErrorValueType.Null
		ElseIf s.Equals(REF_STRING, StringComparison.OrdinalIgnoreCase) = True Then
			ev = ErrorValueType.Ref
		ElseIf s.Equals(VALUE_STRING, StringComparison.OrdinalIgnoreCase) = True Then
			ev = ErrorValueType.Value
		ElseIf s.Equals(NUM_STRING, StringComparison.OrdinalIgnoreCase) = True Then
			ev = ErrorValueType.Num
		Else
			Return Nothing
		End If

		Return New ErrorValueWrapper(ev)
	End Function

	''' <summary>
	''' Formats the inner error value
	''' </summary>
	''' <returns>A string with the formatted error</returns>
	''' <remarks>This method will format an error value similarly to Excel.  For example: the error value Ref will
	''' be formatted as "#REF!"</remarks>
	Public Overrides Function ToString() As String
		Select Case MyErrorValue
			Case ErrorValueType.Div0
				Return DIV0_STRING
			Case ErrorValueType.NA
				Return NA_STRING
			Case ErrorValueType.Name
				Return NAME_STRING
			Case ErrorValueType.Null
				Return NULL_STRING
			Case ErrorValueType.Ref
				Return REF_STRING
			Case ErrorValueType.Value
				Return VALUE_STRING
			Case ErrorValueType.Num
				Return NUM_STRING
			Case Else
				Throw New InvalidOperationException("Unknown value")
		End Select
	End Function

	''' <summary>
	''' Gets the actual error value that the class contains
	''' </summary>
	''' <value>The error value</value>
	''' <remarks>Returns the error value that the wrapper contains.</remarks>
	Public ReadOnly Property ErrorValue() As ErrorValueType
		Get
			Return MyErrorValue
		End Get
	End Property
End Class

''' <summary>
''' Base class for properties that are specific to a reference
''' </summary>
<Serializable()> _
Friend MustInherit Class ReferenceProperties
	Implements ICloneable

	Public Function Clone() As Object Implements System.ICloneable.Clone
		Dim copy As ReferenceProperties = Me.MemberwiseClone()
		Me.InitializeClone(copy)
		Return copy
	End Function

	Protected Overridable Sub InitializeClone(ByVal clone As ReferenceProperties)

	End Sub
End Class

''' <summary>
''' Properties resulting from parsing a reference.  Used to pass information from the analyzer to a formula
''' </summary>
Friend Class ReferenceParseProperties
	Public SheetName As String
End Class

''' <summary>
''' Implemented by references that can have a formula bound to them
''' </summary>
Friend Interface IFormulaSelfReference
	Sub OnFormulaRecalculate(ByVal target As Formula)
End Interface

Friend Structure ArgumentMarshalResult
	Public Success As Boolean
	Public Result As IOperand

	Public Sub New(ByVal success As Boolean, ByVal result As IOperand)
		Me.Success = success
		Me.Result = result
	End Sub
End Structure

<Serializable()> _
Friend Structure ReferencePoolInfo
	Public Target As Reference
	Public Count As Integer

	Public Sub New(ByVal target As Reference)
		Me.Target = target
		Me.Count = 1
	End Sub
End Structure

' Holds all information about a parsed reference
Friend Class ReferenceParseInfo
	Public Target As Reference
	Public Location As System.Drawing.CharacterRange
	Public Properties As ReferenceProperties
	Public ParseProperties As ReferenceParseProperties
End Class

''' <summary>
''' Contains information about the cause of circular references
''' </summary>
Public Class CircularReferenceDetectedEventArgs
	Inherits EventArgs

	Private MyRoots As IReference()

	Friend Sub New(ByVal roots As IReference())
		MyRoots = roots
	End Sub

	''' <summary>
	''' The root of each circular reference
	''' </summary>
	''' <value>An array of the circular reference roots</value>
	''' <remarks>Each reference in this array is the root of a circular reference.  That is, by starting at the reference and
	''' following all its dependents, you will eventually wind up at the same reference.</remarks>
	Public ReadOnly Property Roots() As IReference()
		Get
			Return MyRoots
		End Get
	End Property
End Class

''' <summary>
''' Defines constants to represent different grammars to use when parsing an expression
''' </summary>
''' <remarks>Using the constants in this enumeration, you can specify a particular grammar to use when parsing an expression.</remarks>
Public Enum GrammarType
	''' <summary>A grammar suitable for parsing Excel-style formulas</summary>
	Excel
	''' <summary>A grammar suitable for general expressions which don't require cell and range references</summary>
	General
End Enum

Friend Structure ParseInfo
	Public RootElement As ParseTreeElement
	Public Infos As ReferenceParseInfo()

	Public Sub New(ByVal root As ParseTreeElement, ByVal infos As ReferenceParseInfo())
		Me.RootElement = root
		Me.Infos = infos
	End Sub
End Structure