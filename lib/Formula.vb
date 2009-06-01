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
''' Represents a parsed formula
''' </summary>
''' <remarks>Instances of this class are created by the formula engine after parsing an expression.  The class contains the compiled
''' form of the given expression, exposes some useful properties, and allows you to evaluate the formula and format it.
''' </remarks>
<Serializable()> _
Public Class Formula
	Implements System.Runtime.Serialization.ISerializable

	Private MyComponents As IFormulaComponent()					' The compiled array of our operands and operators
	Private MyOwner As FormulaEngine
	Private MyRawReferences As Reference()						' References used when we format
	Private MyDependencyReferences As Reference()				' References that this formula has
	Private MySelfReference As Reference						' Reference that this formula is bound to
	Private MyTemplate As String								' Template string for formatting
	Private MyReferenceProperties As ReferenceProperties()		' Properties we specific to each reference
	Private MyResultType As OperandType							' Type of result to produce

	Private Const VERSION As Integer = 1

	Private Sub New()

	End Sub

	Friend Sub New(ByVal owner As FormulaEngine, ByVal expression As String, ByVal gt As GrammarType)
		MyOwner = owner
		' By default, we evaluate to a primitive
		MyResultType = OperandType.Primitive

		Dim pi As ParseInfo = Me.Parse(expression, gt)
		Dim rootElement As ParseTreeElement = pi.RootElement
		Dim infos As ReferenceParseInfo() = pi.Infos

		Me.ProcessParseInfos(infos)
		MyTemplate = Me.CreateTemplateString(expression, infos)

		MyComponents = Me.CreateComponents(rootElement)
		Me.ValidateComponents()
		Me.ComputeRawReferenceHashCodes()
		Me.ComputeDependencyReferences()
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyOwner = info.GetValue("Engine", GetType(FormulaEngine))
		MyComponents = info.GetValue("Components", GetType(IFormulaComponent()))
		MyRawReferences = info.GetValue("RawReferences", GetType(Reference()))
		MyDependencyReferences = info.GetValue("DependencyReferences", GetType(Reference()))
		MySelfReference = info.GetValue("SelfReference", GetType(Reference))
		MyTemplate = info.GetString("Template")
		MyReferenceProperties = info.GetValue("ReferenceProperties", GetType(ReferenceProperties()))
	End Sub

	Public Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
		info.AddValue("Engine", MyOwner)
		info.AddValue("Components", MyComponents)
		info.AddValue("RawReferences", MyRawReferences)
		info.AddValue("DependencyReferences", MyDependencyReferences)
		info.AddValue("SelfReference", MySelfReference)
		info.AddValue("Template", MyTemplate)
		info.AddValue("ReferenceProperties", MyReferenceProperties)
	End Sub

	Private Function CreateComponents(ByVal rootElement As ParseTreeElement) As IFormulaComponent()
		Dim list As IList = New ArrayList
		rootElement.AddAsRPN(list)

		Dim arr(list.Count - 1) As IFormulaComponent
		list.CopyTo(arr, 0)
		Return arr
	End Function

	Private Sub ValidateComponents()
		For Each component As IFormulaComponent In MyComponents
			component.Validate(MyOwner)
		Next
	End Sub

	''' <summary>
	''' Create the template string that we use for formatting this formula.  It will have placeholders where all the
	''' references will go.  This way, the formula's formatted value will be updated as our references change.
	''' </summary>
	Private Function CreateTemplateString(ByVal expression As String, ByVal refInfos As ReferenceParseInfo()) As String
		Dim sb As New System.Text.StringBuilder(expression)
		Dim offset As Integer = 0
		Dim ordinal As Integer = 0

		' Create a place holder for each reference
		For i As Integer = 0 To MyRawReferences.Length - 1
			Dim ref As Reference = MyRawReferences(i)
			Dim range As System.Drawing.CharacterRange = refInfos(i).Location
			Dim placeHolder As String = "{" & ordinal & "}"
			range.First += offset
			sb.Remove(range.First, range.Length)
			sb.Insert(range.First, placeHolder)
			offset += placeHolder.Length - range.Length
			ordinal += 1
		Next

		Return sb.ToString()
	End Function

	Private Sub ProcessParseInfos(ByVal infos As ReferenceParseInfo())
		MyRawReferences = New Reference(infos.Length - 1) {}
		MyReferenceProperties = New ReferenceProperties(infos.Length - 1) {}

		For i As Integer = 0 To infos.Length - 1
			Dim info As ReferenceParseInfo = infos(i)
			Dim ref As Reference = info.Target
			Dim props As ReferenceParseProperties = info.ParseProperties
			ref.SetEngine(MyOwner)
			ref.ProcessParseProperties(props, MyOwner)
			MyRawReferences(i) = ref
			MyReferenceProperties(i) = info.Properties
		Next
	End Sub

	Private Sub ComputeRawReferenceHashCodes()
		For i As Integer = 0 To MyRawReferences.Length - 1
			Dim ref As Reference = MyRawReferences(i)
			ref.ComputeHashCode()
		Next
	End Sub

	Friend Function Clone() As Formula
		Dim cloneFormula As Formula = New Formula()
		cloneFormula.MyOwner = MyOwner
		cloneFormula.MyTemplate = MyTemplate
		cloneFormula.MyResultType = MyResultType

		Me.CloneComponents(cloneFormula)
		cloneFormula.ComputeDependencyReferences()
		cloneFormula.MyReferenceProperties = New ReferenceProperties(MyReferenceProperties.Length - 1) {}
		Me.CloneArray(MyReferenceProperties, cloneFormula.MyReferenceProperties)

		Return cloneFormula
	End Function

	Private Sub CloneComponents(ByVal cloneFormula As Formula)
		cloneFormula.MyComponents = New IFormulaComponent(MyComponents.Length - 1) {}
		cloneFormula.MyRawReferences = New Reference(MyRawReferences.Length - 1) {}
		Dim index As Integer = 0

		' We clone all the components and make sure that the component and its corresponding raw reference 
		' are the same since this is how a new formula would be setup.
		For i As Integer = 0 To MyComponents.Length - 1
			Dim component As IFormulaComponent = MyComponents(i)
			Dim cloneComponent As IFormulaComponent = component.Clone()
			cloneFormula.MyComponents(i) = cloneComponent
			Dim ref As Reference = TryCast(cloneComponent, Reference)
			If Not ref Is Nothing Then
				cloneFormula.MyRawReferences(index) = ref
				index += 1
			End If
		Next
	End Sub

	Private Sub CloneArray(ByVal source As Object(), ByVal dest As Object())
		For i As Integer = 0 To source.Length - 1
			dest(i) = DirectCast(source(i), ICloneable).Clone()
		Next
	End Sub

	''' <summary>
	''' Evaluate and return the final operand on the stack
	''' </summary>
	Friend Function EvaluateToOperand() As IOperand
		Dim state As New Stack

		' Simply loop through each component and tell it to evaluate using the provided stack
		For Each component As IFormulaComponent In MyComponents
            MyOwner.CurrentFormula = Me
            component.Evaluate(state, MyOwner)
		Next

		' Get the final operand on the stack
		Debug.Assert(state.Count = 1, "incomplete stack")
		Dim result As IOperand = state.Pop()

		' Get the real value based on our result type
		result = Me.GetResultOperand(result)
		Return result
	End Function

	''' <summary>
	''' Computes the result of the formula
	''' </summary>
	''' <returns>The result of evaluating the formula</returns>
	''' <remarks>This method will compute the result of the current formula.  The formula engine will typically call this method during
	''' a recalculate but you are free to call it anytime you need the latest result of the formula.  Use this method instead of the
	''' formula engine's evaluate method if you have a static expression that you wish to evaluate many times</remarks>
	''' <example>This example shows how you can create a formula and evaluate it.
	''' <code>
	''' Dim engine As New FormulaEngine
	''' Dim f As Formula = engine.CreateFormula("=cos(pi())")
	''' ' result will contain the value -1 as a Double
	''' Dim result As Object = f.Evaluate()
	''' </code>
	''' </example>
	Public Function Evaluate() As Object
		Dim result As IOperand = Me.EvaluateToOperand()
		Return Me.GetFinalValue(result)
	End Function

	''' <summary>
	''' Get the value that we will expose to the outside
	''' </summary>
	Private Function GetFinalValue(ByVal op As IOperand) As Object
		If DirectCast(op, Object).GetType() Is GetType(ErrorValueOperand) Then
			' We want to return error wrappers instead of error enums so that formatting will be automatic
			Return DirectCast(op, ErrorValueOperand).ValueAsErrorWrapper
		Else
			Return op.Value
		End If
	End Function

	''' <summary>
	''' Get the operand based on our result type
	''' </summary>
	Private Function GetResultOperand(ByVal op As IOperand) As IOperand
		op = op.Convert(MyResultType)
		If op Is Nothing Then
			op = New ErrorValueOperand(ErrorValueType.Value)
		End If

		Return op
	End Function

	Friend Sub SetSelfReference(ByVal ref As Reference)
		MySelfReference = ref
	End Sub

	Friend Sub ClearSelfReference()
		MySelfReference = Nothing
	End Sub

	Private Sub ReplaceRawReference(ByVal ordinal As Integer, ByVal newRef As Reference)
		Dim oldRef As Reference = MyRawReferences(ordinal)
		Dim index As Integer = System.Array.IndexOf(MyComponents, oldRef)
		MyComponents(index) = newRef
		MyRawReferences(ordinal) = newRef
	End Sub

	''' <summary>
	''' Compute the references that will be passed onto the dependency manager.
	''' </summary>
	Private Sub ComputeDependencyReferences()
		Dim references As IList = New ArrayList

		For Each c As IFormulaComponent In MyComponents
			c.EvaluateForDependencyReference(references, MyOwner)
		Next

		MyDependencyReferences = Me.GetUniqueValidDependencyReferences(references)
	End Sub

	''' <summary>
	''' Gets only unique and valid references from the computed dependency reference list
	''' </summary>
	Private Function GetUniqueValidDependencyReferences(ByVal refs As IList) As Reference()
		Dim seenReferences As Generic.IDictionary(Of Reference, Reference) = New Generic.Dictionary(Of Reference, Reference)(New ReferenceEqualityComparer)

		For Each ref As Reference In refs
			' Ignore invalid or duplicate references
			If seenReferences.ContainsKey(ref) = False And ref.Valid = True Then
				seenReferences.Add(ref, ref)
			End If
		Next

		Dim arr(seenReferences.Keys.Count - 1) As Reference
		seenReferences.Keys.CopyTo(arr, 0)
		Return arr
	End Function

	Friend Sub SetDependencyReferences(ByVal refs As Reference())
		MyDependencyReferences = refs
		For i As Integer = 0 To refs.Length - 1
			Dim ref As Reference = refs(i)
			Me.ReplaceRawReferences(ref)
		Next
	End Sub

	Friend Function GetDependencyReferences() As Reference()
		Return MyDependencyReferences
	End Function

	Private Sub ReplaceRawReferences(ByVal target As Reference)
		For i As Integer = 0 To MyRawReferences.Length - 1
			Dim ref As Reference = MyRawReferences(i)
			If ref.EqualsReference(target) = True Then
				Me.ReplaceRawReference(i, target)
			End If
		Next
	End Sub

	''' <summary>
	''' Parse an expression and return the root of the parse tree
	''' </summary>
	Private Function Parse(ByVal formula As String, ByVal gt As GrammarType) As ParseInfo
		SyncLock GetType(Object)
			Dim sr As New System.IO.StringReader(formula)
			Dim parser As PerCederberg.Grammatica.Runtime.Parser = ParserFactory.GetParser(gt)
            'parser.Reset(sr)
			parser.Tokenizer.Reset(sr)

			Try
				Dim result As PerCederberg.Grammatica.Runtime.Node = parser.Parse()
				Dim element As ParseTreeElement = result.Values.Item(0)
				Dim analyzer As IAnalyzer = parser.Analyzer
				Dim infos As ReferenceParseInfo() = analyzer.ReferenceInfos
				Dim info As New ParseInfo(element, infos)
				analyzer.Reset()
				Return info
			Catch ex As PerCederberg.Grammatica.Runtime.ParserLogException
				Me.OnParseException(parser)
				Throw New InvalidFormulaException(ex)
			Catch ex As OverflowException
				Me.OnParseException(parser)
				' An number value is too big/small to fit into its type
				Throw New InvalidFormulaException(ex)
			End Try
		End SyncLock
	End Function

	Private Sub OnParseException(ByVal parser As PerCederberg.Grammatica.Runtime.Parser)
		Dim analyzer As IAnalyzer = parser.Analyzer
		analyzer.Reset()
	End Sub

	''' <summary>
	''' Offset all our references for a copy operation
	''' </summary>
	Friend Sub OffsetReferencesForCopy(ByVal destSheet As ISheet, ByVal rowOffset As Integer, ByVal columnOffset As Integer)
		If Not MySelfReference Is Nothing Then
			Throw New InvalidOperationException("Cannot offset references while formula is in engine")
		End If
		For i As Integer = 0 To MyRawReferences.Length - 1
			' Ask each reference to offset itself and pass it its appropriate reference properties
			Dim ref As Reference = MyRawReferences(i)
			ref.OnCopy(rowOffset, columnOffset, destSheet, MyReferenceProperties(i))
			ref.ComputeHashCode()
		Next

		Me.ComputeDependencyReferences()
	End Sub

	Friend Sub GetReferenceProperties(ByVal target As Reference, ByVal dest As IList)
		For i As Integer = 0 To MyRawReferences.Length - 1
			Dim ref As Reference = MyRawReferences(i)
			If target Is ref Or target Is MySelfReference Then
				Dim props As ReferenceProperties = MyReferenceProperties(i)
				dest.Add(props)
			End If
		Next
	End Sub

	''' <summary>
	''' Returns a string representation of the formula
	''' </summary>
	''' <returns>A string representing the formatted value of the formula</returns>
	''' <remarks>This value will be the same as the expression that the formula was created from except that the text for all references
	''' is dynamically updated as the references change.</remarks>
	Public Overrides Function ToString() As String
		Dim refStrings(MyRawReferences.Length - 1) As String

		For i As Integer = 0 To MyRawReferences.Length - 1
			Dim ref As Reference = MyRawReferences(i)
			refStrings(i) = ref.ToStringFormula(MyReferenceProperties(i))
		Next

		Return String.Format(MyTemplate, refStrings)
	End Function

	''' <summary>
	''' Gets the engine that owns this formula
	''' </summary>
	''' <value>The engine that the current formula is bound to</value>
	''' <remarks>All formulas are owned by a formula engine.  This property gets the engine that owns this particular formula.
	''' </remarks>
	Public ReadOnly Property Engine() As FormulaEngine
		Get
			Return MyOwner
		End Get
	End Property

	Friend ReadOnly Property DependencyReferences() As Reference()
		Get
			Return MyDependencyReferences
		End Get
	End Property

	''' <summary>
	''' Gets all the references that this formula uses
	''' </summary>
	''' <value>An array with all the references of the formula</value>
	''' <remarks>Formulas can reference other cells.  The formula engine analyzes each formula for its references so it can
	''' determine dependencies.  This property returns all the references that a formula refers to.</remarks>
	''' <example>A formula such as "=B2+C2" would return an array of two references to cells B2 and C2.  A formula like 
	''' "=cos(pi())" would return a zero length array since it does not reference any cells.</example>
	Public ReadOnly Property References() As IReference()
		Get
			Dim arr(MyRawReferences.Length - 1) As IReference
			System.Array.Copy(MyRawReferences, arr, arr.Length)
			Return arr
		End Get
	End Property

	''' <summary>
	''' Gets the reference that this formula is bound to
	''' </summary>
	''' <value>The reference where this formula is located</value>
	''' <remarks>All formulas are bound to a reference.  This property exposes the reference that this particular formula is bound to.
	''' <note>The value will be null if the formula hasn't been added to a formula engine</note>
	''' </remarks>
	Public ReadOnly Property SelfReference() As IReference
		Get
			Return MySelfReference
		End Get
	End Property

	Friend ReadOnly Property SelfReferenceInternal() As Reference
		Get
			Return MySelfReference
		End Get
	End Property

	''' <summary>
	''' Gets or sets the data type of the formula's result
	''' </summary>
	''' <value>The data type of the result</value>
	''' <remarks>This property gives you control over what the data type of this formula's result will be.
	''' The default is for the formula to evaluate to a primitive.  The most common reason for changing this value is to
	''' have the formula evaluate to a reference instead of the reference's value.  If the result of evaluating the formula
	''' cannot be converted to the requested type, the formula will return a Value error.</remarks>
	''' <example>This sample shows how you can use this property to control what a formula evaluates to:
	''' <code>
	''' Dim engine As New FormulaEngine
	''' Dim f As Formula = Engine.CreateFormula("A1")
	''' f.ResultType = OperandType.Primitive
	''' ' result will have the value in cell A1
	''' Dim result As Object = f.Evaluate()
	''' f.ResultType = OperandType.SheetReference
	''' ' result will be a reference to cell A1
	''' result = f.Evaluate()
	''' </code>
	''' </example>
	Public Property ResultType() As OperandType
		Get
			Return MyResultType
		End Get
		Set(ByVal value As OperandType)
			MyResultType = value
		End Set
	End Property
End Class