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
''' Implements all the functionality of a formula engine
''' </summary>
''' <remarks>This is the main class of this library.  It is responsible for managing formulas, their dependencies, and
''' all recalculations.  You can think of this class as a container for formulas.  It has methods for adding and removing formulas and
''' methods for recalculating all formulas that depend on a given reference.</remarks>
''' <example>This example shows how you can declare a formula that points to a cell and have the formula engine recalculate it
''' when you tell it that the cell has changed:
''' <code>
''' Dim engine As New FormulaEngine
''' ' Declare a reference to cell B2
''' Dim b2Ref As ISheetReference = engine.ReferenceFactory.Parse("B2")
''' ' Add a formula that depends on cell A1 at cell B2
''' engine.AddFormula("A1+1", b2Ref)
''' ' Declare a reference to cell A1
''' Dim a1Ref As ISheetReference = engine.ReferenceFactory.Parse("A1")
''' ' Tell the engine that cell A1 has changed at which point it will recalculate the formula at B2
''' engine.Recalculate(a1Ref)
''' ' Remove the formula at B2
''' engine.RemoveFormulaAt(b2Ref)
''' </code>
''' </example>
<Serializable()> _
Public Class FormulaEngine
	Implements System.Runtime.Serialization.ISerializable

	Private MyFunctionLibrary As FunctionLibrary
	Private MyDependencyManager As DependencyManager
	Private MyReferenceFactory As ReferenceFactory
	Private MyReferencePool As ReferencePool
	Private MyReferenceFormulaMap As IDictionary
	Private MySheets As SheetCollection
    Private MyEngineInfo As FormulaEngineInfo
    Private MyCurrentlyEvaluatingFormula As Formula

	''' <summary>Notifies listeners that the engine has detected one or more circular references</summary>
	''' <param name="sender">An instance of the formula engine</param>
	''' <param name="e">Information about the circular references</param>
	''' <remarks>This event will get fired when the engine detects one or more circular references.  Circular references are allowed
	''' by the engine but will be ignored during any recalculations.  You would typically listen to this event when you want
	''' to notify users that they have caused a circular reference.</remarks>
	Public Event CircularReferenceDetected(ByVal sender As Object, ByVal e As CircularReferenceDetectedEventArgs)
	Private Const VERSION As Integer = 1

	Public Sub New()
		MyReferenceFormulaMap = New Hashtable()
		MyFunctionLibrary = New FunctionLibrary(Me)
		MyDependencyManager = New DependencyManager(Me)
		MyReferenceFactory = New ReferenceFactory(Me)
		MyReferencePool = New ReferencePool(Me)
		MySheets = New SheetCollection(Me)
		MyEngineInfo = New FormulaEngineInfo(Me)
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyDependencyManager = info.GetValue("DependencyManager", GetType(DependencyManager))
		MyReferencePool = info.GetValue("ReferencePool", GetType(ReferencePool))
		MyReferenceFormulaMap = info.GetValue("ReferenceFormulaMap", GetType(IDictionary))
		MySheets = info.GetValue("SheetManager", GetType(SheetCollection))
		MyFunctionLibrary = info.GetValue("FunctionLibrary", GetType(FunctionLibrary))
		MyEngineInfo = New FormulaEngineInfo(Me)
		MyReferenceFactory = New ReferenceFactory(Me)
	End Sub

	Public Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
		info.AddValue("DependencyManager", MyDependencyManager)
		info.AddValue("ReferencePool", MyReferencePool)
		info.AddValue("ReferenceFormulaMap", MyReferenceFormulaMap)
		info.AddValue("SheetManager", MySheets)
		info.AddValue("FunctionLibrary", MyFunctionLibrary)
	End Sub

	''' <summary>
	''' Creates a formula by parsing an expression
	''' </summary>
	''' <param name="expression">The expression to parse</param>
	''' <returns>A formula representing the parsed expression</returns>
	''' <remarks>Works the same way as <see cref="M:ciloci.FormulaEngine.FormulaEngine.CreateFormula(System.String,ciloci.FormulaEngine.GrammarType)"/>
	''' except that the Excel grammar is used.</remarks>
	''' <exception cref="T:ciloci.FormulaEngine.InvalidFormulaException">The formula could not be created</exception>
	Public Function CreateFormula(ByVal expression As String) As Formula
		ValidateNonNull(expression, "expression")
		Dim f As New Formula(Me, expression, GrammarType.Excel)
		Return f
	End Function

	''' <summary>
	''' Creates a formula by parsing an expression using a specific grammar
	''' </summary>
	''' <param name="expression">The expression to parse</param>
	''' <param name="gt">The type of grammar to use when parsing the expression</param>
	''' <returns>A formula representing the parsed expression</returns>
	''' <remarks>This method is used to create instances of the formula class from an expression.  The returned instance is
	''' not part of the formula engine; you must explicitly add it.</remarks>
	''' <exception cref="T:ciloci.FormulaEngine.InvalidFormulaException">The formula could not be created</exception>
	Public Function CreateFormula(ByVal expression As String, ByVal gt As GrammarType) As Formula
		ValidateNonNull(expression, "expression")
		Dim f As New Formula(Me, expression, gt)
		Return f
	End Function

	''' <overloads>Evaluates an expression</overloads>
	''' <summary>
	''' Evaluates an expression using the Excel grammar
	''' </summary>
	''' <param name="expression">The expression to evaluate</param>
	''' <returns>The result of evaluating the expression</returns>
	''' <remarks>Works the same way as <see cref="M:ciloci.FormulaEngine.FormulaEngine.Evaluate(System.String,ciloci.FormulaEngine.GrammarType)"/> except that the Excel grammar
	''' is used.</remarks>
	Public Function Evaluate(ByVal expression As String) As Object
		ValidateNonNull(expression, "expression")
		Dim f As Formula = Me.CreateFormula(expression)
		Dim result As Object = f.Evaluate()
		Return result
	End Function

	''' <summary>
	''' Evaluates an expression after parsing it using a specific grammar
	''' </summary>
	''' <param name="expression">The expression to evaluate</param>
	''' <param name="gt">The type of grammar to use when parsing the expression</param>
	''' <returns>The result of evaluating the expression</returns>
	''' <remarks>You can use this function to evaluate an expression without creating a formula.  If you plan on evaluating the same
	''' expression many times, it is more efficient to create a formula from it first and then call its 
	''' <see cref="M:ciloci.FormulaEngine.Formula.Evaluate"/> method as many times as you need.</remarks>
	Public Function Evaluate(ByVal expression As String, ByVal gt As GrammarType) As Object
		ValidateNonNull(expression, "expression")
		Dim f As Formula = Me.CreateFormula(expression, gt)
		Dim result As Object = f.Evaluate()
		Return result
	End Function

	''' <overloads>Adds a formula to the engine</overloads>
	''' <summary>
	''' Adds a formula parsed from an expression into the engine
	''' </summary>
	''' <param name="expression">The expression to create the formula from</param>
	''' <param name="ref">The reference the formula will be bound to</param>
	''' <returns>The added formula</returns>
	''' <remarks>This method does the same thing as 
	''' <see cref="M:ciloci.FormulaEngine.FormulaEngine.AddFormula(ciloci.FormulaEngine.Formula,ciloci.FormulaEngine.IReference)"/>
	''' except that takes an expression instead of a formula.</remarks>
	Public Function AddFormula(ByVal expression As String, ByVal Ref As IReference) As Formula
		Dim target As Formula = Me.CreateFormula(expression)
		Me.AddFormula(target, Ref)
		Return target
	End Function

	''' <summary>
	''' Adds a formula to the formula engine
	''' </summary>
	''' <param name="target">The formula instance to add</param>
	''' <param name="ref">The reference the formula will be bound to</param>
	''' <remarks>Use this method when you want to add a formula to the engine.  The engine will bind the formula to the given reference 
	''' and analyze its dependencies.  Currently, formulas can only be bound to cell, named, and external references.</remarks>
	''' <exception cref="System.ArgumentException">
	''' <list type="bullet">
	''' <item>The given formula was not created by this engine</item>
	''' <item>The formula cannot be bound to the type of reference given</item>
	''' <item>There is already a formula bound to the given reference</item>
	''' </list>
	''' </exception>
	''' <exception cref="System.ArgumentNullException">target or ref is null</exception>
	Public Sub AddFormula(ByVal target As Formula, ByVal ref As IReference)
		ValidateNonNull(ref, "ref")
		Dim selfRef As Reference = MyReferencePool.GetReference(ref)
		Me.ValidateAddFormula(target, selfRef)

		Me.SetFormulaDependencyReferences(target)
		target.SetSelfReference(selfRef)

		MyDependencyManager.AddFormula(target)
		MyReferenceFormulaMap.Add(selfRef, target)

		Me.CheckFormulaCircularReference(target)
		Me.NotifySelfReference(target)
	End Sub

	Private Sub ValidateAddFormula(ByVal target As Formula, ByVal selfRef As Reference)
		If Not TypeOf (selfRef) Is IFormulaSelfReference Then
			Throw New ArgumentException("A formula cannot be bound to this type of reference")
		End If

		ValidateNonNull(target, "target")
		Me.ValidateFormulaOwner(target)

		If MyReferenceFormulaMap.Contains(selfRef) = True Then
			Throw New ArgumentException(String.Format("A formula already exists at reference {0}", selfRef.ToStringIReference()))
		End If
	End Sub

	Friend Shared Sub ValidateNonNull(ByVal obj As Object, ByVal name As String)
		If obj Is Nothing Then
			Throw New ArgumentNullException(name)
		End If
	End Sub

	Private Sub ValidateFormulaOwner(ByVal target As Formula)
		If Not target.Engine Is Me Then
			Throw New ArgumentException("The formula is not owned by this engine")
		End If
	End Sub

	''' <summary>
	''' Removes a formula from the engine
	''' </summary>
	''' <param name="target">The formula instance to remove</param>
	''' <remarks>This method removes a formula from the engine and removes the formula's references from its dependency graph.</remarks>
	''' <exception cref="System.ArgumentException">
	''' <list type="bullet">
	''' <item>The given formula was not created by this engine</item>
	''' <item>The formula is not contained in the engine</item>
	''' </list>
	''' </exception>
	''' <exception cref="System.ArgumentNullException">The given formula is null</exception>
	Public Sub RemoveFormula(ByVal target As Formula)
		Me.ValidateRemoveFormula(target)
		Dim selfRef As Reference = target.SelfReferenceInternal

		MyDependencyManager.RemoveFormula(target)
		Me.RemoveReferences(target.DependencyReferences)

		Debug.Assert(MyReferenceFormulaMap.Contains(target.SelfReferenceInternal), "formula self ref not in map")
		MyReferenceFormulaMap.Remove(selfRef)
		MyReferencePool.RemoveReference(selfRef)
		target.ClearSelfReference()
	End Sub

	Private Sub ValidateRemoveFormula(ByVal target As Formula)
		ValidateNonNull(target, "target")
		Me.ValidateFormulaOwner(target)

		If target.SelfReferenceInternal Is Nothing Then
			Throw New ArgumentException("The formula is contained in this formula engine")
		End If
	End Sub

	''' <summary>
	''' Copies and adjusts a formula on a sheet
	''' </summary>
	''' <param name="source">The formula to copy</param>
	''' <param name="destRef">The destination of the copied formula</param>
	''' <remarks>This method is used when wishing to implement copying of formulas similarly to Excel.  It makes a copy of the source formula, 
	''' offsets its references by the difference between destRef and source's current location, and adds the copy to the engine.
	''' The engine will only adjust references marked as relative in the copied formula.</remarks>
	''' <exception cref="System.ArgumentException">The source formula is not bound to a sheet reference</exception>
	Public Sub CopySheetFormula(ByVal source As Formula, ByVal destRef As ISheetReference)
		ValidateNonNull(source, "source")
		ValidateNonNull(destRef, "destRef")
		If Not TypeOf (source.SelfReference) Is SheetReference Then
			Throw New ArgumentException("Source formula must be on a sheet")
		End If
		Dim sourceRef As SheetReference = source.SelfReference
		Dim rowOffset As Integer = destRef.Row - sourceRef.Row
		Dim colOffset As Integer = destRef.Column - sourceRef.Column
		Dim clone As Formula = source.Clone()
		clone.OffsetReferencesForCopy(destRef.Sheet, rowOffset, colOffset)
		Me.AddFormula(clone, destRef)
	End Sub

	Private Sub CheckFormulaCircularReference(ByVal f As Formula)
		If MyDependencyManager.FormulaHasCircularReference(f) = True Then
			Me.OnFormulaCircularReference(f)
		End If
	End Sub

	''' <summary>
	''' Set the dependency references of a formula to their pooled equivalents
	''' </summary>
	Private Sub SetFormulaDependencyReferences(ByVal target As Formula)
		Dim depRefs As Reference() = target.GetDependencyReferences()

		For i As Integer = 0 To depRefs.Length - 1
			Dim ref As Reference = depRefs(i)
			ref = MyReferencePool.GetReference(ref)
			depRefs(i) = ref
		Next

		target.SetDependencyReferences(depRefs)
	End Sub

	Friend Sub RemoveReferences(ByVal refs As Reference())
		For Each ref As Reference In refs
			MyReferencePool.RemoveReference(ref)
		Next
	End Sub

	Private Sub OnFormulaCircularReference(ByVal target As Formula)
		Dim list As New ArrayList
		list.Add(target.SelfReferenceInternal)
		Me.OnCircularReferenceDetected(list)
	End Sub

	Friend Sub OnCircularReferenceDetected(ByVal roots As IList)
		Dim arr(roots.Count - 1) As IReference
		roots.CopyTo(arr, 0)
		RaiseEvent CircularReferenceDetected(Me, New CircularReferenceDetectedEventArgs(arr))
	End Sub

	''' <summary>
	''' Recalculates all formulas that depend on a reference
	''' </summary>
	''' <param name="root">The reference whose dependents will be recalculated</param>
	''' <remarks>This is the method that controls all recalculation in the formula engine.  Given a root reference, it will find
	''' all formulas that depend on it and recalculate them in natural order.  If no formulas depend on the given reference then the 
	''' method does nothing.</remarks>
	''' <example>This example demonstrates how to define some formulas and then have the engine recalculate them
	''' <code>
	''' Dim engine As New FormulaEngine
	''' ' Add a formula at B2 that depends on cell A1
	''' engine.AddFormula("a1+1", engine.ReferenceFactory.Parse("B2"))
	''' ' Add a formula at C3 that depends on the formula at cell B2
	''' engine.AddFormula("b2+1", engine.ReferenceFactory.Parse("c3"))
	''' ' Get a reference to cell A1
	''' Dim a1Ref As ISheetReference = engine.ReferenceFactory.Parse("A1")
	''' ' This will recalculate the formula at B2, then the formula at C3
	''' engine.Recalculate(a1Ref)
	''' </code>
	''' </example>
	Public Sub Recalculate(ByVal root As IReference)
		ValidateNonNull(root, "root")
		Me.DoRecalculate(MyDependencyManager.GetReferenceCalculationList(root))
	End Sub

	Friend Sub DoRecalculate(ByVal calculationList As Reference())
		Dim formulas As Formula() = Me.GetFormulasFromReferences(calculationList)
		Me.NotifyRecalculate(formulas)
	End Sub

	Friend Function GetFormulasFromReferences(ByVal references As Reference()) As Formula()
		Dim arr(references.Length - 1) As Formula

		For i As Integer = 0 To arr.Length - 1
			Dim ref As Reference = references(i)
			Debug.Assert(MyReferenceFormulaMap.Contains(ref), "reference not mapped to formula")
			arr(i) = MyReferenceFormulaMap.Item(ref)
		Next

		Return arr
	End Function

	Private Sub NotifyRecalculate(ByVal calculationList As Formula())
		For i As Integer = 0 To calculationList.Length - 1
			Dim f As Formula = calculationList(i)
			Me.NotifySelfReference(f)
		Next
	End Sub

	Private Sub NotifySelfReference(ByVal target As Formula)
		DirectCast(target.SelfReferenceInternal, IFormulaSelfReference).OnFormulaRecalculate(target)
	End Sub

	''' <summary>
	''' Determines if the engine has a formula bound to a particular reference
	''' </summary>
	''' <param name="ref">The reference to test</param>
	''' <returns>True is there is a formula bound to the reference; False if there is not</returns>
	''' <remarks>Use this formula to test if the engine has a formula bound to a given reference</remarks>
	Public Function HasFormulaAt(ByVal ref As IReference) As Boolean
		Return Not Me.GetFormulaAt(ref) Is Nothing
	End Function

	''' <summary>
	''' Gets the formula bound to a particular reference
	''' </summary>
	''' <param name="ref">The reference to find a formula for</param>
	''' <returns>The formula bound to the reference; null if no formula is bound to ref</returns>
	''' <remarks>Use this method when you have a reference and need to get the formula that is bound to it</remarks>
	Public Function GetFormulaAt(ByVal ref As IReference) As Formula
		ValidateNonNull(ref, "ref")
		Dim realRef As Reference = MyReferencePool.GetPooledReference(ref)
		If realRef Is Nothing Then
			Return Nothing
		Else
			Return MyReferenceFormulaMap.Item(realRef)
		End If
	End Function

	''' <summary>
	''' Removes a formula bound to a particular reference
	''' </summary>
	''' <param name="ref">The reference whose formula to remove</param>
	''' <remarks>This method will remove the formula bound to a particular reference</remarks>
	''' <exception cref="System.ArgumentException">No formula is bound to ref</exception>
	Public Sub RemoveFormulaAt(ByVal ref As IReference)
		Dim f As Formula = Me.GetFormulaAt(ref)

		If f Is Nothing Then
			Throw New ArgumentException("No formula exists at given reference")
		Else
			Me.RemoveFormula(f)
		End If
	End Sub

	''' <summary>
	''' Removes all sheet formulas in a given range
	''' </summary>
	''' <param name="range">The range to clear formulas in</param>
	''' <remarks>This method is used when you need to clear all sheet formulas in a given range.  All formulas bound to references
	''' on the same sheet as range and that intersect its area will be removed from the engine.</remarks>
	''' <example>This sample shows how to remove all formulas in the range B2:C4
	''' <code>
	''' Dim engine As New FormulaEngine
	''' ' Declare a reference to all cells from B2 to C4
	''' Dim range As ISheetReference = engine.ReferenceFactory.Parse("B2:C4")
	''' ' Remove all formulas in that range
	''' engine.RemoveFormulasInRange(range)
	''' </code>
	''' </example>
	Public Sub RemoveFormulasInRange(ByVal range As ISheetReference)
		ValidateNonNull(range, "range")
		Dim realRange As SheetReference = range
		Dim toRemove As IList = New ArrayList

		For Each selfRef As Reference In MyReferenceFormulaMap.Keys
			If realRange.Intersects(selfRef) = True Then
				toRemove.Add(MyReferenceFormulaMap.Item(selfRef))
			End If
		Next
        If toRemove.Count > 0 Then
            Me.RemoveFormulas(toRemove)
        End If
    End Sub

	''' <summary>
	''' Remove many formulas in a more efficient manner
	''' </summary>
	Private Sub RemoveFormulas(ByVal formulas As IList)
		' Tell the dependency manager to suspend range link calculation
		MyDependencyManager.RemoveRangeLinks()
		MyDependencyManager.SetSuspendRangeLinks(True)

		' Remove all the formulas
		For Each f As Formula In formulas
			Me.RemoveFormula(f)
		Next

		' Resume range link calculation
		MyDependencyManager.SetSuspendRangeLinks(False)
		MyDependencyManager.AddRangeLinks()
	End Sub

	''' <summary>
	''' Resets the formula engine to an empty state
	''' </summary>
	''' <remarks>Use this function when you need to reset the formula engine to an empty state.  This function will remove all: formulas, dependencies, 
	''' references, and sheets from the engine.  It will <b>not</b> clear the function library.  You should call that class'
	''' <see cref="M:ciloci.FormulaEngine.FunctionLibrary.Clear(System.Boolean)"/> method if you also want to remove all formulas.</remarks>
	Public Sub Clear()
		' Clear everything but functions
		MyDependencyManager.Clear()
		MyReferencePool.Clear()
		MyReferenceFormulaMap.Clear()
		MySheets.Clear()
	End Sub

	''' <summary>
	''' Notifies the engine that columns have been inserted on the active sheet
	''' </summary>
	''' <param name="insertAt">The index of the first inserted column</param>
	''' <param name="count">The number of columns inserted</param>
	''' <remarks>Use this method to notify the engine that columns have been inserted on the active sheet.  The engine will update
	''' all references as necessary.</remarks>
	''' <exception cref="System.ArgumentOutOfRangeException">
	''' <list type="bullet">
	''' <item>insertAt is less than 1</item>
	''' <item>count is negative</item>
	''' </list>
	''' </exception>
	Public Sub OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer)
		Me.ValidateHeaderOpArgs(insertAt, count)
		Me.DoActiveSheetReferenceOperation(New ColumnsInsertedOperator(insertAt, count))
	End Sub

	''' <summary>
	''' Notifies the engine that columns have been removed on the active sheet
	''' </summary>
	''' <param name="removeAt">The index of the first removed column</param>
	''' <param name="count">The number of removed columns</param>
	''' <remarks>Use this method to notify the engine that columns have been removed on the active sheet.  The engine will invalidate
	''' all references in the removed area, recalculate any formulas that depend on those references, and remove all
	''' formulas in the area.</remarks>
	''' <exception cref="System.ArgumentOutOfRangeException">
	''' <list type="bullet">
	''' <item>removeAt is less than 1</item>
	''' <item>count is negative</item>
	''' </list>
	''' </exception>
	Public Sub OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer)
		Me.ValidateHeaderOpArgs(removeAt, count)
		Me.DoActiveSheetReferenceOperation(New ColumnsRemovedOperator(removeAt, count))
	End Sub

	''' <summary>
	''' Notifies the engine that rows have been inserted on the active sheet
	''' </summary>
	''' <param name="insertAt">The index of the first inserted row</param>
	''' <param name="count">The number of rows inserted</param>
	''' <remarks>Use this method to notify the engine that rows have been inserted on the active sheet.  The engine will update
	''' all references as necessary.</remarks>
	''' <exception cref="System.ArgumentOutOfRangeException">
	''' <list type="bullet">
	''' <item>insertAt is less than 1</item>
	''' <item>count is negative</item>
	''' </list>
	''' </exception>
	Public Sub OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer)
		Me.ValidateHeaderOpArgs(insertAt, count)
		Me.DoActiveSheetReferenceOperation(New RowsInsertedOperator(insertAt, count))
	End Sub

	''' <summary>
	''' Notifies the engine that rows have been removed on the active sheet
	''' </summary>
	''' <param name="removeAt">The index of the first removed row</param>
	''' <param name="count">The number of removed rows</param>
	''' <remarks>Use this method to notify the engine that rows have been removed on the active sheet.  The engine will invalidate
	''' all references in the removed area, recalculate any formulas that depend on those references, and remove all
	''' formulas in the area.</remarks>
	''' <exception cref="System.ArgumentOutOfRangeException">
	''' <list type="bullet">
	''' <item>removeAt is less than 1</item>
	''' <item>count is negative</item>
	''' </list>
	''' </exception>
	Public Sub OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer)
		Me.ValidateHeaderOpArgs(removeAt, count)
		Me.DoActiveSheetReferenceOperation(New RowsRemovedOperator(removeAt, count))
	End Sub

	''' <summary>
	''' Notifies the engine that a range has moved
	''' </summary>
	''' <param name="range">The range that has moved</param>
	''' <param name="rowOffset">The number of rows range has moved.  Can be negative.</param>
	''' <param name="colOffset">The number of columns the range has moved.  Can be negative.</param>
	''' <remarks>Use this method to notify the engine that a range on a sheet has moved.  The engine will update all references
	''' in, or that depend on, the moved range accordingly.</remarks>
	''' <exception cref="System.ArgumentOutOfRangeException">The given range, when offset by the given offsets, is not within the bounds
	''' of the active sheet</exception>
	Public Sub OnRangeMoved(ByVal range As ISheetReference, ByVal rowOffset As Integer, ByVal colOffset As Integer)
		ValidateNonNull(range, "range")
		Dim sourceRef As SheetReference = range

		Dim destRect As System.Drawing.Rectangle = range.Area
		destRect.Offset(colOffset, rowOffset)
		Dim destRef As SheetReference = MyReferenceFactory.FromRectangle(destRect)

		Dim pred As ReferencePredicateBase
		If sourceRef.Sheet Is destRef.Sheet Then
			pred = New SheetReferencePredicate(MySheets.ActiveSheet)
		Else
			pred = New CrossSheetReferencePredicate(sourceRef.Sheet, destRef.Sheet)
		End If

		Me.DoReferenceOperation(New RangeMovedOperator(Me, sourceRef, destRef), pred)
	End Sub

	Friend Sub OnSheetRemoved(ByVal sheet As ISheet)
		Me.DoSheetReferenceOperation(New SheetRemovedOperator, sheet)
	End Sub

	Private Sub DoActiveSheetReferenceOperation(ByVal refOp As ReferenceOperator)
		Me.DoSheetReferenceOperation(refOp, MySheets.ActiveSheet)
	End Sub

	Private Sub DoSheetReferenceOperation(ByVal refOp As ReferenceOperator, ByVal sheet As ISheet)
		Me.DoReferenceOperation(refOp, New SheetReferencePredicate(sheet))
	End Sub

	Private Sub DoReferenceOperation(ByVal refOp As ReferenceOperator, ByVal predicate As ReferencePredicateBase)
		Dim targets As IList = MyReferencePool.FindReferences(predicate)
		MyReferencePool.DoReferenceOperation(targets, refOp)
	End Sub

	Friend Sub RecalculateAffectedReferences(ByVal affected As IList)
		Dim sources As IList = MyDependencyManager.GetSources(affected)

		Dim calcList As Reference() = MyDependencyManager.GetCalculationList(sources)
		Me.DoRecalculate(calcList)
	End Sub

	Friend Sub RemoveInvalidFormulas(ByVal invalidRefs As IList)
		For Each ref As Reference In invalidRefs
			Dim f As Formula = MyReferenceFormulaMap.Item(ref)
			If Not f Is Nothing Then
				Me.RemoveFormula(f)
			End If
		Next
	End Sub

	Private Sub ValidateHeaderOpArgs(ByVal location As Integer, ByVal count As Integer)
		If location < 1 Then
			Throw New ArgumentOutOfRangeException("location", "Value must be greater than 0")
		End If

		If count < 0 Then
			Throw New ArgumentOutOfRangeException("count", "Value cannot be negative")
		End If
	End Sub

	''' <summary>
	''' Creates an error wrapper around a specified error type
	''' </summary>
	''' <param name="errorType">The type of error to create a wrapper for</param>
	''' <returns>A wrapper around the error</returns>
	''' <remarks>This function lets you create an error wrapper around a specific error type</remarks>
	Public Shared Function CreateError(ByVal errorType As ErrorValueType) As ErrorValueWrapper
		Return New ErrorValueWrapper(errorType)
	End Function

	''' <summary>
	''' Gets all named references bound to formulas in the engine
	''' </summary>
	''' <returns>An array of named references</returns>
	''' <remarks>Use this function when you need to get all the named references bound to formulas in this engine.</remarks>
	Public Function GetNamedReferences() As INamedReference()
		Dim found As IList = New ArrayList

		For Each de As DictionaryEntry In MyReferenceFormulaMap
			If de.Key.GetType() Is GetType(NamedReference) Then
				found.Add(de.Key)
			End If
		Next

		Dim arr(found.Count - 1) As INamedReference
		found.CopyTo(arr, 0)
		Return arr
	End Function

	''' <summary>
	''' Defines a variable for use in formulas
	''' </summary>
	''' <param name="name">The name of the variable</param>
	''' <returns>A <see cref="T:ciloci.FormulaEngine.Variable"/> instance representing the newly defined variable</returns>
	''' <remarks>Use this method to define a new variable for use in formulas</remarks>
	Public Function DefineVariable(ByVal name As String) As Variable
		Return New Variable(Me, name)
	End Function

	''' <summary>
	''' Recreates the parsers used to parse formulas
	''' </summary>
	''' <remarks>The parsers used to parse formulas are cached to improve performance.  Calling this method will
	''' force the creation of a new instance of the parser.  You typically will not need to
	''' call this method unless you are switching cultures at run-time.</remarks>
	Public Sub RecreateParsers()
		ParserFactory.Clear()
	End Sub

	Friend Function GetReferenceProperties(ByVal target As Reference) As IList
		Dim found As IList = New ArrayList

		For Each de As DictionaryEntry In MyReferenceFormulaMap
			Dim f As Formula = de.Value
			f.GetReferenceProperties(target, found)
		Next

		Return found
	End Function

	''' <summary>
	''' Gets the function library the engine is using
	''' </summary>
	''' <value>An instance of FunctionLibrary</value>
	''' <remarks>This property lets you access the engine's function library</remarks>
	Public ReadOnly Property FunctionLibrary() As FunctionLibrary
		Get
			Return MyFunctionLibrary
		End Get
	End Property

	Friend ReadOnly Property DependencyManager() As DependencyManager
		Get
			Return MyDependencyManager
		End Get
	End Property

	''' <summary>
	''' Gets the number of formulas this engine contains
	''' </summary>
	''' <value>A count of all the formulas</value>
	''' <remarks>Use this property when you need to know how many formulas are contained in the engine</remarks>
	Public ReadOnly Property FormulaCount() As Integer
		Get
			Return MyReferenceFormulaMap.Count
		End Get
	End Property

	''' <summary>Gets the engine's ReferenceFactory instance</summary>
	''' <value>The reference factory of the engine</value>
	''' <remarks>This property lets you access this engine's reference factory</remarks>
	Public ReadOnly Property ReferenceFactory() As ReferenceFactory
		Get
			Return MyReferenceFactory
		End Get
	End Property

	Friend ReadOnly Property ReferencePool() As ReferencePool
		Get
			Return MyReferencePool
		End Get
	End Property

	''' <summary>Gets the engine's SheetCollection instance</summary>
	''' <value>The SheetCollection of engine</value>
	''' <remarks>Use this property when you need to access the engine's SheetCollection</remarks>
	Public ReadOnly Property Sheets() As SheetCollection
		Get
			Return MySheets
		End Get
	End Property

	''' <summary>
	''' Gets the engine's FormulaEngineInfo instance
	''' </summary>
	''' <value>An instance of the FormulaEngineInfo class</value>
	''' <remarks>This property lets you access the engine's FormulaEngineInfo instance</remarks>
	Public ReadOnly Property Info() As FormulaEngineInfo
		Get
			Return MyEngineInfo
		End Get
    End Property

    Public Property CurrentFormula() As Formula
        Get
            Return MyCurrentlyEvaluatingFormula
        End Get
        Friend Set(ByVal value As Formula)
            MyCurrentlyEvaluatingFormula = value
        End Set
    End Property
End Class