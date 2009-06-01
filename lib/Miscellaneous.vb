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

' Various classes that don't deserve an invididual source file

Imports System.Drawing
Imports System.Collections.Generic

Friend MustInherit Class ReferenceOperator

	Public Sub New()

	End Sub

	Public Overridable Sub PreOperate(ByVal references As IList)

	End Sub

	Public Overridable Sub PostOperate(ByVal references As IList)

	End Sub

	Public MustOverride Function Operate(ByVal ref As Reference) As ReferenceOperationResultType
End Class

Friend Class ColumnsInsertedOperator
	Inherits ReferenceOperator

	Private MyInsertAt As Integer
	Private MyCount As Integer

	Public Sub New(ByVal insertAt As Integer, ByVal count As Integer)
		MyInsertAt = insertAt
		MyCount = count
	End Sub

	Public Overrides Function Operate(ByVal ref As Reference) As ReferenceOperationResultType
		Return ref.GridOps.OnColumnsInserted(MyInsertAt, MyCount)
	End Function
End Class

Friend Class RowsInsertedOperator
	Inherits ReferenceOperator

	Private MyInsertAt As Integer
	Private MyCount As Integer

	Public Sub New(ByVal insertAt As Integer, ByVal count As Integer)
		MyInsertAt = insertAt
		MyCount = count
	End Sub

	Public Overrides Function Operate(ByVal ref As Reference) As ReferenceOperationResultType
		Return ref.GridOps.OnRowsInserted(MyInsertAt, MyCount)
	End Function
End Class

Friend Class ColumnsRemovedOperator
	Inherits ReferenceOperator

	Private MyRemoveAt As Integer
	Private MyCount As Integer

	Public Sub New(ByVal removeAt As Integer, ByVal count As Integer)
		MyRemoveAt = removeAt
		MyCount = count
	End Sub

	Public Overrides Function Operate(ByVal ref As Reference) As ReferenceOperationResultType
		Return ref.GridOps.OnColumnsRemoved(MyRemoveAt, MyCount)
	End Function
End Class

Friend Class RowsRemovedOperator
	Inherits ReferenceOperator

	Private MyRemoveAt As Integer
	Private MyCount As Integer

	Public Sub New(ByVal removeAt As Integer, ByVal count As Integer)
		MyRemoveAt = removeAt
		MyCount = count
	End Sub

	Public Overrides Function Operate(ByVal ref As Reference) As ReferenceOperationResultType
		Return ref.GridOps.OnRowsRemoved(MyRemoveAt, MyCount)
	End Function
End Class

Friend Class RangeMovedOperator
	Inherits ReferenceOperator

	Private MyOwner As FormulaEngine
	Private MySource As SheetReference
	Private MyDest As SheetReference

	Public Sub New(ByVal owner As FormulaEngine, ByVal source As SheetReference, ByVal dest As SheetReference)
		MyOwner = owner
		MySource = source
		MyDest = dest
	End Sub

	Public Overrides Function Operate(ByVal ref As Reference) As ReferenceOperationResultType
		Return ref.GridOps.OnRangeMoved(MySource, MyDest)
	End Function

	Public Overrides Sub PreOperate(ByVal references As System.Collections.IList)
		MyOwner.DependencyManager.RemoveRangeLinks()
	End Sub

	Public Overrides Sub PostOperate(ByVal references As System.Collections.IList)
		MyOwner.DependencyManager.AddRangeLinks()
		Dim circularRefs As IList = New ArrayList

		For Each ref As Reference In references
			If MyOwner.DependencyManager.IsCircularReference(ref) = True Then
				circularRefs.Add(ref)
			End If
		Next

		If circularRefs.Count > 0 Then
			MyOwner.OnCircularReferenceDetected(circularRefs)
		End If
	End Sub
End Class

Friend Class SheetRemovedOperator
	Inherits ReferenceOperator

	Public Overrides Function Operate(ByVal ref As Reference) As ReferenceOperationResultType
		Return ReferenceOperationResultType.Invalidated
	End Function
End Class

''' <summary>
''' Creates an operand dynamically based on a value
''' </summary>
Friend Class OperandFactory

	Public Shared Function CreateDynamic(ByVal value As Object) As IOperand
		If value Is Nothing Then
			Return New NullValueOperand
		End If

		Dim t As Type = value.GetType()
		Dim op As IOperand

		If t Is GetType(Double) Then
			op = New DoubleOperand(DirectCast(value, Double))
		ElseIf t Is GetType(String) Then
			op = New StringOperand(DirectCast(value, String))
		ElseIf t Is GetType(Boolean) Then
			op = New BooleanOperand(DirectCast(value, Boolean))
		ElseIf t Is GetType(Integer) Then
			op = New IntegerOperand(DirectCast(value, Integer))
		ElseIf TypeOf (value) Is IReference Then
			op = value
		ElseIf t Is GetType(ErrorValueWrapper) Then
			op = New ErrorValueOperand(DirectCast(value, ErrorValueWrapper))
		ElseIf t Is GetType(DateTime) Then
			op = New DateTimeOperand(DirectCast(value, DateTime))
		Else
			Throw New ArgumentException(String.Format("The type {0} is not supported as an operand", t.Name))
		End If

		Return op
	End Function
End Class

Friend Class SheetReferencePredicate
	Inherits ReferencePredicateBase

	Private MyTarget As ISheet

	Public Sub New(ByVal target As ISheet)
		MyTarget = target
	End Sub

	Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
		Return ref.IsOnSheet(MyTarget)
	End Function
End Class

Friend Class EqualsReferencePredicate
	Inherits ReferencePredicateBase

	Private MyTarget As Reference

	Public Sub New(ByVal target As Reference)
		MyTarget = target
	End Sub

	Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
		Return MyTarget.EqualsReference(ref)
	End Function
End Class

Friend Class CrossSheetReferencePredicate
	Inherits ReferencePredicateBase

	Private MySource As ISheet
	Private MyDestination As ISheet

	Public Sub New(ByVal source As ISheet, ByVal dest As ISheet)
		MySource = source
		MyDestination = dest
	End Sub

	Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
		Return ref.IsOnSheet(MySource) Or ref.IsOnSheet(MyDestination)
	End Function
End Class

Friend Class AlwaysPredicate
	Inherits ReferencePredicateBase

	Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
		Return True
	End Function
End Class

''' <summary>
''' Contains useful information about the formula engine
''' </summary>
''' <remarks>This class allows you to get various information about the formula engine.  Mostly used for development/testing and when
''' you would like to show some statistics to the user.</remarks>
Public Class FormulaEngineInfo

	Private MyOwner As FormulaEngine

	Friend Sub New(ByVal owner As FormulaEngine)
		MyOwner = owner
	End Sub

	''' <summary>
	''' Gets the calculation list for a reference
	''' </summary>
	''' <param name="root">The reference from where to calculate the list</param>
	''' <returns>An array of formulas that would need to be recalculated</returns>
	''' <remarks>Given a reference, this method returns a list of formulas that would need to be recalculated when that reference changes.
	''' The formulas in the list will be in natural order.</remarks>
	Public Function GetCalculationList(ByVal root As IReference) As Formula()
		FormulaEngine.ValidateNonNull(root, "root")
		Dim refs As Reference() = MyOwner.DependencyManager.GetReferenceCalculationList(root)
		Return MyOwner.GetFormulasFromReferences(refs)
	End Function

	''' <summary>
	''' Determines whether a reference is valid
	''' </summary>
	''' <param name="ref">The reference to check</param>
	''' <returns>True if the reference is valid; False if it is not</returns>
	''' <remarks>This function determines whether a given reference is valid.  For example: A reference to column C will become
	''' invalid when column C is removed.</remarks>
	Public Function IsReferenceValid(ByVal ref As IReference) As Boolean
		Dim realRef As Reference = MyOwner.ReferencePool.GetPooledReference(ref)
		If realRef Is Nothing Then
			Return False
		Else
			Return realRef.Valid
		End If
	End Function

	''' <summary>
	''' Gets the number of direct precedents of a reference
	''' </summary>
	''' <param name="ref">The reference whose direct precedents to get</param>
	''' <returns>A count of the number of direct precedents of the reference</returns>
	''' <remarks>The count indicates the number of references that have ref as their dependant</remarks>
	Public Function GetDirectPrecedentsCount(ByVal ref As IReference) As Integer
		FormulaEngine.ValidateNonNull(ref, "ref")
		Dim realRef As Reference = MyOwner.ReferencePool.GetPooledReference(ref)
		If realRef Is Nothing Then
			Return 0
		Else
			Return MyOwner.DependencyManager.GetDirectPrecedentsCount(realRef)
		End If
	End Function

	''' <summary>
	''' Gets the total number of references tracked by the engine
	''' </summary>
	''' <value>A count indicating the total number of references</value>
	''' <remarks>This property lets you get a count of the total number of references that the engine is tracking</remarks>
	Public ReadOnly Property ReferenceCount() As Integer
		Get
			Return MyOwner.ReferencePool.ReferenceCount
		End Get
	End Property

	''' <summary>
	''' Gets the number of dependents in the engine's dependency graph
	''' </summary>
	''' <value>A count of the number of dependents</value>
	''' <remarks>The count indicates the number of references that have other references which depend on them</remarks>
	Public ReadOnly Property DependentsCount() As Integer
		Get
			Return MyOwner.DependencyManager.DependentsCount
		End Get
	End Property

	''' <summary>
	''' Gets the number of precedents in the engine's dependency graph
	''' </summary>
	''' <value>A count of the number of precedents</value>
	''' <remarks>The count indicates the number of references that are dependents of other references</remarks>
	Public ReadOnly Property PrecedentsCount() As Integer
		Get
			Return MyOwner.DependencyManager.PrecedentsCount
		End Get
	End Property

	''' <summary>
	''' Gets a string representation of the engine's dependency graph
	''' </summary>
	''' <value>A string representing all dependencies</value>
	''' <remarks>This property will return a string representation of the engine's dependents graph.  There will be one
	''' line for each dependency and each line will be of the form "Ref1 -> Ref2, Ref3" and reads that a change in Ref1 will 
	''' change Ref2 and Ref3.</remarks>
	Public ReadOnly Property DependencyDump() As String
		Get
			Return MyOwner.DependencyManager.DependencyDump
		End Get
	End Property
End Class

''' <summary>
''' Encapsulates all grid related operations on a reference.  I made this a separate class so that non-grid references don't have
''' to include all these methods as stubs.
''' </summary>
Friend MustInherit Class GridOperationsBase

	Protected Delegate Sub SetRangeCallback(ByVal newRect As Rectangle)

	Protected Sub New()

	End Sub

	Public MustOverride Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
	Public MustOverride Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
	Public MustOverride Function OnRangeMoved(ByVal source As SheetReference, ByVal dest As SheetReference) As ReferenceOperationResultType
	Public MustOverride Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
	Public MustOverride Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType

	Protected Shared Function Affected2Enum(ByVal affected As Boolean) As ReferenceOperationResultType
		If affected = True Then
			Return ReferenceOperationResultType.Affected
		Else
			Return ReferenceOperationResultType.NotAffected
		End If
	End Function

	''' <summary>
	''' Handles a range move and tries to do it the same way as Excel.  This function is way too complicated but Excel has some very
	''' weird rules with regards to range moves and this is the only way I can think of emulating them.
	''' </summary>
	Protected Function HandleRangeMoved(ByVal current As SheetReference, ByVal source As SheetReference, ByVal dest As SheetReference, ByVal callback As SetRangeCallback) As ReferenceOperationResultType
		Dim destRect As Rectangle = dest.Range
		Dim sourceRect As Rectangle = source.Range
		Dim rowOffset, colOffset As Integer
		rowOffset = destRect.Top - sourceRect.Top
		colOffset = destRect.Left - sourceRect.Left

		Dim myRect As Rectangle = current.Range
		Dim isOnSourceSheet, isOnDestSheet As Boolean

		isOnSourceSheet = source.Sheet Is current.Sheet
		isOnDestSheet = dest.Sheet Is current.Sheet

		Dim sameSheet As Boolean = isOnSourceSheet And isOnDestSheet

		If isOnSourceSheet = True And isOnDestSheet = False AndAlso Me.IsEdgeMove(myRect, sourceRect) = True Then
			' Move of one of our edges to another sheet
			callback(Me.SubtractRectangle(myRect, sourceRect))
			Return ReferenceOperationResultType.Affected
		ElseIf isOnSourceSheet = False And isOnDestSheet = True AndAlso Me.IsEdgeDestroyMove(myRect, sourceRect, destRect) = True Then
			callback(Me.SubtractRectangle(myRect, destRect))
			Return ReferenceOperationResultType.Affected
		ElseIf sameSheet = True AndAlso Me.IsEdgeExpandMove(myRect, sourceRect, rowOffset, colOffset) = True Then
			callback(Me.GetEdgeExpandRectangle(myRect, sourceRect, destRect))
			Return ReferenceOperationResultType.Affected
		ElseIf sameSheet = True AndAlso Me.IsEdgeShrinkMove(myRect, sourceRect, rowOffset, colOffset) = True Then
			callback(Me.GetEdgeShrinkRectangle(myRect, sourceRect, rowOffset, colOffset))
			Return ReferenceOperationResultType.Affected
		ElseIf sameSheet = True AndAlso Me.IsEdgeDestroyMoveSameSheet(myRect, sourceRect, destRect) = True Then
			callback(Me.SubtractRectangle(myRect, destRect))
			Return ReferenceOperationResultType.Affected
		ElseIf isOnSourceSheet = True AndAlso sourceRect.Contains(myRect) = True Then
			' Move of the whole range
			myRect.Offset(colOffset, rowOffset)
			current.SetSheetForRangeMove(dest.Sheet)
			callback(myRect)
			Return ReferenceOperationResultType.Affected
		ElseIf isOnDestSheet = True AndAlso destRect.Contains(myRect) = True Then
			' We are overwritten by the move
			Return ReferenceOperationResultType.Invalidated
		Else
			' We are affected only if the moved range intersects us
			If current.Intersects(dest) = True Then
				Return ReferenceOperationResultType.Affected
			Else
				Return ReferenceOperationResultType.NotAffected
			End If
		End If
	End Function

	' Determines if the move is of the type that pulls an edge outwards and expands the range
	' Ex: range c4:e6; move c4:c6 -1col -> b4:e6
	Private Function IsEdgeExpandMove(ByVal currentRect As Rectangle, ByVal sourceRect As Rectangle, ByVal rowOffset As Integer, ByVal colOffset As Integer) As Boolean
		Dim destRect As Rectangle = sourceRect
		destRect.Offset(colOffset, rowOffset)

		Dim leftEdge As Rectangle = Me.GetLeftEdge(currentRect)
		Dim rightEdge As Rectangle = Me.GetRightEdge(currentRect)
		Dim topEdge As Rectangle = Me.GetTopEdge(currentRect)
		Dim botEdge As Rectangle = Me.GetBottomEdge(currentRect)

		Dim isLeftRightEdgeMove As Boolean = (rowOffset = 0) And (sourceRect.Contains(leftEdge) Or sourceRect.Contains(rightEdge))
		Dim isTopBotEdgeMove As Boolean = (colOffset = 0) And (sourceRect.Contains(topEdge) Or sourceRect.Contains(botEdge))

		Dim isGoodRect As Boolean = currentRect.IntersectsWith(sourceRect) = True And sourceRect.Contains(currentRect) = False And currentRect.IntersectsWith(destRect) = False

		Return (isLeftRightEdgeMove Or isTopBotEdgeMove) And isGoodRect
	End Function

	' Determines if the move is of the type that pulls an edge inwards and shrinks the range
	' Ex: range c4:e6; move c4:c6 +1col -> d4:e6
	Private Function IsEdgeShrinkMove(ByVal currentRect As Rectangle, ByVal sourceRect As Rectangle, ByVal rowOffset As Integer, ByVal colOffset As Integer) As Boolean
		Dim destRect As Rectangle = sourceRect
		destRect.Offset(colOffset, rowOffset)

		Dim leftEdge As Rectangle = Me.GetLeftEdge(currentRect)
		Dim rightEdge As Rectangle = Me.GetRightEdge(currentRect)
		Dim topEdge As Rectangle = Me.GetTopEdge(currentRect)
		Dim botEdge As Rectangle = Me.GetBottomEdge(currentRect)

		Dim isLeftRightEdgeMove As Boolean = (rowOffset = 0) And (sourceRect.Contains(leftEdge) Or sourceRect.Contains(rightEdge))
		Dim isTopBotEdgeMove As Boolean = (colOffset = 0) And (sourceRect.Contains(topEdge) Or sourceRect.Contains(botEdge))

		Dim isGoodRect As Boolean = currentRect.IntersectsWith(destRect) = True And sourceRect.Contains(currentRect) = False

		Return (isLeftRightEdgeMove Or isTopBotEdgeMove) And isGoodRect
	End Function

	' Determines if the move is of the type that has a range moved on top of the edge of our range thus
	' chopping off that edge
	' Ex: range c4:e6; move g4:h6 -2col -> c4:d6
	Private Function IsEdgeDestroyMove(ByVal currentRect As Rectangle, ByVal sourceRect As Rectangle, ByVal destRect As Rectangle) As Boolean
		Dim leftEdge As Rectangle = Me.GetLeftEdge(currentRect)
		Dim rightEdge As Rectangle = Me.GetRightEdge(currentRect)
		Dim topEdge As Rectangle = Me.GetTopEdge(currentRect)
		Dim botEdge As Rectangle = Me.GetBottomEdge(currentRect)

		Dim leftEdgeGood As Boolean = destRect.Contains(leftEdge) And destRect.Left < leftEdge.Left
		Dim rightEdgeGood As Boolean = destRect.Contains(rightEdge) And destRect.Right > rightEdge.Right
		Dim topEdgeGood As Boolean = destRect.Contains(topEdge) And destRect.Top < topEdge.Top
		Dim botEdgeGood As Boolean = destRect.Contains(botEdge) And destRect.Bottom > botEdge.Bottom

		Dim xGood As Boolean = (leftEdgeGood Or rightEdgeGood) And sourceRect.Width > 1
		Dim yGood As Boolean = (topEdgeGood Or botEdgeGood) And sourceRect.Height > 1

		Dim edgeGood As Boolean = xGood Or yGood
		Dim rectGood As Boolean = destRect.Contains(currentRect) = False

		Return edgeGood And rectGood
	End Function

	Private Function IsEdgeDestroyMoveSameSheet(ByVal currentRect As Rectangle, ByVal sourceRect As Rectangle, ByVal destRect As Rectangle) As Boolean
		Dim isMoveGood As Boolean = Me.IsEdgeDestroyMove(currentRect, sourceRect, destRect)
		Dim isRectGood As Boolean = currentRect.Contains(destRect) = False And sourceRect.IntersectsWith(currentRect) = False
		Return isRectGood And isMoveGood
	End Function

	Private Function IsCrossSheetEdgeDestroyMove(ByVal currentRect As Rectangle, ByVal sourceRect As Rectangle, ByVal rowOffset As Integer, ByVal colOffset As Integer) As Boolean

	End Function

	' Gets the new rectangle after one of our edges is pulled away
	Private Function GetEdgeExpandRectangle(ByVal currentRect As Rectangle, ByVal sourceRect As Rectangle, ByVal destRect As Rectangle) As Rectangle
		Dim leftEdge As Rectangle = Me.GetLeftEdge(currentRect)
		Dim rightEdge As Rectangle = Me.GetRightEdge(currentRect)
		Dim topEdge As Rectangle = Me.GetTopEdge(currentRect)
		Dim botEdge As Rectangle = Me.GetBottomEdge(currentRect)
		Dim newRect As Rectangle = Rectangle.Empty

		If sourceRect.Contains(leftEdge) = True Then
			newRect = Rectangle.FromLTRB(destRect.Left, currentRect.Top, currentRect.Right, currentRect.Bottom)
		ElseIf sourceRect.Contains(rightEdge) Then
			newRect = Rectangle.FromLTRB(currentRect.Left, currentRect.Top, destRect.Right, currentRect.Bottom)
		ElseIf sourceRect.Contains(topEdge) Then
			newRect = Rectangle.FromLTRB(currentRect.Left, destRect.Top, destRect.Right, currentRect.Bottom)
		ElseIf sourceRect.Contains(botEdge) Then
			newRect = Rectangle.FromLTRB(currentRect.Left, currentRect.Top, destRect.Right, destRect.Bottom)
		End If

		Debug.Assert(newRect.IsEmpty = False)
		Return newRect
	End Function

	' Gets the new rectangle after one of our edges is pulled in
	Private Function GetEdgeShrinkRectangle(ByVal currentRect As Rectangle, ByVal sourceRect As Rectangle, ByVal rowOffset As Integer, ByVal colOffset As Integer) As Rectangle
		Dim newRect As Rectangle
		Dim destRect As Rectangle

		sourceRect.Intersect(currentRect)

		destRect = sourceRect
		destRect.Offset(colOffset, rowOffset)
		newRect = Me.SubtractRectangle(currentRect, sourceRect)
		newRect = Rectangle.Union(destRect, newRect)

		Return newRect
	End Function

	Private Function IsEdgeMove(ByVal currentRect As Rectangle, ByVal movedRect As Rectangle) As Boolean
		Dim leftEdge As Rectangle = Me.GetLeftEdge(currentRect)
		Dim rightEdge As Rectangle = Me.GetRightEdge(currentRect)
		Dim topEdge As Rectangle = Me.GetTopEdge(currentRect)
		Dim botEdge As Rectangle = Me.GetBottomEdge(currentRect)

		Dim touchesEdge As Boolean = False
		touchesEdge = touchesEdge Or movedRect.Contains(leftEdge)
		touchesEdge = touchesEdge Or movedRect.Contains(rightEdge)
		touchesEdge = touchesEdge Or movedRect.Contains(topEdge)
		touchesEdge = touchesEdge Or movedRect.Contains(botEdge)

		Dim movedDoesNotContainCurrent As Boolean = movedRect.Contains(currentRect) = False

		Return touchesEdge And movedDoesNotContainCurrent
	End Function

	Private Function SubtractRectangle(ByVal target As Rectangle, ByVal toSubtract As Rectangle) As Rectangle
		Dim leftEdge As Rectangle = Me.GetLeftEdge(target)
		Dim rightEdge As Rectangle = Me.GetRightEdge(target)
		Dim topEdge As Rectangle = Me.GetTopEdge(target)
		Dim botEdge As Rectangle = Me.GetBottomEdge(target)
		Dim newRect As Rectangle = Rectangle.Empty

		If toSubtract.Contains(leftEdge) = True Then
			newRect = Rectangle.FromLTRB(toSubtract.Right, target.Top, target.Right, target.Bottom)
		ElseIf toSubtract.Contains(rightEdge) = True Then
			newRect = Rectangle.FromLTRB(target.Left, target.Top, toSubtract.Left, target.Bottom)
		ElseIf toSubtract.Contains(topEdge) = True Then
			newRect = Rectangle.FromLTRB(target.Left, toSubtract.Bottom, target.Right, target.Bottom)
		ElseIf toSubtract.Contains(botEdge) = True Then
			newRect = Rectangle.FromLTRB(target.Left, target.Top, target.Right, toSubtract.Top)
		End If

		Return newRect
	End Function

	Private Function GetTopEdge(ByVal rect As Rectangle) As Rectangle
		Return New Rectangle(rect.Left, rect.Top, rect.Width, 1)
	End Function

	Private Function GetBottomEdge(ByVal rect As Rectangle) As Rectangle
		Return New Rectangle(rect.Left, rect.Bottom - 1, rect.Width, 1)
	End Function

	Private Function GetLeftEdge(ByVal rect As Rectangle) As Rectangle
		Return New Rectangle(rect.Left, rect.Top, 1, rect.Height)
	End Function

	Private Function GetRightEdge(ByVal rect As Rectangle) As Rectangle
		Return New Rectangle(rect.Right - 1, rect.Top, 1, rect.Height)
	End Function

	Protected Shared Function HandleRangeInserted(ByVal start As Integer, ByVal finish As Integer, ByVal insertAt As Integer, ByVal count As Integer) As Point
		start = HandleInsert(start, insertAt, count)
		finish = HandleInsert(finish, insertAt, count)
		Return New Point(start, finish)
	End Function

	Private Shared Function HandleInsert(ByVal value As Integer, ByVal insertAt As Integer, ByVal count As Integer) As Integer
		If value >= insertAt Then
			Return value + count
		Else
			Return value
		End If
	End Function

	Private Shared Function HandleRemove(ByVal value As Integer, ByVal removeAt As Integer, ByVal count As Integer) As Integer
		If value > removeAt + count - 1 Then
			' We are below the removed hole
			Return value - count
		Else
			Return value
		End If
	End Function

	Protected Shared Function HandleRangeRemoved(ByVal start As Integer, ByVal finish As Integer, ByVal removeAt As Integer, ByVal count As Integer) As Point
		Dim containsStart, containsFinish As Boolean
		containsStart = start >= removeAt And start < removeAt + count
		containsFinish = finish >= removeAt And finish < removeAt + count

		If containsStart And containsFinish Then
			' The entire range we cover is removed so we are invalid
			Return Point.Empty
		ElseIf containsStart = True Then
			start = removeAt
			finish = HandleRemove(finish, removeAt, count)
		ElseIf containsFinish = True Then
			start = HandleRemove(start, removeAt, count)
			finish = removeAt - 1
		Else
			start = HandleRemove(start, removeAt, count)
			finish = HandleRemove(finish, removeAt, count)
		End If

		Dim p As New Point(start, finish)
		Return p
	End Function
End Class

''' <summary>
''' Returned by references that don't interact with the grid
''' </summary>
Friend Class NullGridOps
	Inherits GridOperationsBase

	Public Overrides Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType

	End Function

	Public Overrides Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType

	End Function

	Public Overrides Function OnRangeMoved(ByVal source As SheetReference, ByVal dest As SheetReference) As ReferenceOperationResultType

	End Function

	Public Overrides Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType

	End Function

	Public Overrides Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType

	End Function
End Class

''' <summary>
''' Provides various utility functions
''' </summary>
''' <remarks>This class provides various methods for doing common tasks when dealing with formulas</remarks>
Public NotInheritable Class Utility

	Private Shared OurNumericTypes As Type() = CreateNumericTypes()

	Private Sub New()

	End Sub

	Private Shared Function CreateNumericTypes() As Type()
		Return New Type() {GetType(Double), GetType(Integer)}
	End Function

	''' <summary>
	''' Tries to convert a string into a value similarly to Excel
	''' </summary>
	''' <param name="text">The string to parse</param>
	''' <returns>A value from the parsed string</returns>
	''' <remarks>This method will try to parse text into a value.  It will try to convert the text into a Boolean, ErrorValueWrapper, 
	''' DateTime, Integer, Double, or if all of the previous conversions fail, a string.</remarks>
	Public Shared Function Parse(ByVal text As String) As Object
		FormulaEngine.ValidateNonNull(text, "text")
		Dim b As Boolean
		If Boolean.TryParse(text, b) = True Then
			Return b
		End If

		Dim wrapper As ErrorValueWrapper = ErrorValueWrapper.TryParse(text)

		If Not wrapper Is Nothing Then
			Return wrapper
		End If

		Dim dt As DateTime

		If DateTime.TryParseExact(text, New String() {"D", "d", "G", "g", "t", "T"}, Nothing, Globalization.DateTimeStyles.AllowWhiteSpaces, dt) = True Then
			Return dt
		End If

		Dim d As Double
		Dim success As Boolean
		success = Double.TryParse(text, Globalization.NumberStyles.Integer, Nothing, d)

		If success = True And d >= Integer.MinValue And d <= Integer.MaxValue Then
			Return CInt(d)
		End If

		success = Double.TryParse(text, Globalization.NumberStyles.Float, Nothing, d)

		If success = True Then
			Return d
		End If

		Return text
	End Function

	''' <summary>
	''' Determines whether a type is a numeric type
	''' </summary>
	''' <param name="t">The type to test</param>
	''' <returns>True if the type is numeric; False otherwise</returns>
	''' <remarks>Useful when you are processing sheet values and need to know that a type is numeric</remarks>
	Public Shared Function IsNumericType(ByVal t As Type) As Boolean
		Return System.Array.IndexOf(OurNumericTypes, t) <> -1
	End Function

	''' <summary>
	''' Determines whether a value is of a numeric type
	''' </summary>
	''' <param name="value">The value to test</param>
	''' <returns>True if the value is numeric; False otherwise</returns>
	''' <remarks>This does the same thing as <see cref="M:ciloci.FormulaEngine.Utility.IsNumericType(System.Type)"/> except that it acts
	''' on a value.</remarks>
	Public Shared Function IsNumericValue(ByVal value As Object) As Boolean
		If value Is Nothing Then
			Return False
		Else
			Return IsNumericType(value.GetType())
		End If
	End Function

	''' <summary>
	''' Converts a numeric value to a double
	''' </summary>
	''' <param name="value">The value to convert</param>
	''' <returns>The value converted to a double</returns>
	''' <remarks>Since there are many types of numeric values in .NET, there exists a need to have a common denominator format that
	''' they all can be converted to.  The type chosen here is the Double.</remarks>
	''' <exception cref="System.ArgumentException">The value is not of a numeric type</exception>
	Public Shared Function NormalizeNumericValue(ByVal value As Object) As Double
		If IsNumericValue(value) = False Then
			Throw New ArgumentException("Value is not numeric")
		End If
		Return DirectCast(value, IConvertible).ToDouble(Nothing)
	End Function

	''' <summary>
	''' Normalizes a value if it is of a numeric type
	''' </summary>
	''' <param name="value">The value to try to normalize</param>
	''' <returns>The value normalized value</returns>
	''' <remarks>This function normalizes value if it is of a numeric type; otherwise it returns the value unchanged</remarks>
	Public Shared Function NormalizeIfNumericValue(ByVal value As Object) As Object
		If value Is Nothing Then
			Return Nothing
		ElseIf IsNumericType(value.GetType()) = True Then
			Return DirectCast(value, IConvertible).ToDouble(Nothing)
		Else
			Return value
		End If
	End Function

	''' <summary>
	''' Gets the regular expression equivalent pattern of an excel wildcard expression
	''' </summary>
	''' <param name="pattern">The pattern to convert</param>
	''' <returns>A regular expression representation of pattern</returns>
	''' <remarks>Excel has its own syntax for pattern matching that many functions use.  This method converts such an expression
	''' into its regular expression equivalent.</remarks>
	Public Shared Function Wildcard2Regex(ByVal pattern As String) As String
		FormulaEngine.ValidateNonNull(pattern, "pattern")
		pattern = EscapeRegex(pattern)
		Dim sb As New System.Text.StringBuilder(pattern.Length)
		Dim ignoreChar As Boolean = False

		For i As Integer = 0 To pattern.Length - 1
			Dim c As Char = pattern.Chars(i)
			If ignoreChar = True Then
				ignoreChar = False
			ElseIf c = "~" Then
				' Escape char
				If i = pattern.Length - 1 Then
					' If the escape char is the last char then just match it
					sb.Append("~"c)
				Else
					Dim nextChar As Char = pattern.Chars(i + 1)
					If nextChar = "?" Then
						sb.Append("\?")
					ElseIf nextChar = "*" Then
						sb.Append("\*")
					Else
						sb.Append(nextChar)
					End If

					ignoreChar = True
				End If
			ElseIf c = "?" Then
				sb.Append(".")
			ElseIf c = "*" Then
				sb.Append(".*")
			Else
				sb.Append(c)
			End If
		Next

		Return sb.ToString()
	End Function

	Private Shared Function EscapeRegex(ByVal pattern As String) As String
		Dim sb As New System.Text.StringBuilder(pattern)
		sb.Replace("\", "\\")
		sb.Replace("[", "\[")
		sb.Replace("^", "\^")
		sb.Replace("$", "\$")
		sb.Replace(".", "\.")
		sb.Replace("|", "\|")
		sb.Replace("+", "\+")
		sb.Replace("(", "\(")
		sb.Replace(")", "\)")
		Return sb.ToString()
	End Function

	''' <summary>
	''' Determines whether a rectangle is inside the bounds of a given sheet
	''' </summary>
	''' <param name="rect">The rectangle to test</param>
	''' <param name="sheet">The sheet to use</param>
	''' <returns>True if the sheet contains the rectangle; False otherwise</returns>
	''' <remarks>Use this function when you have a rectangle and a sheet and need to know if the rectangle is inside the sheet's bounds.</remarks>
	Public Shared Function IsRectangleInSheet(ByVal rect As Rectangle, ByVal sheet As ISheet) As Boolean
		FormulaEngine.ValidateNonNull(sheet, "sheet")
		Return SheetReference.IsRectangleInSheet(rect, sheet)
	End Function

	''' <summary>
	''' Gets a row from a table of values
	''' </summary>
	''' <param name="table">The table to get the row from</param>
	''' <param name="rowIndex">The index of the row to get</param>
	''' <returns>An array containing the values from the requested row</returns>
	''' <remarks>This method is used when you have a table of values (like the ones returned by <see cref="M:ciloci.FormulaEngine.ISheetReference.GetValuesTable"/>) 
	''' and you need to get the values from a row.</remarks>
	Public Shared Function GetTableRow(ByVal table As Object(,), ByVal rowIndex As Integer) As Object()
		FormulaEngine.ValidateNonNull(table, "table")
		Dim arr(table.GetLength(1) - 1) As Object

		For col As Integer = 0 To arr.Length - 1
			arr(col) = table(rowIndex, col)
		Next

		Return arr
	End Function

	''' <summary>
	''' Gets a column from a table of values
	''' </summary>
	''' <param name="table">The table to get the column from</param>
	''' <param name="columnIndex">The index of the column to get</param>
	''' <returns>An array containing the values from the requested column</returns>
	''' <remarks>This method is used when you have a table of values (like the ones returned by <see cref="M:ciloci.FormulaEngine.ISheetReference.GetValuesTable"/>) 
	''' and you need to get the values from a column.</remarks>
	Public Shared Function GetTableColumn(ByVal table As Object(,), ByVal columnIndex As Integer) As Object()
		FormulaEngine.ValidateNonNull(table, "table")
		Dim arr(table.GetLength(0) - 1) As Object

		For row As Integer = 0 To arr.Length - 1
			arr(row) = table(row, columnIndex)
		Next

		Return arr
	End Function

	''' <summary>
	''' Gets the label for a column index
	''' </summary>
	''' <param name="columnIndex">The index whose label you wish to get</param>
	''' <returns>A string with the colum label</returns>
	''' <remarks>This function is handy when you have a column index and you want to get its associated label.</remarks>
	''' <example>
	''' <list type="table">
	''' <listheader><term>Column index</term><description>Resultant label</description></listheader>
	''' <item><term>1</term><description>"A"</description></item>
	''' <item><term>14</term><description>"N"</description></item>
	''' <item><term>123</term><description>"DS"</description></item>
	''' <item><term>256</term><description>"IV"</description></item>
	''' </list>
	''' </example>
	Public Shared Function ColumnIndex2Label(ByVal columnIndex As Integer) As String
		Return SheetReference.ColumnIndex2Label(columnIndex)
	End Function

	''' <summary>
	''' Gets the column index from a column label
	''' </summary>
	''' <param name="label">The label whose column index you wish to get</param>
	''' <returns>An index representing the label</returns>
	''' <remarks>This function is handy when you have a column label and you want to get its associated index.</remarks>
	''' <example>
	''' <list type="table">
	''' <listheader><term>Column label</term><description>Resultant index</description></listheader>
	''' <item><term>"A"</term><description>1</description></item>
	''' <item><term>"N"</term><description>14</description></item>
	''' <item><term>"DS"</term><description>123</description></item>
	''' <item><term>"IV"</term><description>256</description></item>
	''' </list>
	''' </example>
	Public Shared Function ColumnLabel2Index(ByVal label As String) As Integer
		FormulaEngine.ValidateNonNull(label, "label")
		If label.Length < 1 Or label.Length > 2 Then
			Throw New ArgumentException("The given label must be one or two characters long")
		End If
		Dim c2 As Char
		If label.Length = 2 Then
			c2 = label.Chars(1)
		End If
		Return SheetReference.ColumnLabel2Index(label.Chars(0), c2)
	End Function
End Class

<Serializable()> _
Friend Class ReferenceEqualityComparer
	Implements System.Collections.Generic.IEqualityComparer(Of Reference)

	Public Function Equals1(ByVal x As Reference, ByVal y As Reference) As Boolean Implements System.Collections.Generic.IEqualityComparer(Of Reference).Equals
		Return x.EqualsReference(y)
	End Function

	Public Function GetHashCode1(ByVal obj As Reference) As Integer Implements System.Collections.Generic.IEqualityComparer(Of Reference).GetHashCode
		Return obj.GetReferenceHashCode()
	End Function
End Class

''' <summary>
''' Represents a named value in a formula
''' </summary>
''' <remarks>This class encapsulates a named value that can be used in formulas.
''' You define a variable using the <see cref="M:ciloci.FormulaEngine.FormulaEngine.DefineVariable(System.String)"/> method, set its value, and
''' then reference it from other formulas.  Once you are finished with the variable, you call its Dispose method to remove it from
''' the formula engine.</remarks>
''' <example>This example shows how to define a variable, use it in a formula, and then undefine it:
''' <code>
''' Dim engine As New FormulaEngine
''' ' Create a new variable called 'x'
''' Dim v As Variable = engine.DefineVariable("x")
''' ' Give it a value of 100
''' v.Value = 100
''' ' Create a formula that uses the variable
''' Dim f As Formula = engine.CreateFormula("=x*cos(x)")
''' ' Evaluate the formula to get a result
''' Dim result As Object = f.Evaluate()
''' ' Undefine the variable
''' v.Dispose()
''' </code>
''' </example>
Public Class Variable

	Private MyEngine As FormulaEngine
	Private MyReference As NamedReference

	Friend Sub New(ByVal engine As FormulaEngine, ByVal name As String)
		MyEngine = engine
		MyReference = engine.ReferenceFactory.Named(name)
		engine.AddFormula("=0", MyReference)
		MyReference = engine.ReferencePool.GetPooledReference(MyReference)
		MyReference.OperandValue = Nothing
	End Sub

	''' <summary>
	''' Recalculates any formulas that depend on this variable
	''' </summary>
	''' <remarks>Call this method after you have changed the value of a variable and you want the formula engine
	''' to recalculate any dependant formulas.</remarks>
	Public Sub Recalculate()
		MyEngine.Recalculate(MyReference)
	End Sub

	''' <summary>
	''' Undefines the variable
	''' </summary>
	''' <remarks>Call this method when you are finished using the variable and wish to remove it from the formula engine</remarks>
	Public Sub Dispose()
		MyEngine.RemoveFormulaAt(MyReference)
		MyReference = Nothing
		MyEngine = Nothing
	End Sub

	''' <summary>
	''' Gets the name of the variable
	''' </summary>
	''' <value>The variable's name</value>
	''' <remarks>Use this property to get the name of the variable</remarks>
	Public ReadOnly Property Name() As String
		Get
			Return MyReference.Name
		End Get
	End Property

	''' <summary>
	''' Gets or sets the value of the variable
	''' </summary>
	''' <value>The value of the variable</value>
	''' <remarks>Use this property to get the value of the variable or to assign a new value to it.  Once a new value is assigned, 
	''' all formulas that reference it will use the new value once they are recalculated.</remarks>
	Public Property Value() As Object
		Get
			Return MyReference.OperandValue
		End Get
		Set(ByVal value As Object)
			MyReference.OperandValue = value
		End Set
	End Property

	Friend ReadOnly Property Reference() As NamedReference
		Get
			Return MyReference
		End Get
	End Property
End Class

Friend Interface IAnalyzer
	Sub Reset()
	ReadOnly Property ReferenceInfos() As ReferenceParseInfo()
End Interface

Friend Class ParserFactory

	Private Shared OurParserMap As Generic.IDictionary(Of GrammarType, PerCederberg.Grammatica.Runtime.Parser)

	Shared Sub New()
		OurParserMap = New Generic.Dictionary(Of GrammarType, PerCederberg.Grammatica.Runtime.Parser)
	End Sub

	Private Sub New()

	End Sub

	Public Shared Function GetParser(ByVal gt As GrammarType) As PerCederberg.Grammatica.Runtime.Parser
		If OurParserMap.ContainsKey(gt) = False Then
			OurParserMap.Add(gt, CreateParser(gt))
		End If

		Return OurParserMap.Item(gt)
	End Function

	Public Shared Sub Clear()
		OurParserMap.Clear()
	End Sub

	Private Shared Function CreateParser(ByVal type As GrammarType) As PerCederberg.Grammatica.Runtime.Parser
		Select Case type
			Case GrammarType.Excel
				Dim sr As New System.IO.StringReader(String.Empty)
				Dim analyzer As New CustomExcelAnalyzer
				Return New ExcelParser(sr, analyzer)
			Case GrammarType.General
				Dim sr As New System.IO.StringReader(String.Empty)
				Dim analyzer As New CustomGeneralAnalyzer
				Return New GeneralParser(sr, analyzer)
			Case Else
				Throw New NotSupportedException("Unknown grammar type")
		End Select
	End Function
End Class