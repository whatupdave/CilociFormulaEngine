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

Imports System.Drawing

''' <summary>
''' Creates all references used by the formula engine
''' </summary>
''' <remarks>This class is responsible for creating all references that you will use when interacting with the formula engine.
''' It has methods for creating sheet and non-sheet references.  Sheet references can be created from a rectangle, string, or
''' integer indices.  There are overloads for creating references to a specific sheet or to the currently active sheet.  Non-grid references
''' such as named and external references are also created by this class.
''' </remarks>
Public NotInheritable Class ReferenceFactory

	Private MyOwner As FormulaEngine

	Friend Sub New(ByVal owner As FormulaEngine)
		MyOwner = owner
	End Sub

	''' <summary>
	''' Creates a sheet reference from a string
	''' </summary>
	''' <param name="s">A string that contains a sheet reference expression</param>
	''' <returns>A sheet reference parsed from the given string</returns>
	''' <remarks>This method creates a sheet reference by parsing a given string.
	''' The method accepts strings with the following syntax:
	''' <list type="table">
	''' <listheader><term>String format</term><description>Resultant reference</description></listheader>
	''' <item>
	''' <term>Column letter followed by a row number: "C3"</term><description>Cell</description>
	''' </item>
	''' <item>
	''' <term>Two column letter and row number pairs separated by a colon: "C3:D4"</term><description>Cell range</description>
	''' </item>
	''' <item>
	''' <term>Two column letters separated by a colon: "E:G"</term><description>Columns</description>
	''' </item>
	''' <item>
	''' <term>Two row numbers separated by a colon: "4:6"</term><description>Rows</description>
	''' </item>
	''' </list>
	''' All of the above formats can specify a specific sheet by prefixing the reference with a sheet name
	''' followed by an exclamation point (ie: "Sheet2!E:G")
	''' If no sheet name is specified, the currently active sheet is used.
	''' </remarks>
	''' <example>This example shows how you would create sheet references from various strings
	''' <code >
	''' ' Get a reference to cell A1
	''' Dim cellRef As ISheetReference = factory.Parse("A1")
	''' ' Get a reference to cells B2 through E4
	''' Dim rangeRef As ISheetReference = factory.Parse("b2:e4")
	''' ' Get a reference to columns D through F
	''' Dim colsRef As ISheetReference = factory.Parse("D:F")
	''' 'Get a reference to rows 4 through 6
	''' Dim rowsRef As ISheetReference = factory.Parse("4:6")
	''' ' Get a reference to cell C4 on sheet 'Sheet4'
	''' Dim cellRef As ISheetReference = factory.Parse("Sheet4!C4")
	''' </code>
	''' </example>
	''' <exception cref="T:System.ArgumentException">
	''' <para>The given string could not be parsed into a sheet reference</para>
	''' <para>The string references a sheet name that is not defined</para>
	''' <para>The resulting sheet reference is not within the bounds of its sheet</para>
	''' </exception>
	Public Function Parse(ByVal s As String) As ISheetReference
		FormulaEngine.ValidateNonNull(s, "s")
		Dim ref As SheetReference = Me.TryParseGridReference(s)
		If ref Is Nothing Then
			Me.OnInvalidReferenceString()
		End If

		Me.InitializeParsedGridReference(ref, s)

		Return ref
	End Function

	Private Function TryParseGridReference(ByVal s As String) As SheetReference
		If CellReference.IsValidString(s) = True Then
			Return CellReference.FromString(s)
		ElseIf CellRangeReference.IsValidString(s) = True Then
			Return CellRangeReference.FromString(s)
		ElseIf ColumnReference.IsValidString(s) = True Then
			Return ColumnReference.FromString(s)
		ElseIf RowReference.IsValidString(s) = True Then
			Return RowReference.FromString(s)
		Else
			Return Nothing
		End If
	End Function

	Private Sub InitializeParsedGridReference(ByVal ref As SheetReference, ByVal s As String)
		Dim parseProps As ReferenceParseProperties = SheetReference.CreateParseProperties(s)

		ref.SetEngine(MyOwner)
		ref.ProcessParseProperties(parseProps, MyOwner)
		ref.Validate()
		ref.ComputeHashCode()
	End Sub

	''' <summary>
	''' Creates a sheet reference from a rectangle
	''' </summary>
	''' <param name="rect">The rectangle to create the sheet reference from</param>
	''' <returns>A sheet reference that matches the given rectangle</returns>
	''' <remarks>This method is identical to <see cref="M:ciloci.FormulaEngine.ReferenceFactory.FromRectangle(ciloci.FormulaEngine.ISheet,System.Drawing.Rectangle)"/>
	''' except that the returned reference is on the currently active sheet.
	''' </remarks>
	Public Function FromRectangle(ByVal rect As Rectangle) As ISheetReference
		Return FromRectangle(MyOwner.Sheets.ActiveSheet, rect)
	End Function

	''' <summary>
	''' Creates a sheet reference from a rectangle and a sheet
	''' </summary>
	''' <param name="rect">The rectangle to create the sheet reference from</param>
	''' <param name="sheet">The sheet that the reference will be on</param>
	''' <returns>A sheet reference that matches the given rectangle and is on the given sheet</returns>
	''' <remarks>Use this method when you have a rectangle that you would like translated into a sheet reference.  Note that the
	''' top-left corner of the sheet is (1,1).  The method will try to create the appropriate type of reference based on the
	''' dimensions of the rectangle.  For example: A rectangle 1 unit wide and 1 unit tall will be translated into a cell reference</remarks>
	''' <exception cref="T:System.ArgumentException">
	''' <para>The resulting sheet reference is not within the bounds of its sheet</para>
	''' <para>The given sheet argument is not registered with the SheetManager</para>
	''' </exception>
	''' <example>The following code creates a reference to the range A1:B2 on the currently active sheet
	''' <code>
	''' Dim engine As New FormulaEngine
	''' Dim rect As New Rectangle(1, 1, 2, 2)
	''' Dim ref As ISheetReference = factory.FromRectangle(rect)
	''' </code>
	''' </example>
	Public Function FromRectangle(ByVal sheet As ISheet, ByVal rect As Rectangle) As ISheetReference
		FormulaEngine.ValidateNonNull(sheet, "sheet")
		Dim ref As SheetReference
		Dim sheetRect As Rectangle = SheetReference.GetSheetRectangle(sheet)

		If rect.Width = 1 And rect.Height = 1 Then
			ref = New CellReference(rect.Top, rect.Left)
		ElseIf rect.Height = sheetRect.Height Then
			ref = New ColumnReference(rect.Left, rect.Right - 1)
		ElseIf rect.Width = sheetRect.Width Then
			ref = New RowReference(rect.Top, rect.Bottom - 1)
		Else
			ref = New CellRangeReference(rect)
		End If

		Me.InitializeGridReference(ref, sheet)
		Return ref
	End Function

	Private Sub OnInvalidReferenceString()
		Throw New ArgumentException("The value could not be parsed into a reference")
	End Sub

	Private Sub InitializeGridReference(ByVal ref As SheetReference, ByVal sheet As ISheet)
		ref.SetEngine(MyOwner)
		Me.ValidateSheet(sheet)
		ref.SetSheet(sheet)
		ref.Validate()
		ref.ComputeHashCode()
	End Sub

	Private Sub ValidateSheet(ByVal sheet As ISheet)
		If MyOwner.Sheets.Contains(sheet) = False Then
			Throw New ArgumentException("The sheet does not exist in the sheets collection")
		End If
	End Sub

	''' <summary>
	''' Creates a sheet reference to a specific cell on the active sheet
	''' </summary>
	''' <param name="row">The row of the cell; first row is 1</param>
	''' <param name="column">The column of the cell; first column is 1</param>
	''' <returns>A sheet reference to the specified row and column on the active sheet</returns>
	''' <remarks>This method behaves exactly like <see cref="M:ciloci.FormulaEngine.ReferenceFactory.Cell(ciloci.FormulaEngine.ISheet,System.Int32,System.Int32)"/>
	''' except that it uses the currently active sheet</remarks>
	Public Function Cell(ByVal row As Integer, ByVal column As Integer) As ISheetReference
		Return Me.Cell(MyOwner.Sheets.ActiveSheet, row, column)
	End Function

	''' <summary>
	''' Creates a sheet reference to a cell on a specific sheet
	''' </summary>
	''' <param name="sheet">The sheet the reference will be on</param>
	''' <param name="row">The row of the cell; first row is 1</param>
	''' <param name="column">The column of the cell; first column is 1</param>
	''' <returns>A sheet reference to the specified row and column and on the given sheet</returns>
	''' <remarks>Use this method when you need a sheet reference to a specific cell and sheet and have
	''' the row and column indices handy.
	''' You usually need a cell reference when you wish to bind a formula to a cell</remarks>
	''' <exception cref="T:System.ArgumentException">
	''' <para>The cell at row,col is not within the bounds of the given sheet</para>
	''' <para>The given sheet argument is not registered with the SheetManager</para>
	''' </exception>
	''' <example>The following code creates a reference to the cell C3 on the currently active sheet
	''' <code>
	''' Dim engine As New FormulaEngine
	''' Dim ref As ISheetReference = engine.ReferenceFactory.Cell(3, 3)
	''' </code>
	''' </example>
	Public Function Cell(ByVal sheet As ISheet, ByVal row As Integer, ByVal column As Integer) As ISheetReference
		FormulaEngine.ValidateNonNull(sheet, "sheet")
		Dim ref As New CellReference(row, column)
		Me.InitializeGridReference(ref, sheet)
		Return ref
	End Function

	''' <summary>
	''' Creates a sheet reference to a range of cells on the currently active sheet
	''' </summary>
	''' <param name="startRow">The top of the range</param>
	''' <param name="startColumn">The left of the range</param>
	''' <param name="endRow">The right of the range</param>
	''' <param name="endColumn">The bottom of the range</param>
	''' <returns>A sheet reference to the specified range on the currently active sheet</returns>
	''' <remarks>This method behaves exactly like <see cref="M:ciloci.FormulaEngine.ReferenceFactory.Cells(ciloci.FormulaEngine.ISheet,System.Int32,System.Int32,System.Int32,System.Int32)"/>
	''' except that it uses the currently active sheet</remarks>
	Public Function Cells(ByVal startRow As Integer, ByVal startColumn As Integer, ByVal endRow As Integer, ByVal endColumn As Integer) As ISheetReference
		Return Me.Cells(MyOwner.Sheets.ActiveSheet, startRow, startColumn, endRow, endColumn)
	End Function

	''' <summary>
	''' Creates a sheet reference to a range of cells on a given sheet
	''' </summary>
	''' <param name="sheet">The sheet the reference will be on</param>
	''' <param name="startRow">The top row of the range</param>
	''' <param name="startColumn">The left column of the range</param>
	''' <param name="endRow">The bottom row of the range</param>
	''' <param name="endColumn">The right column of the range</param>
	''' <returns>A sheet reference to the specified range on the specified sheet</returns>
	''' <remarks>Use this method when you need a sheet reference to a range of cells on a specific sheet and have the four indices handy.</remarks>
	''' <exception cref="T:System.ArgumentException">
	''' <para>The resultant range is not within the bounds of the given sheet</para>
	''' <para>The given sheet argument is not registered with the SheetManager</para>
	''' </exception>
	''' <example>The following example creates a reference to the range C3:E4 on the currently active sheet
	''' <code>
	''' Dim engine As New FormulaEngine
	''' Dim ref As ISheetReference = engine.ReferenceFactory.Cells(3, 3, 4, 5)
	''' </code>
	''' </example>
	Public Function Cells(ByVal sheet As ISheet, ByVal startRow As Integer, ByVal startColumn As Integer, ByVal endRow As Integer, ByVal endColumn As Integer) As ISheetReference
		FormulaEngine.ValidateNonNull(sheet, "sheet")
		Dim ref As New CellRangeReference(startRow, startColumn, endRow, endColumn)
		Me.InitializeGridReference(ref, sheet)
		Return ref
	End Function

	''' <summary>
	''' Creates a reference to a range of rows on the active sheet
	''' </summary>
	''' <param name="start">The top row of the range</param>
	''' <param name="finish">The bottom row of the range</param>
	''' <returns>A sheet reference to the range of rows on the active sheet</returns>
	''' <remarks>This method behaves exactly like <see cref="M:ciloci.FormulaEngine.ReferenceFactory.Rows(ciloci.FormulaEngine.ISheet,System.Int32,System.Int32)"/>
	''' except that it uses the currently active sheet</remarks>
	Public Function Rows(ByVal start As Integer, ByVal finish As Integer) As ISheetReference
		Return Me.Rows(MyOwner.Sheets.ActiveSheet, start, finish)
	End Function

	''' <summary>
	''' Creates a reference to a range of rows on a given sheet
	''' </summary>
	''' <param name="sheet">The sheet the reference will use</param>
	''' <param name="start">The top row of the range</param>
	''' <param name="finish">The bottom row of the range</param>
	''' <returns>A sheet reference to the range of rows on the given sheet</returns>
	''' <remarks>This method will create a sheet reference to an entire range of rows on the given sheet.  Use it when
	''' you wish to reference entire rows and have the two indices handy.</remarks>
	''' <exception cref="T:System.ArgumentException">
	''' <para>The resultant range is not within the bounds of the given sheet</para>
	''' <para>The given sheet argument is not registered with the SheetManager</para>
	''' </exception>
	''' <example>The following example creates a reference to rows 5 through 7 on the currently active sheet
	''' <code>
	''' Dim engine As New FormulaEngine
	''' Dim ref As ISheetReference = engine.ReferenceFactory.Rows(5, 7)
	''' </code>
	''' </example>
	Public Function Rows(ByVal sheet As ISheet, ByVal start As Integer, ByVal finish As Integer) As ISheetReference
		FormulaEngine.ValidateNonNull(sheet, "sheet")
		Dim ref As New RowReference(start, finish)
		Me.InitializeGridReference(ref, sheet)
		Return ref
	End Function

	''' <summary>
	''' Creates a reference to a range of columns on the active sheet
	''' </summary>
	''' <param name="start">The left column of the range</param>
	''' <param name="finish">The right column of the range</param>
	''' <returns>A sheet reference to the range of columns on the given sheet</returns>
	''' <remarks>This method behaves exactly like <see cref="M:ciloci.FormulaEngine.ReferenceFactory.Columns(ciloci.FormulaEngine.ISheet,System.Int32,System.Int32)"/>
	''' except that it uses the currently active sheet</remarks>
	Public Function Columns(ByVal start As Integer, ByVal finish As Integer) As ISheetReference
		Return Me.Columns(MyOwner.Sheets.ActiveSheet, start, finish)
	End Function

	''' <summary>
	''' Creates a reference to a range of columns on a given sheet
	''' </summary>
	''' <param name="sheet">The sheet the reference will use</param>
	''' <param name="start">The left column of the range</param>
	''' <param name="finish">The right column of the range</param>
	''' <returns>A sheet reference to the range of columns on the given sheet</returns>
	''' <remarks>This method will create a sheet reference to an entire range of columns on the given sheet.  Use it when you
	''' want to reference entire columns and have the two indices handy.</remarks>
	''' <exception cref="T:System.ArgumentException">
	''' <para>The resultant range is not within the bounds of the given sheet</para>
	''' <para>The given sheet argument is not registered with the SheetManager</para>
	''' </exception>
	''' <example>The following example creates a reference to columns A through C on the currently active sheet
	''' <code>
	''' Dim engine As New FormulaEngine
	''' Dim ref As ISheetReference = engine.ReferenceFactory.Columns(1, 3)
	''' </code>
	''' </example>
	Public Function Columns(ByVal sheet As ISheet, ByVal start As Integer, ByVal finish As Integer) As ISheetReference
		FormulaEngine.ValidateNonNull(sheet, "sheet")
		Dim ref As New ColumnReference(start, finish)
		Me.InitializeGridReference(ref, sheet)
		Return ref
	End Function

	''' <summary>
	''' Creates a named reference
	''' </summary>
	''' <param name="name">The name of the reference</param>
	''' <returns>A reference to the name</returns>
	''' <remarks>A named reference lets you refer to a formula by a name and lets you refer to that name in other formulas.
	''' A valid name must start with an underscore or letter and can be followed by any combination of underscores, letters, and numbers.</remarks>
	''' <exception cref="T:System.ArgumentException">
	''' <para>The name argument is not in the proper format for a named reference</para>
	''' </exception>
	Public Function Named(ByVal name As String) As INamedReference
		FormulaEngine.ValidateNonNull(name, "name")
		If NamedReference.IsValidName(name) = False Then
			Me.OnInvalidReferenceString()
		End If
		Dim ref As New NamedReference(name)
		ref.SetEngine(MyOwner)
		ref.ComputeHashCode()
		Return ref
	End Function

	''' <summary>
	''' Creates an external reference
	''' </summary>
	''' <returns>An external reference</returns>
	''' <remarks>External references are useful when you need to have many formulas outside of a grid (hence the name) and don't
	''' want to create unique names for each formula</remarks>
	Public Function External() As IExternalReference
		Dim ref As New ExternalReference()
		ref.SetEngine(MyOwner)
		ref.ComputeHashCode()
		Return ref
	End Function

	''' <summary>
	''' Creates a reference that is a union of a list of references
	''' </summary>
	''' <param name="references">The references to union</param>
	''' <returns>A reference representing the union of the given references</returns>
	''' <remarks>This method is useful when you need one reference that can represent multiple other references.
	''' For example: If you want to recalculate cells A1 and B2, you would use this method to get a reference representing both those cells
	''' and then pass it to the Recalculate method.</remarks>
	''' <example>This example shows how to create a reference representing the union of cells A1 and B2 and then use that reference
	''' in one call to the recalculate method:
	''' <code>
	''' Dim engine As New FormulaEngine()
	''' ' Assume we have some formulas that depend on cells A1 and B2
	''' ' Create references to those cells
	''' Dim refA1 As IReference = engine.ReferenceFactory.Cell(1, 1)
	''' Dim refB2 As IReference = engine.ReferenceFactory.Cell(1, 2)
	''' ' Union them into one reference
	''' Dim refA1AndA2 As IReference = engine.ReferenceFactory.Union(refA1, refB2)
	''' ' This single call will recalculate all dependents of A1 and B2
	''' engine.Recalculate(refA1AndA2)
	''' </code>
	''' </example>
	Public Function Union(ByVal ParamArray references As IReference()) As IReference
		Dim ref As New ReferenceList(references)
		ref.SetEngine(MyOwner)
		ref.ComputeHashCode()
		Return ref
	End Function

	''' <summary>
	''' Creates a reference that is a union of a list of variables
	''' </summary>
	''' <param name="variables">The variables to union</param>
	''' <returns>A reference representing the union of the given variables</returns>
	''' <remarks>Works the same way as <see cref="M:ciloci.FormulaEngine.ReferenceFactory.Union(ciloci.FormulaEngine.IReference[])"/>
	''' except that it works on a list of variables.</remarks>
	Public Function Union(ByVal ParamArray variables As Variable()) As IReference
		Dim refs(variables.Length - 1) As IReference
		For i As Integer = 0 To refs.Length - 1
			refs(i) = variables(i).Reference
		Next
		Return Me.Union(refs)
	End Function
End Class
