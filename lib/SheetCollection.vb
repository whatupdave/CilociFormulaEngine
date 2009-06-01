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
''' Manages all worksheets within the formula engine
''' </summary>
''' <remarks>This class is responsible for managing the worksheets that are referenced by formulas.  It provides various collection-oriented methods
''' for managing sheets.  Any sheet that you wish to be able to be used in a formula must be registered with this class.</remarks>
<Serializable()> _
Public Class SheetCollection
	Implements System.Runtime.Serialization.ISerializable

	Private MyOwner As FormulaEngine
	Private MyActiveSheet As ISheet
	Private MySheets As IList
	Private Const VERSION As Integer = 1

	Friend Sub New(ByVal owner As FormulaEngine)
		MyOwner = owner
		MySheets = New ArrayList
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyOwner = info.GetValue("Engine", GetType(FormulaEngine))
		MyActiveSheet = info.GetValue("ActiveSheet", GetType(ISheet))
		MySheets = info.GetValue("Sheets", GetType(IList))
	End Sub

	Public Overridable Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
		info.AddValue("Engine", MyOwner)
		info.AddValue("Sheets", MySheets)
		info.AddValue("ActiveSheet", MyActiveSheet)
	End Sub

	''' <summary>
	''' Adds a sheet to the formula engine
	''' </summary>
	''' <param name="sheet">The sheet you wish to add</param>
	''' <remarks>This method registers a sheet with the formula engine.  If sheet is the first sheet in the collection, then it is
	''' set as the <see cref="P:ciloci.FormulaEngine.SheetCollection.ActiveSheet"/>.</remarks>
	''' <exception cref="System.ArgumentNullException">
	''' sheet is null
	''' <para>The sheet's name is null</para>
	''' </exception>
	''' <exception cref="System.ArgumentException">The sheet already exists in the collection
	''' <para>A sheet with the same name already exists in the collection</para>
	''' </exception>
	Public Sub Add(ByVal sheet As ISheet)
		Me.InsertSheet(sheet, Me.Count)
	End Sub

	Private Sub InsertSheet(ByVal sheet As ISheet, ByVal index As Integer)
		FormulaEngine.ValidateNonNull(sheet, "sheet")

		If Me.Contains(sheet) = True Then
			Throw New ArgumentException("The sheet is already contained in the collection")
		End If

		FormulaEngine.ValidateNonNull(sheet.Name, "sheet.Name")

		If Not Me.GetSheetByName(sheet.Name) Is Nothing Then
			Throw New ArgumentException("A sheet with that name already exists")
		End If

		MySheets.Insert(index, sheet)

		If Me.Count = 1 Then
			MyActiveSheet = sheet
		End If
	End Sub

	''' <summary>
	''' Removes a sheet from the collection
	''' </summary>
	''' <param name="sheet">The sheet to remove</param>
	''' <remarks>This method unregisters a sheet from the formula engine.  All references on the removed sheet will become invalid
	''' and all formulas using those references will be recalculated.</remarks>
	''' <exception cref="System.ArgumentException">The sheet is not contained in the collection</exception>
	Public Sub Remove(ByVal sheet As ISheet)
		FormulaEngine.ValidateNonNull(sheet, "sheet")
		If Me.Contains(sheet) = False Then
			Throw New ArgumentException("Sheet not in list")
		End If
		MySheets.Remove(sheet)
		MyOwner.OnSheetRemoved(sheet)
		If Me.Count = 0 Then
			MyActiveSheet = Nothing
		End If
	End Sub

	''' <summary>
	''' Inserts a sheet into the collection
	''' </summary>
	''' <param name="index">The index to insert the sheet at</param>
	''' <param name="sheet">The sheet to insert</param>
	''' <remarks>This method lets you insert a sheet at a particular index</remarks>
	''' <exception cref="System.ArgumentOutOfRangeException">index is negative or greater than the sheet count</exception>
	''' <exception cref="System.ArgumentNullException">
	''' sheet is null
	''' <para>The sheet's name is null</para>
	''' </exception>
	''' <exception cref="System.ArgumentException">The sheet already exists in the collection
	''' <para>A sheet with the same name already exists in the collection</para>
	''' </exception>
	Public Sub Insert(ByVal index As Integer, ByVal sheet As ISheet)
		If index < 0 Or index > Me.Count Then
			Throw New ArgumentOutOfRangeException("index")
		End If
		Me.InsertSheet(sheet, index)
	End Sub

	''' <summary>
	''' Determines if the collection contains a sheet
	''' </summary>
	''' <param name="sheet">The sheet to test</param>
	''' <returns>True if the collection contains the sheet; False otherwise</returns>
	''' <remarks>This method lets you test whether a particular sheet is contained in the collection</remarks>
	Public Function Contains(ByVal sheet As ISheet) As Boolean
		Return MySheets.Contains(sheet)
	End Function

	''' <summary>
	''' Gets a sheet by name
	''' </summary>
	''' <param name="name">The name of the desired sheet</param>
	''' <returns>An instance of a sheet with the same name; a null reference if no sheet with the name is found</returns>
	''' <remarks>This method lets you get an instance of a sheet by specifying its name</remarks>
	Public Function GetSheetByName(ByVal name As String) As ISheet
		FormulaEngine.ValidateNonNull(name, "name")
		For Each sheet As ISheet In MySheets
			If name.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase) = True Then
				Return sheet
			End If
		Next
		Return Nothing
	End Function

	''' <summary>
	''' Gets the index of a sheet in the collection
	''' </summary>
	''' <param name="sheet">The sheet whose index you wish to get</param>
	''' <returns>The index of the sheet; -1 if sheet is not contained in the collection</returns>
	''' <remarks>A simple method that allows you to get the index of a sheet</remarks>
	Public Function IndexOf(ByVal sheet As ISheet) As Integer
		For i As Integer = 0 To MySheets.Count - 1
			If sheet Is MySheets.Item(i) Then
				Return i
			End If
		Next
		Return -1
	End Function

	Friend Sub Clear()
		MySheets.Clear()
		MyActiveSheet = Nothing
	End Sub

	''' <summary>
	''' Gets or sets the active sheet of the formula engine
	''' </summary>
	''' <value>The active sheet</value>
	''' <remarks>The active sheet is the sheet used when none is specified.  When creating references via the <see cref="T:ciloci.FormulaEngine.ReferenceFactory"/>,
	''' this is the sheet that will be used unless one is specified.  The same applies when creating formulas
	''' that have sheet references that do not explicitly specify a sheet.  This property can be modified at any time.
	''' <note>The value of this property can only be set to a sheet already contained in the collection</note>
	''' </remarks>
	''' <exception cref="System.ArgumentException">The property was assigned a sheet that is not in the collection</exception>
	''' <exception cref="System.ArgumentNullException">The property was assigned a null reference</exception>
	Public Property ActiveSheet() As ISheet
		Get
			Return MyActiveSheet
		End Get
		Set(ByVal value As ISheet)
			FormulaEngine.ValidateNonNull(value, "value")
			If Me.Contains(value) = False Then
				Throw New ArgumentException("The active sheet must exist in the sheet collection")
			End If
			MyActiveSheet = value
		End Set
	End Property

	''' <summary>
	''' Gets a sheet at an index in the collection
	''' </summary>
	''' <param name="index">The index of the sheet to retrieve</param>
	''' <value>The sheet at the specified index</value>
	''' <remarks>This property is a simple indexer into the collection</remarks>
	''' <exception cref="System.ArgumentOutOfRangeException">index is negative or greater than the index of the last sheet</exception>
	Public ReadOnly Property Item(ByVal index As Integer) As ISheet
		Get
			If index < 0 Or index >= MySheets.Count Then
				Throw New ArgumentOutOfRangeException("index")
			End If
			Return MySheets.Item(index)
		End Get
	End Property

	''' <summary>
	''' Gets the number of sheets in the collection
	''' </summary>
	''' <value>The number of sheets in the collection</value>
	''' <remarks>Use this property when you need to get a count of the number of sheets registered with the formula engine</remarks>
	Public ReadOnly Property Count() As Integer
		Get
			Return MySheets.Count
		End Get
	End Property
End Class