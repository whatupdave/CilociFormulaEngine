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

' Implementations of all reference types

Imports System.Drawing

''' <summary>
''' Base class for all references
''' </summary>
<Serializable()> _
Friend MustInherit Class Reference
	Implements IFormulaComponent
	Implements IOperand
	Implements IReference
	Implements System.Runtime.Serialization.ISerializable

	Private MyEngine As FormulaEngine
	Private MyValid As Boolean
	Protected Const REFERENCE_HASH_SIZE As Integer = 4
	Private MyHashCode As Integer
    Private Const VERSION As Integer = 1

    Public Event Recalculated As EventHandler Implements IReference.Recalculated

	Protected Sub New()
		MyValid = True
	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyValid = info.GetBoolean("Valid")
		MyEngine = info.GetValue("Engine", GetType(FormulaEngine))
	End Sub

	Public Overridable Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
		info.AddValue("Valid", MyValid)
		info.AddValue("Engine", MyEngine)
	End Sub

	Public Sub SetEngine(ByVal engine As FormulaEngine)
		MyEngine = engine
		Me.OnEngineSet(engine)
	End Sub

	Protected MustOverride Function CreateGridOps() As GridOperationsBase

	Public Overridable Sub ProcessParseProperties(ByVal props As ReferenceParseProperties, ByVal engine As FormulaEngine)

	End Sub

	Public Shared Function References2String(ByVal refs As IList) As String
		Dim arr(refs.Count - 1) As String

		For i As Integer = 0 To refs.Count - 1
			arr(i) = refs(i).ToString()
		Next

		Return String.Join(", ", arr)
	End Function

	Protected Overridable Sub OnEngineSet(ByVal engine As FormulaEngine)

	End Sub

	Protected Overridable Sub CopyToReference(ByVal target As Reference)
		target.SetEngine(MyEngine)
	End Sub

	Public Overridable Sub Validate()

	End Sub

	Public Function GetReferenceHashCode() As Integer
		Debug.Assert(MyHashCode <> 0, "hash code not set")
		Return MyHashCode
	End Function

	Public Sub ComputeHashCode()
		MyHashCode = Me.ComputeHashCodeInternal()
	End Sub

	Protected Overridable Function ComputeHashCodeInternal() As Integer
		Dim hashData As Byte() = Me.GetHashData()
		Return Me.ComputeJooatHash(hashData)
	End Function

	Protected MustOverride Function GetHashData() As Byte()

	Protected Overridable Sub GetBaseHashData(ByVal bytes As Byte())
		Dim typeHashCode As Integer = Me.GetType().GetHashCode()
		Me.IntegerToBytes(typeHashCode, bytes, 0)
	End Sub

	Protected Sub IntegerToBytes(ByVal i As Integer, ByVal bytes As Byte(), ByVal startIndex As Integer)
		bytes(startIndex) = CByte(i)
		bytes(startIndex + 1) = CByte(i >> 8)
		bytes(startIndex + 2) = CByte(i >> 16)
		bytes(startIndex + 3) = CByte(i >> 24)
	End Sub

	Protected Sub RowColumnIndexToBytes(ByVal index As Integer, ByVal dest As Byte(), ByVal startIndex As Integer)
		dest(startIndex) = CByte(index)
		dest(startIndex + 1) = CByte(index >> 8)
	End Sub

	Private Function ComputeJooatHash(ByVal key As Byte()) As Integer
		Dim hash As Integer = 0

		For i As Integer = 0 To key.Length - 1
			hash += key(i)
			hash += hash << 10
			hash = hash Xor (hash >> 6)
		Next

		hash += hash << 3
		hash = hash Xor (hash >> 11)
		hash += hash << 15

		Return hash
	End Function

	Public MustOverride Function IsOnSheet(ByVal sheet As ISheet) As Boolean
	Public MustOverride Function Intersects(ByVal ref As Reference) As Boolean

	Public Function EqualsReference(ByVal ref As Reference) As Boolean
		If MyValid = False Or ref.MyValid = False Then
			' Invalid references are not equal to anything
			Return False
		ElseIf ref.GetType() Is Me.GetType() Then
			' Only test references of the same type
			Return Me.EqualsReferenceInternal(ref)
		Else
			Return False
		End If
	End Function

	Protected MustOverride Function EqualsReferenceInternal(ByVal ref As Reference) As Boolean

	Public Function EqualsIReference(ByVal ref As IReference) As Boolean Implements IReference.Equals
		Dim realRef As Reference = ref
		Return Me.EqualsReference(realRef)
	End Function

	Protected Sub OnRecalculated()
		RaiseEvent Recalculated(Me, EventArgs.Empty)
	End Sub

	Protected MustOverride Function Format() As String
	Protected MustOverride Function FormatWithProps(ByVal props As ReferenceProperties) As String

	Public Function ToStringIReference() As String Implements IReference.ToString
		If MyValid = False Then
			Return Me.GetRefString()
		Else
			Return Me.Format()
		End If
	End Function

	Public Overridable Function ToStringFormula(ByVal props As ReferenceProperties) As String
		If MyValid = False Then
			Return Me.GetRefString()
		Else
			Return Me.FormatWithProps(props)
		End If
	End Function

	Public Overrides Function ToString() As String
		Return Me.ToStringIReference()
	End Function

	Private Function GetRefString() As String
		Return New ErrorValueWrapper(ErrorValueType.Ref).ToString()
	End Function

	Public MustOverride Sub OnCopy(ByVal rowOffset As Integer, ByVal colOffset As Integer, ByVal destSheet As ISheet, ByVal props As ReferenceProperties)
	Public MustOverride Function IsReferenceEqualForCircularReference(ByVal ref As Reference) As Boolean

	Public Sub Evaluate(ByVal state As System.Collections.Stack, ByVal engine As FormulaEngine) Implements IFormulaComponent.Evaluate
		state.Push(Me)
	End Sub

	Public Sub EvaluateForDependencyReference(ByVal references As System.Collections.IList, ByVal engine As FormulaEngine) Implements IFormulaComponent.EvaluateForDependencyReference
		references.Add(Me)
	End Sub

	Public Function Clone() As Object Implements System.ICloneable.Clone
		Dim refClone As Reference = Me.MemberwiseClone()
		Me.InitializeClone(refClone)
		Return refClone
	End Function

	Protected Overridable Sub InitializeClone(ByVal clone As Reference)

	End Sub

	Public Function Convert(ByVal convertType As OperandType) As IOperand Implements IOperand.Convert
		If convertType = OperandType.Reference Or convertType = OperandType.Self Then
			' Converting to a reference or self takes priority over everything else
			Return Me
		ElseIf convertType = OperandType.SheetReference Then
			Return TryCast(Me, SheetReference)
		ElseIf MyValid = False Then
			Return Me.ConvertWhenInvalid(convertType)
		Else
			Return Me.ConvertInternal(convertType)
		End If
	End Function

	Private Function ConvertWhenInvalid(ByVal convertType As OperandType) As IOperand
		If convertType = OperandType.Error Then
			Return New ErrorValueOperand(ErrorValueType.Ref)
		Else
			Return Nothing
		End If
	End Function

	Protected Overridable Function ConvertInternal(ByVal convertType As OperandType) As IOperand
		Return Nothing
	End Function

	Public Sub GetReferenceValues(ByVal processor As IReferenceValueProcessor) Implements IReference.GetReferenceValues
		Debug.Assert(MyValid = True, "invalid reference should not be getting here")
		Me.GetReferenceValuesInternal(processor)
	End Sub

	Protected Overridable Sub GetReferenceValuesInternal(ByVal processor As IReferenceValueProcessor)

	End Sub

	Public Sub MarkAsInvalid()
		MyValid = False
	End Sub

	Public Overridable Sub Validate(ByVal engine As FormulaEngine) Implements IFormulaComponent.Validate

	End Sub

	Public ReadOnly Property Valid() As Boolean
		Get
			Return MyValid
		End Get
	End Property

	Public Overridable ReadOnly Property CanRangeLink() As Boolean
		Get
			Return False
		End Get
	End Property

	Protected ReadOnly Property Engine() As FormulaEngine
		Get
			Return MyEngine
		End Get
	End Property

	Public ReadOnly Property Value() As Object Implements IOperand.Value
		Get
			Return Me
		End Get
	End Property

	Public ReadOnly Property GridOps() As GridOperationsBase
		Get
			' Don't cache this for now
			Return Me.CreateGridOps()
		End Get
	End Property

	Public ReadOnly Property NativeType() As OperandType Implements IOperand.NativeType
		Get
			Return OperandType.Reference
		End Get
	End Property
End Class

''' <summary>
''' Base class for references that are on a sheet
''' </summary>
<Serializable()> _
Friend MustInherit Class SheetReference
	Inherits Reference
	Implements ISheetReference

    Private _left As Integer = -1

	<Serializable()> _
	Protected MustInherit Class GridReferenceProperties
		Inherits ReferenceProperties

		' Was the sheet name explicitly specified in the formula?
		Public ImplicitSheet As Boolean
	End Class

	Protected Const COLUMN_REGEX As String = "[a-zA-Z]{1,2}"
	Protected Const ROW_REGEX As String = "[0-9]{1,5}"
	Protected Const SHEET_REGEX As String = "([_A-Za-z][\w]*!)?"
	Protected Const MAX_COLUMN As Integer = (26 ^ 2) + 26
	Protected Const MAX_ROW As Integer = UInt16.MaxValue
	Protected Const GRID_REFERENCE_HASH_SIZE As Integer = REFERENCE_HASH_SIZE + 4
	Private MySheet As ISheet

	Protected Sub New()

	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MySheet = info.GetValue("Sheet", GetType(ISheet))
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Sheet", MySheet)
	End Sub

	Public Shared Function IsValidColumnIndex(ByVal columnIndex As Integer) As Boolean
		Return Not (columnIndex < 1 Or columnIndex > MAX_COLUMN)
	End Function

	Public Shared Function IsValidRowIndex(ByVal rowIndex As Integer) As Boolean
		Return Not (rowIndex < 1 Or rowIndex > MAX_ROW)
	End Function

	Protected Shared Sub ValidateColumnIndex(ByVal columnIndex As Integer)
		If IsValidColumnIndex(columnIndex) = False Then
			Throw New ArgumentOutOfRangeException("columnIndex", "Value must be greater than 0 and less than " & MAX_COLUMN)
		End If
	End Sub

	Protected Shared Sub ValidateRowIndex(ByVal rowIndex As Integer)
		If IsValidRowIndex(rowIndex) = False Then
			Throw New ArgumentOutOfRangeException("RowIndex", "Value must be greater than 0 and less than " & MAX_ROW)
		End If
	End Sub

	Public Shared Sub GetProperties(ByVal implicitSheet As Boolean, ByVal props As ReferenceProperties)
		DirectCast(props, GridReferenceProperties).ImplicitSheet = implicitSheet
	End Sub

	Protected Shared Sub GetProperties(ByVal implicitSheet As Boolean, ByVal props As GridReferenceProperties)
		props.ImplicitSheet = implicitSheet
	End Sub

	Protected Shared Function PrepareParseString(ByVal s As String) As String
		s = RemoveSheetReference(s)
		Return s.Replace("$", String.Empty)
	End Function

	Protected Shared Function RemoveSheetReference(ByVal s As String) As String
		Dim bangIndex As Integer = s.IndexOf("!"c)
		Dim count As Integer = System.Math.Max(0, bangIndex + 1)
		s = s.Remove(0, count)
		Return s
	End Function

	Public Shared Function CreateParseProperties(ByVal s As String) As ReferenceParseProperties
		Dim props As New ReferenceParseProperties

		Dim bangIndex As Integer = s.IndexOf("!"c)

		If bangIndex <> -1 Then
			props.SheetName = s.Substring(0, bangIndex)
		End If

		Return props
	End Function

	Public Overrides Sub ProcessParseProperties(ByVal props As ReferenceParseProperties, ByVal engine As FormulaEngine)
		Dim sheet As ISheet
		If props.SheetName Is Nothing Then
			sheet = engine.Sheets.ActiveSheet
		Else
			sheet = engine.Sheets.GetSheetByName(props.SheetName)
		End If

		If Not sheet Is Nothing Then
			Me.SetSheet(sheet)
		End If
	End Sub

	Private Function GetValidateException() As Exception
		Dim ex As Exception

		If Me.Engine.Sheets.Count = 0 Then
			ex = New InvalidOperationException("The formula has a sheet reference but no sheets are defined")
		ElseIf MySheet Is Nothing Then
			ex = New ArgumentException("No sheet with that name exists")
		ElseIf Me.IsInGrid = False Then
			ex = New ArgumentException(String.Format("Grid reference {0} is not contained in sheet {1}", Me.ToString(), MySheet.Name))
		Else
			ex = Nothing
		End If

		Return ex
	End Function

	' Validate for reference factory
	Public Overrides Sub Validate()
		Dim ex As Exception = Me.GetValidateException()

		If Not ex Is Nothing Then
			Throw ex
		End If
	End Sub

	' Validate for formula
	Public Overrides Sub Validate(ByVal engine As FormulaEngine)
		MyBase.Validate(engine)

		Dim ex As Exception = Me.GetValidateException()

		If Not ex Is Nothing Then
			Throw New InvalidFormulaException(ex)
		End If
	End Sub

	Protected Overrides Sub CopyToReference(ByVal target As Reference)
		MyBase.CopyToReference(target)
		Dim gRef As SheetReference = target
		gRef.SetSheet(MySheet)
	End Sub

	Public Shared Function ColumnIndex2Label(ByVal columnIndex As Integer) As String
		SheetReference.ValidateColumnIndex(columnIndex)

		Dim firstIndex, secondIndex As Integer

		columnIndex -= 1

		firstIndex = columnIndex \ 26
		secondIndex = columnIndex Mod 26

		If firstIndex = 0 Then
			Return System.Convert.ToChar(65 + secondIndex).ToString()
		Else
			Return System.Convert.ToChar(65 + firstIndex - 1) & System.Convert.ToChar(65 + secondIndex)
		End If
	End Function

	Public Shared Function ColumnLabel2Index(ByVal char1 As Char, Optional ByVal char2 As Char = Char.MinValue) As Integer
		Dim index1, index2 As Integer
		char1 = Char.ToUpper(char1)
		index1 = System.Convert.ToInt32(char1) - 65 + 1

		If char2 <> Char.MinValue Then
			char2 = Char.ToUpper(char2)
			index2 = System.Convert.ToInt32(char2) - 65
			Return index1 * 26 + (index2 + 1)
		Else
			Return index1
		End If
	End Function

	Protected Shared Function GetAbsoluteString(ByVal absolute As Boolean) As String
		If absolute = True Then
			Return "$"
		Else
			Return String.Empty
		End If
	End Function

	Protected Shared Function OffsetIndex(ByVal index As String, ByVal offset As Integer, ByVal isAbsolute As Boolean) As Integer
		If isAbsolute = False Then
			Return index + offset
		Else
			Return index
		End If
	End Function

	Public NotOverridable Overrides Sub OnCopy(ByVal rowOffset As Integer, ByVal colOffset As Integer, ByVal destSheet As ISheet, ByVal props As ReferenceProperties)
		Dim gridProps As GridReferenceProperties = props

		If gridProps.ImplicitSheet = True Then
			Me.SetSheet(destSheet)
		End If

		Me.OnCopyInternal(rowOffset, colOffset, destSheet, props)

		If Me.IsInGrid() = False Then
			Me.MarkAsInvalid()
		End If
	End Sub

	Protected MustOverride Sub OnCopyInternal(ByVal rowOffset As Integer, ByVal colOffset As Integer, ByVal destSheet As ISheet, ByVal props As ReferenceProperties)

	Protected NotOverridable Overrides Function FormatWithProps(ByVal props As ReferenceProperties) As String
		Dim realProps As GridReferenceProperties = props
		Dim refString As String = Me.FormatWithPropsInternal(props)
		Dim sheetString As String = String.Empty
		Dim bang As String = String.Empty

		If realProps.ImplicitSheet = False Then
			sheetString = MySheet.Name
			bang = "!"
		End If

		Return String.Concat(sheetString, bang, refString)
	End Function

	Protected MustOverride Function FormatWithPropsInternal(ByVal props As ReferenceProperties) As String
	Protected MustOverride Function FormatInternal() As String

	Protected NotOverridable Overrides Function Format() As String
		Dim refString As String = Me.FormatInternal()
		Dim sheetName As String = MySheet.Name
		Return String.Concat(sheetName, "!", refString)
	End Function

	Protected Overrides Sub GetReferenceValuesInternal(ByVal processor As IReferenceValueProcessor)
		Dim rect As Rectangle = Me.Range

		For row As Integer = rect.Top To rect.Bottom - 1
			For col As Integer = rect.Left To rect.Right - 1
				Dim value As Object = MySheet.GetCellValue(row, col)
				If processor.ProcessValue(value) = False Then
					Return
				End If
			Next
		Next
	End Sub

	Public Function GetValuesTable() As Object(,) Implements ISheetReference.GetValuesTable
		Dim rect As Rectangle = Me.Range
		Dim arr(rect.Height - 1, rect.Width - 1) As Object

		For row As Integer = rect.Top To rect.Bottom - 1
			For col As Integer = rect.Left To rect.Right - 1
				Dim value As Object = MySheet.GetCellValue(row, col)
				arr(row - rect.Top, col - rect.Left) = value
			Next
		Next

		Return arr
	End Function

	Public NotOverridable Overrides Function Intersects(ByVal ref As Reference) As Boolean
		If Not TypeOf (ref) Is SheetReference Then
			Return False
		End If

		Dim gref As SheetReference = DirectCast(ref, SheetReference)

		If Not MySheet Is gref.MySheet Then
			Return False
		Else
			Return Me.Range.IntersectsWith(gref.Range)
		End If
	End Function

	Public Sub SetSheet(ByVal sheet As ISheet)
		MySheet = sheet
		Me.OnSheetSet(sheet)
	End Sub

	Protected Overridable Sub OnSheetSet(ByVal sheet As ISheet)

	End Sub

	Public Sub SetSheetForRangeMove(ByVal sheet As ISheet)
		If MySheet Is sheet Then
			Return
		End If

		Me.SetSheet(sheet)

		Dim props As IList = Me.Engine.GetReferenceProperties(Me)
		Me.ClearImplicitSheet(props)
	End Sub

	Private Sub ClearImplicitSheet(ByVal props As IList)
		For Each prop As GridReferenceProperties In props
			prop.ImplicitSheet = False
		Next
	End Sub

	Protected Function IsInGrid() As Boolean
		Return Me.GridRectangle.Contains(Me.Range)
	End Function

	Public Shared Function IsRectangleInSheet(ByVal rect As Rectangle, ByVal sheet As ISheet) As Boolean
		Dim gridRect As New System.Drawing.Rectangle(1, 1, sheet.ColumnCount, sheet.RowCount)
		Return gridRect.Contains(rect)
	End Function

	Public Overrides Function IsOnSheet(ByVal sheet As ISheet) As Boolean
		Return MySheet Is sheet
	End Function

	Protected Overrides Sub GetBaseHashData(ByVal bytes() As Byte)
		MyBase.GetBaseHashData(bytes)
		Dim sheetHashCode As Integer = DirectCast(MySheet, Object).GetHashCode()
		MyBase.IntegerToBytes(sheetHashCode, bytes, REFERENCE_HASH_SIZE)
	End Sub

	Protected NotOverridable Overrides Function EqualsReferenceInternal(ByVal ref As Reference) As Boolean
		Dim gridRef As SheetReference = ref
		If Me.IsOnSheet(gridRef.MySheet) = False Then
			Return False
		Else
			Return Me.EqualsGridReference(gridRef)
		End If
	End Function

	Public Shared Function GetSheetRectangle(ByVal sheet As ISheet) As Rectangle
		Return New System.Drawing.Rectangle(1, 1, sheet.ColumnCount, sheet.RowCount)
	End Function

	Protected MustOverride Function EqualsGridReference(ByVal ref As SheetReference) As Boolean

	Public Overridable ReadOnly Property IsEmptyIntersection() As Boolean
		Get
			Return False
		End Get
	End Property

	Public MustOverride ReadOnly Property Range() As Rectangle Implements ISheetReference.Area

	Public ReadOnly Property GridRectangle() As System.Drawing.Rectangle
		Get
			Return New System.Drawing.Rectangle(1, 1, MySheet.ColumnCount, MySheet.RowCount)
		End Get
	End Property

	Public ReadOnly Property Sheet() As ISheet Implements ISheetReference.Sheet
		Get
			Return MySheet
		End Get
	End Property

	Public ReadOnly Property Row() As Integer Implements ISheetReference.Row
		Get
			Return Me.Range.Top
		End Get
	End Property

	Public ReadOnly Property Column() As Integer Implements ISheetReference.Column
        Get
            If _left = -1 Then
                _left = Me.Range.Left
            End If

            Return _left
        End Get
	End Property

	Public ReadOnly Property Height() As Integer Implements ISheetReference.Height
		Get
			Return Me.Range.Height
		End Get
	End Property

	Public ReadOnly Property Width() As Integer Implements ISheetReference.Width
		Get
			Return Me.Range.Width
		End Get
	End Property
End Class

<Serializable()> _
Friend Class CellReference
	Inherits SheetReference
	Implements IFormulaSelfReference

	<Serializable()> _
	Private Class Properties
		Inherits GridReferenceProperties

		Public ColumnAbsolute As Boolean
		Public RowAbsolute As Boolean
	End Class

	Private Class CellGridOps
		Inherits GridOperationsBase

		Private MyOwner As CellReference

		Public Sub New(ByVal owner As CellReference)
			MyOwner = owner
		End Sub

		Public Overrides Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			If MyOwner.MyColumnIndex >= insertAt Then
				' We are in the inserted area; we must adjust
				MyOwner.MyColumnIndex += count
				Return ReferenceOperationResultType.Affected
			Else
				Return ReferenceOperationResultType.NotAffected
			End If
		End Function

		Public Overrides Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			If MyOwner.MyColumnIndex > removeAt + count - 1 Then
				' We are to the right of the removed hole
				MyOwner.MyColumnIndex -= count
				Return ReferenceOperationResultType.Affected
			ElseIf MyOwner.MyColumnIndex >= removeAt Then
				' We are in the removed hole
				Return ReferenceOperationResultType.Invalidated
			Else
				Return ReferenceOperationResultType.NotAffected
			End If
		End Function

		Public Overrides Function OnRangeMoved(ByVal source As SheetReference, ByVal dest As SheetReference) As ReferenceOperationResultType
			Dim sourceRect As Rectangle = source.Range
			Dim destRect As Rectangle = dest.Range
			Dim rowOffset, colOffset As Integer
			rowOffset = destRect.Top - sourceRect.Top
			colOffset = destRect.Left - sourceRect.Left
			Dim myRect As Rectangle = MyOwner.Range

			Dim isContainedInSource As Boolean = sourceRect.Contains(myRect) And MyOwner.IsOnSheet(source.Sheet)
			Dim isContainedInDest As Boolean = destRect.Contains(myRect) And MyOwner.IsOnSheet(dest.Sheet)

			If isContainedInSource = True Then
				' We are in the moved range so we have to adjust
				MyOwner.MyRowIndex += rowOffset
				MyOwner.MyColumnIndex += colOffset
				MyOwner.SetSheetForRangeMove(dest.Sheet)
				Return ReferenceOperationResultType.Affected
			ElseIf isContainedInDest = True Then
				' We are overwritten by the moved range
				Return ReferenceOperationResultType.Invalidated
			Else
				' We are not affected
				Return ReferenceOperationResultType.NotAffected
			End If
		End Function

		Public Overrides Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			If MyOwner.MyRowIndex >= insertAt Then
				MyOwner.MyRowIndex += count
				Return ReferenceOperationResultType.Affected
			Else
				Return ReferenceOperationResultType.NotAffected
			End If
		End Function

		Public Overrides Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			If MyOwner.MyRowIndex > removeAt + count - 1 Then
				' We are below the removed hole
				MyOwner.MyRowIndex -= count
				Return ReferenceOperationResultType.Affected
			ElseIf MyOwner.MyRowIndex >= removeAt Then
				' We are in the removed hole
				Return ReferenceOperationResultType.Invalidated
			Else
				Return ReferenceOperationResultType.NotAffected
			End If
		End Function
	End Class

	Private MyRowIndex, MyColumnIndex As Integer
	Private Shared OurRegex As System.Text.RegularExpressions.Regex = CreateRegex()

	Public Sub New(ByVal rowIndex As Integer, ByVal colIndex As Integer)
		SheetReference.ValidateColumnIndex(colIndex)
		SheetReference.ValidateRowIndex(rowIndex)
		MyRowIndex = rowIndex
		MyColumnIndex = colIndex
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyRowIndex = info.GetInt32("RowIndex")
		MyColumnIndex = info.GetInt32("ColumnIndex")
		Me.ComputeHashCode()
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("RowIndex", MyRowIndex)
		info.AddValue("ColumnIndex", MyColumnIndex)
	End Sub

	Protected Overrides Function CreateGridOps() As GridOperationsBase
		Return New CellGridOps(Me)
	End Function

	Public Shared Function FromString(ByVal image As String) As CellReference
		image = SheetReference.PrepareParseString(image)
		Dim c1 As Char = image.Chars(0)
		Dim c2 As Char = image.Chars(1)

		Dim rowStringIndex As Integer
		Dim rowIndex, columnIndex As String

		If Char.IsLetter(c2) = True Then
			columnIndex = ColumnLabel2Index(c1, c2)
			rowStringIndex = 2
		Else
			columnIndex = ColumnLabel2Index(c1)
			rowStringIndex = 1
		End If

		Dim rowRef As String = image.Substring(rowStringIndex)

		rowIndex = Integer.Parse(rowRef)

		Return New CellReference(rowIndex, columnIndex)
	End Function

	Public Shared Function CreateProperties(ByVal implicitSheet As Boolean, ByVal image As String) As ReferenceProperties
		Dim props As New Properties
		SheetReference.GetProperties(implicitSheet, props)

		props.ColumnAbsolute = image.StartsWith("$")
		props.RowAbsolute = image.IndexOf("$", 1) <> -1

		Return props
	End Function

	Private Shared Function CreateRegex() As System.Text.RegularExpressions.Regex
		Dim exp As String = String.Format("^{0}{1}{2}$", SheetReference.SHEET_REGEX, SheetReference.COLUMN_REGEX, SheetReference.ROW_REGEX)
		Return New System.Text.RegularExpressions.Regex(exp)
	End Function

	Public Shared Function IsValidString(ByVal s As String) As Boolean
		Return OurRegex.IsMatch(s)
	End Function

	Public Shared Sub SwapRowProperties(ByVal props1 As ReferenceProperties, ByVal props2 As ReferenceProperties)
		Dim realProps1, realProps2 As Properties
		realProps1 = props1
		realProps2 = props2

		Dim temp As Boolean = realProps1.RowAbsolute
		realProps1.RowAbsolute = realProps2.RowAbsolute
		realProps2.RowAbsolute = temp
	End Sub

	Public Shared Sub SwapColumnProperties(ByVal props1 As ReferenceProperties, ByVal props2 As ReferenceProperties)
		Dim realProps1, realProps2 As Properties
		realProps1 = props1
		realProps2 = props2

		Dim temp As Boolean = realProps1.ColumnAbsolute
		realProps1.ColumnAbsolute = realProps2.ColumnAbsolute
		realProps2.ColumnAbsolute = temp
	End Sub

	Private Function GetConvertedTargetOperandValue(ByVal convertType As OperandType) As IOperand
		Return Me.TargetOperand.Convert(convertType)
	End Function

	Protected Overrides Sub OnCopyInternal(ByVal rowOffset As Integer, ByVal colOffset As Integer, ByVal destSheet As ISheet, ByVal props As ReferenceProperties)
		Dim realProps As Properties = props

		MyRowIndex = SheetReference.OffsetIndex(MyRowIndex, rowOffset, realProps.RowAbsolute)
		MyColumnIndex = SheetReference.OffsetIndex(MyColumnIndex, colOffset, realProps.ColumnAbsolute)
	End Sub

	Public Overrides Function IsReferenceEqualForCircularReference(ByVal ref As Reference) As Boolean
		Return ref Is Me
	End Function

	Protected Overrides Function FormatInternal() As String
		Return String.Concat(ColumnIndex2Label(MyColumnIndex), MyRowIndex.ToString())
	End Function

	Public Function FormatPlain() As String
		Return Me.FormatInternal()
	End Function

	Protected Overloads Overrides Function FormatWithPropsInternal(ByVal props As ReferenceProperties) As String
		Dim realProps As Properties = props

		Dim rowString, colString As String
		Dim rowAbsolute, colAbsolute As String

		rowString = MyRowIndex.ToString()
		colString = ColumnIndex2Label(MyColumnIndex)
		rowAbsolute = GetAbsoluteString(realProps.RowAbsolute)
		colAbsolute = GetAbsoluteString(realProps.ColumnAbsolute)

		Return String.Concat(colAbsolute, colString, rowAbsolute, rowString)
	End Function

	Friend Sub Offset(ByVal rowOffset As Integer, ByVal colOffset As Integer)
		MyRowIndex += rowOffset
		MyColumnIndex += colOffset
	End Sub

	Friend Sub SetColumn(ByVal column As Integer)
		MyColumnIndex = column
	End Sub

	Friend Sub SetRow(ByVal row As Integer)
		MyRowIndex = row
	End Sub

	Protected Overrides Function ConvertInternal(ByVal convertType As OperandType) As IOperand
		Return Me.TargetOperand.Convert(convertType)
	End Function

	Protected Overrides Function EqualsGridReference(ByVal ref As SheetReference) As Boolean
		Dim cellRef As CellReference = ref
		Return MyRowIndex = cellRef.MyRowIndex And MyColumnIndex = cellRef.MyColumnIndex
	End Function

	Public Sub OnFormulaRecalculate(ByVal target As Formula) Implements IFormulaSelfReference.OnFormulaRecalculate
		Dim result As Object = target.Evaluate()
		Me.Sheet.SetFormulaResult(result, MyRowIndex, MyColumnIndex)
		MyBase.OnRecalculated()
	End Sub

	Protected Overrides Function GetHashData() As Byte()
		Dim bytes(GRID_REFERENCE_HASH_SIZE + 2 + 2 - 1) As Byte
		MyBase.GetBaseHashData(bytes)
		MyBase.RowColumnIndexToBytes(MyRowIndex, bytes, GRID_REFERENCE_HASH_SIZE)
		MyBase.RowColumnIndexToBytes(MyColumnIndex, bytes, GRID_REFERENCE_HASH_SIZE + 2)
		Return bytes
	End Function

	Private ReadOnly Property DependencyManager() As DependencyManager
		Get
			Return Me.Engine.DependencyManager
		End Get
	End Property

	Private ReadOnly Property TargetOperand() As IOperand
		Get
			If Me.Valid = False Then
				Return New ErrorValueOperand(ErrorValueType.Ref)
			Else
				Return OperandFactory.CreateDynamic(Me.TargetCellValue)
			End If
		End Get
	End Property

	Private ReadOnly Property TargetCellValue() As Object
		Get
			Return Me.Sheet.GetCellValue(MyRowIndex, MyColumnIndex)
		End Get
	End Property

	Public Overrides ReadOnly Property Range() As System.Drawing.Rectangle
		Get
			Return New System.Drawing.Rectangle(MyColumnIndex, MyRowIndex, 1, 1)
		End Get
	End Property

	Public ReadOnly Property RowIndex() As Integer
		Get
			Return MyRowIndex
		End Get
	End Property

	Public ReadOnly Property ColumnIndex() As Integer
		Get
			Return MyColumnIndex
		End Get
	End Property
End Class

<Serializable()> _
Friend Class CellRangeReference
	Inherits SheetReference

	<Serializable()> _
	Private Class Properties
		Inherits GridReferenceProperties

		Public MyStartProps As ReferenceProperties
		Public MyFinishProps As ReferenceProperties

		Protected Overrides Sub InitializeClone(ByVal clone As ReferenceProperties)
			Dim cloneProps As Properties = clone
			cloneProps.MyStartProps = MyStartProps.Clone()
			cloneProps.MyFinishProps = MyFinishProps.Clone()
		End Sub
	End Class

	Private Class CellRangeGridOps
		Inherits GridOperationsBase

		Private MyOwner As CellRangeReference

		Public Sub New(ByVal owner As CellRangeReference)
			MyOwner = owner
		End Sub

		Public Overrides Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Dim startAffected, finishAffected As ReferenceOperationResultType
			startAffected = MyOwner.MyStart.GridOps.OnColumnsInserted(insertAt, count)
			finishAffected = MyOwner.MyFinish.GridOps.OnColumnsInserted(insertAt, count)

			Return startAffected Or finishAffected
		End Function

		Public Overrides Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Dim result As Point = GridOperationsBase.HandleRangeRemoved(MyOwner.MyStart.Column, MyOwner.MyFinish.Column, removeAt, count)

			If result.IsEmpty = True Then
				Return ReferenceOperationResultType.Invalidated
			Else
				Dim startAffected, finishAffected As Boolean
				startAffected = result.X <> MyOwner.MyStart.Column
				finishAffected = result.Y <> MyOwner.MyFinish.Column

				MyOwner.MyStart.SetColumn(result.X)
				MyOwner.MyFinish.SetColumn(result.Y)

				Return GridOperationsBase.Affected2Enum(startAffected Or finishAffected)
			End If
		End Function

		Public Overrides Function OnRangeMoved(ByVal source As SheetReference, ByVal dest As SheetReference) As ReferenceOperationResultType
			Return MyBase.HandleRangeMoved(MyOwner, source, dest, AddressOf OnSetMovedRangeResult)
		End Function

		Private Sub OnSetMovedRangeResult(ByVal result As Rectangle)
			MyOwner.SetRange(result)
		End Sub

		Public Overrides Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Dim startAffected, finishAffected As ReferenceOperationResultType
			startAffected = MyOwner.MyStart.GridOps.OnRowsInserted(insertAt, count)
			finishAffected = MyOwner.MyFinish.GridOps.OnRowsInserted(insertAt, count)

			Return startAffected Or finishAffected
		End Function

		Public Overrides Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Dim result As Point = GridOperationsBase.HandleRangeRemoved(MyOwner.MyStart.Row, MyOwner.MyFinish.Row, removeAt, count)

			If result.IsEmpty = True Then
				Return ReferenceOperationResultType.Invalidated
			Else
				Dim startAffected, finishAffected As Boolean
				startAffected = result.X <> MyOwner.MyStart.Row
				finishAffected = result.Y <> MyOwner.MyFinish.Row

				MyOwner.MyStart.SetRow(result.X)
				MyOwner.MyFinish.SetRow(result.Y)
				Return GridOperationsBase.Affected2Enum(startAffected Or finishAffected)
			End If
		End Function
	End Class

	Private MyStart, MyFinish As CellReference
	Private Shared OurRegex As System.Text.RegularExpressions.Regex = CreateRegex()

	Private Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyStart = info.GetValue("Start", GetType(CellReference))
		MyFinish = info.GetValue("Finish", GetType(CellReference))
		Me.ComputeHashCode()
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Start", MyStart)
		info.AddValue("Finish", MyFinish)
	End Sub

	Public Shared Function FromString(ByVal image As String) As CellRangeReference
		image = SheetReference.PrepareParseString(image)
		Dim arr As String() = image.Split(":")
		Dim start, finish As CellReference
		start = CellReference.FromString(arr(0))
		finish = CellReference.FromString(arr(1))

		' Handle backwards references
		If start.Row > finish.Row Then
			Dim tempRow As Integer = start.Row
			start.SetRow(finish.Row)
			finish.SetRow(tempRow)
		End If

		If start.Column > finish.Column Then
			Dim tempcol As Integer = start.Column
			start.SetColumn(finish.Column)
			finish.SetColumn(tempcol)
		End If

		Dim ref As New CellRangeReference
		ref.MyStart = start
		ref.MyFinish = finish
		Return ref
	End Function

	Public Shared Function CreateProperties(ByVal implicitSheet As Boolean, ByVal image As String) As ReferenceProperties
		Dim props As New Properties
		SheetReference.GetProperties(implicitSheet, props)

		Dim parts As String() = image.Split(":")

		props.MyStartProps = CellReference.CreateProperties(True, parts(0))
		props.MyFinishProps = CellReference.CreateProperties(True, parts(1))

		Return props
	End Function

	Private Shared Function CreateRegex() As System.Text.RegularExpressions.Regex
		Dim cellExp As String = String.Concat(SheetReference.COLUMN_REGEX, SheetReference.ROW_REGEX)
		Dim exp As String = String.Format("^{0}{1}:{1}$", SheetReference.SHEET_REGEX, cellExp)
		Return New System.Text.RegularExpressions.Regex(exp)
	End Function

	Public Shared Function IsValidString(ByVal s As String) As Boolean
		Return OurRegex.IsMatch(s)
	End Function

	Protected Overrides Function CreateGridOps() As GridOperationsBase
		Return New CellRangeGridOps(Me)
	End Function

	Public Sub New(ByVal startRow As Integer, ByVal startCol As Integer, ByVal finishRow As Integer, ByVal finishCol As Integer)
		Me.SetRange(startRow, startCol, finishRow, finishCol)
	End Sub

	Public Sub New(ByVal rect As Rectangle)
		Me.SetRange(rect)
	End Sub

	Private Sub SetRange(ByVal startRow As Integer, ByVal startCol As Integer, ByVal finishRow As Integer, ByVal finishCol As Integer)
		MyStart = New CellReference(startRow, startCol)
		MyFinish = New CellReference(finishRow, finishCol)
		Me.InitializeInnerRefs()
	End Sub

	Private Sub SetRange(ByVal rect As Rectangle)
		Me.SetRange(rect.Top, rect.Left, rect.Bottom - 1, rect.Right - 1)
	End Sub

	Protected Overrides Sub OnEngineSet(ByVal engine As FormulaEngine)
		Me.InitializeInnerRefs()
	End Sub

	Protected Overrides Sub OnSheetSet(ByVal sheet As ISheet)
		MyStart.SetSheet(sheet)
		MyFinish.SetSheet(sheet)
	End Sub

	Private Sub InitializeInnerRefs()
		MyStart.SetEngine(Me.Engine)
		MyFinish.SetEngine(Me.Engine)
		MyStart.SetSheet(Me.Sheet)
		MyFinish.SetSheet(Me.Sheet)
	End Sub

	Protected Overrides Function EqualsGridReference(ByVal ref As SheetReference) As Boolean
		Dim rangeRef As CellRangeReference = ref
		Return MyStart.EqualsReference(rangeRef.MyStart) And MyFinish.EqualsReference(rangeRef.MyFinish)
	End Function

	Protected Overrides Sub OnCopyInternal(ByVal rowOffset As Integer, ByVal colOffset As Integer, ByVal destSheet As ISheet, ByVal props As ReferenceProperties)
		Dim realProps As Properties = props
		MyStart.OnCopy(rowOffset, colOffset, destSheet, realProps.MyStartProps)
		MyFinish.OnCopy(rowOffset, colOffset, destSheet, realProps.MyFinishProps)
		Me.OrderReferencesAfterOffset(props)
	End Sub

	Private Sub OrderReferencesAfterOffset(ByVal props As Properties)
		Dim startColumn, finishColumn As Integer
		Dim startRow, finishRow As Integer

		startColumn = MyStart.Column
		finishColumn = MyFinish.Column

		startRow = MyStart.Row
		finishRow = MyFinish.Row

		If startColumn > finishColumn Then
			Dim tempcol As Integer = startColumn
			MyStart.SetColumn(finishColumn)
			MyFinish.SetColumn(tempcol)

			CellReference.SwapColumnProperties(props.MyStartProps, props.MyFinishProps)
		End If

		If startRow > finishRow Then
			Dim tempRow As Integer = startRow
			MyStart.SetRow(finishRow)
			MyFinish.SetRow(tempRow)
			CellReference.SwapRowProperties(props.MyStartProps, props.MyFinishProps)
		End If
	End Sub

	Public Overrides Function IsReferenceEqualForCircularReference(ByVal ref As Reference) As Boolean
		Return Me.Intersects(ref)
	End Function

	Protected Overrides Sub InitializeClone(ByVal clone As Reference)
		Dim rangeClone As CellRangeReference = clone
		rangeClone.MyStart = MyStart.Clone()
		rangeClone.MyFinish = MyFinish.Clone()
	End Sub

	Protected Overrides Function FormatInternal() As String
		Dim startString, finishString As String
		startString = MyStart.FormatPlain()
		finishString = MyFinish.FormatPlain()
		Return String.Concat(startString, ":", finishString)
	End Function

	Protected Overloads Overrides Function FormatWithPropsInternal(ByVal props As ReferenceProperties) As String
		Dim realProps As Properties = props
		Dim startString, finishString As String
		startString = MyStart.ToStringFormula(realProps.MyStartProps)
		finishString = MyFinish.ToStringFormula(realProps.MyFinishProps)
		Return String.Concat(startString, ":", finishString)
	End Function

	Protected Overrides Function GetHashData() As Byte()
		Dim bytes(GRID_REFERENCE_HASH_SIZE + 4 + 4 - 1) As Byte
		MyBase.GetBaseHashData(bytes)

		MyBase.RowColumnIndexToBytes(MyStart.RowIndex, bytes, GRID_REFERENCE_HASH_SIZE)
		MyBase.RowColumnIndexToBytes(MyStart.ColumnIndex, bytes, GRID_REFERENCE_HASH_SIZE + 2)

		MyBase.RowColumnIndexToBytes(MyFinish.RowIndex, bytes, GRID_REFERENCE_HASH_SIZE + 4)
		MyBase.RowColumnIndexToBytes(MyFinish.ColumnIndex, bytes, GRID_REFERENCE_HASH_SIZE + 6)

		Return bytes
	End Function

	Public Overrides ReadOnly Property CanRangeLink() As Boolean
		Get
			Return True
		End Get
	End Property

    Private _range As System.Drawing.Rectangle = Nothing
	Public Overrides ReadOnly Property Range() As System.Drawing.Rectangle
        Get
            If _range = Nothing Then
                Dim h, w As Integer
                h = MyFinish.Row - MyStart.Row + 1
                w = MyFinish.Column - MyStart.Column + 1
                _range = New System.Drawing.Rectangle(MyStart.Column, MyStart.Row, w, h)
            End If
            Return _range
        End Get
	End Property
End Class

<Serializable()> _
Friend MustInherit Class RowColumnReference
	Inherits SheetReference

	<Serializable()> _
	Private Class Properties
		Inherits GridReferenceProperties

		Public StartAbsolute As Boolean
		Public FinishAbsolute As Boolean
	End Class

	Protected Class RowColumnGridOps
		Inherits GridOperationsBase

		Private MyOwner As RowColumnReference

		Public Sub New(ByVal owner As RowColumnReference)
			MyOwner = owner
		End Sub

		Public Function OnRowColumnInsert(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Dim result As Point = GridOperationsBase.HandleRangeInserted(MyOwner.Start, MyOwner.Finish, insertAt, count)
			Dim startAffected, finishAffected As Boolean
			startAffected = MyOwner.Start <> result.X
			finishAffected = MyOwner.Finish <> result.Y

			MyOwner.Start = result.X
			MyOwner.Finish = result.Y

			Return GridOperationsBase.Affected2Enum(startAffected Or finishAffected)
		End Function

		Public Function OnRowColumnDelete(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Dim result As Point = GridOperationsBase.HandleRangeRemoved(MyOwner.Start, MyOwner.Finish, removeAt, count)

			If result.IsEmpty = True Then
				Return ReferenceOperationResultType.Invalidated
			Else
				Dim startAffected, finishAffected As Boolean
				startAffected = MyOwner.Start <> result.X
				finishAffected = MyOwner.Finish <> result.Y

				MyOwner.Start = result.X
				MyOwner.Finish = result.Y

				Return GridOperationsBase.Affected2Enum(startAffected Or finishAffected)
			End If
		End Function

		Public Overrides Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Return MyOwner.OnColumnsInserted(insertAt, count, Me)
		End Function

		Public Overrides Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Return MyOwner.OnColumnsRemoved(removeAt, count, Me)
		End Function

		Public Overrides Function OnRangeMoved(ByVal source As SheetReference, ByVal dest As SheetReference) As ReferenceOperationResultType
			Return MyBase.HandleRangeMoved(MyOwner, source, dest, AddressOf OnSetMovedRangeResult)
		End Function

		Private Sub OnSetMovedRangeResult(ByVal result As Rectangle)
			MyOwner.MyStart = MyOwner.GetStartIndex(result)
			MyOwner.MyFinish = MyOwner.GetFinishIndex(result) - 1
		End Sub

		Public Overrides Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Return MyOwner.OnRowsInserted(insertAt, count, Me)
		End Function

		Public Overrides Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer) As ReferenceOperationResultType
			Return MyOwner.OnRowsRemoved(removeAt, count, Me)
		End Function
	End Class

	Private MyStart, MyFinish As Integer

	Protected Sub New(ByVal start As Integer, ByVal finish As Integer)
		Me.ValidateIndex(start)
		Me.ValidateIndex(finish)

		If start > finish Then
			Dim temp As Integer = start
			start = finish
			finish = temp
		End If

		MyStart = start
		MyFinish = finish
	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyStart = info.GetInt32("Start")
		MyFinish = info.GetUInt32("Finish")
		Me.ComputeHashCode()
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Start", MyStart)
		info.AddValue("Finish", MyFinish)
	End Sub

	Protected MustOverride Sub ValidateIndex(ByVal index As Integer)

	Protected Overrides Function CreateGridOps() As GridOperationsBase
		Return New RowColumnGridOps(Me)
	End Function

	Public Shared Function CreateProperties(ByVal implicitSheet As Boolean, ByVal image As String) As ReferenceProperties
		Dim props As New Properties

		SheetReference.GetProperties(implicitSheet, props)

		props.StartAbsolute = image.StartsWith("$")
		props.FinishAbsolute = image.IndexOf("$", 1) <> -1

		Return props
	End Function

	Protected Overrides Function EqualsGridReference(ByVal ref As SheetReference) As Boolean
		Dim realRef As RowColumnReference = ref
		Return MyStart = realRef.MyStart And MyFinish = realRef.MyFinish
	End Function

	Public Overrides Function IsReferenceEqualForCircularReference(ByVal ref As Reference) As Boolean
		Return Me.Intersects(ref)
	End Function

	Protected MustOverride Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType
	Protected MustOverride Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType
	Protected MustOverride Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType
	Protected MustOverride Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType

	Protected MustOverride Function GetStartIndex(ByVal rect As Rectangle) As Integer
	Protected MustOverride Function GetFinishIndex(ByVal rect As Rectangle) As Integer

	Protected MustOverride Function FormatIndex(ByVal index As Integer) As String

	Protected Overrides Function FormatInternal() As String
		Dim startString, finishString As String
		startString = Me.FormatIndex(MyStart)
		finishString = Me.FormatIndex(MyFinish)
		Return String.Concat(startString, ":", finishString)
	End Function

	Protected Overloads Overrides Function FormatWithPropsInternal(ByVal props As ReferenceProperties) As String
		Dim realProps As Properties = props

		Dim startString, finishString As String
		Dim startAbsolute, finishAbsolute As String

		startString = Me.FormatIndex(MyStart)
		finishString = Me.FormatIndex(MyFinish)
		startAbsolute = GetAbsoluteString(realProps.StartAbsolute)
		finishAbsolute = GetAbsoluteString(realProps.FinishAbsolute)

		Return String.Concat(startAbsolute, startString, ":", finishAbsolute, finishString)
	End Function

	Protected MustOverride Function GetCopyOffset(ByVal rowOffset As Integer, ByVal colOffset As Integer) As Integer

	Protected Overrides Sub OnCopyInternal(ByVal rowOffset As Integer, ByVal colOffset As Integer, ByVal destsheet As ISheet, ByVal props As ReferenceProperties)
		Dim offset As Integer = Me.GetCopyOffset(rowOffset, colOffset)
		Dim realProps As Properties = props
		MyStart = SheetReference.OffsetIndex(MyStart, offset, realProps.StartAbsolute)
		MyFinish = SheetReference.OffsetIndex(MyFinish, offset, realProps.FinishAbsolute)

		If MyStart > MyFinish Then
			Dim tempIndex As Integer = MyStart
			MyStart = MyFinish
			MyFinish = tempIndex

			Dim tempBool As Boolean = realProps.StartAbsolute
			realProps.StartAbsolute = realProps.FinishAbsolute
			realProps.FinishAbsolute = tempBool
		End If
	End Sub

	Protected Overrides Function GetHashData() As Byte()
		Dim bytes(GRID_REFERENCE_HASH_SIZE + 2 + 2 - 1) As Byte
		MyBase.GetBaseHashData(bytes)
		MyBase.RowColumnIndexToBytes(MyStart, bytes, GRID_REFERENCE_HASH_SIZE)
		MyBase.RowColumnIndexToBytes(MyFinish, bytes, GRID_REFERENCE_HASH_SIZE + 2)
		Return bytes
	End Function

	Public Overrides ReadOnly Property CanRangeLink() As Boolean
		Get
			Return True
		End Get
	End Property

	Protected Property Start() As Integer
		Get
			Return MyStart
		End Get
		Set(ByVal Value As Integer)
			MyStart = Value
		End Set
	End Property

	Protected Property Finish() As Integer
		Get
			Return MyFinish
		End Get
		Set(ByVal Value As Integer)
			MyFinish = Value
		End Set
	End Property
End Class

<Serializable()> _
Friend Class ColumnReference
	Inherits RowColumnReference

	Private Shared OurRegex As System.Text.RegularExpressions.Regex = CreateRegex()

	Public Sub New(ByVal start As Integer, ByVal finish As Integer)
		MyBase.New(start, finish)
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Sub ValidateIndex(ByVal index As Integer)
		SheetReference.ValidateColumnIndex(index)
	End Sub

	Public Shared Function IsValidString(ByVal s As String) As Boolean
		Return OurRegex.IsMatch(s)
	End Function

	Public Shared Function FromString(ByVal image As String) As ColumnReference
		image = SheetReference.PrepareParseString(image)
		Dim parts As String() = image.Split(":")
		Dim start, finish As Integer

		start = PartToColumnIndex(parts(0))
		finish = PartToColumnIndex(parts(1))

		Return New ColumnReference(start, finish)
	End Function

	Private Shared Function PartToColumnIndex(ByVal part As String) As Integer
		If part.Length = 1 Then
			Return ColumnLabel2Index(part.Chars(0))
		Else
			Return ColumnLabel2Index(part.Chars(0), part.Chars(1))
		End If
	End Function

	Private Shared Function CreateRegex() As System.Text.RegularExpressions.Regex
		Dim exp As String = String.Format("^{0}{1}:{1}$", SheetReference.SHEET_REGEX, SheetReference.COLUMN_REGEX)
		Return New System.Text.RegularExpressions.Regex(exp)
	End Function

	Protected Overrides Function GetCopyOffset(ByVal rowOffset As Integer, ByVal colOffset As Integer) As Integer
		Return colOffset
	End Function

	Protected Overloads Overrides Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer, ByVal ops As RowColumnReference.RowColumnGridOps) As ReferenceOperationResultType
		Return ops.OnRowColumnInsert(insertAt, count)
	End Function

	Protected Overloads Overrides Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer, ByVal ops As RowColumnReference.RowColumnGridOps) As ReferenceOperationResultType
		Return ops.OnRowColumnDelete(removeAt, count)
	End Function

	Protected Overloads Overrides Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer, ByVal ops As RowColumnReference.RowColumnGridOps) As ReferenceOperationResultType
		Return ReferenceOperationResultType.Affected
	End Function

	Protected Overloads Overrides Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer, ByVal ops As RowColumnReference.RowColumnGridOps) As ReferenceOperationResultType
		Return ReferenceOperationResultType.Affected
	End Function

	Protected Overrides Function FormatIndex(ByVal index As Integer) As String
		Return ColumnIndex2Label(index)
	End Function

	Protected Overrides Function GetStartIndex(ByVal rect As System.Drawing.Rectangle) As Integer
		Return rect.Left
	End Function

	Protected Overrides Function GetFinishIndex(ByVal rect As System.Drawing.Rectangle) As Integer
		Return rect.Right
	End Function

	Public Overrides ReadOnly Property Range() As System.Drawing.Rectangle
		Get
			Dim h, w As Integer
			h = Me.Sheet.RowCount
			w = Me.Finish - Me.Start + 1
			Return New System.Drawing.Rectangle(Me.Start, 1, w, h)
		End Get
	End Property
End Class

<Serializable()> _
Friend Class RowReference
	Inherits RowColumnReference

	Private Shared OurRegex As System.Text.RegularExpressions.Regex = CreateRegex()

	Public Sub New(ByVal start As Integer, ByVal finish As Integer)
		MyBase.New(start, finish)
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Sub ValidateIndex(ByVal index As Integer)
		SheetReference.ValidateRowIndex(index)
	End Sub

	Public Shared Function IsValidString(ByVal s As String) As Boolean
		Return OurRegex.IsMatch(s)
	End Function

	Public Shared Function FromString(ByVal image As String) As RowReference
		image = SheetReference.PrepareParseString(image)
		Dim parts As String() = image.Split(":")
		Dim start, finish As Integer

		start = Integer.Parse(parts(0))
		finish = Integer.Parse(parts(1))

		Return New RowReference(start, finish)
	End Function

	Private Shared Function CreateRegex() As System.Text.RegularExpressions.Regex
		Dim exp As String = String.Format("^{0}{1}:{1}$", SheetReference.SHEET_REGEX, SheetReference.ROW_REGEX)
		Return New System.Text.RegularExpressions.Regex(exp)
	End Function

	Protected Overrides Function OnColumnsInserted(ByVal insertAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType
		Return ReferenceOperationResultType.Affected
	End Function

	Protected Overrides Function OnColumnsRemoved(ByVal removeAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType
		Return ReferenceOperationResultType.Affected
	End Function

	Protected Overrides Function GetStartIndex(ByVal rect As System.Drawing.Rectangle) As Integer
		Return rect.Top
	End Function

	Protected Overrides Function GetFinishIndex(ByVal rect As System.Drawing.Rectangle) As Integer
		Return rect.Bottom
	End Function

	Protected Overrides Function OnRowsInserted(ByVal insertAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType
		Return ops.OnRowColumnInsert(insertAt, count)
	End Function

	Protected Overrides Function OnRowsRemoved(ByVal removeAt As Integer, ByVal count As Integer, ByVal ops As RowColumnGridOps) As ReferenceOperationResultType
		Return ops.OnRowColumnDelete(removeAt, count)
	End Function

	Protected Overrides Function FormatIndex(ByVal index As Integer) As String
		Return index.ToString()
	End Function

	Protected Overrides Function GetCopyOffset(ByVal rowOffset As Integer, ByVal colOffset As Integer) As Integer
		Return rowOffset
	End Function

	Public Overrides ReadOnly Property Range() As System.Drawing.Rectangle
		Get
			Dim h, w As Integer
			h = Me.Finish - Me.Start + 1
			w = Me.Sheet.ColumnCount
			Return New System.Drawing.Rectangle(1, Me.Start, w, h)
		End Get
	End Property
End Class

''' <summary>
''' Base class for references that aren't on a sheet
''' </summary>
<Serializable()> _
Friend MustInherit Class NonGridReference
	Inherits Reference

	Protected Sub New()

	End Sub

	Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
	End Sub

	Protected Overrides Function FormatWithProps(ByVal props As ReferenceProperties) As String
		Debug.Assert(False, "should not be called")
		Return Nothing
	End Function

	Public Overrides Function Intersects(ByVal ref As Reference) As Boolean
		Return Me.EqualsReference(ref)
	End Function

	Public Overrides Function IsOnSheet(ByVal sheet As ISheet) As Boolean
		Return False
	End Function

	Public Overrides Function IsReferenceEqualForCircularReference(ByVal ref As Reference) As Boolean
		Return Me.EqualsReference(ref)
	End Function

	Public Overrides Sub OnCopy(ByVal rowOffset As Integer, ByVal colOffset As Integer, ByVal destSheet As ISheet, ByVal props As ReferenceProperties)

	End Sub

	Protected Overrides Function CreateGridOps() As GridOperationsBase
		Return New NullGridOps
	End Function
End Class

<Serializable()> _
Friend Class NamedReference
	Inherits NonGridReference
	Implements IFormulaSelfReference
	Implements INamedReference

	Private MyName As String
	Private Shared OurRegex As System.Text.RegularExpressions.Regex = CreateRegex()
	Private MyValueOperand As IOperand

	Public Sub New(ByVal name As String)
		MyName = name
		MyValueOperand = New NullValueOperand()
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyName = info.GetString("Name")
		MyValueOperand = info.GetValue("ValueOperand", GetType(IOperand))
		Me.ComputeHashCode()
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Name", MyName)
		info.AddValue("ValueOperand", MyValueOperand)
	End Sub

	Public Shared Function IsValidName(ByVal name As String) As Boolean
		Return OurRegex.IsMatch(name)
	End Function

	Private Shared Function CreateRegex() As System.Text.RegularExpressions.Regex
		Dim exp As String = "^[_A-Za-z][_0-9A-Za-z]*$"
		Return New System.Text.RegularExpressions.Regex(exp)
	End Function

	Protected Overrides Function EqualsReferenceInternal(ByVal ref As Reference) As Boolean
		Dim realRef As NamedReference = ref
		Return String.Equals(MyName, realRef.MyName, StringComparison.OrdinalIgnoreCase)
	End Function

	Protected Overrides Function Format() As String
		Return MyName
	End Function

	Public Sub OnFormulaRecalculate(ByVal target As Formula) Implements IFormulaSelfReference.OnFormulaRecalculate
		MyValueOperand = target.EvaluateToOperand()
		MyBase.OnRecalculated()
	End Sub

	Protected Overrides Function GetHashData() As Byte()
		Dim nameLower As String = MyName.ToLowerInvariant()
		Dim nameBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(nameLower)

		Dim bytes(REFERENCE_HASH_SIZE + nameBytes.Length - 1) As Byte
		MyBase.GetBaseHashData(bytes)
		System.Array.Copy(nameBytes, 0, bytes, REFERENCE_HASH_SIZE, nameBytes.Length)
		Return bytes
	End Function

	Public Property OperandValue() As Object
		Get
			Return MyValueOperand.Value
		End Get
		Set(ByVal value As Object)
			MyValueOperand = OperandFactory.CreateDynamic(value)
		End Set
	End Property

	Public ReadOnly Property ValueOperand() As IOperand
		Get
			Return MyValueOperand
		End Get
	End Property

	Public ReadOnly Property Name() As String Implements INamedReference.Name
		Get
			Return MyName
		End Get
	End Property

	Public ReadOnly Property Result() As Object Implements INamedReference.Result
		Get
			Return Me.OperandValue
		End Get
	End Property
End Class

Friend Class VolatileFunctionReference
	Inherits NonGridReference

	Public Sub New()
		Me.ComputeHashCode()
	End Sub

	Protected Overrides Function EqualsReferenceInternal(ByVal ref As Reference) As Boolean
		Return ref.GetType() Is GetType(VolatileFunctionReference)
	End Function

	Protected Overrides Function Format() As String
		Return "VolatileFunction"
	End Function

	Protected Overrides Function GetHashData() As Byte()
		Dim typeHashCode As Integer = GetType(VolatileFunctionReference).GetHashCode()
		Return BitConverter.GetBytes(typeHashCode)
	End Function
End Class

<Serializable()> _
Friend Class ExternalReference
	Inherits NonGridReference
	Implements IFormulaSelfReference
	Implements IExternalReference

	Private MyResult As Object

	Public Sub New()

	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyResult = info.GetValue("Result", GetType(Object))
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("Result", MyResult)
	End Sub

	Protected Overrides Function EqualsReferenceInternal(ByVal ref As Reference) As Boolean
		Return ref Is Me
	End Function

	Protected Overrides Function Format() As String
		Dim hashCode As Integer = Me.GetHashCode()
		Return String.Format("ExternalRef_{0:x}", hashCode)
	End Function

	Protected Overrides Function ComputeHashCodeInternal() As Integer
		Return Me.GetHashCode()
	End Function

	Protected Overrides Function GetHashData() As Byte()
		Return New Byte() {}
	End Function

	Public Sub OnFormulaRecalculate(ByVal target As Formula) Implements IFormulaSelfReference.OnFormulaRecalculate
		MyResult = target.Evaluate()
		MyBase.OnRecalculated()
	End Sub

	Public ReadOnly Property Result() As Object Implements IExternalReference.Result
		Get
			Return MyResult
		End Get
	End Property
End Class

<Serializable()> _
Friend Class ReferenceList
	Inherits NonGridReference

	Private MyReferences As Reference()

	Public Sub New(ByVal references As IReference())
		Dim arr(references.Length - 1) As Reference
		For i As Integer = 0 To arr.Length - 1
			arr(i) = references(i)
		Next
		MyReferences = arr
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.New(info, context)
		MyReferences = info.GetValue("References", GetType(IReference()))
	End Sub

	Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyBase.GetObjectData(info, context)
		info.AddValue("References", MyReferences)
	End Sub

	Protected Overrides Sub OnEngineSet(ByVal engine As FormulaEngine)
		MyBase.OnEngineSet(engine)
		For Each ref As Reference In MyReferences
			ref.SetEngine(engine)
		Next
	End Sub

	Protected Overrides Function EqualsReferenceInternal(ByVal ref As Reference) As Boolean
		Dim rhsListRef As ReferenceList = ref
		If MyReferences.Length <> rhsListRef.MyReferences.Length Then
			Return False
		End If

		For i As Integer = 0 To MyReferences.Length - 1
			If MyReferences(i).EqualsReference(rhsListRef.MyReferences(i)) = False Then
				Return False
			End If
		Next

		Return True
	End Function

	Public Overrides Function Intersects(ByVal ref As Reference) As Boolean
		For Each listRef As Reference In MyReferences
			If listRef.Intersects(ref) = True Then
				Return True
			End If
		Next
		Return False
	End Function

	Protected Overrides Sub GetReferenceValuesInternal(ByVal processor As IReferenceValueProcessor)
		For Each ref As Reference In MyReferences
			ref.GetReferenceValues(processor)
		Next
	End Sub

	Protected Overrides Function Format() As String
		Dim arr(MyReferences.Length - 1) As String

		For i As Integer = 0 To arr.Length - 1
			arr(i) = MyReferences(i).ToString()
		Next

		Return String.Join(",", arr)
	End Function

	Protected Overrides Function GetHashData() As Byte()
		Dim bytes(Reference.REFERENCE_HASH_SIZE - 1) As Byte
		MyBase.GetBaseHashData(bytes)
		Return bytes
	End Function
End Class