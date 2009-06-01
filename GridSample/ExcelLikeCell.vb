Imports ciloci.FormulaEngine

' Implements a custom datagridviewcell that behaves like excel
Friend Class ExcelLikeCell
	Inherits DataGridViewTextBoxCell

	Private MyOwner As MainForm

	Public Sub New()

	End Sub

	Public Sub New(ByVal owner As MainForm)
		MyOwner = owner
	End Sub

	Public Overrides Function Clone() As Object
		Dim copy As ExcelLikeCell = MyBase.Clone()
		copy.MyOwner = MyOwner
		Return copy
	End Function

	' Get a sheet reference to this cell
	Private Function GetCellReference(ByVal rowIndex As Integer) As ISheetReference
		Return Me.Engine.ReferenceFactory.Cell(rowIndex + 1, Me.ColumnIndex + 1)
	End Function

	Protected Overrides Function GetFormattedValue(ByVal value As Object, ByVal rowIndex As Integer, ByRef cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal valueTypeConverter As System.ComponentModel.TypeConverter, ByVal formattedValueTypeConverter As System.ComponentModel.TypeConverter, ByVal context As System.Windows.Forms.DataGridViewDataErrorContexts) As Object
		If TypeOf (value) Is DateTime Then
			' Try to display a shorter format if time is zero
			Dim dt As DateTime = DirectCast(value, DateTime)
			If dt.Hour = 0 And dt.Minute = 0 And dt.Second = 0 Then
				Return dt.ToShortDateString()
			Else
				Return dt.ToString()
			End If
		Else
			Return MyBase.GetFormattedValue(value, rowIndex, cellStyle, valueTypeConverter, formattedValueTypeConverter, context)
		End If
	End Function

	Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, ByVal initialFormattedValue As Object, ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle)
		MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)
		' Get a reference to this cell
		Dim ref As ISheetReference = Me.GetCellReference(rowIndex)
		Dim f As Formula = Me.Engine.GetFormulaAt(ref)

		' Is there a formula here?
		If Not f Is Nothing Then
			' If so then initialize the editing control with its text
			Me.DataGridView.EditingControl.Text = f.ToString()
		End If
	End Sub

	' Set the value of the cell from a string
	Public Function SetTextValue(ByVal text As String) As Boolean
		' Get a reference to our cell
		Dim ref As ISheetReference = Me.GetCellReference(Me.RowIndex)
		Dim success As Boolean = True

		If text Is Nothing Then
			' Clear our value
			MyBase.SetValue(Me.RowIndex, Nothing)
		ElseIf text.StartsWith("=") Then
			' Add a formula at our cell
			success = Me.AddFormula(text, ref)
		Else
			' Parse the text into a value and store it in our cell
			Dim value As Object = ciloci.FormulaEngine.Utility.Parse(text)
			' Make sure we remove any formula at this cell
			Me.Engine.RemoveFormulasInRange(ref)
			' Store the value
			MyBase.SetValue(Me.RowIndex, value)
		End If

		If success = True Then
			' Have the engine recalculate this cell
			Me.Engine.Recalculate(ref)
		End If

		Return success
	End Function

	' Add a formula at this cell
	Private Function AddFormula(ByVal expression As String, ByVal ref As ISheetReference) As Boolean
		' Try to create the formula
		Dim f As Formula = MyOwner.CreateFormula(expression)
		If f Is Nothing Then
			' Invalid formula
			Return False
		Else
			' Valid formula
			' Remove any previous formula
			Me.Engine.RemoveFormulasInRange(ref)
			' Set the new formula at this cell
			Me.Engine.AddFormula(f, ref)
			Return True
		End If
	End Function

	Private ReadOnly Property Engine() As ciloci.FormulaEngine.FormulaEngine
		Get
			Return MyOwner.Engine
		End Get
	End Property

	Private ReadOnly Property Owner() As MainForm
		Get
			Return MyOwner
		End Get
	End Property
End Class