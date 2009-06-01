Imports ciloci.FormulaEngine

' A demo of external references using a chart.  It demonstrates how you can create a formula that references
' sheet data but that is not itself on a sheet.  By creating a formula that evaluates to a reference, you can effectively
' 'watch' values on a sheet and be notified when the reference to those values changes.  This form sets up 5 formulas
' that evaluate to references to the sample data on the sheet.  As the engine recalculates, we are notified and change
' the chart accordingly.
Friend Class ChartForm

	' A wrapper around a chart series
	Private Class CustomSeries

		Private MyOwner As ChartForm
		Private MyNameReference As IExternalReference			' The external reference to our series name
		Private MyValuesReference As IExternalReference			' The external reference to our series values
		Private MySeries As ZedGraph.BarItem
		Private MyNameLabel As Label
		Private MyValuesLabel As Label

		Public Sub New(ByVal owner As ChartForm, ByVal nameLabel As Label, ByVal valuesLabel As Label)
			MyOwner = owner
			MyNameLabel = nameLabel
			MyValuesLabel = valuesLabel

			Me.SetupReferences()

			MySeries = New ZedGraph.BarItem("")
			owner.ZedGraphControl1.GraphPane.CurveList.Add(MySeries)
		End Sub

		Private Sub SetupReferences()
			' Create two external references.  We have to keep track of the returned references
			' since we'll need them to refer to their formulas later
			MyNameReference = MyOwner.Engine.ReferenceFactory.External()
			MyValuesReference = MyOwner.Engine.ReferenceFactory.External()
			' Listen for when the formulas recalculate
			AddHandler MyNameReference.Recalculated, AddressOf OnNameReferenceRecalculated
			AddHandler MyValuesReference.Recalculated, AddressOf OnValuesReferenceRecalculated
		End Sub

		Public Sub Initialize(ByVal column As Integer, ByVal seriesColor As Color)
			' Create a reference to cell that holds our series' name
			Dim ref As ISheetReference = MyOwner.Engine.ReferenceFactory.Cell(1, column)
			' Create a formula that evaluates to that reference
			Dim f As Formula = MyOwner.Engine.CreateFormula(String.Format("={0}", ref.ToString()))
			' Add the formula to the engine using our external reference
			MyOwner.Engine.AddFormula(f, MyNameReference)

			' Create a reference to our series values
			ref = MyOwner.Engine.ReferenceFactory.Cells(2, column, 4, column)
			' Create a formula that evaluates to that reference
			f = MyOwner.Engine.CreateFormula(String.Format("={0}", ref.ToString()))
			' Add the formula to the engine using our values external reference
			MyOwner.Engine.AddFormula(f, MyValuesReference)

			MySeries.Bar.Fill.Color = seriesColor
		End Sub

		' Called by the formula engine when our reference is recalculated.  This could be for several reasons:
		' The reference was moved, or invalidated, or values inside of it have changed.
		Private Sub OnNameReferenceRecalculated(ByVal sender As Object, ByVal e As EventArgs)
			' Get the formula
			Dim f As Formula = MyOwner.Engine.GetFormulaAt(MyNameReference)
			' Get the formula's first reference.  This will be a reference to our series name
			Dim ref As ISheetReference = f.References(0)
			' Update our label
			MyNameLabel.Text = f.ToString()

			' Has the reference been invalidated?
			If MyOwner.Engine.Info.IsReferenceValid(ref) = False Then
				' Update our text with the error value
				MySeries.Label.Text = f.ToString()
				MyOwner.InvalidateChart()
				' Exit.  The engine will not call this handler again.
				Return
			End If

			' Get the reference's values and update our series
			Dim table As Object(,) = ref.GetValuesTable()
			Dim label As String = String.Empty

			Dim value As Object = table(0, 0)
			If Not value Is Nothing Then
				label = value.ToString()
			End If

			MySeries.Label.Text = label
			MyOwner.InvalidateChart()
		End Sub

		Private Sub OnValuesReferenceRecalculated(ByVal sender As Object, ByVal e As EventArgs)
			' Get the formula at our values reference
			Dim f As Formula = MyOwner.Engine.GetFormulaAt(MyValuesReference)
			' Get the range we are watching
			Dim range As ISheetReference = f.References(0)
			MyValuesLabel.Text = f.ToString()

			' Is the reference invalid?
			If MyOwner.Engine.Info.IsReferenceValid(range) = False Then
				' Clear our series and quit
				MySeries.Clear()
				MyOwner.InvalidateChart()
				Return
			End If

			' Get the reference's values and update our series with it
			Dim table As Object(,) = range.GetValuesTable()
			Me.SetPointCount(table.GetLength(0))
			Me.SetValues(table)
			MyOwner.ZedGraphControl1.AxisChange()
			MyOwner.InvalidateChart()
		End Sub

		Private Sub SetValues(ByVal values As Object(,))
			Dim list As New ZedGraph.PointPairList()

			For i As Integer = 0 To values.GetLength(0) - 1
				Dim value As Object = values(i, 0)
				value = Utility.NormalizeIfNumericValue(value)

				If TypeOf (value) Is Double Then
					list.Add(i, DirectCast(value, Double))
				Else
					list.Add(i, 0)
				End If
			Next

			MySeries.Points = list
		End Sub

		Private Sub SetPointCount(ByVal count As Integer)
			Dim currentCount As Integer = MySeries.Points.Count
			If currentCount < count Then
				For i As Integer = currentCount To count - 1
					MySeries.AddPoint(0, 0)
				Next
			ElseIf currentCount > count Then
				For i As Integer = count To currentCount - 1
					MySeries.RemovePoint(count)
				Next
			End If
		End Sub

		Public Sub Dispose()
			' Stop listening for recalculation
			RemoveHandler MyNameReference.Recalculated, AddressOf OnNameReferenceRecalculated
			RemoveHandler MyValuesReference.Recalculated, AddressOf OnValuesReferenceRecalculated
			' Make sure we clean up by having the formula engine remove the formulas at our references.  The engine
			' will also stop tracking our external references since it will figure out that nothing is using them.
			MyOwner.Engine.RemoveFormulaAt(MyNameReference)
			MyOwner.Engine.RemoveFormulaAt(MyValuesReference)
		End Sub
	End Class

	Private MyXAxisValuesReference As IExternalReference		' The external reference to our X-Axis values
	Private MySeries As CustomSeries()
	Private WithEvents MyDocumentStatusService As DocumentStatusService

	Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
		MyBase.OnLoad(e)
		Me.EnableMainFormCommmands(False)
		MyDocumentStatusService = Me.Site.GetService(GetType(DocumentStatusService))

		Me.SetupChart()
		
		Dim arr(2 - 1) As CustomSeries
		arr(0) = New CustomSeries(Me, Me.lblSeries1Name, Me.lblSeries1Values)
		arr(1) = New CustomSeries(Me, Me.lblSeries2Name, Me.lblSeries2Values)
		MySeries = arr

		' Get an external reference for our X-Axis values
		MyXAxisValuesReference = Me.Engine.ReferenceFactory.External()
		' Listen for recalculation
		AddHandler MyXAxisValuesReference.Recalculated, AddressOf OnXAxisValuesRecalculated
		' Get a reference to the cells that will be our x-axis values
		Dim ref As ISheetReference = Me.Engine.ReferenceFactory.Cells(2, 1, 4, 1)
		' Create a formula that evaluates to that reference
		Dim f As Formula = Me.Engine.CreateFormula("=" & ref.ToString())
		' Add it to the engine
		Me.Engine.AddFormula(f, MyXAxisValuesReference)

		arr(0).Initialize(2, Color.DodgerBlue)
		arr(1).Initialize(3, Color.Gold)
	End Sub

	Private Sub EnableMainFormCommmands(ByVal enabled As Boolean)
		Dim mf As MainForm = Me.Site.GetService(GetType(MainForm))
		mf.SaveAsToolStripMenuItem.Enabled = enabled
		mf.SaveToolStripMenuItem.Enabled = enabled
		mf.OpenToolStripMenuItem.Enabled = enabled
		mf.NewToolStripMenuItem.Enabled = enabled
		mf.ChartToolStripMenuItem.Enabled = enabled
	End Sub

	Private Sub SetupChart()
		Me.ZedGraphControl1.GraphPane.Title.Text = "Chart Demo"
		Me.ZedGraphControl1.GraphPane.XAxis.Title.Text = "Month"
		Me.ZedGraphControl1.GraphPane.XAxis.Type = ZedGraph.AxisType.Text
		Me.ZedGraphControl1.GraphPane.YAxis.Title.Text = "Sales"
	End Sub

	Private Sub InvalidateChart()
		Me.ZedGraphControl1.Invalidate()
	End Sub

	Protected Overrides Sub OnClosed(ByVal e As System.EventArgs)
		For Each series As CustomSeries In MySeries
			series.Dispose()
		Next
		' Remove our formula from the engine
		Me.Engine.RemoveFormulaAt(MyXAxisValuesReference)
		' Stop listening
		RemoveHandler MyXAxisValuesReference.Recalculated, AddressOf OnXAxisValuesRecalculated
		MyDocumentStatusService = Nothing
		Me.EnableMainFormCommmands(True)
		MyBase.OnClosed(e)
	End Sub

	Private Sub MyDocumentStatusService_NewOrOpenDocument(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyDocumentStatusService.NewDocument, MyDocumentStatusService.OpenDocument
		Me.Close()
	End Sub

	' Called by the engine when our formula has been recalculated
	Private Sub OnXAxisValuesRecalculated(ByVal sender As Object, ByVal e As EventArgs)
		' Get the formula at our reference
		Dim f As Formula = Me.Engine.GetFormulaAt(MyXAxisValuesReference)
		' Get its first reference
		Dim ref As ISheetReference = f.References(0)

		Me.lblXaxisFormula.Text = f.ToString()

		If Me.Engine.Info.IsReferenceValid(ref) = False Then
			' Our reference has been invalidated so set our x-axis labels to an empty array
			Me.ZedGraphControl1.GraphPane.XAxis.Scale.TextLabels = Nothing
			Me.InvalidateChart()
		Else
			' Update our labels
			Me.SetXAxisLabels(ref)
		End If
	End Sub

	Private Sub SetXAxisLabels(ByVal ref As ISheetReference)
		Dim table As Object(,) = ref.GetValuesTable()

		Dim arr(table.GetLength(0) - 1) As String

		For i As Integer = 0 To arr.Length - 1
			Dim value As Object = table(i, 0)
			Dim label As String = String.Empty
			If Not value Is Nothing Then
				label = value.ToString()
			End If
			arr(i) = label
		Next

		Me.ZedGraphControl1.GraphPane.XAxis.Scale.TextLabels = arr
		Me.InvalidateChart()
	End Sub

	Private ReadOnly Property Engine() As FormulaEngine
		Get
			Return Me.Site.GetService(GetType(FormulaEngine))
		End Get
	End Property
End Class