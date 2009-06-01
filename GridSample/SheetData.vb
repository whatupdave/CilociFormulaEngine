' Stores the contents of a range of cells from a sheet
<Serializable()> _
Friend Class SheetData

	' Loads data from a range of cells into a table
	Private Class LoadDataTransform
		Inherits CellTransform

		Private MyData As Object(,)
		Private MyStart As Point

		Public Sub New(ByVal data As Object(,), ByVal start As Point)
			MyData = data
			MyStart = start
		End Sub

		Public Overrides Sub TransformCell(ByVal target As ExcelLikeCell, ByVal row As Integer, ByVal col As Integer)
			MyData(row - MyStart.Y, col - MyStart.X) = target.Value
		End Sub
	End Class

	' Writes data from a table into a range of cells
	Private Class WriteDataTransform
		Inherits CellTransform

		Private MyData As Object(,)
		Private MyStart As Point

		Public Sub New(ByVal data As Object(,), ByVal start As Point)
			MyData = data
			MyStart = start
		End Sub

		Public Overrides Sub TransformCell(ByVal target As ExcelLikeCell, ByVal row As Integer, ByVal col As Integer)
			target.Value = MyData(row - MyStart.Y, col - MyStart.X)
		End Sub
	End Class


	Private MyData As Object(,)				' The actual table of cell values
	<NonSerialized()> _
	Private MySourceRange As Rectangle		' The range where the data was loaded from
	<NonSerialized()> _
	Private MySourceSheet As Worksheet		' The sheet where the data was loaded from

	Public Sub New()

	End Sub

	' Loads the values of a range of cells from a worksheet
	Public Sub LoadData(ByVal source As Worksheet, ByVal sourceRange As Rectangle)
		MySourceSheet = source
		MyData = New Object(sourceRange.Height - 1, sourceRange.Width - 1) {}
		MySourceRange = sourceRange
		source.Transform(sourceRange, New LoadDataTransform(MyData, sourceRange.Location))
	End Sub

	' Writes the data to a range of cells on a worksheet
	Public Sub WriteData(ByVal target As Worksheet, ByVal range As Rectangle)
		target.Transform(range, New WriteDataTransform(MyData, range.Location))
	End Sub

	Public ReadOnly Property SourceRange() As Rectangle
		Get
			Return MySourceRange
		End Get
	End Property

	Public ReadOnly Property SourceSheet() As Worksheet
		Get
			Return MySourceSheet
		End Get
	End Property

	Public ReadOnly Property DataSize() As System.Drawing.Size
		Get
			Return New Size(MyData.GetLength(1), MyData.GetLength(0))
		End Get
	End Property
End Class