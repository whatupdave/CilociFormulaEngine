<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChartForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Me.Label1 = New System.Windows.Forms.Label
		Me.Panel1 = New System.Windows.Forms.Panel
		Me.lblSeries2Values = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.lblSeries2Name = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.lblSeries1Values = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.lblSeries1Name = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.lblXaxisFormula = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.ZedGraphControl1 = New ZedGraph.ZedGraphControl
		Me.Panel1.SuspendLayout()
		Me.SuspendLayout()
		'
		'Label1
		'
		Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(526, 36)
		Me.Label1.TabIndex = 1
		Me.Label1.Text = "This chart is linked to the sample values you see on the grid.  Try changing the " & _
			"values on the grid to see the changes reflected in the chart."
		'
		'Panel1
		'
		Me.Panel1.Controls.Add(Me.lblSeries2Values)
		Me.Panel1.Controls.Add(Me.Label6)
		Me.Panel1.Controls.Add(Me.lblSeries2Name)
		Me.Panel1.Controls.Add(Me.Label8)
		Me.Panel1.Controls.Add(Me.lblSeries1Values)
		Me.Panel1.Controls.Add(Me.Label4)
		Me.Panel1.Controls.Add(Me.lblSeries1Name)
		Me.Panel1.Controls.Add(Me.Label3)
		Me.Panel1.Controls.Add(Me.lblXaxisFormula)
		Me.Panel1.Controls.Add(Me.Label2)
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
		Me.Panel1.Location = New System.Drawing.Point(0, 36)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(526, 33)
		Me.Panel1.TabIndex = 2
		'
		'lblSeries2Values
		'
		Me.lblSeries2Values.AutoSize = True
		Me.lblSeries2Values.BackColor = System.Drawing.Color.White
		Me.lblSeries2Values.Location = New System.Drawing.Point(428, 13)
		Me.lblSeries2Values.Name = "lblSeries2Values"
		Me.lblSeries2Values.Size = New System.Drawing.Size(39, 13)
		Me.lblSeries2Values.TabIndex = 9
		Me.lblSeries2Values.Text = "Label5"
		'
		'Label6
		'
		Me.Label6.AutoSize = True
		Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Location = New System.Drawing.Point(323, 13)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(99, 13)
		Me.Label6.TabIndex = 8
		Me.Label6.Text = "Series 2 Values:"
		'
		'lblSeries2Name
		'
		Me.lblSeries2Name.AutoSize = True
		Me.lblSeries2Name.BackColor = System.Drawing.Color.White
		Me.lblSeries2Name.Location = New System.Drawing.Point(428, 0)
		Me.lblSeries2Name.Name = "lblSeries2Name"
		Me.lblSeries2Name.Size = New System.Drawing.Size(39, 13)
		Me.lblSeries2Name.TabIndex = 7
		Me.lblSeries2Name.Text = "Label4"
		'
		'Label8
		'
		Me.Label8.AutoSize = True
		Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.Location = New System.Drawing.Point(323, 0)
		Me.Label8.Name = "Label8"
		Me.Label8.Size = New System.Drawing.Size(93, 13)
		Me.Label8.TabIndex = 6
		Me.Label8.Text = "Series 2 Name:"
		'
		'lblSeries1Values
		'
		Me.lblSeries1Values.AutoSize = True
		Me.lblSeries1Values.BackColor = System.Drawing.Color.White
		Me.lblSeries1Values.Location = New System.Drawing.Point(220, 13)
		Me.lblSeries1Values.Name = "lblSeries1Values"
		Me.lblSeries1Values.Size = New System.Drawing.Size(39, 13)
		Me.lblSeries1Values.TabIndex = 5
		Me.lblSeries1Values.Text = "Label5"
		'
		'Label4
		'
		Me.Label4.AutoSize = True
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(115, 13)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(99, 13)
		Me.Label4.TabIndex = 4
		Me.Label4.Text = "Series 1 Values:"
		'
		'lblSeries1Name
		'
		Me.lblSeries1Name.AutoSize = True
		Me.lblSeries1Name.BackColor = System.Drawing.Color.White
		Me.lblSeries1Name.Location = New System.Drawing.Point(220, 0)
		Me.lblSeries1Name.Name = "lblSeries1Name"
		Me.lblSeries1Name.Size = New System.Drawing.Size(39, 13)
		Me.lblSeries1Name.TabIndex = 3
		Me.lblSeries1Name.Text = "Label4"
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(115, 0)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(93, 13)
		Me.Label3.TabIndex = 2
		Me.Label3.Text = "Series 1 Name:"
		'
		'lblXaxisFormula
		'
		Me.lblXaxisFormula.AutoSize = True
		Me.lblXaxisFormula.BackColor = System.Drawing.Color.White
		Me.lblXaxisFormula.Location = New System.Drawing.Point(3, 13)
		Me.lblXaxisFormula.Name = "lblXaxisFormula"
		Me.lblXaxisFormula.Size = New System.Drawing.Size(39, 13)
		Me.lblXaxisFormula.TabIndex = 1
		Me.lblXaxisFormula.Text = "Label3"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(3, 0)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(82, 13)
		Me.Label2.TabIndex = 0
		Me.Label2.Text = "X axis labels:"
		'
		'ZedGraphControl1
		'
		Me.ZedGraphControl1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.ZedGraphControl1.Location = New System.Drawing.Point(0, 69)
		Me.ZedGraphControl1.Name = "ZedGraphControl1"
		Me.ZedGraphControl1.ScrollMaxX = 0
		Me.ZedGraphControl1.ScrollMaxY = 0
		Me.ZedGraphControl1.ScrollMaxY2 = 0
		Me.ZedGraphControl1.ScrollMinX = 0
		Me.ZedGraphControl1.ScrollMinY = 0
		Me.ZedGraphControl1.ScrollMinY2 = 0
		Me.ZedGraphControl1.Size = New System.Drawing.Size(526, 334)
		Me.ZedGraphControl1.TabIndex = 10
		'
		'ChartForm
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(526, 403)
		Me.Controls.Add(Me.ZedGraphControl1)
		Me.Controls.Add(Me.Panel1)
		Me.Controls.Add(Me.Label1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "ChartForm"
		Me.ShowInTaskbar = False
		Me.Text = "Chart demo"
		Me.Panel1.ResumeLayout(False)
		Me.Panel1.PerformLayout()
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents lblSeries1Values As System.Windows.Forms.Label
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents lblSeries1Name As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents lblXaxisFormula As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents lblSeries2Values As System.Windows.Forms.Label
	Friend WithEvents Label6 As System.Windows.Forms.Label
	Friend WithEvents lblSeries2Name As System.Windows.Forms.Label
	Friend WithEvents Label8 As System.Windows.Forms.Label
	Friend WithEvents ZedGraphControl1 As ZedGraph.ZedGraphControl
End Class
