<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChild
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChild))
		Me.imlToolbar = New System.Windows.Forms.ImageList(Me.components)
		Me.tlbGeometry = New DGApp.DGToolbar
		Me.SuspendLayout()
		'
		'imlToolbar
		'
		Me.imlToolbar.ImageStream = CType(resources.GetObject("imlToolbar.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.imlToolbar.TransparentColor = System.Drawing.Color.Magenta
		Me.imlToolbar.Images.SetKeyName(0, "DGPointer.bmp")
		Me.imlToolbar.Images.SetKeyName(1, "DGPoint.bmp")
		Me.imlToolbar.Images.SetKeyName(2, "DGSegment.bmp")
		Me.imlToolbar.Images.SetKeyName(3, "DGRay.bmp")
		Me.imlToolbar.Images.SetKeyName(4, "DGLine.bmp")
		Me.imlToolbar.Images.SetKeyName(5, "DGLineParallel.bmp")
		Me.imlToolbar.Images.SetKeyName(6, "DGLinePerpendicular.bmp")
		Me.imlToolbar.Images.SetKeyName(7, "DGBisector.bmp")
		Me.imlToolbar.Images.SetKeyName(8, "DGCircle.bmp")
		Me.imlToolbar.Images.SetKeyName(9, "DGCircleByRadius.bmp")
		Me.imlToolbar.Images.SetKeyName(10, "DGArc.bmp")
		Me.imlToolbar.Images.SetKeyName(11, "DGMidPoint.bmp")
		'
		'tlbGeometry
		'
		Me.tlbGeometry.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
		Me.tlbGeometry.DropDownArrows = True
		Me.tlbGeometry.ImageList = Me.imlToolbar
		Me.tlbGeometry.Location = New System.Drawing.Point(0, 0)
		Me.tlbGeometry.Name = "tlbGeometry"
		Me.tlbGeometry.ShowToolTips = True
		Me.tlbGeometry.Size = New System.Drawing.Size(292, 42)
		Me.tlbGeometry.TabIndex = 0
		'
		'frmChild
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(292, 268)
		Me.Controls.Add(Me.tlbGeometry)
		Me.Name = "frmChild"
		Me.Text = "Untitled"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents tlbGeometry As DGApp.DGToolbar
	Friend WithEvents imlToolbar As System.Windows.Forms.ImageList
End Class
