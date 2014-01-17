<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmRender
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
        Me.NesScreen = New System.Windows.Forms.PictureBox()
        CType(Me.NesScreen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NesScreen
        '
        Me.NesScreen.BackColor = System.Drawing.Color.Black
        Me.NesScreen.Location = New System.Drawing.Point(26, 23)
        Me.NesScreen.Margin = New System.Windows.Forms.Padding(0)
        Me.NesScreen.Name = "NesScreen"
        Me.NesScreen.Size = New System.Drawing.Size(256, 240)
        Me.NesScreen.TabIndex = 1
        Me.NesScreen.TabStop = False
        '
        'FrmRender
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(640, 480)
        Me.Controls.Add(Me.NesScreen)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FrmRender"
        Me.Text = "FrmMain"
        CType(Me.NesScreen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents NesScreen As System.Windows.Forms.PictureBox
End Class
