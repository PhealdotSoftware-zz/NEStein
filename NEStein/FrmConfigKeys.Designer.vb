<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmConfigKeys
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
        Me.lblKeyToPress = New System.Windows.Forms.Label()
        Me.lblPressedKey = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblKeyToPress
        '
        Me.lblKeyToPress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKeyToPress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKeyToPress.Location = New System.Drawing.Point(12, 72)
        Me.lblKeyToPress.Name = "lblKeyToPress"
        Me.lblKeyToPress.Size = New System.Drawing.Size(270, 32)
        Me.lblKeyToPress.TabIndex = 4
        Me.lblKeyToPress.Text = "Loading..."
        Me.lblKeyToPress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblPressedKey
        '
        Me.lblPressedKey.Location = New System.Drawing.Point(12, 116)
        Me.lblPressedKey.Name = "lblPressedKey"
        Me.lblPressedKey.Size = New System.Drawing.Size(270, 15)
        Me.lblPressedKey.TabIndex = 5
        Me.lblPressedKey.Text = "..."
        Me.lblPressedKey.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FrmConfigKeys
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(294, 176)
        Me.Controls.Add(Me.lblPressedKey)
        Me.Controls.Add(Me.lblKeyToPress)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.KeyPreview = True
        Me.Name = "FrmConfigKeys"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configurar Controles"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblKeyToPress As System.Windows.Forms.Label
    Friend WithEvents lblPressedKey As System.Windows.Forms.Label
End Class
