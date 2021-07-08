<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmYearEnd
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.cmdCreateTable = New System.Windows.Forms.Button()
        Me.cmdUpdateParticipant = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkPen = New System.Windows.Forms.CheckBox()
        Me.chkBenef = New System.Windows.Forms.CheckBox()
        Me.chkActPart = New System.Windows.Forms.CheckBox()
        Me.chkInActVest = New System.Windows.Forms.CheckBox()
        Me.chkInActNonVest = New System.Windows.Forms.CheckBox()
        Me.chk70Pen = New System.Windows.Forms.CheckBox()
        Me.chkQdro = New System.Windows.Forms.CheckBox()
        Me.chkKey = New System.Windows.Forms.CheckBox()
        Me.chkReconc = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.cmdExport = New System.Windows.Forms.Button()
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdCreateTable
        '
        Me.cmdCreateTable.Location = New System.Drawing.Point(13, 28)
        Me.cmdCreateTable.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdCreateTable.Name = "cmdCreateTable"
        Me.cmdCreateTable.Size = New System.Drawing.Size(164, 32)
        Me.cmdCreateTable.TabIndex = 0
        Me.cmdCreateTable.Text = "Create Actuarial Table"
        Me.cmdCreateTable.UseVisualStyleBackColor = True
        '
        'cmdUpdateParticipant
        '
        Me.cmdUpdateParticipant.Location = New System.Drawing.Point(193, 28)
        Me.cmdUpdateParticipant.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdUpdateParticipant.Name = "cmdUpdateParticipant"
        Me.cmdUpdateParticipant.Size = New System.Drawing.Size(139, 32)
        Me.cmdUpdateParticipant.TabIndex = 1
        Me.cmdUpdateParticipant.Text = "Update Participant Info"
        Me.cmdUpdateParticipant.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 537)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "lblStatus"
        '
        'chkPen
        '
        Me.chkPen.AutoSize = True
        Me.chkPen.Location = New System.Drawing.Point(25, 127)
        Me.chkPen.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkPen.Name = "chkPen"
        Me.chkPen.Size = New System.Drawing.Size(101, 21)
        Me.chkPen.TabIndex = 1
        Me.chkPen.Tag = "v_PensionersExport"
        Me.chkPen.Text = "Pensioners"
        Me.chkPen.UseVisualStyleBackColor = True
        '
        'chkBenef
        '
        Me.chkBenef.AutoSize = True
        Me.chkBenef.Location = New System.Drawing.Point(25, 155)
        Me.chkBenef.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkBenef.Name = "chkBenef"
        Me.chkBenef.Size = New System.Drawing.Size(111, 21)
        Me.chkBenef.TabIndex = 2
        Me.chkBenef.Tag = "v_BeneficiaryExport"
        Me.chkBenef.Text = "Beneficiaries"
        Me.chkBenef.UseVisualStyleBackColor = True
        '
        'chkActPart
        '
        Me.chkActPart.AutoSize = True
        Me.chkActPart.Location = New System.Drawing.Point(25, 183)
        Me.chkActPart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkActPart.Name = "chkActPart"
        Me.chkActPart.Size = New System.Drawing.Size(146, 21)
        Me.chkActPart.TabIndex = 3
        Me.chkActPart.Tag = "v_ActiveMembersExport"
        Me.chkActPart.Text = "Active Participants"
        Me.chkActPart.UseVisualStyleBackColor = True
        '
        'chkInActVest
        '
        Me.chkInActVest.AutoSize = True
        Me.chkInActVest.Location = New System.Drawing.Point(25, 212)
        Me.chkInActVest.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkInActVest.Name = "chkInActVest"
        Me.chkInActVest.Size = New System.Drawing.Size(126, 21)
        Me.chkInActVest.TabIndex = 4
        Me.chkInActVest.Tag = "v_InactiveVestedExport"
        Me.chkInActVest.Text = "Inactive Vested"
        Me.chkInActVest.UseVisualStyleBackColor = True
        '
        'chkInActNonVest
        '
        Me.chkInActNonVest.AutoSize = True
        Me.chkInActNonVest.Location = New System.Drawing.Point(25, 241)
        Me.chkInActNonVest.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkInActNonVest.Name = "chkInActNonVest"
        Me.chkInActNonVest.Size = New System.Drawing.Size(156, 21)
        Me.chkInActNonVest.TabIndex = 5
        Me.chkInActNonVest.Tag = "v_InactiveNotVestedExport"
        Me.chkInActNonVest.Text = "Inactive Non Vested"
        Me.chkInActNonVest.UseVisualStyleBackColor = True
        '
        'chk70Pen
        '
        Me.chk70Pen.AutoSize = True
        Me.chk70Pen.Location = New System.Drawing.Point(25, 270)
        Me.chk70Pen.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chk70Pen.Name = "chk70Pen"
        Me.chk70Pen.Size = New System.Drawing.Size(242, 21)
        Me.chk70Pen.TabIndex = 6
        Me.chk70Pen.Tag = "v_ActiveParticpant_846_Months_AndOlder"
        Me.chk70Pen.Text = "Age Seventy and Half Pensioners"
        Me.chk70Pen.UseVisualStyleBackColor = True
        '
        'chkQdro
        '
        Me.chkQdro.AutoSize = True
        Me.chkQdro.Location = New System.Drawing.Point(25, 299)
        Me.chkQdro.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkQdro.Name = "chkQdro"
        Me.chkQdro.Size = New System.Drawing.Size(145, 21)
        Me.chkQdro.TabIndex = 7
        Me.chkQdro.Tag = "v_QDROExport"
        Me.chkQdro.Text = "QDRO Employees"
        Me.chkQdro.UseVisualStyleBackColor = True
        '
        'chkKey
        '
        Me.chkKey.AutoSize = True
        Me.chkKey.Location = New System.Drawing.Point(25, 329)
        Me.chkKey.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkKey.Name = "chkKey"
        Me.chkKey.Size = New System.Drawing.Size(127, 21)
        Me.chkKey.TabIndex = 8
        Me.chkKey.Tag = "v_KeyEmployees"
        Me.chkKey.Text = "Key Employees"
        Me.chkKey.UseVisualStyleBackColor = True
        '
        'chkReconc
        '
        Me.chkReconc.AutoSize = True
        Me.chkReconc.Location = New System.Drawing.Point(25, 358)
        Me.chkReconc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkReconc.Name = "chkReconc"
        Me.chkReconc.Size = New System.Drawing.Size(205, 21)
        Me.chkReconc.TabIndex = 9
        Me.chkReconc.Tag = "v_ReconPenBen"
        Me.chkReconc.Text = "Reconciliation Report Detail"
        Me.chkReconc.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(21, 81)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 20)
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "Export Files"
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(25, 471)
        Me.txtPath.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(445, 22)
        Me.txtPath.TabIndex = 27
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(25, 436)
        Me.cmdExport.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(164, 28)
        Me.cmdExport.TabIndex = 28
        Me.cmdExport.Text = "Export Selected Files"
        Me.cmdExport.UseVisualStyleBackColor = True
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Location = New System.Drawing.Point(480, 469)
        Me.cmdBrowse.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(53, 28)
        Me.cmdBrowse.TabIndex = 29
        Me.cmdBrowse.Text = "..."
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(340, 28)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(108, 32)
        Me.cmdExit.TabIndex = 30
        Me.cmdExit.Text = "&Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'frmYearEnd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(763, 559)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.cmdExport)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.chkReconc)
        Me.Controls.Add(Me.chkKey)
        Me.Controls.Add(Me.chkQdro)
        Me.Controls.Add(Me.chk70Pen)
        Me.Controls.Add(Me.chkInActNonVest)
        Me.Controls.Add(Me.chkInActVest)
        Me.Controls.Add(Me.chkActPart)
        Me.Controls.Add(Me.chkBenef)
        Me.Controls.Add(Me.chkPen)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdUpdateParticipant)
        Me.Controls.Add(Me.cmdCreateTable)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmYearEnd"
        Me.Text = "Year End - Actuarial Process"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdCreateTable As Button
    Friend WithEvents cmdUpdateParticipant As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents chkPen As CheckBox
    Friend WithEvents chkBenef As CheckBox
    Friend WithEvents chkActPart As CheckBox
    Friend WithEvents chkInActVest As CheckBox
    Friend WithEvents chkInActNonVest As CheckBox
    Friend WithEvents chk70Pen As CheckBox
    Friend WithEvents chkQdro As CheckBox
    Friend WithEvents chkKey As CheckBox
    Friend WithEvents chkReconc As CheckBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtPath As TextBox
    Friend WithEvents cmdExport As Button
    Friend WithEvents cmdBrowse As Button
    Friend WithEvents cmdExit As Button
End Class
