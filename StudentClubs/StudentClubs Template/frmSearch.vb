'Author: Kyle K. Loewenhagen
'Date: ##/##/###
'Purpose: Select a student to access the Student Clubs system.

Option Strict On
Option Explicit On

Public Class frmStudentSearch
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents imgWelcome As System.Windows.Forms.PictureBox
    Friend WithEvents imgStudents As System.Windows.Forms.PictureBox
    Friend WithEvents txtStudentsearch As System.Windows.Forms.Label
    Friend WithEvents cboStudents As System.Windows.Forms.ComboBox
    Friend WithEvents txtVersion As System.Windows.Forms.Label
    Friend WithEvents imgSmallwelcome As System.Windows.Forms.PictureBox
    Friend WithEvents btnSelectStudent As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStudentSearch))
        Me.imgWelcome = New System.Windows.Forms.PictureBox()
        Me.imgStudents = New System.Windows.Forms.PictureBox()
        Me.btnSelectStudent = New System.Windows.Forms.Button()
        Me.txtStudentsearch = New System.Windows.Forms.Label()
        Me.cboStudents = New System.Windows.Forms.ComboBox()
        Me.txtVersion = New System.Windows.Forms.Label()
        Me.imgSmallwelcome = New System.Windows.Forms.PictureBox()
        CType(Me.imgWelcome, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgStudents, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgSmallwelcome, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'imgWelcome
        '
        Me.imgWelcome.Image = CType(resources.GetObject("imgWelcome.Image"), System.Drawing.Image)
        Me.imgWelcome.Location = New System.Drawing.Point(8, 16)
        Me.imgWelcome.Name = "imgWelcome"
        Me.imgWelcome.Size = New System.Drawing.Size(320, 88)
        Me.imgWelcome.TabIndex = 0
        Me.imgWelcome.TabStop = False
        '
        'imgStudents
        '
        Me.imgStudents.Image = CType(resources.GetObject("imgStudents.Image"), System.Drawing.Image)
        Me.imgStudents.Location = New System.Drawing.Point(0, 104)
        Me.imgStudents.Name = "imgStudents"
        Me.imgStudents.Size = New System.Drawing.Size(328, 192)
        Me.imgStudents.TabIndex = 1
        Me.imgStudents.TabStop = False
        '
        'btnSelectStudent
        '
        Me.btnSelectStudent.BackColor = System.Drawing.Color.RosyBrown
        Me.btnSelectStudent.Location = New System.Drawing.Point(352, 136)
        Me.btnSelectStudent.Name = "btnSelectStudent"
        Me.btnSelectStudent.Size = New System.Drawing.Size(136, 23)
        Me.btnSelectStudent.TabIndex = 2
        Me.btnSelectStudent.Text = "Select Student!"
        Me.btnSelectStudent.UseVisualStyleBackColor = False
        '
        'txtStudentsearch
        '
        Me.txtStudentsearch.Font = New System.Drawing.Font("Garamond", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStudentsearch.Location = New System.Drawing.Point(304, 24)
        Me.txtStudentsearch.Name = "txtStudentsearch"
        Me.txtStudentsearch.Size = New System.Drawing.Size(208, 40)
        Me.txtStudentsearch.TabIndex = 3
        Me.txtStudentsearch.Text = "Student Search"
        '
        'cboStudents
        '
        Me.cboStudents.Location = New System.Drawing.Point(344, 88)
        Me.cboStudents.Name = "cboStudents"
        Me.cboStudents.Size = New System.Drawing.Size(152, 21)
        Me.cboStudents.TabIndex = 4
        Me.cboStudents.Text = "Jeremiah Isaacson "
        '
        'txtVersion
        '
        Me.txtVersion.Location = New System.Drawing.Point(464, 264)
        Me.txtVersion.Name = "txtVersion"
        Me.txtVersion.Size = New System.Drawing.Size(48, 32)
        Me.txtVersion.TabIndex = 5
        Me.txtVersion.Text = "Version 1.1"
        '
        'imgSmallwelcome
        '
        Me.imgSmallwelcome.Image = CType(resources.GetObject("imgSmallwelcome.Image"), System.Drawing.Image)
        Me.imgSmallwelcome.Location = New System.Drawing.Point(352, 168)
        Me.imgSmallwelcome.Name = "imgSmallwelcome"
        Me.imgSmallwelcome.Size = New System.Drawing.Size(152, 72)
        Me.imgSmallwelcome.TabIndex = 6
        Me.imgSmallwelcome.TabStop = False
        '
        'frmStudentSearch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(512, 294)
        Me.Controls.Add(Me.imgSmallwelcome)
        Me.Controls.Add(Me.txtVersion)
        Me.Controls.Add(Me.cboStudents)
        Me.Controls.Add(Me.txtStudentsearch)
        Me.Controls.Add(Me.btnSelectStudent)
        Me.Controls.Add(Me.imgStudents)
        Me.Controls.Add(Me.imgWelcome)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmStudentSearch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CVTC Student Search"
        CType(Me.imgWelcome, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgStudents, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgSmallwelcome, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmStudentSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Instantiate Student Class
        Dim objStudent As New clsStudent()

        'Execute Method
        objStudent.Select_Student_List()

        'Populate Combobox
        cboStudents.DataSource = objStudent.Student_Roster.Tables("students")
        cboStudents.DisplayMember = "Full_Name"
        cboStudents.ValueMember = "Student_ID"

    End Sub

    Private Sub btnSelectStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectStudent.Click

        'Instantiate the frmClubs
        Dim objClubs As New frmStudentClubs
        'Populate Student ID Property
        objClubs.StudentID = cboStudents.SelectedValue.ToString
        'Open the frmClubs
        objClubs.ShowDialog()

    End Sub
End Class
