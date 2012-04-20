'Author: Kyle K. Loewenhagen
'Date: ##/##/####
'Purpose: Select Clubs for the student to enroll into.

Option Strict On
Option Explicit On

Imports System.IO

Public Class frmStudentClubs
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
    Friend WithEvents imgSmallwelcome As System.Windows.Forms.PictureBox
    Friend WithEvents lstClubs As System.Windows.Forms.ListBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cboStudentClubs As System.Windows.Forms.ComboBox
    Friend WithEvents lblInstructions1 As System.Windows.Forms.Label
    Friend WithEvents lblInstructions2 As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblPhone As System.Windows.Forms.Label
    Friend WithEvents lblCityStreet As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents grpInformation As System.Windows.Forms.GroupBox
    Friend WithEvents lblNameDB As System.Windows.Forms.Label
    Friend WithEvents lblAddressDB As System.Windows.Forms.Label
    Friend WithEvents lblCityDB As System.Windows.Forms.Label
    Friend WithEvents lblPhoneDB As System.Windows.Forms.Label
    Friend WithEvents lblEmailDB As System.Windows.Forms.Label
    Friend WithEvents pbxStudent As System.Windows.Forms.PictureBox
    Friend WithEvents btnSwitchStudent As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStudentClubs))
        Me.imgSmallwelcome = New System.Windows.Forms.PictureBox
        Me.pbxStudent = New System.Windows.Forms.PictureBox
        Me.lstClubs = New System.Windows.Forms.ListBox
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.cboStudentClubs = New System.Windows.Forms.ComboBox
        Me.lblInstructions1 = New System.Windows.Forms.Label
        Me.lblInstructions2 = New System.Windows.Forms.Label
        Me.lblEmail = New System.Windows.Forms.Label
        Me.lblPhone = New System.Windows.Forms.Label
        Me.lblCityStreet = New System.Windows.Forms.Label
        Me.lblAddress = New System.Windows.Forms.Label
        Me.lblName = New System.Windows.Forms.Label
        Me.grpInformation = New System.Windows.Forms.GroupBox
        Me.lblEmailDB = New System.Windows.Forms.Label
        Me.lblPhoneDB = New System.Windows.Forms.Label
        Me.lblCityDB = New System.Windows.Forms.Label
        Me.lblAddressDB = New System.Windows.Forms.Label
        Me.lblNameDB = New System.Windows.Forms.Label
        Me.btnSwitchStudent = New System.Windows.Forms.Button
        CType(Me.imgSmallwelcome, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbxStudent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpInformation.SuspendLayout()
        Me.SuspendLayout()
        '
        'imgSmallwelcome
        '
        Me.imgSmallwelcome.Image = CType(resources.GetObject("imgSmallwelcome.Image"), System.Drawing.Image)
        Me.imgSmallwelcome.Location = New System.Drawing.Point(0, 0)
        Me.imgSmallwelcome.Name = "imgSmallwelcome"
        Me.imgSmallwelcome.Size = New System.Drawing.Size(152, 72)
        Me.imgSmallwelcome.TabIndex = 7
        Me.imgSmallwelcome.TabStop = False
        '
        'pbxStudent
        '
        Me.pbxStudent.Location = New System.Drawing.Point(24, 56)
        Me.pbxStudent.Name = "pbxStudent"
        Me.pbxStudent.Size = New System.Drawing.Size(288, 273)
        Me.pbxStudent.TabIndex = 8
        Me.pbxStudent.TabStop = False
        '
        'lstClubs
        '
        Me.lstClubs.Location = New System.Drawing.Point(188, 407)
        Me.lstClubs.Name = "lstClubs"
        Me.lstClubs.Size = New System.Drawing.Size(248, 95)
        Me.lstClubs.TabIndex = 10
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnAdd.Location = New System.Drawing.Point(460, 439)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 11
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnDelete.Location = New System.Drawing.Point(556, 439)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 12
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'cboStudentClubs
        '
        Me.cboStudentClubs.Location = New System.Drawing.Point(148, 367)
        Me.cboStudentClubs.Name = "cboStudentClubs"
        Me.cboStudentClubs.Size = New System.Drawing.Size(344, 21)
        Me.cboStudentClubs.TabIndex = 13
        '
        'lblInstructions1
        '
        Me.lblInstructions1.Location = New System.Drawing.Point(52, 343)
        Me.lblInstructions1.Name = "lblInstructions1"
        Me.lblInstructions1.Size = New System.Drawing.Size(368, 24)
        Me.lblInstructions1.TabIndex = 14
        Me.lblInstructions1.Text = "Below is a list of clubs that the current student is involved with:"
        '
        'lblInstructions2
        '
        Me.lblInstructions2.Location = New System.Drawing.Point(452, 407)
        Me.lblInstructions2.Name = "lblInstructions2"
        Me.lblInstructions2.Size = New System.Drawing.Size(184, 24)
        Me.lblInstructions2.TabIndex = 15
        Me.lblInstructions2.Text = "You may add or delete a club here:"
        '
        'lblEmail
        '
        Me.lblEmail.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblEmail.Location = New System.Drawing.Point(28, 269)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(40, 23)
        Me.lblEmail.TabIndex = 28
        Me.lblEmail.Text = "Email:"
        '
        'lblPhone
        '
        Me.lblPhone.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblPhone.Location = New System.Drawing.Point(20, 221)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(40, 23)
        Me.lblPhone.TabIndex = 27
        Me.lblPhone.Text = "Phone:"
        '
        'lblCityStreet
        '
        Me.lblCityStreet.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblCityStreet.Location = New System.Drawing.Point(4, 165)
        Me.lblCityStreet.Name = "lblCityStreet"
        Me.lblCityStreet.Size = New System.Drawing.Size(64, 23)
        Me.lblCityStreet.TabIndex = 26
        Me.lblCityStreet.Text = "City, State:"
        '
        'lblAddress
        '
        Me.lblAddress.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblAddress.Location = New System.Drawing.Point(12, 109)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(56, 23)
        Me.lblAddress.TabIndex = 25
        Me.lblAddress.Text = "Address:"
        '
        'lblName
        '
        Me.lblName.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblName.Location = New System.Drawing.Point(20, 61)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(40, 24)
        Me.lblName.TabIndex = 24
        Me.lblName.Text = "Name:"
        '
        'grpInformation
        '
        Me.grpInformation.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.grpInformation.Controls.Add(Me.lblEmailDB)
        Me.grpInformation.Controls.Add(Me.lblPhoneDB)
        Me.grpInformation.Controls.Add(Me.lblCityDB)
        Me.grpInformation.Controls.Add(Me.lblAddressDB)
        Me.grpInformation.Controls.Add(Me.lblNameDB)
        Me.grpInformation.Controls.Add(Me.lblEmail)
        Me.grpInformation.Controls.Add(Me.lblPhone)
        Me.grpInformation.Controls.Add(Me.lblCityStreet)
        Me.grpInformation.Controls.Add(Me.lblAddress)
        Me.grpInformation.Controls.Add(Me.lblName)
        Me.grpInformation.Location = New System.Drawing.Point(352, 16)
        Me.grpInformation.Name = "grpInformation"
        Me.grpInformation.Size = New System.Drawing.Size(296, 313)
        Me.grpInformation.TabIndex = 9
        Me.grpInformation.TabStop = False
        Me.grpInformation.Text = "Member Information"
        '
        'lblEmailDB
        '
        Me.lblEmailDB.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblEmailDB.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmailDB.Location = New System.Drawing.Point(72, 272)
        Me.lblEmailDB.Name = "lblEmailDB"
        Me.lblEmailDB.Size = New System.Drawing.Size(216, 24)
        Me.lblEmailDB.TabIndex = 33
        '
        'lblPhoneDB
        '
        Me.lblPhoneDB.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblPhoneDB.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhoneDB.Location = New System.Drawing.Point(72, 224)
        Me.lblPhoneDB.Name = "lblPhoneDB"
        Me.lblPhoneDB.Size = New System.Drawing.Size(216, 24)
        Me.lblPhoneDB.TabIndex = 32
        '
        'lblCityDB
        '
        Me.lblCityDB.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblCityDB.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCityDB.Location = New System.Drawing.Point(72, 168)
        Me.lblCityDB.Name = "lblCityDB"
        Me.lblCityDB.Size = New System.Drawing.Size(216, 24)
        Me.lblCityDB.TabIndex = 31
        '
        'lblAddressDB
        '
        Me.lblAddressDB.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblAddressDB.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddressDB.Location = New System.Drawing.Point(72, 112)
        Me.lblAddressDB.Name = "lblAddressDB"
        Me.lblAddressDB.Size = New System.Drawing.Size(208, 24)
        Me.lblAddressDB.TabIndex = 30
        '
        'lblNameDB
        '
        Me.lblNameDB.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblNameDB.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNameDB.Location = New System.Drawing.Point(72, 64)
        Me.lblNameDB.Name = "lblNameDB"
        Me.lblNameDB.Size = New System.Drawing.Size(208, 24)
        Me.lblNameDB.TabIndex = 29
        '
        'btnSwitchStudent
        '
        Me.btnSwitchStudent.Location = New System.Drawing.Point(468, 479)
        Me.btnSwitchStudent.Name = "btnSwitchStudent"
        Me.btnSwitchStudent.Size = New System.Drawing.Size(160, 24)
        Me.btnSwitchStudent.TabIndex = 16
        Me.btnSwitchStudent.Text = "Switch Students"
        '
        'frmStudentClubs
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(656, 519)
        Me.Controls.Add(Me.btnSwitchStudent)
        Me.Controls.Add(Me.lblInstructions2)
        Me.Controls.Add(Me.lblInstructions1)
        Me.Controls.Add(Me.cboStudentClubs)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.lstClubs)
        Me.Controls.Add(Me.grpInformation)
        Me.Controls.Add(Me.pbxStudent)
        Me.Controls.Add(Me.imgSmallwelcome)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmStudentClubs"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main"
        CType(Me.imgSmallwelcome, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbxStudent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpInformation.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Module Level Variables
    Dim mstrStudentID As String

    'Public Properties
    Public Property StudentID() As String
        Get
            Return mstrStudentID
        End Get
        Set(ByVal strValue As String)
            mstrStudentID = strValue
        End Set
    End Property

    Private Sub frmStudentClubs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Populate Form Text Property
        Me.Text = StudentID

        'Populate Picture
        Dim imgPath As String = "../pictures/" & mstrStudentID & ".jpg"
        If File.Exists(imgPath) Then
            pbxStudent.Image = Image.FromFile(imgPath)
        End If

        'Populate Student Data Labels
        'Instantiate Student Class
        Dim objStudent As New clsStudent(StudentID)
        'Call Select Student Data Subroutine
        objStudent.Select_Student_Data()
        'Populate the values of Student Properties into the Label Text Properties
        lblNameDB.Text = objStudent.Display_Full_Name
        lblAddressDB.Text = objStudent.Address
        lblCityDB.Text = objStudent.Display_Full_Address
        lblPhoneDB.Text = objStudent.Telephone
        lblEmailDB.Text = objStudent.Email

        'Populate Clubs Combobox
        'Instantiate Clubs Class
        Dim objClubs As New clsClubs(StudentID)
        'Call Select Club List Subroutine
        objClubs.Select_Club_List()
        'Populate Clubs Combobox
        cboStudentClubs.DataSource = objClubs.Club_List.Tables("clubs")
        cboStudentClubs.DisplayMember = "CLUB_NAME"
        cboStudentClubs.ValueMember = "CLUB_ID"

        'Populate Clubs ListBox
        'Call Select_Enrolled_Club subroutine
        objClubs.Select_Enrolled_Club()
        For Each strClub As String In objClubs.Enrolled_Clubs
            lstClubs.Items.Add(strClub)
        Next

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        'Instantiate clsClubs class and use overload to pass in Student and Club ID
        Dim objClubs As New clsClubs(mstrStudentID, cboStudentClubs.SelectedValue.ToString)

        'Call Club Add Method
        objClubs.Club_Add()

        lstClubs.Items.Add(cboStudentClubs.Text)
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        'Instantiate clsClubs class and use overload to pass in Student and Club ID
        Dim objClubs As New clsClubs(mstrStudentID, lstClubs.SelectedItem.ToString.Substring(0, 3))

        'Call Club Add Method
        objClubs.Club_Remove()

        'Add current select Club into lstClubs
        lstClubs.Items.Remove(lstClubs.SelectedItem)
    End Sub

    Private Sub btnSwitchStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSwitchStudent.Click
        frmStudentClubs.ActiveForm.Close()
    End Sub

End Class
