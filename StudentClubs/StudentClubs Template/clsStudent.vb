' Project name:         Student Class
' Project purpose:      Manage Student Data.
' Created/revisd by:    <your name> on <current date>

Option Explicit On
Option Strict On

Imports System.Data.OleDb

Public Class clsStudent
    'Module Level Variables
    Private mstrCN As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student_Clubs.accdb; Persist Security Info=False;"
    Private mstrSQL As String
    'Module Level Variable for Properties
    Private mstrStudentID As String
    Private mstrFirstName As String
    Private mstrLastName As String
    Private mstrAddress As String
    Private mstrCity As String
    Private mstrState As String
    Private mstrZip As String
    Private mstrTelephone As String
    Private mstrEmail As String
    Private mdstStudent As New DataSet

    'Public Properties 
    Public Property Student_Roster As DataSet
        Get
            Return mdstStudent
        End Get
        Set(value As DataSet)
            mdstStudent = value
        End Set
    End Property

    'Default Constructor
    Public Sub New()
        mstrStudentID = String.Empty
        mstrFirstName = String.Empty
        mstrLastName = String.Empty
        mstrAddress = String.Empty
        mstrCity = String.Empty
        mstrState = String.Empty
        mstrZip = String.Empty
        mstrTelephone = String.Empty
        mstrEmail = String.Empty
        mdstStudent.Tables.Clear()
    End Sub

    Public Sub New(ByVal strID As String)
        mstrStudentID = strID
        mstrFirstName = String.Empty
        mstrLastName = String.Empty
        mstrAddress = String.Empty
        mstrCity = String.Empty
        mstrState = String.Empty
        mstrZip = String.Empty
        mstrTelephone = String.Empty
        mstrEmail = String.Empty
        mdstStudent.Tables.Clear()
    End Sub

    'Methods
    Public Sub Select_Student_List()
        'Instantiate Connection
        Dim objConnection As New OleDbConnection(mstrCN)
        'Open the Database
        objConnection.Open()

        'Define SQL
        mstrSQL = "Select Student_ID, First_Name & ' ' & Last_Name as Full_Name from " &
            "Student_Data_tbl order by First_Name"

        MessageBox.Show(mstrSQL)

        'Instantiate DataAdapter
        Dim objDA As New OleDbDataAdapter(mstrSQL, objConnection)

        'Populate DataSet by using the DataAdapter
        objDA.Fill(mdstStudent, "students")

        'Close Objects
        objConnection.Close()
        objDA.Dispose()
        objConnection.Dispose()

    End Sub


End Class
