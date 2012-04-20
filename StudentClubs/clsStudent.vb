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

    Public Property StudentID As String
        Get
            Return mstrStudentID
        End Get
        Set(value As String)
            mstrStudentID = value
        End Set
    End Property

    Public Property FirstName As String
        Get
            Return mstrFirstName
        End Get
        Set(value As String)
            mstrFirstName = value
        End Set
    End Property

    Public Property LastName As String
        Get
            Return mstrLastName
        End Get
        Set(value As String)
            mstrLastName = value
        End Set
    End Property

    Public Property Address As String
        Get
            Return mstrAddress
        End Get
        Set(value As String)
            mstrAddress = value
        End Set
    End Property

    Public Property City As String
        Get
            Return mstrCity
        End Get
        Set(value As String)
            mstrCity = value
        End Set
    End Property

    Public Property State As String
        Get
            Return mstrState
        End Get
        Set(value As String)
            mstrState = value
        End Set
    End Property

    Public Property Zip As String
        Get
            Return mstrZip
        End Get
        Set(value As String)
            mstrZip = value
        End Set
    End Property

    Public Property Telephone As String
        Get
            Return mstrTelephone
        End Get
        Set(value As String)
            mstrTelephone = value
        End Set
    End Property

    Public Property Email As String
        Get
            Return mstrEmail
        End Get
        Set(value As String)
            mstrEmail = value
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

    'Overload Constructor
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
    Public Function Display_Full_Name() As String
        Return FirstName & " " & LastName
    End Function

    Public Function Display_Full_Address() As String
        Return City & ", " & State & " " & Zip
    End Function

    Public Sub Select_Student_Data()
        'Instantiate Connection
        Dim objConnection As New OleDbConnection(mstrCN)

        'Create SQL Statement
        mstrSQL = "Select First_Name, Last_Name, Address1,City,State,Zip," &
            "Telephone,Email_Address from Student_Data_Tbl where student_ID = '" &
            mstrStudentID & "'"

        'Instantiate Command Class
        Dim objCommand As New OleDbCommand(mstrSQL, objConnection)

        'Open the Database
        objCommand.Connection.Open()

        'Instantiate DataReader Class
        Dim objDataReader As OleDbDataReader

        'Execute SQL
        objDataReader = objCommand.ExecuteReader

        'Read SQL Results from DataReader
        Do While (objDataReader.Read)
            FirstName = objDataReader.Item("First_Name").ToString
            LastName = objDataReader.Item("Last_Name").ToString
            Address = objDataReader.Item("Address1").ToString
            City = objDataReader.Item("City").ToString
            State = objDataReader.Item("State").ToString
            Zip = objDataReader.Item("Zip").ToString
            Telephone = objDataReader.Item("Telephone").ToString
            Email = objDataReader.Item("Email_Address").ToString
        Loop

        'Close Objects
        objCommand.Dispose()
        objDataReader.Close()
        objConnection.Close()
        objConnection.Dispose()

    End Sub

    Public Sub Select_Student_List()
        'Instantiate Connection
        Dim objConnection As New OleDbConnection(mstrCN)
        'Open the Database
        objConnection.Open()

        'Create SQL Statement
        mstrSQL = "Select Student_ID, First_Name & ' ' & Last_Name as Full_Name from " &
            "Student_Data_tbl order by First_Name"

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
