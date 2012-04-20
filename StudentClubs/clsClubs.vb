' Project name:         Clubs Class
' Project purpose:      Manage Clubs Data.
' Created/revisd by:    <your name> on <current date>

Option Explicit On
Option Strict On

Imports System.Data.OleDb
Imports System.Collections.Generic

Public Class clsClubs
    'Module Level Variables
    Private mstrCN As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student_Clubs.accdb; Persist Security Info=False;"
    Private mstrSQL As String
    'Module Level Variable for Properties
    Private mstrStudentID As String
    Private mstrClubID As String
    Private mstrClubName As String
    Private mstrClubStatus As String
    Private mdstClubs As New DataSet
    Private mlstEnrolledClubs As New List(Of String)

    'Public Properties 
    Public Property Club_List As DataSet
        Get
            Return mdstClubs
        End Get
        Set(value As DataSet)
            mdstClubs = value
        End Set
    End Property

    Public Property Enrolled_Clubs As List(Of String)
        Get
            Return mlstEnrolledClubs
        End Get
        Set(value As List(Of String))
            mlstEnrolledClubs = value
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

    Public Property ClubID As String
        Get
            Return mstrClubID
        End Get
        Set(value As String)
            mstrClubID = value
        End Set
    End Property

    Public Property ClubName As String
        Get
            Return mstrClubName
        End Get
        Set(value As String)
            mstrClubName = value
        End Set
    End Property

    Public Property ClubStatus As String
        Get
            Return mstrClubStatus
        End Get
        Set(value As String)
            mstrClubStatus = value
        End Set
    End Property

    'Default Constructor
    Public Sub New()
        mstrStudentID = String.Empty
        mstrClubID = String.Empty
        mstrClubName = String.Empty
        mstrClubStatus = String.Empty
        mdstClubs.Tables.Clear()
    End Sub

    'Overload Constructor
    Public Sub New(ByVal strStudentID As String)
        mstrStudentID = strStudentID
        mstrClubID = String.Empty
        mstrClubName = String.Empty
        mstrClubStatus = String.Empty
        mdstClubs.Tables.Clear()
    End Sub

    Public Sub New(ByVal strStudentID As String, ByVal strClubID As String)
        mstrStudentID = strStudentID
        mstrClubID = strClubID
        mstrClubName = String.Empty
        mstrClubStatus = String.Empty
        mdstClubs.Tables.Clear()
    End Sub


    'Methods
    Public Sub Select_Club_List()
        'Instantiate Connection
        Dim objConnection As New OleDbConnection(mstrCN)
        'Open the Database
        objConnection.Open()

        'Create SQL Statement
        mstrSQL = "Select Club_ID, Status, Club_ID & ' - ' & Student_Clubs as Club_Name " &
            "from Club_Setup_Tbl where Status = 'A'"

        'Instantiate DataAdapter
        Dim objDA As New OleDbDataAdapter(mstrSQL, objConnection)

        'Populate DataSet by using the DataAdapter
        objDA.Fill(mdstClubs, "clubs")

        'Close Objects
        objConnection.Close()
        objDA.Dispose()
        objConnection.Dispose()

    End Sub

    Public Sub Select_Enrolled_Club()
        'Instantiate Connection
        Dim objConnection As New OleDbConnection(mstrCN)

        'Create SQL Statement
        mstrSQL = "Select A.Club_ID, A.Student_ID, A.Club_ID & ' - ' & B.Student_Clubs as Club_Name " &
            "from Student_Clubs_Tbl A, Club_Setup_Tbl B where A.Club_ID = B.Club_ID " &
            "and A.Student_ID = '" & mstrStudentID & "'"

        'Instantiate Command Class
        Dim objCommand As New OleDbCommand(mstrSQL, objConnection)

        'Open the Database
        objCommand.Connection.Open()

        'Instantiate DataReader Class
        Dim objDR As OleDbDataReader = objCommand.ExecuteReader

        'Read SQL Results from DataReader
        Do While (objDR.Read)
            mlstEnrolledClubs.Add(objDR.Item("Club_Name").ToString())
        Loop

        'Close Objects
        objConnection.Close()
        objCommand.Dispose()
        objDR.Close()
        objConnection.Dispose()
    End Sub

    'Wrapper method for external use
    Public Sub Club_Add()
        Call Club_Alter("Insert")
    End Sub

    'Wrapper method for external use
    Public Sub Club_Remove()
        Call Club_Alter("Delete")
    End Sub

    'Internal method that actually performs the work when something needs to be altered.
    Private Sub Club_Alter(action As String)
        ' Null mstrSQL so nothing gets 'accidentally' done
        mstrSQL = String.Empty

        If String.Compare(action, "Insert") = 0 Then
            'Create SQL Statement
            mstrSQL = "Insert into Student_Clubs_tbl(Student_ID, Club_ID) values('" &
                mstrStudentID & "', '" & mstrClubID & "')"
        ElseIf String.Compare(action, "Delete") = 0 Then
            'Create SQL Statement
            mstrSQL = "Delete from Student_Clubs_tbl where Student_ID = '" &
                mstrStudentID & "' and Club_ID = '" & mstrClubID & "'"
        End If


        'Instantiate Connection
        Dim objConnection As New OleDbConnection(mstrCN)
        'Open the Database
        objConnection.Open()

        'Instantiate Command Class
        Dim objCommand As New OleDbCommand(mstrSQL, objConnection)

        'Execute Command Method
        objCommand.ExecuteNonQuery()

        'Close Objects
        objConnection.Close()
        objCommand.Dispose()
        objConnection.Dispose()
    End Sub

End Class
