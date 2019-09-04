' DBL.vb
'         Title: DBL - Data Base Layer for Piecework Payroll
' Last Modified: October-24-2018
'    Written By: Rahulkumar Patel
' Adapted from PieceworkWorker by Kyle Chapman, October 2017
' 
' This is a module with a set of classes allowing for interaction between
' Piecework Worker data objects and a database.


Imports System.Data
Imports System.Data.SqlClient

Public Class DBL

#Region "Connection String"

    Friend Class Conn

        ''' <summary>
        ''' Return connection string
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Function GetConnectionString() As String
            Return My.Settings.dbConnection
        End Function
    End Class

#End Region

#Region "SQL Statements"

    ''' <summary>
    ''' Prepare SQL statements used to perform necessary actions in the database
    ''' </summary>
    Friend Class SQLStatements
        Friend Const SelectById As String = "SELECT TOP 1 * FROM [tblEntries] WHERE [EntryId] = @entryId"
        Friend Const SelectAll As String = "SELECT * FROM [tblEntries]"
        Friend Const InsertNew As String = "INSERT INTO tblEntries VALUES(@firstName, @lastName, @messages, @pay, @entryDate)"
        Friend Const UpdateExisting As String = "UPDATE tblEntries Set FirstName = @firstName, LastName = @lastName, Messages = @messages, Pay = @pay WHERE EntryId = @entryId"
        Friend Const DeleteExisting As String = "DELETE FROM [tblEntries] WHERE [EntryId] = @entryId"

        ' These additional statements may be used to replace the summary values used in the class
        Friend Const TotalPay As String = "SELECT SUM(Pay) FROM tblEmployee"
        Friend Const TotalMessages As String = "SELECT SUM(Messages) FROM tblEmployee"
        Friend Const TotalEmployees As String = "SELECT COUNT(EntryId) FROM tblEmployee"

    End Class

#End Region

#Region "Methods"

    ''' <summary>
    ''' Function used to select one row from database, takes workerID as the primary key
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns></returns>
    Friend Shared Function GetOneRow(id As Integer) As PieceworkWorker

        ' Declare new worker object and dbConnection
        Dim returnWorker As New PieceworkWorker()
        Dim dbConnection As New SqlConnection
        dbConnection.ConnectionString = Conn.GetConnectionString

        ' Create new SQL command, assign it prepared statement
        Dim command As New SqlCommand
        command.Connection = dbConnection
        command.CommandType = CommandType.Text
        command.CommandText = SQLStatements.SelectById
        command.Parameters.AddWithValue("@entryId", id)

        ' Try to connect to the database, create a datareader. If successful, read from the database and fill created row
        ' with information from matching record
        Try
            dbConnection.Open()
            Dim reader As IDataReader = command.ExecuteReader
            If reader.Read() Then

                returnWorker = New PieceworkWorker(reader("FirstName"), reader("LastName"), reader("Messages"))
                returnWorker.Id = id

            End If
        Catch ex As Exception
            MessageBox.Show("A database error has been encountered:" & vbCrLf & ex.Message, "Database Error")
        Finally
            dbConnection.Close()
        End Try

        ' Return the populated row
        Return returnWorker
    End Function

    ''' <summary>
    ''' Returns a list of every row found in the table
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function GetEmployeeList()

        ' Declare the Connection
        Dim dbConnection = New SqlConnection(Conn.GetConnectionString())

        ' Create new SQL command, assign it prepared statement
        Dim commandString = New SqlCommand(SQLStatements.SelectAll, dbConnection)
        Dim adapter = New SqlDataAdapter(commandString)

        ' Declare a DataTable object that will hold the return value
        Dim employeeTable = New DataTable()

        ' Try to connect to the database, and use the adapter to fill the table
        Try
            dbConnection.Open()
            adapter.Fill(employeeTable)
        Catch ex As Exception
            MessageBox.Show("A database error has been encountered:" & vbCrLf & ex.Message, "Database Error")
        Finally
            dbConnection.Close()
        End Try

        ' Return the populated DataTable's DataView
        Return employeeTable.DefaultView

    End Function

    ''' <summary>
    ''' Insert a new record into the database
    ''' </summary>
    ''' <param name="insertWorker"></param>
    ''' <returns></returns>
    Friend Shared Function InsertNewRecord(insertWorker As PieceworkWorker) As Boolean

        ' Create return value and dbConnection
        Dim returnValue As Boolean = False
        Dim dbConnection As New SqlConnection
        dbConnection.ConnectionString = Conn.GetConnectionString

        ' Create new command, assign it prepared statement, and assign it paramaters
        Dim command As New SqlCommand
        command.Connection = dbConnection
        command.CommandType = CommandType.Text
        command.CommandText = SQLStatements.InsertNew
        command.Parameters.AddWithValue("@firstName", insertWorker.FirstName)
        command.Parameters.AddWithValue("@lastName", insertWorker.LastName)
        command.Parameters.AddWithValue("@messages", insertWorker.Messages)
        command.Parameters.AddWithValue("@pay", insertWorker.Pay)
        command.Parameters.AddWithValue("@entryDate", insertWorker.EntryDate)

        ' Try to insert the new record, return result
        Try
            dbConnection.Open()
            returnValue = (command.ExecuteNonQuery = 1)
        Catch ex As Exception
            MessageBox.Show("A database error has been encountered:" & vbCrLf & ex.Message, "Database Error")
        Finally
            dbConnection.Close()
        End Try

        Return returnValue
    End Function

    ''' <summary>
    ''' Delete record from the database
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns></returns>
    Friend Shared Function DeleteRow(id As Integer) As Boolean

        ' Create return value and dbConnection
        Dim returnValue As Boolean = False
        Dim dbConnection As New SqlConnection
        dbConnection.ConnectionString = Conn.GetConnectionString

        ' Create new command, assign it a prepared statement and add the paramater of workerID (PK)
        Dim command As New SqlCommand
        command.Connection = dbConnection
        command.CommandType = CommandType.Text
        command.CommandText = SQLStatements.DeleteExisting
        command.Parameters.AddWithValue("@entryId", id)

        ' Attempt to open connection to DB, return result if delete was successful
        Try
            dbConnection.Open()
            returnValue = command.ExecuteNonQuery() > 0
        Catch ex As Exception
            MessageBox.Show("A database error has been encountered:" & vbCrLf & ex.Message, "Database Error")
        Finally
            dbConnection.Close()
        End Try

        Return returnValue
    End Function

    ''' <summary>
    ''' Updating an already existing row in the database
    ''' </summary>
    ''' <param name="updateWorker"></param>
    ''' <returns></returns>
    Friend Shared Function UpdateExistingRow(updateWorker As PieceworkWorker) As Boolean

        ' Create return value
        Dim returnValue As Boolean = False

        ' If row exists, create dbConnection
        If updateWorker.Id > 0 Then
            Dim dbConnection As New SqlConnection
            dbConnection.ConnectionString = Conn.GetConnectionString

            ' Create new command, assign it a prepared SQL statement and assign it paramaters
            Dim command As New SqlCommand
            command.Connection = dbConnection
            command.CommandType = CommandType.Text
            command.CommandText = SQLStatements.UpdateExisting
            command.Parameters.AddWithValue("@workerId", updateWorker.Id)
            command.Parameters.AddWithValue("@firstName", updateWorker.FirstName)
            command.Parameters.AddWithValue("@lastName", updateWorker.LastName)
            command.Parameters.AddWithValue("@messages", updateWorker.Messages)
            command.Parameters.AddWithValue("@pay", updateWorker.Pay)
            command.Parameters.AddWithValue("@entryDate", updateWorker.EntryDate)

            ' Try to open a connection to the database and update the record. Return result.
            Try
                dbConnection.Open()
                If command.ExecuteNonQuery > 0 Then returnValue = True
            Catch ex As Exception
                MessageBox.Show("A database error has been encountered: " & Environment.NewLine & ex.Message, "Database Error")
            Finally
                dbConnection.Close()
            End Try

            ' If the worker does Not exist, attempt to insert it instead
        Else
            If CInt(InsertNewRecord(updateWorker)) > 0 Then returnValue = True
        End If

        ' Returns true if the query executed; always false if the row is invalid
        Return returnValue
    End Function

#End Region

End Class