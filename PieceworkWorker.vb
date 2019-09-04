' PieceworkWorker.vb
'         Title: IncInc Payroll (Piecework)
' Last Modified: 18-09-2018
'    Written By: Rahulkumar Patel
' Adapted from PieceworkWorker by Kyle Chapman, October 2017
' 
' This is a class representing individual worker objects. Each stores
' their own name and number of messages and the class methods allow for
' calculation of the worker's pay and for updating of shared summary
' values. Name and messages are received as strings.
' This is being used as part of a piecework payroll application.

' This is currently incomplete; note the three large comment blocks
' below that begin with "TO DO"


Option Strict On

Public Class PieceworkWorker

#Region "Variable declarations"

    ' Instance variables
    Private employeeName As String
    Private employeeFirstName As String
    Private employeeLastName As String
    Private employeeMessages As Integer
    Private employeePay As Decimal
    Private employeeId As Integer
    Private creationDate As DateTime

    Private isValid As Boolean = True

    ' Shared class variables
    Private Shared overallNumberOfEmployees As Integer
    Private Shared overallMessages As Integer
    Private Shared overallPayroll As Decimal
    Private Shared result As Decimal


#End Region

#Region "Constructors"

    ''' <summary>
    ''' PieceworkWorker constructor: accepts a worker's name and number of
    ''' messages, sets and calculates values as appropriate.
    ''' </summary>
    ''' <param name="nameValue">a worker's name</param>
    ''' <param name="messagesValue">a worker's number of messages sent</param>
    Friend Sub New(ByVal nameValue As String, messagesValue As String)

        ' Validate and set the worker's name
        Me.FullName = nameValue
        ' Validate and set the worker's number of messages
        Me.Messages = messagesValue
        Dim result As Double = 0.0
        ' Calculcate the worker's pay and update all summary values
        findPay()

    End Sub

    ''' <summary>
    ''' PieceworkWorker constructor: accepts a worker's first and last name and number of
    ''' messages, sets and calculates values as appropriate.
    ''' </summary>
    ''' <param name="firstNameValue">a worker's first name</param>
    ''' <param name="lastNameValue">a worker's last name</param>
    ''' <param name="messagesValue">a worker's number of messages sent</param>
    Friend Sub New(ByVal firstNameValue As String, ByVal lastNameValue As String, messagesValue As String)

        ' Validate and set the worker's name
        Me.FirstName = firstNameValue
        Me.LastName = lastNameValue
        ' Validate and set the worker's number of messages
        Me.Messages = messagesValue
        ' Calculate the worker's pay and update all summary values
        findPay()
    End Sub
    ''' <summary>
    ''' PieceworkWorker constructor: empty constructor used strictly for
    ''' inheritance and instantiation
    ''' </summary>
    Friend Sub New()

    End Sub

#End Region

#Region "Class methods"

    ''' <summary>
    ''' Currently called in the constructor, the findPay() method is
    ''' used to calculate a worker's pay using threshold values to
    ''' change how much a worker is paid per message. This also updates
    ''' all summary values.
    ''' </summary>
    Private Sub findPay()

        ' TO DOTTT
        ' Fill in this entire method by following the instructions provided
        ' in the NETD 3202 Lab 1 handout
        ' It is suggested that you use the requirements as a checklist in
        ' order to ensure you don't miss any requirements.
        ' Only calculate if the data entered was valid (otherwise, pay is 0)
        If isValid Then
            If (employeeMessages <= 2499) Then
                employeePay = CDec((employeeMessages * 0.022))
            ElseIf (employeeMessages >= 2500 And employeeMessages <= 4999) Then
                employeePay = CDec((employeeMessages * 0.024))
            ElseIf (employeeMessages >= 5000 And employeeMessages <= 7499) Then
                employeePay = CDec((employeeMessages * 0.027))
            ElseIf (employeeMessages >= 7500 And employeeMessages <= 1000) Then
                employeePay = CDec((employeeMessages * 0.031))
            ElseIf (employeeMessages > 1000 And employeeMessages <= 2000) Then 'Set Upper Bouund value
                employeePay = CDec((employeeMessages * 0.035))
            End If
        End If
        overallMessages += employeeMessages
        overallPayroll += employeePay

        creationDate = Now()
        DBL.InsertNewRecord(Me)
    End Sub

    ''' <summary>
    ''' This returns a list of workers as a DataView object
    ''' for use with a DataGrid at the presentation level
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function ReturnAllWorkers() As Object

        Return DBL.GetEmployeeList

    End Function
#End Region

#Region "Property Procedures"

    ''' <summary>
    ''' Gets and sets a worker's first name
    ''' </summary>
    ''' <returns>an employee's first name</returns>

    Friend Property FirstName() As String
        Get
            Return employeeFirstName
        End Get
        Set(nameValue As String)

            ' Add validation; this cannot be blank
            employeeLastName = nameValue

        End Set
    End Property

    ''' <summary>
    ''' Gets and sets a worker's last name
    ''' </summary>
    ''' <returns>an employee's last name</returns>
    Friend Property LastName() As String
        Get
            Return employeeLastName
        End Get
        Set(nameValue As String)

            ' Add validation; this cannot be blank
            employeeLastName = nameValue

        End Set
    End Property

    Friend Property FullName() As String
        Get
            Return employeeFirstName & " " & employeeLastName
        End Get
        Set(nameValue As String)

            ' Declare the splitNames array and set it equal to all names entered separated by spaces
            Dim splitNames As String() = nameValue.Split(" "c)
            ' firstNameValue is equal to the first name entered (separated by spaces)
            Dim firstNameValue As String = splitNames(0)
            ' firstNameValue is equal to the last name entered (separated by spaces)
            Dim lastNameValue As String = splitNames(splitNames.Length - 1)

            ' If there was only one (or fewer) names entered, throw an exception
            If splitNames.Length < 1 Then
                Dim ex As New ArgumentException("First and last names required", "name")
                Throw ex
            End If

            ' Concatenate middle names to the middle name
            For nameIndex As Integer = 1 To splitNames.Length - 2
                firstNameValue &= " " & splitNames(nameIndex)
            Next

            ' Set the first name, then set the last name, using existing Property Procedures
            Me.FirstName = firstNameValue
            Me.LastName = lastNameValue

        End Set
    End Property
    ''' <summary>
    ''' Gets and sets the number of messages sent by a worker
    ''' </summary>
    ''' <returns>an employee's number of messages</returns>
    Friend Property Messages() As String
        Get
            Return employeeMessages.ToString()
        End Get
        Set(messagesValue As String)

            ' TO DO
            ' Add validation for the number of messages based on the
            ' requirements document
            If Not Decimal.TryParse(messagesValue, result) Then
                isValid = False
                MessageBox.Show("Please enter number of message as a decimal value.", "Entry Error")
            Else
                employeeMessages = CInt(messagesValue)
            End If
        End Set
    End Property

    ''' <summary>
    ''' Pay(): Gets the worker's pay
    ''' </summary>
    ''' <returns>a worker's pay</returns>
    Friend ReadOnly Property Pay() As Decimal
        Get
            Return employeePay
        End Get
    End Property

    ''' <summary>
    ''' Gets the overall number of workers
    ''' </summary>
    ''' <returns>the overall number of workers</returns>
    Friend Shared ReadOnly Property TotalWorkers() As Integer
        Get
            Return overallNumberOfEmployees
        End Get
    End Property
    ''' <summary>
    ''' Gets the overall number of messages sent
    ''' </summary>
    ''' <returns>the overall number of messages sent</returns>
    Friend Shared ReadOnly Property TotalMessages() As Integer
        Get
            Return overallMessages
        End Get
    End Property
    ''' <summary>
    ''' Calculates and returns an average pay among all workers
    ''' </summary>
    ''' <returns>the average pay among all workers</returns>
    Friend Shared ReadOnly Property AveragePay() As Decimal
        Get
            If overallNumberOfEmployees = 0 Then
                Return 0
            Else
                Return overallPayroll / CDec(overallNumberOfEmployees)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Id(): Gets the worker's id
    ''' </summary>
    ''' <returns>a worker's id</returns>
    Friend Property Id() As Integer
        Get
            Return employeeId
        End Get
        Set(value As Integer)
            employeeId = value
        End Set
    End Property
    ''' <summary>
    ''' EntryDate(): Gets the worker's entry date
    ''' </summary>
    ''' <returns>a worker's pay</returns>
    Friend ReadOnly Property EntryDate() As DateTime
        Get
            Return creationDate
        End Get
    End Property
    ''' <summary>
    ''' TotalPay(): Gets the overall total pay among all workers
    ''' </summary>
    ''' <returns>the overall total pay among all workers</returns>
    Friend Shared ReadOnly Property TotalPay() As Decimal
        Get
            Return overallPayroll
        End Get
    End Property

#End Region

End Class
