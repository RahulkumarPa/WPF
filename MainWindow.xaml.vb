'MainWindow.xaml.vb
'Title: IncInc Payroll (Piecework)
' Last Modified: 18-10-2018
'  Written By: Rahulkumar 
' This is a main class of Incinc Payroll system In which I call the class
' PieceWorker from the file PieceworkWorker.vb.
Class MainWindow

    Dim lastClickTab As TabItem

    'event for the button Calculate PAy
    Private Sub btnCalculatedPay_Click(sender As Object, e As RoutedEventArgs) Handles btnCalculatedPay.Click

        'unhighlight textbox
        unhighlightTextbox(txtWorkerName)
        unhighlightTextbox(txtPiecesProduced)

        txtWorkerName.BorderBrush = Brushes.Red
        txtPiecesProduced.BorderBrush = Brushes.Red

        Dim worker As New PieceworkWorker(txtWorkerName.Text, txtPiecesProduced.Text)    'set user input with the class PieceWorker

        Try
            ' Output the worker's calculated pay to the form
            lblPay.Content = worker.Pay().ToString("c")

            ' Disable all input controls other than the Clear button, so as to force a Clear before further data entry
            txtWorkerName.IsEnabled = False
            txtPiecesProduced.IsEnabled = False
            btnCalculatedPay.IsEnabled = False
            'write worker to the status bar
            statusUpdate("worker " & worker.FirstName & " created, with pay of " & worker.Pay.ToString("c"))

            ' Set focus to Clear button
            btnClear.Focus()

        Catch exx As ArgumentOutOfRangeException
            ' Put an error message in the label
            lblError_Name.Content = exx.Message
            lblError_Value.Content = exx.Message

            'Display which field have an error
            If (exx.ParamName = "name") Then
                highlightTextbox(txtWorkerName)
            ElseIf (exx.ParamName = "messages") Then
                highlightTextbox(txtPiecesProduced)
            End If

        Catch ex As ArgumentException
            ' Put an error message in the label
            lblError_Name.Content = ex.Message
            lblError_Value.Content = ex.Message
            If (ex.ParamName = "name") Then
                highlightTextbox(txtWorkerName)
            ElseIf (ex.ParamName = "messages") Then
                highlightTextbox(txtPiecesProduced)
            End If

        Catch ex As Exception
            ' Catch the general exception And return a generic (exhaustive!) error message
            MessageBox.Show("A critical error has occured! Please contact your IT department and provide the following information:" + Environment.NewLine + Environment.NewLine &
                            ex.Message + Environment.NewLine + Environment.NewLine &
                            ex.Source + Environment.NewLine + Environment.NewLine &
                            ex.StackTrace, "Flagrant Error!")
            lblError_Name.Content = "Error!" + ex.Message
            lblError_Value.Content = "Error!" + ex.Message

        End Try
    End Sub




    ' 'If user's message is a negative number it's shoe s error Else it's dispaly user's pa in lblPay
    ' If (worker.Messages > 0) Then
    ' Output the worker's calculated pay to the form
    ' lblPay.Content = worker.Pay().ToString("c")
    'Else
    'MessageBox.Show("Please enter number of message as a decimal value.", "Entry Error")
    'End If

    ' Disable all input controls other than the Clear button, so as to force a Clear before further data entry
    'txtWorkerName.IsEnabled = False
    'txtPiecesProduced.IsEnabled = False
    'btnCalculatedPay.IsEnabled = False
    ' Set focus to Clear button
    'btnClear.Focus()


    'event for the button Clear
    Private Sub btnClear_Click(sender As Object, e As RoutedEventArgs) Handles btnClear.Click
        setDefaults() 'Call the function of the method setDefaults()

        statusUpdate(" Cleared all payroll entry fields")

    End Sub

    'method for the setDefaults
    Private Sub setDefaults()

        ' Clear all user Inpput
        txtWorkerName.Text = " "
        txtPiecesProduced.Text = " "
        lblPay.Content = ""

        ' Enable all input controls
        txtWorkerName.IsEnabled = True
        txtPiecesProduced.IsEnabled = True
        btnCalculatedPay.IsEnabled = True


        ' Set focus to the first control
        txtWorkerName.Focus()


    End Sub

    Private Sub highlightTextbox(thisBox As TextBox)
        thisBox.BorderBrush = Brushes.Red
        thisBox.SelectAll()
        thisBox.Focus()
    End Sub

    Private Sub unhighlightTextbox(thisBox As TextBox)
        thisBox.BorderBrush = Brushes.Black
    End Sub

    Private Sub statusUpdate(updateMessage As String)
        lblStatus.Content = Date.Now().ToString() & ":" & updateMessage
    End Sub

    Private Sub txtWorkerName_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtWorkerName.TextChanged

    End Sub

    'for a summary tab

    Private Sub refreshViews(sender As Object, e As RoutedEventArgs) Handles tbcNavigation.SelectionChanged

        If (tbcNavigation.SelectedItem Is tbiPayrollEntry And lastClickTab IsNot tbiPayrollEntry) Then

            statusUpdate("Viewed Payroll Entry")

        ElseIf (tbcNavigation.SelectedItem Is tbiSummary And lastClickTab IsNot tbiSummary) Then
            'Set all label's value form PieceworkWorker
            lblNumberOfWorkersDisplay.Content = PieceworkWorker.TotalWorkers.ToString()
            lblTotalMessagesDisplay.Content = PieceworkWorker.TotalMessages.ToString()
            lblTotalPayDisplay.Content = PieceworkWorker.TotalPay.ToString("c")
            lblAveragePayDisplay.Content = PieceworkWorker.AveragePay.ToString("c")

            statusUpdate("Viewed Summary Tab")

        ElseIf (tbcNavigation.SelectedItem Is tbiEmployeeList And lastClickTab IsNot tbiEmployeeList) Then
            dgWorkerGrid.ItemsSource = PieceworkWorker.ReturnAllWorkers

            statusUpdate("Viewed Employee list")

        End If

        lastClickTab = CType(tbcNavigation.SelectedItem, TabItem)

    End Sub

    ''' <summary>
    ''' for a employe list view or a grid vie
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub viewEmployeeList(sender As Object, e As EventArgs) Handles tbiEmployeeList.GotFocus, Me.Loaded

        dgWorkerGrid.ItemsSource = PieceworkWorker.ReturnAllWorkers

    End Sub
    Private Sub dgWorkerGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgWorkerGrid.SelectionChanged

        Try
            If dgWorkerGrid.SelectedItem IsNot Nothing Then
                txtEmployeeIDPrompt.Text = dgWorkerGrid.SelectedItem.Row.ItemArray(0)
                txtFirstName.Text = dgWorkerGrid.SelectedItem.Row.ItemArray(1)
                txtLastName.Text = dgWorkerGrid.SelectedItem.Row.ItemArray(2)
                txtMessages.Text = dgWorkerGrid.SelectedItem.Row.ItemArray(3)
                txtTotalPay.Text = dgWorkerGrid.SelectedItem.Row.ItemArray(4)
                txtEntryDate.Text = dgWorkerGrid.SelectedItem.Row.ItemArray(5)

                statusUpdate("View details for workeer" & txtFirstName.Text & "( ID #" & txtEmployeeIDPrompt.Text & ")")

            End If

        Catch ex As Exception
            MessageBox.Show("Please select a valid worker for display.", "No worker Selected")

        End Try
    End Sub

    'To close the application Exit button
    Private Sub btnExit_Click(sender As Object, e As RoutedEventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class
