'Fallon Smith
'RCET 0265
'Spring 2022
'Car Rental
'

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm

    Dim Miles As Integer
    Dim MitoKilo As Integer
    Dim Cost As Integer
    Dim Discount As Double
    Dim customers As Integer
    Dim miless As Integer
    Dim charge As Integer

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "Car Rental Information"
        TotalMilesTextBox.Enabled = False
        MileageChargeTextBox.Enabled = False
        DayChargeTextBox.Enabled = False
        TotalDiscountTextBox.Enabled = False
        TotalChargeTextBox.Enabled = False
        MilesradioButton.Checked = True


    End Sub

    Function ValidateInput() As Boolean
        Dim Validate As Boolean = True

        If NameTextBox.Text = "" Then
            AccumulateMessage("Name is empty")
            NameTextBox.Focus()
            Validate = False
        ElseIf CInt(NameTextBox.Text) = CInt(NameTextBox.Text) Then
            AccumulateMessage("Name can not be numbers")
            NameTextBox.Focus()
            Validate = False
        End If

        If AddressTextBox.Text = "" Then
            AccumulateMessage("Address cannot be empty")
            AddressTextBox.Focus()
            Validate = False
        End If

        If CityTextBox.Text = "" Then
            AccumulateMessage("City is empty")
            CityTextBox.Focus()
            Validate = False
        ElseIf CInt(CityTextBox.Text) = CInt(CityTextBox.Text) Then
            AccumulateMessage("City can not be numbers")
            CityTextBox.Focus()
            Validate = False
        End If

        If StateTextBox.Text = "" Then
            AccumulateMessage("State is empty")
            StateTextBox.Focus()
            Validate = False
        ElseIf CInt(StateTextBox.Text) = CInt(StateTextBox.Text) Then
            AccumulateMessage("State can not be numbers")
            StateTextBox.Focus()
            Validate = False
        End If

        Try
            ZipCodeTextBox.Text = CStr(CInt(ZipCodeTextBox.Text))

        Catch ex As Exception
            AccumulateMessage("Zipcode must be a number")
            ZipCodeTextBox.Focus()
            Validate = False

        End Try

        Try
            BeginOdometerTextBox.Text = CStr(CInt(BeginOdometerTextBox.Text))
            EndOdometerTextBox.Text = CStr(CInt(EndOdometerTextBox.Text))
            Miles = CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)
            If BeginOdometerTextBox.Text > EndOdometerTextBox.Text Then
                AccumulateMessage("Ending Odometer must be greater then start milage")
                BeginOdometerTextBox.Focus()
                Validate = False

            End If

        Catch ex As Exception
            Validate = False
            AccumulateMessage("Odometers must read as a number")
            BeginOdometerTextBox.Focus()
        End Try

        Select Case CInt(DaysTextBox.Text)
            Case 1 To 45
                Validate = True
            Case Else
                AccumulateMessage("Days must be between 1 - 45")

        End Select

        Return Validate
    End Function
    Private Function AccumulateMessage(Optional NewMessage As String = "", Optional clear As Boolean = False) As String
        Static _message As String

        Select Case clear
            Case False
                If NewMessage <> "" Then
                    _message &= NewMessage & vbCrLf
                End If
            Case Else
                _message = ""
        End Select

        Return _message

    End Function

    Function totalMiles() As Double
        Dim milecost As Double = 0
        If KilometersradioButton.Checked = True Then
            MitoKilo = CInt(Miles * 0.62)
        End If

        Select Case Miles
            Case 0 To 200

            Case 201 To 500
                milecost += Miles * 0.12
            Case Else

                milecost += Miles * 0.1
        End Select


        Return milecost


    End Function


    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim close As Integer = 0
        close = Msgbox("Are you sure would like to exit?", msgboxstyle.question.YesNo)

        
        
        If close = 6 Then 

         Me.close

        End If
        
        End Sub



    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False

    End Sub

    Private Sub Seniorcheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles Seniorcheckbox.CheckedChanged
        Discount = 0.03
    End Sub

    Private Sub AAAcheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles AAAcheckbox.CheckedChanged
        Discount = 0.05
    End Sub
    Function discounts() As Double
        If Seniorcheckbox.Checked And AAAcheckbox.Checked Then
            Discount = 0.08
        End If
    End Function
    Function CalculateCost() As Integer
        Dim errr As Integer
        Cost = 0
        Cost += CInt(DaysTextBox.Text) * 15
        Cost += CInt(totalMiles())
        Cost = CInt(Cost - (Cost * Discount))
        errr = MsgBox($"Your information is :" + vbCrLf +
                      $"Name: {NameTextBox.Text} " + vbCrLf +
                      $"Address: {AddressTextBox.Text}" + vbCrLf +
                      $"City: {CityTextBox.Text} " + vbCrLf +
                      $"State: {StateTextBox.Text} " + vbCrLf +
                      $"Zipcode: {ZipCodeTextBox.Text}" + vbCrLf +
                      $"{TotalMilesTextBox.Text} miles over {DaysTextBox.Text} day(s) using the {AAAcheckbox}   {Seniorcheckbox} discount applied " + vbCrLf +
                      "If the informaion above is correct, Press yes to submit your information ", MsgBoxStyle.Question.YesNo, "Submit form")

        If errr = 6 Then
            summery()
        End If
        Return Cost
    End Function
    Sub DisplayCharge()
        TotalMilesTextBox.Text = CStr(Miles)
        MileageChargeTextBox.Text = CStr(CInt(totalMiles()))
        DayChargeTextBox.Text = CStr(CInt(DaysTextBox.Text) * 15)
        TotalDiscountTextBox.Text = CStr(Discount)
        TotalChargeTextBox.Text = CStr(Cost)

    End Sub
    Sub Summery()
        customers += 1 + customers
        miless += Miles + miless
        charge += Cost + charge
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If Validate() = False Then
            ValidateInput()
        Else
            CalculateCost()
            DisplayCharge()
        End If
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        MsgBox($"Total Customers : {customers} " + vbCrLf +
               $"Total miles Driven: {miless} mi " + vbCrLf +
               $"Total charge: ${charge}" + vbCrLf +
               "would you like to clear this history?", MsgBoxStyle.Question.YesNo, "Clear Form")

    End Sub
End Class
