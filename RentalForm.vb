'Laura Riley
'RCET 0265
'Spring 2021
'Car Rental
'https://github.com/rilelaur/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    'closes the form
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub
    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    'clears the form when the clear button is selected
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()

        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()

        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub

    Private Sub ClearToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem1.Click
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()

        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()

        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click

        Static totalDays As Integer
        Static daysCharge As Double
        Static beginingOdometerReading As Integer
        Static endingOdometerReading As Integer
        Static totalMilesDriven As Double
        Static milageCharge As Double
        Static problem As Boolean = True

        If NameTextBox.Text = "" Then
            MsgBox("Please enter your name")
            NameTextBox.Focus()
            problem = True
        ElseIf AddressTextBox.Text = "" Then
            MsgBox("Please enter your address")
            AddressTextBox.Focus()
            problem = True
        ElseIf CityTextBox.Text = "" Then
            MsgBox("Please enter your city")
            CityTextBox.Focus()
            problem = True
        ElseIf StateTextBox.Text = "" Then
            MsgBox("Please enter your state")
            StateTextBox.Focus()
            problem = True
        ElseIf ZipCodeTextBox.Text = "" Then
            MsgBox("Please enter your zipcode")
            ZipCodeTextBox.Focus()
            problem = True
        ElseIf BeginOdometerTextBox.Text = "" Then
            MsgBox("Please enter a number")
            BeginOdometerTextBox.Focus()
            problem = True
        ElseIf EndOdometerTextBox.Text = "" Then
            MsgBox("Please enter a number")
            EndOdometerTextBox.Focus()
            problem = True
        ElseIf DaysTextBox.Text = "" Then
            MsgBox("Please enter your address")
            DaysTextBox.Focus()
            problem = True
        End If

        Try
            totalDays = Convert.ToInt32(DaysTextBox.Text)
        Catch ex As Exception
            MsgBox("Please enter a number between 1 and 45")
            DaysTextBox.Focus()
            problem = True
        End Try

        If totalDays > 0 And totalDays <= 45 Then
            daysCharge = totalDays * 15
            DayChargeTextBox.Text = CStr(daysCharge)
        Else
            MsgBox("Please enter a number between 1 and 45")
            DaysTextBox.Focus()
            problem = True
        End If

        Try
            beginingOdometerReading = Convert.ToInt32(BeginOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Please enter a valid number for your begining odometer reading.")
            problem = True
        End Try

        Try
            endingOdometerReading = Convert.ToInt32(EndOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Please enter a valid number for your end odometer reading.")
            problem = True
        End Try

        If beginingOdometerReading < endingOdometerReading And MilesradioButton.Checked Then
            totalMilesDriven = (endingOdometerReading - beginingOdometerReading)
            TotalMilesTextBox.Text = CStr(totalMilesDriven) & " Miles"
        ElseIf beginingOdometerReading < endingOdometerReading And KilometersradioButton.Checked Then
            totalMilesDriven = ((endingOdometerReading - beginingOdometerReading) / 1.609)
            TotalMilesTextBox.Text = CStr(totalMilesDriven) & " Miles"
        Else
            MsgBox("Make sure that the ending odometer reading is greater than the begining odometer reading")
            problem = True
        End If

        If totalMilesDriven < 201 Then
            MileageChargeTextBox.Text = CStr(0)
        ElseIf totalMilesDriven > 200 And totalMilesDriven < 501 Then
            milageCharge = ((totalMilesDriven - 200) * 0.12)
            MileageChargeTextBox.Text = CStr(milageCharge)
        ElseIf totalMilesDriven > 500 Then
            milageCharge = ((totalMilesDriven - 200) * 0.1)
            MileageChargeTextBox.Text = CStr(milageCharge)
        End If

        If problem = False Then
            SummaryButton.Enabled = True
        End If
    End Sub
End Class
