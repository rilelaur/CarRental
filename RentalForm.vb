'Laura Riley
'RCET 0265
'Spring 2021
'Car Rental
'https://github.com/rilelaur/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Dim customertotal As Integer
    Dim totalCharges As Double
    Dim totalDistanceDriven As Integer = 0

    'closes the form
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        MsgBox("Would you like to exit?", CType(vbYesNo, MsgBoxStyle))
        If CBool(DialogResult.Yes) Then
            Me.Close()
        End If
    End Sub
    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        MsgBox("Would you like to exit?", CType(vbYesNo, MsgBoxStyle))
        If CBool(DialogResult.Yes) Then
            Me.Close()
        End If
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

    'clears the program
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
        Static daysCharge As Decimal
        Static _daysCharge As Decimal
        Static beginingOdometerReading As Integer
        Static endingOdometerReading As Integer
        Static totalMilesDriven As Double
        Static milageCharge As Double
        Static _milageCharge As Decimal
        Static totalCharge As Double
        Static _totalCharge As Decimal
        Static discount As Double
        Static _discount As Decimal
        'Static customertotal As Integer
        'Static totalCharges As Double
        'Static totalDistanceDriven As Integer = 0
        Static problem As Boolean = True

        'validates that all textboxes have something in them
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

        'trys to convert the number of days to a number
        Try
            totalDays = Convert.ToInt32(DaysTextBox.Text)
        Catch ex As Exception
            MsgBox("Please enter a number between 1 and 45")
            DaysTextBox.Focus()
            problem = True
        End Try

        'if it passed the try catch it will then charge $15 per day if the number 
        'is between 1 and 45 if not it will prompt the user to enter a valid number
        If totalDays > 0 And totalDays <= 45 Then
            daysCharge = totalDays * 15
            _daysCharge = Math.Round(daysCharge, 2)
            DayChargeTextBox.Text = CStr($"${_daysCharge}")
        Else
            MsgBox("Please enter a number between 1 and 45")
            DaysTextBox.Focus()
            problem = True
        End If

        'trys to convert the begining odometer reading to a number
        Try
            beginingOdometerReading = Convert.ToInt32(BeginOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Please enter a valid number for your begining odometer reading.")
            problem = True
        End Try

        'trys to convert the ending odometer reading to a number
        Try
            endingOdometerReading = Convert.ToInt32(EndOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Please enter a valid number for your end odometer reading.")
            problem = True
        End Try

        'checks to make sure that then begining odometer reading is less than the ending odometer reading
        'if not then it will prompt the user to fix that
        'if the kilometers radiobutton is selected it will then convert km to miles
        If beginingOdometerReading < endingOdometerReading And MilesradioButton.Checked Then
            totalMilesDriven = (endingOdometerReading - beginingOdometerReading)
            TotalMilesTextBox.Text = CStr(totalMilesDriven) & " mi"
        ElseIf beginingOdometerReading < endingOdometerReading And KilometersradioButton.Checked Then
            totalMilesDriven = ((endingOdometerReading - beginingOdometerReading) / 1.609)
            TotalMilesTextBox.Text = CStr(totalMilesDriven) & " mi"
        Else
            MsgBox("Make sure that the ending odometer reading is greater than the begining odometer reading")
            problem = True
        End If

        'charges the proper amount based off of the total miles driven
        If totalMilesDriven < 201 Then
            MileageChargeTextBox.Text = ("$0.00")
        ElseIf totalMilesDriven > 200 And totalMilesDriven < 501 Then
            milageCharge = CDbl(totalMilesDriven - 200) * 0.12
            _milageCharge = CDec(Math.Round(milageCharge, 2))
            MileageChargeTextBox.Text = CStr($"${_milageCharge}")
        ElseIf totalMilesDriven > 500 Then
            milageCharge = ((totalMilesDriven - 200) * 0.1)
            _milageCharge = CDec(Math.Round(milageCharge, 2))
            MileageChargeTextBox.Text = CStr($"${_milageCharge}")
        End If

        'checks to see if checkboxes were check and applies the discount accordingly
        If AAAcheckbox.Checked And Seniorcheckbox.Checked Then
            totalCharge = (milageCharge + daysCharge) - (milageCharge + daysCharge) * 0.08
            _totalCharge = CDec(Math.Round(totalCharge, 2))
            discount = (milageCharge + daysCharge) * 0.08
            _discount = CDec(Math.Round(discount, 2))
            TotalDiscountTextBox.Text = CStr($"${_discount}")
            TotalChargeTextBox.Text = CStr($"${_totalCharge}")
        ElseIf AAAcheckbox.Checked Then
            totalCharge = (milageCharge + daysCharge) - (milageCharge + daysCharge) * 0.05
            _totalCharge = CDec(Math.Round(totalCharge, 2))
            discount = (milageCharge + daysCharge) * 0.05
            _discount = CDec(Math.Round(discount, 2))
            TotalDiscountTextBox.Text = CStr($"${_discount}")
            TotalChargeTextBox.Text = CStr($"${_totalCharge}")
        ElseIf Seniorcheckbox.Checked Then
            totalCharge = (milageCharge + daysCharge) - (milageCharge + daysCharge) * 0.03
            _totalCharge = CDec(Math.Round(totalCharge, 2))
            discount = (milageCharge + daysCharge) * 0.03
            _discount = CDec(Math.Round(discount, 2))
            TotalDiscountTextBox.Text = CStr($"${_discount}")
            TotalChargeTextBox.Text = CStr($"${_totalCharge}")
        Else
            totalCharge = milageCharge + daysCharge
            _totalCharge = CDec(Math.Round(totalCharge, 2))
            TotalDiscountTextBox.Text = "$0.00"
            TotalChargeTextBox.Text = CStr($"${_totalCharge}")
        End If

        If TotalMilesTextBox.Text = "" Then
            problem = True
        ElseIf MileageChargeTextBox.Text = "" Then
            problem = True
        ElseIf DayChargeTextBox.Text = "" Then
            problem = True
        ElseIf TotalDiscountTextBox.Text = "" Then
            problem = True
        ElseIf TotalChargeTextBox.Text = "" Then
            problem = True
        Else
            problem = False
            customertotal += 1
            totalCharges += 1
            totalDistanceDriven = CInt(totalMilesDriven) + totalDistanceDriven
        End If

        If problem = False Then
            SummaryButton.Enabled = True
        End If

    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        MsgBox("Total Customers: " & customertotal & vbCrLf & "Total Charges: " & totalcharges & vbCrLf & "Total miles Driven: " & Totaldistancedriven & " miles")
    End Sub

End Class
