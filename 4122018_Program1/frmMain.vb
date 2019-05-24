Imports System.IO

Public Class frmMain
    'Variables that keep track of the number of tests, vehicles tested, total miles travelled, total fuel cells used, total cost to be displayed in the summary form
    Dim numOfTests As Integer = 0
    Dim vehiclesTested As Integer = 0
    Dim totalMiles As Integer = 0
    Dim totalFuelCellsUsed As Integer = 0
    Dim totalCost As Integer = 0

    'Variable that keeps the cost per fuel cell
    Private Const PER_FUEL_CELL_COST = 22

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        'Clears test info group box controls
        For Each Control As Control In Me.grpTestInfo.Controls
            If TypeOf (Control) Is TextBox Then
                TryCast(Control, TextBox).Text = ""
            End If
        Next

        'Clears inventor info group box controls
        For Each Control As Control In Me.grpInventorInfo.Controls
            If TypeOf (Control) Is TextBox Then
                TryCast(Control, TextBox).Text = ""
            End If
        Next

        'Clears driver info group box controls
        For Each Control As Control In Me.grpDriverInfo.Controls
            If TypeOf (Control) Is TextBox Then
                TryCast(Control, TextBox).Text = ""
            ElseIf TypeOf (Control) Is ComboBox Then
                TryCast(Control, ComboBox).SelectedIndex = -1
            ElseIf TypeOf (Control) Is PictureBox Then
                TryCast(Control, PictureBox).Image = Nothing
            End If
        Next

        'Clears data panel controsl
        For Each Control As Control In Me.pnlData.Controls
            If TypeOf (Control) Is TextBox Then
                TryCast(Control, TextBox).Text = ""
            ElseIf TypeOf (Control) Is ComboBox Then
                TryCast(Control, ComboBox).SelectedIndex = -1
            End If
        Next

        'Clears results panel controls
        For Each Control As Control In Me.pnlResults.Controls
            If TypeOf (Control) Is TextBox Then
                TryCast(Control, TextBox).Text = ""
            End If
        Next

    End Sub

    Private Sub cmbID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbID.SelectedIndexChanged

        'Adds driver information to the appropiate text boxes and picture box depending on what ID the user selects
        If cmbID.SelectedIndex = 0 Then
            txtDriverFirstName.Text = "Bart"
            txtDriverLastName.Text = "Simpson"
            txtDriverPhone.Text = "909-888-7777"
            pctDriverImage.Image = Image.FromFile(Directory.GetCurrentDirectory() & "\images\100" & ".png")
        ElseIf cmbID.SelectedIndex = 1 Then
            txtDriverFirstName.Text = "Homer"
            txtDriverLastName.Text = "Simpson"
            txtDriverPhone.Text = "909-666-5555"
            pctDriverImage.Image = Image.FromFile(Directory.GetCurrentDirectory() & "\images\200" & ".jpg")
        ElseIf cmbID.SelectedIndex = 2 Then
            txtDriverFirstName.Text = "Marge"
            txtDriverLastName.Text = "Simpson"
            txtDriverPhone.Text = "909-111-3333"
            pctDriverImage.Image = Image.FromFile(Directory.GetCurrentDirectory() & "\images\300" & ".jpg")
        ElseIf cmbID.SelectedIndex = 3 Then
            txtDriverFirstName.Text = "Lisa"
            txtDriverLastName.Text = "Simpson"
            txtDriverPhone.Text = "909-333-6666"
            pctDriverImage.Image = Image.FromFile(Directory.GetCurrentDirectory() & "\images\400" & ".jpg")
        End If
    End Sub


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        'On form load add items to the ID combo box
        cmbID.Items.Add(100)
        cmbID.Items.Add(200)
        cmbID.Items.Add(300)
        cmbID.Items.Add(400)

        'On form load add items to the first Vehicle combo box
        cmbVehicleOne.Items.Add("Ghost")
        cmbVehicleOne.Items.Add("Banshee")
        cmbVehicleOne.Items.Add("Hornet")
        cmbVehicleOne.Items.Add("Chopper")

        'On form load add items to the second Vehicle combo box
        cmbVehicleTwo.Items.Add("Ghost")
        cmbVehicleTwo.Items.Add("Banshee")
        cmbVehicleTwo.Items.Add("Hornet")
        cmbVehicleTwo.Items.Add("Chopper")

        'On form load add items to the third Vehicle combo box
        cmbVehicleThree.Items.Add("Ghost")
        cmbVehicleThree.Items.Add("Banshee")
        cmbVehicleThree.Items.Add("Hornet")
        cmbVehicleThree.Items.Add("Chopper")

        'On form load add items to the fourth Vehicle combo box
        cmbVehicleFour.Items.Add("Ghost")
        cmbVehicleFour.Items.Add("Banshee")
        cmbVehicleFour.Items.Add("Hornet")
        cmbVehicleFour.Items.Add("Chopper")

    End Sub

    Private Sub btnProcess_Click(sender As Object, e As EventArgs) Handles btnProcess.Click

        'Check form before we proceed
        If (isValidForm() IsNot "") Then
            MessageBox.Show(isValidForm())
            Exit Sub
        End If

        'Loop that goes through the four vehicle combo boxes and retrieves the data to do the caculations
        'It then displays the calcuations to the appropiate text boxes in the RESULTS panel
        For Each ctrl As Control In Me.pnlData.Controls
            If TypeOf (ctrl) Is ComboBox Then
                If (TryCast(ctrl, ComboBox).SelectedIndex = 0) Then
                    'Displays Vehicle Name
                    txtVehicleOne.Text = cmbVehicleOne.SelectedItem.ToString

                    'Displays Vehicle Family type
                    txtFamilyOne.Text = family(cmbVehicleOne.SelectedIndex)

                    'Displays Mileage
                    txtMileageOne.Text = FormatNumber(calcMileage(CInt(txtStartMileageOne.Text), CInt(txtEndMileageOne.Text)).ToString)

                    'Displays Fuel Used
                    txtFuelUsedOne.Text = FormatNumber(calcFuelUsed(CInt(txtStartFuelOne.Text), CInt(txtEndFuelOne.Text)).ToString)

                    'Displays Fuel Cost
                    txtFuelCostOne.Text = FormatCurrency(calcFuelCost(calcFuelUsed(CInt(txtStartFuelOne.Text), CInt(txtEndFuelOne.Text))).ToString)

                    'Displays MPFC Rating
                    txtMPFCRatingOne.Text = calcMPFC(calcMileage(CInt(txtStartMileageOne.Text), CInt(txtEndMileageOne.Text)), calcFuelUsed(CInt(txtStartFuelOne.Text), CInt(txtEndFuelOne.Text)))

                    'Incremeting variables used in Summary form
                    vehiclesTested += 1
                    totalMiles += calcMileage(CInt(txtStartMileageOne.Text), CInt(txtEndMileageOne.Text))
                    totalFuelCellsUsed += calcFuelUsed(CInt(txtStartFuelOne.Text), CInt(txtEndFuelOne.Text))
                    totalCost += calcFuelCost(calcFuelUsed(CInt(txtStartFuelOne.Text), CInt(txtEndFuelOne.Text)))
                ElseIf (TryCast(ctrl, ComboBox).SelectedIndex = 1) Then
                    'Displays Vehicle Name
                    txtVehicleTwo.Text = cmbVehicleTwo.SelectedItem.ToString

                    'Displays Vehicle Family type
                    txtFamilyTwo.Text = family(cmbVehicleTwo.SelectedIndex)

                    'Displays Mileage
                    txtMileageTwo.Text = FormatNumber(calcMileage(CInt(txtStartMileageTwo.Text), CInt(txtEndMileageTwo.Text)).ToString)

                    'Displays Fuel Used
                    txtFuelUsedTwo.Text = FormatNumber(calcFuelUsed(CInt(txtStartFuelTwo.Text), CInt(txtEndFuelTwo.Text)).ToString)

                    'Displays Fuel Cost
                    txtFuelCostTwo.Text = FormatCurrency(calcFuelCost(calcFuelUsed(CInt(txtStartFuelTwo.Text), CInt(txtEndFuelTwo.Text))).ToString)

                    'Displays MPFC Rating
                    txtMPFCRatingTwo.Text = calcMPFC(calcMileage(CInt(txtStartMileageTwo.Text), CInt(txtEndMileageTwo.Text)), calcFuelUsed(CInt(txtStartFuelTwo.Text), CInt(txtEndFuelTwo.Text)))

                    'Incremeting variables used in Summary form
                    vehiclesTested += 1
                    totalMiles += calcMileage(CInt(txtStartMileageTwo.Text), CInt(txtEndMileageTwo.Text))
                    totalFuelCellsUsed += calcFuelUsed(CInt(txtStartFuelTwo.Text), CInt(txtEndFuelTwo.Text))
                    totalCost += calcFuelCost(calcFuelUsed(CInt(txtStartFuelTwo.Text), CInt(txtEndFuelTwo.Text)))
                ElseIf (TryCast(ctrl, ComboBox).SelectedIndex = 2) Then
                    'Displays Vehicle Name
                    txtVehicleThree.Text = cmbVehicleThree.SelectedItem.ToString

                    'Displays Vehicle Family type
                    txtFamilyThree.Text = family(cmbVehicleThree.SelectedIndex)

                    'Displays Mileage
                    txtMileageThree.Text = FormatNumber(calcMileage(CInt(txtStartMileageThree.Text), CInt(txtEndMileageThree.Text)).ToString)

                    'Displays Fuel Used
                    txtFuelUsedThree.Text = FormatNumber(calcFuelUsed(CInt(txtStartFuelThree.Text), CInt(txtEndFuelThree.Text)).ToString)

                    'Displays Fuel Cost
                    txtFuelCostThree.Text = FormatCurrency(calcFuelCost(calcFuelUsed(CInt(txtStartFuelThree.Text), CInt(txtEndFuelThree.Text))).ToString)

                    'Displays MPFC Rating
                    txtMPFCRatingThree.Text = calcMPFC(calcMileage(CInt(txtStartMileageThree.Text), CInt(txtEndMileageThree.Text)), calcFuelUsed(CInt(txtStartFuelThree.Text), CInt(txtEndFuelThree.Text)))

                    'Incremeting variables used in Summary form
                    vehiclesTested += 1
                    totalMiles += calcMileage(CInt(txtStartMileageThree.Text), CInt(txtEndMileageThree.Text))
                    totalFuelCellsUsed += calcFuelUsed(CInt(txtStartFuelThree.Text), CInt(txtEndFuelThree.Text))
                    totalCost += calcFuelCost(calcFuelUsed(CInt(txtStartFuelThree.Text), CInt(txtEndFuelThree.Text)))
                ElseIf (TryCast(ctrl, ComboBox).SelectedIndex = 3) Then
                    'Displays Vehicle Name
                    txtVehicleFour.Text = cmbVehicleFour.SelectedItem.ToString

                    'Displays Vehicle Family type
                    txtFamilyFour.Text = family(cmbVehicleFour.SelectedIndex)

                    'Displays Mileage
                    txtMileageFour.Text = FormatNumber(calcMileage(CInt(txtStartMileageFour.Text), CInt(txtEndMileageFour.Text)).ToString)

                    'Displays Fuel Used
                    txtFuelUsedFour.Text = FormatNumber(calcFuelUsed(CInt(txtStartFuelFour.Text), CInt(txtEndFuelFour.Text)).ToString)

                    'Displays Fuel Cost
                    txtFuelCostFour.Text = FormatCurrency(calcFuelCost(calcFuelUsed(CInt(txtStartFuelFour.Text), CInt(txtEndFuelFour.Text))).ToString)

                    'Displays MPFC Rating
                    txtMPFCRatingFour.Text = calcMPFC(calcMileage(CInt(txtStartMileageFour.Text), CInt(txtEndMileageFour.Text)), calcFuelUsed(CInt(txtStartFuelFour.Text), CInt(txtEndFuelFour.Text)))

                    'Incremeting variables used in Summary form
                    vehiclesTested += 1
                    totalMiles += calcMileage(CInt(txtStartMileageFour.Text), CInt(txtEndMileageFour.Text))
                    totalFuelCellsUsed += calcFuelUsed(CInt(txtStartFuelFour.Text), CInt(txtEndFuelFour.Text))
                    totalCost += calcFuelCost(calcFuelUsed(CInt(txtStartFuelFour.Text), CInt(txtEndFuelFour.Text)))
                End If
            End If
        Next

        'Increment the number of tests variable by one
        numOfTests += 1

    End Sub

    'Function that validates user entries
    Private Function isValidForm() As String
        'Validate
        Dim result As String = ""

        'Validates test # text box
        If (txtTestNum.TextLength = 0) Then
            result &= "Test # for vehicle one not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtTestNum.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Test # for vehicle one has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle one Test #, Please re-enter whole number" & vbCrLf
            End If
        End If


        'Validates date text box
        If (txtDate.TextLength = 0) Then
            result &= "Date not entered, Please enter" & vbCrLf
        Else
            Dim aDate As Date

            If (Date.TryParse(txtDate.Text, aDate) = True) Then

            Else
                result &= "Please re-enter date. Entry has to be a valid date" & vbCrLf
            End If
        End If

        'Validates inventor first name text box
        If (txtFirstName.TextLength = 0) Then
            result &= "Inventor first name not entered, Please enter" & vbCrLf
        End If

        'Validates inventor last name text box
        If (txtLastName.TextLength = 0) Then
            result &= "Inventor last name not entered, Please enter" & vbCrLf
        End If

        'Valides drive ID combo box
        If (cmbID.SelectedIndex = -1) Then
            result &= "Please select an ID" & vbCrLf
        End If

        'Validates the vehicles combo boxes
        If (cmbVehicleOne.SelectedIndex = -1) Then
            result &= "Please select a first vehicle" & vbCrLf
        End If

        If (cmbVehicleTwo.SelectedIndex = -1) Then
            result &= "Please select a second vehicle" & vbCrLf
        End If

        If (cmbVehicleThree.SelectedIndex = -1) Then
            result &= "Please select a third vehicle" & vbCrLf
        End If

        If (cmbVehicleFour.SelectedIndex = -1) Then
            result &= "Please select a fourth vehicle" & vbCrLf
        End If

        'Validates Start Mileage text boxes
        If (txtStartMileageOne.TextLength = 0) Then
            result &= "Start Mileage for vehicle one not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartMileageOne.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Mileage for vehicle one has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle one Start Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If


        If (txtStartMileageTwo.TextLength = 0) Then
            result &= "Start Mileage for vehicle two not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartMileageTwo.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Mileage for vehicle two has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle two Start Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If

        If (txtStartMileageThree.TextLength = 0) Then
            result &= "Start Mileage for vehicle three not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartMileageThree.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Mileage for vehicle three has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle three Start Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If


        If (txtStartMileageFour.TextLength = 0) Then
            result &= "Start Mileage for vehicle four not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartMileageOne.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Mileage for vehicle four has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle four Start Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If


        'Validates End Mileage text boxes
        If (txtEndMileageOne.TextLength = 0) Then
            result &= "End Mileage for vehicle one not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndMileageOne.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Mileage for vehicle one has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle one End Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If


        If (txtEndMileageTwo.TextLength = 0) Then
            result &= "End Mileage for vehicle two not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndMileageTwo.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Mileage for vehicle two has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle two End Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If

        If (txtEndMileageThree.TextLength = 0) Then
            result &= "End Mileage for vehicle three not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndMileageThree.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Mileage for vehicle three has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle three End Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If


        If (txtEndMileageFour.TextLength = 0) Then
            result &= "End Mileage for vehicle four not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndMileageFour.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Mileage for vehicle four has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle four End Mileage, Please re-enter whole number" & vbCrLf
            End If
        End If


        'Validates Start Fuel text boxes
        If (txtStartFuelOne.TextLength = 0) Then
            result &= "Start Fuel for vehicle one not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartFuelOne.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Fuel for vehicle one has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle one Start Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If


        If (txtStartFuelTwo.TextLength = 0) Then
            result &= "Start Fuel for vehicle two not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartFuelTwo.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Fuel for vehicle two has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle two Start Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If

        If (txtStartFuelThree.TextLength = 0) Then
            result &= "Start Fuel for vehicle three not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartFuelThree.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Fuel for vehicle three has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle three Start Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If

        If (txtStartFuelFour.TextLength = 0) Then
            result &= "Start Fuel for vehicle four not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtStartFuelFour.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "Start Fuel for vehicle four has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle four Start Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If

        'Validates End Fuel text boxes
        If (txtEndFuelOne.TextLength = 0) Then
            result &= "End Fuel for vehicle one not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndFuelOne.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Fuel for vehicle one has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle one End Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If

        If (txtEndFuelTwo.TextLength = 0) Then
            result &= "End Fuel for vehicle two not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndFuelTwo.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Fuel for vehicle two has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle two End Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If

        If (txtEndFuelThree.TextLength = 0) Then
            result &= "End Fuel for vehicle three not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndFuelThree.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Fuel for vehicle three has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle three End Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If

        If (txtEndFuelFour.TextLength = 0) Then
            result &= "End Fuel for vehicle four not entered, Please enter" & vbCrLf
        Else
            Dim aNum As Integer = 0

            If (Integer.TryParse(txtEndFuelFour.Text, aNum)) Then
                If (aNum < 0) Then
                    result &= "End Fuel for vehicle four has to be greater than 0, Please re-enter" & vbCrLf
                End If
            Else
                result &= "Invalid entry for vehicle four End Fuel, Please re-enter whole number" & vbCrLf
            End If
        End If


        Return result

    End Function

    'Function that determines the family type of the vehicle and returns it
    Private Function family(vehicle As Integer) As String
        If (vehicle = 0 Or vehicle = 1 Or vehicle = 3) Then
            Return "Covenant"
        Else
            Return "United Nations Space Command"
        End If
    End Function

    'Function that calculates Mileage and returns it
    Private Function calcMileage(startMileage As Integer, endMileage As Integer) As Integer
        Return endMileage - startMileage
    End Function

    'Function that calculates fuel used and returns it
    Private Function calcFuelUsed(startFuel As Integer, endFuel As Integer) As Integer
        Return startFuel - endFuel
    End Function

    'Function that calcuates fuel cost and returns it
    Private Function calcFuelCost(fuelCellsUsed As Integer) As Integer
        Return fuelCellsUsed * PER_FUEL_CELL_COST
    End Function

    'Function that calculates and returns MPFC Rating
    Private Function calcMPFC(mileage As Integer, fuelCellsUsed As Integer) As String
        Dim MPFC As Integer = mileage / fuelCellsUsed

        If (MPFC < 200) Then
            Return "Guzzler"
        ElseIf (MPFC >= 200 And MPFC <= 300) Then
            Return "Economical"
        ElseIf (MPFC > 300) Then
            Return "Amazing!!"
        Else
            Return "Error! No MPFC Rating could be calculated!"
        End If
    End Function

    'Function that displays summary information in a new window
    Private Sub btnAllTests_Click(sender As Object, e As EventArgs) Handles btnAllTests.Click
        Dim summaryFrm As New frmSummary

        summaryFrm.txtSummary.Text = "Number of Tests: " & numOfTests & vbCrLf _
                                        & "Vehicles Tested: " & vehiclesTested & vbCrLf _
                                        & vbCrLf _
                                        & "Total Miles Travelled: " & FormatNumber(totalMiles) & vbCrLf _
                                        & "Total Fuel Cells Used: " & FormatNumber(totalFuelCellsUsed) & vbCrLf _
                                        & "Total Cost: " & FormatCurrency(totalCost)

        summaryFrm.ShowDialog()

    End Sub
End Class