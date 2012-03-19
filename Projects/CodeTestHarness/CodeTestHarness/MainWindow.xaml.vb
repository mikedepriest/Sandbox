Imports System.Text.RegularExpressions
Imports System.Data

Class MainWindow

    Dim orderTable As DataTable = New DataTable


    Private Sub primaryButton_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles primaryButton.Click
        validatePrimary(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text)
    End Sub
    Private Sub primaryButtonA_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles primaryButtonA.Click
        validateMaster(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeMarkingsTextBox.Text, "", New DataTable, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, codePullDateTextBox.Text, currentDowTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputPullDateTextBox.Text, inputDowTextBox.Text, inputPlantCodeTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text)
    End Sub

    Private Sub primary3Button_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles primary3Button.Click
        validateP3(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text, inputPullDateTextBox.Text, inputDowTextBox.Text, inputPlantCodeTextBox.Text, codePullDateTextBox.Text, currentDowTextBox.Text)
    End Sub
    Private Sub primary3ButtonA_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles primary3ButtonA.Click
        validateMaster(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeMarkingsTextBox.Text, "", New DataTable, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, codePullDateTextBox.Text, currentDowTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputPullDateTextBox.Text, inputDowTextBox.Text, inputPlantCodeTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text)
    End Sub

    Private Sub secondaryButton_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles secondaryButton.Click
        validateSecondary(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text, codeMarkingsTextBox.Text)
    End Sub
    Private Sub secondaryButtonA_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles secondaryButtonA.Click
        validateMaster(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeMarkingsTextBox.Text, "", New DataTable, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, codePullDateTextBox.Text, currentDowTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputPullDateTextBox.Text, inputDowTextBox.Text, inputPlantCodeTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text)
    End Sub
    Private Sub kegButton_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles kegButton.Click
        validateKeg(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text, orderTable, currentOrderTextBox.Text, codeMarkingsTextBox.Text)
    End Sub
    Private Sub kegButtonA_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        validateMaster(stringValueInTextBox.Text, codeExpectedTextBox.Text, codeDowTextBox.Text, codePlantCodeTextBox.Text, codeLineTextBox.Text, codeHourTextBox.Text, codeMarkingsTextBox.Text, currentOrderTextBox.Text, orderTable, codeFrontMatterTextBox.Text, codeBackMatterTextBox.Text, codePullDateTextBox.Text, currentDowTextBox.Text, inputFrontMatterTextBox.Text, inputBackMatterTextBox.Text, inputPullDateTextBox.Text, inputDowTextBox.Text, inputPlantCodeTextBox.Text, inputLineCodeTextBox.Text, inputHourTextBox.Text)
    End Sub

    Private Sub setupOrders()
        Dim Column1 As DataColumn = New DataColumn("PP_ID")
        Column1.DataType = System.Type.GetType("System.String")
        orderTable.Columns.Add(Column1)
        Dim Column2 As DataColumn = New DataColumn("Brand")
        Column2.DataType = System.Type.GetType("System.String")
        orderTable.Columns.Add(Column2)

        Dim newRow1 As DataRow = orderTable.NewRow()
        newRow1.Item("PP_ID") = "12345"
        newRow1.Item("BRAND") = "COORS LT"
        orderTable.Rows.Add(newRow1)

    End Sub

    Private Sub validateMaster(stringValue_In As String,
                               CodeExpected As String,
                               CodeDOW As String,
                               CodePlantCode As String,
                               CodeLine As String,
                               CodeHour As String,
                               CodeMarkings As String,
                               CurrentOrder As String,
                               AvailableOrders As DataTable,
                               ByRef CodeFrontMatter As String,
                               ByRef CodeBackMatter As String,
                               ByRef CodePullDate As String,
                               ByRef CurrentDOW As String,
                               ByRef InputFrontMatter As String,
                               ByRef InputBackMatter As String,
                               ByRef InputPullDate As String,
                               ByRef InputDOW As String,
                               ByRef InputPlantCode As String,
                               ByRef InputLineCode As String,
                               ByRef InputHour As String
                               )

        ' Ticket 13 2012-03-07 Add support for Julian and Molson export date formats.
        ' Extensive changes are not individually marked.
        '
        ' ASSUMPTION: Because this is a generic parser, we DEPEND on the "Code" inputs being in the
        ' proper formats as they are currently issued from ePAC (3/2012). Specifically, we expect that
        ' the CodeExpected does not have the first line (brand code) of the keg marking, just the pull
        ' date and subsequent material.
        '
        Dim mask As String

        Dim codeFormat As Integer
        Const codeFormatCount As Integer = 4
        Const codeFormatsDefault As Integer = 0
        Const codeFormatsNormal As Integer = 1
        Const codeFormatsJulian As Integer = 2
        Const codeFormatsMolson As Integer = 3

        Dim masklist(codeFormatCount) As String
        masklist(codeFormatsDefault) = "..........................." ' Longer than any valid format
        masklist(codeFormatsNormal) = "........................."    ' "APR*01113.2%BEERA10610123"
        masklist(codeFormatsJulian) = "......................."      ' "066*123.2%BEERA10610123"
        masklist(codeFormatsMolson) = "........................."    ' "03-01-123.2%BEERA10610123"

        Dim maskSecondarySuffix As String = "..........."         ' "*TTTT*32899" per PK-ST-01
        Dim maskKegBrandCode As String = "................"       ' Longer than any valid brand code

        Dim depositMarkLocator(codeFormatCount) As Integer
        depositMarkLocator(codeFormatsNormal) = 4 ' Won't be there anyway
        depositMarkLocator(codeFormatsNormal) = 4 ' Just after the month
        depositMarkLocator(codeFormatsJulian) = 3 ' Just after the day number
        depositMarkLocator(codeFormatsMolson) = 1 ' No deposit mark, this will never find anything

        Dim depositMarkCharacter As String = "*"   ' This is what's expected from the MES or PLC

        Dim pullDateLength(codeFormatCount) As Integer
        pullDateLength(codeFormatsDefault) = 7 ' DDMONYYYY (a nonsense format)
        pullDateLength(codeFormatsNormal) = 7  ' 01APR12
        pullDateLength(codeFormatsJulian) = 5  ' 06612
        pullDateLength(codeFormatsMolson) = 8  ' 04-01-12

        Dim svi As String 'StringValue_In
        Dim ifm As String 'InputFrontMatter
        Dim ibm As String 'InputBackMatter
        Dim ipd As String 'InputPullDate
        Dim imk As String 'InputMarking (not used)
        Dim idw As String 'InputDOW
        Dim ipc As String 'InputPlantCode
        Dim ilc As String 'InputLineCode
        Dim ihr As String 'InputHour
        Dim ibc As String 'InputBrandCode

        Dim cex As String 'CodeExpected
        Dim cpd As String 'CodePullDate
        Dim cfm As String 'CodeFrontMatter
        Dim cbm As String 'CodeBackMatter
        Dim cbc As String 'CodeBrandCode

        Dim oti As DataTable 'OrderTableIn

        Dim rgx As New Regex("\s+") ' match one or more whitespace characters
        Dim rgx2 As New Regex("[A-Z]") ' match upper case letters
        Dim rgx3 As New Regex("[0-9]") ' match numbers
        Dim rgx4 As New Regex("[^0-9]") ' match non-numbers
        Dim rgx5 As New Regex("\*+") ' match one or more asterisks

        '
        ' Condition the inputs
        '
        ' Replace whitepace with nothing, normalize case
        svi = rgx.Replace(stringValue_In.ToUpper, "")
        cex = rgx.Replace(CodeExpected.ToUpper, "")

        '
        ' If an order table is passed in, this is "probably" a keg code. Let's see if we
        ' can figure out what the brand code is. Strip whitespace and normalize while
        ' we are at it.
        '
        cbc = ""
        ibc = ""
        Try
            If (orderTable.IsInitialized) Then
                oti = AvailableOrders
                For Each orc As DataRow In oti.Rows
                    If CurrentOrder.Equals(orc.Item("PP_ID")) Then
                        cbc = rgx.Replace(orc.Item("Brand").ToString.ToUpper, "")
                    End If
                Next
            End If
        Catch ex As Exception
            cbc = ""
            ibc = ""
        End Try

        ' Evaluate the Code Expected to see which code format it matches
        '
        ' Simple test:
        '  * if the first character is a letter, it's normal
        '  * if the third character is a hyphen, it's Molson
        '  * if the fourth character is a number, it's Julian
        '  * otherwise, assume it's screwed up and take the default
        '
        If (rgx2.IsMatch(cex.Substring(0, 1))) Then
            codeFormat = codeFormatsNormal
        ElseIf (cex.Substring(2, 1).Contains("-")) Then
            codeFormat = codeFormatsMolson
        ElseIf (rgx3.IsMatch(cex.Substring(0, 3))) Then
            codeFormat = codeFormatsJulian
        Else
            codeFormat = codeFormatsDefault
        End If
        mask = masklist(codeFormat) + maskSecondarySuffix ' provides enough input for selected format

        '  Ensure there is enough input by appending the mask and then
        '  truncating to the mask length (this will right fill with 
        '  mask characters)
        '
        cex = cex + mask + maskKegBrandCode
        cex = cex.Substring(0, cbc.Length + mask.Length)

        svi = svi + mask + maskKegBrandCode
        svi = svi.Substring(0, cbc.Length + mask.Length)

        ' Extract conditioned code parts
        ' Use the code format and the deposit mark to get the pull date

        If (cex.Substring(0, depositMarkLocator(codeFormat)).Contains(depositMarkCharacter)) Then
            cpd = cex.Substring(0, pullDateLength(codeFormat) + 1) ' marker found
        Else
            cpd = cex.Substring(0, pullDateLength(codeFormat))   ' no marker found
        End If

        ' Code Front Matter is the portion of the code that extends to the plant code
        ' For secondary containers and kegs it may include a %BEER marking.
        ' For kegs it includes the brand code. For other containers there's not one.
        cfm = cbc + cpd.ToUpper + CodeMarkings.ToUpper + CodeDOW.ToUpper + CodePlantCode.ToUpper

        ' Code Back Matter is the last part of the conditioned code,
        ' starting at position 2+(cfm.Length+CodeHour.length+1) (skip the minutes)
        cbm = cex.Substring((cfm.Length + CodeLine.Length + CodeHour.Length + 2), (cex.Length - cfm.Length - CodeLine.Length - CodeHour.Length - 2))
        cbm = cbm.Trim(mask.ToCharArray)

        ' Get the component parts of the input
        '  Pull date follows the brand code on kegs.
        If cbc.Length > 0 Then
            ibc = svi.Substring(0, cbc.Length)
        Else
            ibc = ""
        End If

        If (svi.Substring(0, cbc.Length + depositMarkLocator(codeFormat)).Contains(depositMarkCharacter)) Then
            ipd = svi.Substring(cbc.Length, pullDateLength(codeFormat) + 1)
            svi = svi.Substring(cbc.Length + pullDateLength(codeFormat) + 1, (svi.Length - cbc.Length - (pullDateLength(codeFormat) + 1)))
        Else
            ipd = svi.Substring(cbc.Length, pullDateLength(codeFormat))
            svi = svi.Substring(cbc.Length + pullDateLength(codeFormat), (svi.Length - cbc.Length - pullDateLength(codeFormat)))
        End If

        ' Primary code doesn't have the ABV marking
        ' Secondary and keg codes may have it.
        '  ASSUMPTION: "Marking" always has a "%" character if it is present
        '  and a length of 8 (3.2%BEER)
        If (svi.Substring(0, 4).Contains("%")) Then
            imk = svi.Substring(0, 8)
            svi = svi.Substring(8, (svi.Length - 8))
        Else
            imk = ""
        End If

        '
        ' Ticket 13: from this point forward, there's no difference between the various code formats
        '
        '  The next 3 characters are the DOW and plant code,
        '  line code is the next two, with the 
        '  following 2 as the hour, for a total of 7
        idw = svi.Substring(0, 1)
        ipc = svi.Substring(1, 2)

        ilc = svi.Substring(3, 2)
        ihr = svi.Substring(5, 2)
        svi = svi.Substring(7, (svi.Length - 7))

        '  The next two characters are the minute, so just skip them
        '  and make the rest of the string the input back matter
        '  (note the different length calc than we have been using)
        '  then remove any remaining mask characters
        If (svi.Length > 2) Then
            ibm = svi.Substring(2, (svi.Length - 2))
        Else
            ibm = ""
        End If
        ibm = ibm.Trim(mask.ToCharArray)

        '  GLD: On lines with multiple packers an asterisk can be
        '  added to the back matter on one of the packers. Remove this
        '  from the back matter for consideration. This won't affect codes
        '  from other breweries.
        cbm = rgx5.Replace(cbm, "")
        ibm = rgx5.Replace(ibm, "")

        ' Build input front matter
        ifm = ibc + ipd + imk + idw + ipc

        InputFrontMatter = ifm
        InputBackMatter = ibm
        ' Make sure these are forced to valid numerics
        InputLineCode = rgx4.Replace(ilc.Replace(".", "0"), "0")
        InputHour = rgx4.Replace(ihr.Replace(".", "0"), "0")
        InputDOW = idw
        InputPullDate = ipd
        InputPlantCode = ipc


        CodeFrontMatter = cfm
        CodeBackMatter = cbm
        CodeLine = CodeLine.ToUpper
        CodeHour = CodeHour.ToUpper
        CodePullDate = cpd

        CurrentDOW = WeekdayName(Weekday(Now))

    End Sub


    Private Sub validatePrimary(stringValue_In As String, CodeExpected As String, CodeDOW As String, CodePlantCode As String, CodeLine As String, CodeHour As String, ByRef CodeFrontMatter As String, ByRef CodeBackMatter As String, ByRef InputFrontMatter As String, ByRef InputBackMatter As String, ByRef InputLineCode As String, ByRef InputHour As String)
        ' Ticket 13 2012-03-07 Add support for Julian and Molson export date formats.
        ' Extensive changes are not individually marked.
        '
        Dim mask As String

        Dim codeFormat As Integer
        Const codeFormatCount As Integer = 3
        Const codeFormatsNormal As Integer = 1
        Const codeFormatsJulian As Integer = 2
        Const codeFormatsMolson As Integer = 3

        Dim masklist(codeFormatCount) As String
        masklist(codeFormatsNormal) = "................." '"APR*0111A10610123"
        masklist(codeFormatsJulian) = "................." '"066*12A10610123"
        masklist(codeFormatsMolson) = "................." '"03-01-12A10610123"

        Dim depositMarkLocator(codeFormatCount) As Integer
        depositMarkLocator(codeFormatsNormal) = 4 ' Just after the month
        depositMarkLocator(codeFormatsJulian) = 3 ' Just after the day number
        depositMarkLocator(codeFormatsMolson) = 1 ' No deposit mark, this will never find anything

        Dim depositMarkCharacter As String = "*"   ' This is what's expected from the MES or PLC

        Dim pullDateLength(codeFormatCount) As Integer
        pullDateLength(codeFormatsNormal) = 7 ' 01APR12
        pullDateLength(codeFormatsJulian) = 5 ' 06612
        pullDateLength(codeFormatsMolson) = 8 ' 04-01-12

        Dim svi As String 'StringValue_In
        Dim ifm As String 'InputFrontMatter
        Dim ibm As String 'InputBackMatter
        Dim ipd As String
        Dim imk As String
        Dim idwpc As String
        Dim ilc As String 'InputLineCode
        Dim ihr As String 'InputHour

        Dim cex As String 'CodeExpected
        Dim cpd As String
        Dim cfm As String 'CodeFrontMatter
        Dim cbm As String 'CodeBackMatter

        Dim rgx As New Regex("\s+") ' match one or more whitespace characters
        Dim rgx2 As New Regex("[A-Z]") ' match upper case letters
        Dim rgx3 As New Regex("[^0-9]") ' match non-numbers

        ' Build conditioned code front matter
        cex = rgx.Replace(CodeExpected.ToUpper, "") ' replace whitespace with nothing and normalize case

        ' Evaluate the Code Expected to see which code format it matches 
        '
        ' Simple test:
        '  * if the third character is a hyphen, it's Molson
        '  * if the first character is a letter, it's normal
        '  * otherwise, assume it's Julian
        '
        If (cex.Substring(2, 1).Contains("-")) Then
            codeFormat = codeFormatsMolson
        ElseIf (rgx2.IsMatch(cex.Substring(0, 1))) Then
            codeFormat = codeFormatsNormal
        Else
            codeFormat = codeFormatsJulian
        End If

        ' Use the code format and the deposit mark to get the pull date

        If (cex.Substring(0, depositMarkLocator(codeFormat)).Contains(depositMarkCharacter)) Then
            cpd = cex.Substring(0, pullDateLength(codeFormat) + 1) ' marker found
        Else
            cpd = cex.Substring(0, pullDateLength(codeFormat))   ' no marker found
        End If

        ' Leave the line code off for Primary Codes, it is evaluated later
        cfm = cpd.ToUpper + CodeDOW.ToUpper + CodePlantCode.ToUpper

        ' Code Back Matter is the last part of the conditioned code,
        ' starting at position 2+(cfm.Length+CodeHour.length+1) (skip the minutes)
        cbm = cex.Substring((cfm.Length + CodeLine.Length + CodeHour.Length + 2), (cex.Length - cfm.Length - CodeLine.Length - CodeHour.Length - 2))

        '
        ' Condition the input
        '

        '  Eliminate excess whitespace
        svi = rgx.Replace(stringValue_In.ToUpper, "") ' replace whitespace with nothing

        '  Ensure there is enough input by appending the mask and then
        '  truncating to the mask length (this will right fill with 
        '  mask characters)

        mask = masklist(codeFormat)

        Dim i As Integer = mask.Length
        svi = svi + mask
        svi = svi.Substring(0, i)

        ' Get the component parts of the input
        '  Pull date (7 or 8 characters depending on presence of "*"
        If (svi.Substring(0, depositMarkLocator(codeFormat)).Contains(depositMarkCharacter)) Then
            ipd = svi.Substring(0, pullDateLength(codeFormat) + 1)
            svi = svi.Substring(pullDateLength(codeFormat) + 1, (svi.Length - (pullDateLength(codeFormat) + 1)))
        Else
            ipd = svi.Substring(0, pullDateLength(codeFormat))
            svi = svi.Substring(pullDateLength(codeFormat), (svi.Length - pullDateLength(codeFormat)))
        End If

        '
        ' Ticket 13: from this point forward, there's no difference between the various code formats
        '

        ' Primary code doesn't have the ABV marking
        imk = ""

        '  The next 3 characters are the DOW and plant code,
        '  line code is the next two, with the 
        '  following 2 as the hour, for a total of 7
        idwpc = svi.Substring(0, 3)
        ilc = svi.Substring(3, 2)
        ihr = svi.Substring(5, 2)
        svi = svi.Substring(7, (svi.Length - 7))

        '  The next two characters are the minute, so just skip them
        '  and make the rest of the string the input back matter
        '  (note the different length calc than we have been using)
        '  then remove any remaining mask characters
        ' Begin Ticket 213 2011-02-21 add conditional based on length
        If (svi.Length > 2) Then
            ibm = svi.Substring(2, (svi.Length - 2))
        Else
            ibm = ""
        End If
        ' End Ticket 213 2011-02-21
        ibm = ibm.Trim(mask.ToCharArray)

        ' Build input front matter
        ifm = ipd + imk + idwpc

        InputFrontMatter = ifm
        InputBackMatter = ibm
        ' Make sure these are forced to valid numerics
        InputLineCode = rgx3.Replace(ilc.Replace(".", "0"), "0")
        InputHour = rgx3.Replace(ihr.Replace(".", "0"), "0")

        CodeFrontMatter = cfm
        CodeBackMatter = cbm
        CodeLine = CodeLine.ToUpper
        CodeHour = CodeHour.ToUpper

    End Sub

    Private Sub validateSecondary(stringValue_In As String, CodeExpected As String, CodeDOW As String, CodePlantCode As String, CodeLine As String, CodeHour As String, ByRef CodeFrontMatter As String, ByRef CodeBackMatter As String, ByRef InputFrontMatter As String, ByRef InputBackMatter As String, ByRef InputLineCode As String, ByRef InputHour As String, CodeMarkings As String)
        Dim mask As String = "....................................." '"APR*01113.2%BEERA10610123xxxxxxx32899"
        Dim svi As String 'StringValue_In
        Dim ifm As String 'InputFrontMatter
        Dim ibm As String 'InputBackMatter
        Dim ipd As String
        Dim imk As String
        Dim idwpclc As String
        Dim ilc As String 'InputLineCode
        Dim ihr As String 'InputHour

        Dim cex As String 'CodeExpected
        Dim cpd As String
        Dim cfm As String 'CodeFrontMatter
        Dim cbm As String 'CodeBackMatter

        Dim ivf As Boolean 'IsValid_Flag
        ivf = False

        ' Condition the input:
        '  Eliminate excess whitespace
        Dim rgx As New Regex("\s+") ' match one or more whitespace characters
        ' Begin Ticket 213 2011-02-21 and others
        Dim rgx2 As New Regex("[^0-9]") ' match one or more non-numbers
        ' End Ticket 213 2011-02-21
        svi = rgx.Replace(stringValue_In.ToUpper, "") ' replace whitespace with nothing
        '  Ensure there is enough input by appending the mask and then
        '  truncating to the mask length (this will right fill with 
        '  mask characters)
        Dim i As Integer = mask.Length
        svi = svi + mask
        svi = svi.Substring(0, i)

        ' Get the component parts of the input
        '  Pull date (7 or 8 characters depending on presence of "*"
        If (svi.Substring(0, 4).Contains("*")) Then
            ipd = svi.Substring(0, 8)
            svi = svi.Substring(8, (svi.Length - 8))
        Else
            ipd = svi.Substring(0, 7)
            svi = svi.Substring(7, (svi.Length - 7))
        End If

        '  ASSUMPTION: "Marking" always has a "%" character if it is present
        '  and a length of 8 (3.2%BEER)
        If (svi.Substring(0, 4).Contains("%")) Then
            imk = svi.Substring(0, 8)
            svi = svi.Substring(8, (svi.Length - 8))
        Else
            imk = ""
        End If

        '  The next 3 characters are the DOW and plant code,
        '  line code is the next two, with the 
        '  following 2 as the hour, for a total of 7
        idwpclc = svi.Substring(0, 5)
        ilc = svi.Substring(3, 2)
        ihr = svi.Substring(5, 2)
        svi = svi.Substring(7, (svi.Length - 7))

        '  The next two characters are the minute, so just skip them
        '  and make the rest of the string the input back matter
        '  (note the different length calc than we have been using)
        '  then remove any remaining mask characters
        ibm = svi.Substring(2, (svi.Length - 3))
        ibm = ibm.Trim(mask.ToCharArray)

        '  GLD: On lines with multiple packers an asterisk can be
        '  added to the back matter on one of the packers. Remove this
        '  from the back matter for consideration
        Dim rgxm As New Regex("\*+") ' match one or more asterisks
        ibm = rgxm.Replace(ibm, "")

        ' Build input front matter
        ifm = ipd + imk + idwpclc


        ' Build conditioned code front matter to match against ePAC
        Dim rgx3 As New Regex("\s+") ' match one or more whitespace characters
        cex = rgx3.Replace(CodeExpected.ToUpper, "") ' replace whitespace with nothing
        If (cex.Substring(0, 4).Contains("*")) Then
            cpd = cex.Substring(0, 8)
            'cex = cex.Substring(8, (cex.Length - 8))
        Else
            cpd = cex.Substring(0, 7)
            'cex = cex.Substring(7, (cex.Length - 7))
        End If
        cfm = cpd.ToUpper + CodeMarkings.ToUpper + CodeDOW.ToUpper + CodePlantCode.ToUpper + CodeLine.ToUpper

        ' Code Back Matter is the last part of the conditioned code,
        ' starting at position 2+(cfm.Length+CodeHour.length+1) (skip the minutes)
        cbm = cex.Substring((cfm.Length + CodeHour.Length + 2), (cex.Length - cfm.Length - CodeHour.Length - 2))

        InputFrontMatter = ifm
        InputBackMatter = ibm
        ' Begin Ticket 213 2011-02-21 - make sure these are forced to valid numerics
        InputLineCode = rgx2.Replace(ilc.Replace(".", "0"), "0")
        InputHour = rgx2.Replace(ihr.Replace(".", "0"), "0")
        ' End Ticket 213 2011-02-21


        CodeFrontMatter = cfm
        CodeBackMatter = cbm
        CodeLine = CodeLine.ToUpper
        CodeHour = CodeHour.ToUpper

    End Sub

    Private Sub validateP3(stringValue_In As String, CodeExpected As String, CodeDOW As String, CodePlantCode As String, CodeLine As String, CodeHour As String, ByRef CodeFrontMatter As String, ByRef CodeBackMatter As String, ByRef InputFrontMatter As String, ByRef InputBackMatter As String, ByRef InputLineCode As String, ByRef InputHour As String, ByRef InputPullDate As String, ByRef InputDOW As String, ByRef InputPlantCode As String, ByRef CodePullDate As String, ByRef CurrentDOW As String)
        Dim mask As String = "................." '"APR*0111A10610123"

        Dim svi As String 'StringValue_In
        Dim ifm As String 'InputFrontMatter
        Dim ibm As String 'InputBackMatter
        Dim ipd As String 'InputPullDate
        Dim imk As String 'InputMarking (not used)
        Dim idw As String 'InputDOW
        Dim ipc As String 'InputPlantCode
        Dim ilc As String 'InputLineCode
        Dim ihr As String 'InputHour

        Dim cex As String 'CodeExpected
        Dim cpd As String 'CodePullDate
        Dim cfm As String 'CodeFrontMatter
        Dim cbm As String 'CodeBackMatter

        ' Condition the input:
        '  Eliminate excess whitespace
        Dim rgx As New Regex("\s+") ' match one or more whitespace characters
        Dim rgx2 As New Regex("[^0-9]") ' match one or more non-numbers
        svi = rgx.Replace(stringValue_In.ToUpper, "") ' replace whitespace with nothing
        '  Ensure there is enough input by appending the mask and then
        '  truncating to the mask length (this will right fill with 
        '  mask characters)
        Dim i As Integer = mask.Length
        svi = svi + mask
        svi = svi.Substring(0, i)

        ' Get the component parts of the input
        '  Pull date (7 or 8 characters depending on presence of "*"
        If (svi.Substring(0, 4).Contains("*")) Then
            ipd = svi.Substring(0, 8)
            svi = svi.Substring(8, (svi.Length - 8))
        Else
            ipd = svi.Substring(0, 7)
            svi = svi.Substring(7, (svi.Length - 7))
        End If

        ' Primary code doesn't have marking, comment out this section (Ticket 184 refers)
        ' but leave it as a placeholder in case this changes.
        ' '  ASSUMPTION: "Marking" always has a "%" character if it is present
        ' '  and a length of 8 (3.2%BEER)
        ' If (svi.Substring(0, 4).Contains("%")) Then
        '    imk = svi.Substring(0, 8)
        '    svi = svi.Substring(8, (svi.Length - 8))
        ' Else
        imk = ""
        ' End If

        '  The next character is the DOW, then 2 for plant code.
        idw = svi.Substring(0, 1)
        ipc = svi.Substring(1, 2)
        '  Line code is the next two, with the 
        '  following 2 as the hour, for a total of 7
        ilc = svi.Substring(3, 2)
        ihr = svi.Substring(5, 2)
        svi = svi.Substring(7, (svi.Length - 7))

        '  The next two characters are the minute, so just skip them
        '  and make the rest of the string the input back matter
        '  (note the different length calc than we have been using)
        '  then remove any remaining mask characters
        If (svi.Length > 2) Then
            ibm = svi.Substring(2, (svi.Length - 2))
        Else
            ibm = ""
        End If
        ibm = ibm.Trim(mask.ToCharArray)

        ' Build input front matter
        ifm = ipd + imk + idw + ipc

        ' Build conditioned code front matter to match against ePAC
        Dim rgx3 As New Regex("\s+") ' match one or more whitespace characters
        cex = rgx3.Replace(CodeExpected.ToUpper, "") ' replace whitespace with nothing
        If (cex.Substring(0, 4).Contains("*")) Then
            cpd = cex.Substring(0, 8)
        Else
            cpd = cex.Substring(0, 7)
        End If
        ' Leave the line code off for Primary Codes, it is evaluated later
        ' Markings don't appear on primary (Ticket 184 refers) but keep this in case that changes
        ' Was:
        ' cfm = cpd.ToUpper + CodeMarkings.ToUpper + CodeDOW.ToUpper + CodePlantCode.ToUpper
        cfm = cpd.ToUpper + CodeDOW.ToUpper + CodePlantCode.ToUpper

        ' Code Back Matter is the last part of the conditioned code,
        ' starting at position 2+(cfm.Length+CodeHour.length+1) (skip the minutes)
        cbm = cex.Substring((cfm.Length + CodeLine.Length + CodeHour.Length + 2), (cex.Length - cfm.Length - CodeLine.Length - CodeHour.Length - 2))

        InputFrontMatter = ifm
        InputPullDate = ipd
        InputDOW = idw
        ' The regex ensures that these are valid numbers
        InputPlantCode = rgx2.Replace(ipc.Replace(".", "0"), "0")
        InputLineCode = rgx2.Replace(ilc.Replace(".", "0"), "0")
        InputHour = rgx2.Replace(ihr.Replace(".", "0"), "0")
        InputBackMatter = ibm

        CodeFrontMatter = cfm
        CodePullDate = cpd
        CodeDOW = CodeDOW.ToUpper
        CodePlantCode = CodePlantCode.ToUpper
        CodeLine = CodeLine.ToUpper
        CodeHour = CodeHour.ToUpper
        CodeBackMatter = cbm

        CurrentDOW = WeekdayName(Weekday(Now))

    End Sub

    Private Sub validateKeg(stringValue_In As String, CodeExpected As String, CodeDOW As String, CodePlantCode As String, CodeLine As String, CodeHour As String, ByRef CodeFrontMatter As String, ByRef CodeBackMatter As String, ByRef InputFrontMatter As String, ByRef InputBackMatter As String, ByRef InputLineCode As String, ByRef InputHour As String, ByRef AvailableOrders As DataTable, ByRef CurrentOrder As String, CodeMarkings As String)
        Dim mask As String = "..................................."
        '"COORSLTAPR0111A10610123Z*TTTT*SSSSS"
        ' Note the brand code is of variable length

        Dim oti As DataTable 'OrderTableIn

        Dim svi As String = "" 'StringValue_In
        Dim ibl As Int32 = 0 'InputBrandLength
        Dim ibc As String = "" 'InputBrandCode
        Dim ifm As String = "" 'InputFrontMatter
        Dim ibm As String = "" 'InputBackMatter
        Dim ipd As String = "" 'InputPullDate
        Dim imk As String = "" 'InputMarkings (not used)
        Dim idwpclc As String = "" 'InputPlantCodeLineCode
        Dim ilc As String = "" 'InputLineCode
        Dim ihr As String = "" 'InputHour

        Dim cbc As String = "" 'CodeBrandCode
        Dim cex As String = "" 'CodeExpected
        Dim cpd As String = "" 'CodePullDate
        Dim cfm As String = "" 'CodeFrontMatter
        Dim cbm As String = "" 'CodeBackMatter

        Dim ivf As Boolean 'IsValid_Flag
        ivf = False

        '        cbc = "COORS LT"
        ibc = ""
        ' Go over the list of available orders, find the one that matches, 
        ' then get the brand info.
        oti = AvailableOrders
        For Each orc As DataRow In oti.Rows
            If CurrentOrder.Equals(orc.Item("PP_ID")) Then
                cbc = orc.Item("Brand").ToString.ToUpper
            End If
        Next

        ' Condition the input:
        '  Eliminate excess whitespace
        Dim rgx As New Regex("\s+") ' match one or more whitespace characters
        ' Begin Ticket 213 2011-02-21 and others
        Dim rgx2 As New Regex("[^0-9]") ' match one or more non-numbers
        ' End Ticket 213 2011-02-21
        svi = rgx.Replace(stringValue_In.ToUpper, "") ' replace whitespace with nothing
        cbc = rgx.Replace(cbc.ToUpper, "")
        ibl = cbc.Length

        '  Ensure there is enough input by appending the mask and then
        '  truncating to the mask length (this will right fill with 
        '  mask characters)
        Dim i As Integer = mask.Length
        svi = svi + mask
        svi = svi.Substring(0, i)

        ' Get the component parts of the input
        ' Brand (length of expected brand)
        If ibl > 0 Then
            ibc = svi.Substring(0, ibl)
            svi = svi.Substring(ibl, (svi.Length - ibl))
        End If
        '  Pull date (7 characters)
        ipd = svi.Substring(0, 7)
        svi = svi.Substring(7, (svi.Length - 7))

        imk = ""

        '  The next 3 characters are the DOW and plant code,
        '  line code is the next two, with the 
        '  following 2 as the hour, for a total of 7
        idwpclc = svi.Substring(0, 5)
        ilc = svi.Substring(3, 2)
        ihr = svi.Substring(5, 2)
        svi = svi.Substring(7, (svi.Length - 7))

        '  The next two characters are the minute, so just skip them
        '  and make the rest of the string the input back matter
        '  (note the different length calc than we have been using)
        '  then remove any remaining mask characters
        ibm = svi.Substring(2, (svi.Length - 3))
        ibm = ibm.Trim(mask.ToCharArray)

        '  GLD: On lines with multiple packers an asterisk can be
        '  added to the back matter on one of the packers. Remove this
        '  from the back matter for consideration
        Dim rgxm As New Regex("\*+") ' match one or more asterisks
        ibm = rgxm.Replace(ibm, "")

        ' Build input front matter
        ifm = ibc + ipd + imk + idwpclc


        ' Build conditioned code front matter to match against ePAC
        Dim rgx3 As New Regex("\s+") ' match one or more whitespace characters
        cex = rgx3.Replace(CodeExpected.ToUpper, "") ' replace whitespace with nothing
        cpd = cex.Substring(ibl, 7)
        cfm = cbc + cpd.ToUpper + CodeMarkings.ToUpper + CodeDOW.ToUpper + CodePlantCode.ToUpper + CodeLine.ToUpper

        ' Code Back Matter is the last part of the conditioned code,
        ' starting at position 2+(cfm.Length+CodeHour.length+1) (skip the minutes)
        cbm = cex.Substring((cfm.Length + CodeHour.Length + 2), (cex.Length - cfm.Length - CodeHour.Length - 2))

        InputFrontMatter = ifm
        InputBackMatter = ibm
        ' Begin Ticket 213 2011-02-21 - make sure these are forced to valid numerics
        InputLineCode = rgx2.Replace(ilc.Replace(".", "0"), "0")
        InputHour = rgx2.Replace(ihr.Replace(".", "0"), "0")
        ' End Ticket 213 2011-02-21


        CodeFrontMatter = cfm
        CodeBackMatter = cbm
        CodeLine = CodeLine.ToUpper
        CodeHour = CodeHour.ToUpper


    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        setupOrders()

    End Sub

End Class
