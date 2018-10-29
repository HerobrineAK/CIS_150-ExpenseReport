Module Module1

    '   Employee Name & Number
    Private EmpName As String
    Private EmpNumber As String

    '   Employee Beginning and Ending Balance
    Private BegBalance As Decimal
    Private EndBalance As Decimal

    '   Employee Beginning and Ending Odometer
    Private BegOdometer As Decimal
    Private EndOdometer As Decimal

    '   Employee's Lunch, Dinner, Mileage, and the Total Expense
    Private Lunch As Decimal
    Private Dinner As Decimal
    Private MileageCost As Decimal
    Private TotalExpense As Decimal

    '   Expense File Details
    Private CurrentRecord() As String
    Private ExpenseFile As New Microsoft.VisualBasic.FileIO.TextFieldParser("EXPENF18.TXT")

    '   Other Variables
    Private MilesDriven As Integer
    Private Const COST_PER_MILE As Decimal = 0.545

    Sub Main()
        Call HouseKeeping()
        Do While Not (ExpenseFile).EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Sub HouseKeeping()
        Call SetFileDelimiters()
        Call WriteHeadings()
    End Sub

    Sub SetFileDelimiters()
        ExpenseFile.TextFieldType = FileIO.FieldType.Delimited
        ExpenseFile.SetDelimiters(",")
    End Sub

    Sub WriteHeadings()
        Console.WriteLine()
        Console.WriteLine(Space(29) & "Learn More Corporation")
        Console.WriteLine(Space(33) & "Expense Report")
        Console.WriteLine()
        Console.WriteLine(" Employee   Beginning   -----  Expenses  -----    Number      Total      Ending")
        Console.WriteLine(" Name         Balance   Lunch  Dinner  Mileage  of Miles   Expenses     Balance")
    End Sub

    Sub ProcessRecords()
        Call ReadFile()
        Call DetailCalculation()
        Call WriteDetailLine()
    End Sub

    Sub ReadFile()
        CurrentRecord = ExpenseFile.ReadFields()

        EmpName = CurrentRecord(1)
        EmpNumber = CurrentRecord(0)

        BegBalance = CurrentRecord(2)

        BegOdometer = CurrentRecord(3)
        EndOdometer = CurrentRecord(4)

        Lunch = CurrentRecord(5)
        Dinner = CurrentRecord(6)
    End Sub

    Sub DetailCalculation()
        MilesDriven = EndOdometer - BegOdometer
        MileageCost = MilesDriven * COST_PER_MILE
        TotalExpense = Lunch + Dinner + MileageCost
        EndBalance = BegBalance - TotalExpense
    End Sub

    Sub WriteDetailLine()
        Console.WriteLine()
        Console.WriteLine(Space(1) &
                          EmpName.PadRight(10) &
                          Space(2) &
                          BegBalance.ToString("n").PadLeft(8) &
                          Space(3) &
                          Lunch.ToString("n").PadLeft(5) &
                          Space(3) &
                          Dinner.ToString("n").PadLeft(5) &
                          Space(3) &
                          MileageCost.ToString("n").PadLeft(6) &
                          Space(5) &
                          MilesDriven.ToString().PadLeft(5) &
                          Space(2) &
                          TotalExpense.ToString("c").PadLeft(9) &
                          Space(2) &
                          EndBalance.ToString("c").PadLeft(10))

    End Sub
    Sub EndOfJob()
        Call SummaryOutput()
        Call CloseFile()
    End Sub

    Sub SummaryOutput()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(23) & "End Of EXPENSE REPORT by David Rees")
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(30) & "Press -Enter- To Exit")
    End Sub

    Sub CloseFile()
        Console.ReadLine()
        ExpenseFile.Close()
    End Sub

End Module
