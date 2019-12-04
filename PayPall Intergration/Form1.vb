Imports System
Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports System.Configuration

Public Class Form1
    Public conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim Month As String = ComboBox1.Text
        Dim Year As String = ComboBox2.Text
        Dim JDate As Date = DateTimePicker1.Value

        Dim MonthVal As String
        MonthVal = 0

        'set month value

        If Month = "January" Then
            MonthVal = 1
        ElseIf Month = "February" Then
            MonthVal = 2
        ElseIf Month = "March" Then
            MonthVal = 3
        ElseIf Month = "April" Then
            MonthVal = 4
        ElseIf Month = "May" Then
            MonthVal = 5
        ElseIf Month = "June" Then
            MonthVal = 6
        ElseIf Month = "July" Then
            MonthVal = 7
        ElseIf Month = "August" Then
            MonthVal = 8
        ElseIf Month = "September" Then
            MonthVal = 9
        ElseIf Month = "October" Then
            MonthVal = 10
        ElseIf Month = "November" Then
            MonthVal = 11
        ElseIf Month = "December" Then
            MonthVal = 12
        End If

        'set progress bar value
        ProgressBar1.Value = 10

        'set output file location

        My.Computer.FileSystem.CreateDirectory("D:\Payroll " + Month + " " + Year)

        'GOU research
        Dim Query As String = "SELECT ISNULL(SUM(dbo.Salary_Generation.Net_Amount),0), " & _
        "ISNULL(SUM(0.05 * dbo.Salary_Generation.Total_Earn_Amount),0), " & _
        "ISNULL(SUM(0.1 * dbo.Salary_Generation.Total_Earn_Amount),0), " & _
        "ISNULL(SUM(dbo.Salary_Generation.Goods_Amount),0), " & _
        "ISNULL(SUM(dbo.Salary_Generation.LST_Amount),0) " & _
"FROM dbo.Branch_Master INNER JOIN " & _
                      "dbo.Increment ON dbo.Branch_Master.Branch_ID = dbo.Increment.Branch_ID INNER JOIN " & _
                      "dbo.Salary_Generation ON dbo.Increment.Increment_Id = dbo.Salary_Generation.Increment_ID INNER JOIN " & _
                      "dbo.PRODUCT_MASTER ON dbo.Increment.Product_ID = dbo.PRODUCT_MASTER.Product_ID " & _
"WHERE (dbo.Salary_Generation.Month = " + MonthVal + ") AND (dbo.Salary_Generation.Year = " + Year + ") AND (dbo.PRODUCT_MASTER.Product_Name = 'Research') AND " & _
                      "(dbo.Branch_Master.Branch_Name = 'GOU')"

        Dim QueryCmd As New SqlCommand(Query, conn)
        Dim dataadapter As New SqlDataAdapter(QueryCmd)
        Dim ds As New DataSet()
        conn.Open()
        dataadapter.Fill(ds, "tblSalary")
        conn.Close()

        ProgressBar1.Value = 20

        Dim GOUR As String = "SELECT SUM(Net_Amount), SUM([5%]), SUM([10%]), SUM(PAYE), SUM(LST_Amount) " & _
"FROM dbo.ppiGOU_Research WHERE (Month = " + MonthVal + ") AND (Year = " + Year + ")"

        Dim GOURCmd As New SqlCommand(GOUR, conn)
        Dim GOURda As New SqlDataAdapter(GOURCmd)
        Dim GOURds As New DataSet()
        conn.Open()
        GOURda.Fill(ds, "tblSalaryGOUR")
        conn.Close()

        ProgressBar1.Value = 30


        Dim GOUS As String = "SELECT SUM(dbo.Salary_Generation.Net_Amount), SUM(0.05 * dbo.Salary_Generation.Total_Earn_Amount), SUM(0.1 * dbo.Salary_Generation.Total_Earn_Amount), SUM(dbo.Salary_Generation.Goods_Amount), SUM(dbo.Salary_Generation.LST_Amount) " & _
"FROM dbo.Branch_Master INNER JOIN " & _
                  "dbo.Increment ON dbo.Branch_Master.Branch_ID = dbo.Increment.Branch_ID INNER JOIN " & _
                  "dbo.Salary_Generation ON dbo.Increment.Increment_Id = dbo.Salary_Generation.Increment_ID INNER JOIN " & _
                  "dbo.PRODUCT_MASTER ON dbo.Increment.Product_ID = dbo.PRODUCT_MASTER.Product_ID " & _
"WHERE (dbo.Salary_Generation.Month = " + MonthVal + ") AND (dbo.Salary_Generation.Year = " + Year + ") AND (dbo.PRODUCT_MASTER.Product_Name = 'Support') AND " & _
                  "(dbo.Branch_Master.Branch_Name = 'GOU')"

        Dim GOUSCmd As New SqlCommand(GOUS, conn)
        Dim GOUSda As New SqlDataAdapter(GOUSCmd)
        Dim GOUSds As New DataSet()
        conn.Open()
        GOUSda.Fill(ds, "tblSalaryGOUS")
        conn.Close()


        ProgressBar1.Value = 70

        Dim Gratuity As String = "SELECT round(SUM(0.11 * dbo.GRATUITY.Basic),0)" & _
"FROM dbo.GRATUITY WHERE (dbo.GRATUITY.Month = " + MonthVal + ") AND (dbo.GRATUITY.Year = " + Year + ")"

        Dim GratuityCmd As New SqlCommand(Gratuity, conn)
        Dim GratuityDa As New SqlDataAdapter(GratuityCmd)
        Dim GratuityDs As New DataSet()
        conn.Open()
        GratuityDa.Fill(GratuityDs, "tblGratuity")
        conn.Close()

        Dim EDLoanDeduction As String = "SELECT ISNULL(dbo.Emp_Loan_Payment_Detail.Amount,0) AS Loan " & _
"FROM dbo.Emp_Loan_Payment INNER JOIN " & _
"dbo.Emp_Loan_Payment_Detail ON dbo.Emp_Loan_Payment.Loan_Payment_Id = dbo.Emp_Loan_Payment_Detail.Loan_Payment_Id " & _
"WHERE (dbo.Emp_Loan_Payment.Emp_Id = 1) AND (DATEPART(MM, dbo.Emp_Loan_Payment.Payment_Date) = " + MonthVal + ") AND (DATEPART(YYYY, dbo.Emp_Loan_Payment.Payment_Date) = " + Year + ")"

        Dim EDLoanDeductionCmd As New SqlCommand(EDLoanDeduction, conn)
        Dim dsEDLoanDedu As New DataSet()
        Dim daEDLoanDedu As New SqlDataAdapter(EDLoanDeductionCmd)
        conn.Open()
        daEDLoanDedu.Fill(dsEDLoanDedu, "tblEDLoan")
        conn.Close()

        Dim EDHalfLoan As Decimal
        Dim recordCount = dsEDLoanDedu.Tables("tblEDLoan").Rows.Count
        If recordCount = 0 Then
            EDHalfLoan = 0
        Else
            EDHalfLoan = dsEDLoanDedu.Tables("tblEDLoan").Rows(0).Item(0)
        End If


        '-----------------------NET PAY-------------------------------

        Dim NETpath As String = "D:\Payroll " + Month + " " + Year + "\Net Pay.txt"
        Dim fs As FileStream = File.Create(NETpath)

        Dim NetLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "Net Pay For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        fs.Write(NetLine1, 0, NetLine1.Length)

        Dim GOURNet As Decimal = ds.Tables("tblSalary").Rows(0).Item(0)
        Dim GOURNetNew As Decimal = (GOURNet - EDHalfLoan).ToString


        Dim NetLine4 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + GOURNetNew.ToString + Environment.NewLine)
        fs.Write(NetLine4, 0, NetLine4.Length)

        Dim NetLine5 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + GOURNetNew.ToString + Environment.NewLine)
        fs.Write(NetLine5, 0, NetLine5.Length)

        Dim NetLine10 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine10, 0, NetLine10.Length)

        Dim NetLine11 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine11, 0, NetLine11.Length)


        '----------------------PAYE----------------------------


        Dim PAYEPath As String = "D:\Payroll " + Month + " " + Year + "\PAYE.txt"
        Dim PAYEFs As FileStream = File.Create(PAYEPath)

        Dim PAYELine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "PAYE For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        PAYEFs.Write(PAYELine1, 0, PAYELine1.Length)

        Dim PAYELine2 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine2, 0, PAYELine2.Length)

        Dim PAYELine3 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine3, 0, PAYELine3.Length)


        Dim PAYELine4 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalaryGOUR").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine4, 0, PAYELine4.Length)

        Dim PAYELine5 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryGOUR").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine5, 0, PAYELine5.Length)


        ProgressBar1.Value = 80
        '-------------------------------NSSF---------------------------

        Dim NSSFPath As String = "D:\Payroll " + Month + " " + Year + "\NSSF.txt"
        Dim NSSFFs As FileStream = File.Create(NSSFPath)

        Dim NSSFLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "NSSF For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        NSSFFs.Write(NSSFLine1, 0, NSSFLine1.Length)

        Dim NSSF5Line2 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalary").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line2, 0, NSSF5Line2.Length)

        Dim NSSF5Line3 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalary").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line3, 0, NSSF5Line3.Length)


        Dim NSSF10Line4 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalary").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line4, 0, NSSF10Line4.Length)

        Dim NSSF10Line5 As Byte() = New UTF8Encoding(True).GetBytes("20000007," + ds.Tables("tblSalary").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line5, 0, NSSF10Line5.Length)


        Dim NSSF5Line6 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line6, 0, NSSF5Line6.Length)

        Dim NSSF5Line7 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line7, 0, NSSF5Line7.Length)


        Dim NSSF10Line8 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line8, 0, NSSF10Line8.Length)

        Dim NSSF10Line9 As Byte() = New UTF8Encoding(True).GetBytes("20000007," + ds.Tables("tblSalaryGOUS").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line9, 0, NSSF10Line9.Length)


        '---------------------------LST-----------------------------------

        ProgressBar1.Value = 90

        Dim LSTPath As String = "D:\Payroll " + Month + " " + Year + "\LST.txt"
        Dim LSTFs As FileStream = File.Create(LSTPath)

        Dim LSTLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "LST For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        LSTFs.Write(LSTLine1, 0, LSTLine1.Length)

        Dim LSTLine2 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalary").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine2, 0, LSTLine2.Length)

        Dim LSTLine3 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalary").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine3, 0, LSTLine3.Length)


        Dim LSTLine4 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalaryGOUR").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine4, 0, LSTLine4.Length)

        Dim LSTLine5 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryGOUR").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine5, 0, LSTLine5.Length)

        Dim LSTLine10 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine10, 0, LSTLine10.Length)

        Dim LSTLine11 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine11, 0, LSTLine11.Length)


        ProgressBar1.Value = 100


        '---------------------------Gratuity--------------------------


        Dim GratuityPath As String = "D:\Payroll " + Month + " " + Year + "\Gratuity.txt"
        Dim GratuityFs As FileStream = File.Create(GratuityPath)

        Dim GratuityLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "Gratuity " + Month + " " + Year + Chr(34) + Environment.NewLine)
        GratuityFs.Write(GratuityLine1, 0, GratuityLine1.Length)

        Dim GratuityLine2 As Byte() = New UTF8Encoding(True).GetBytes("94000000," + "-" + GratuityDs.Tables("tblGratuity").Rows(0).Item(0).ToString + Environment.NewLine)
        GratuityFs.Write(GratuityLine2, 0, GratuityLine2.Length)

        Dim GratuityLine3 As Byte() = New UTF8Encoding(True).GetBytes("24000010," + GratuityDs.Tables("tblGratuity").Rows(0).Item(0).ToString)
        GratuityFs.Write(GratuityLine3, 0, GratuityLine3.Length)
        GratuityFs.Close()



        '---------------------------Advance--------------------------


        Dim AdvancePath As String = "D:\Payroll " + Month + " " + Year + "\Advance.txt"
        Dim AdvanceFs As FileStream = File.Create(AdvancePath)

        Dim AdvanceLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "Salary Advance Recovery - " + Month + " " + Year + Chr(34) + Environment.NewLine)
        AdvanceFs.Write(AdvanceLine1, 0, AdvanceLine1.Length)


        Dim GOURadvanceDs As New DataSet
        Dim GOURadvanceQuery As String = "select Account,Loan_Amount from [Salary Advances] where Branch_Name = 'gou' and Product_Name= 'Research' and Month = " + MonthVal + " and Year =" + Year
        Dim GOURadvanceCmd As New SqlCommand(GOURadvanceQuery, conn)
        Dim GOURadvanceDa As New SqlDataAdapter(GOURadvanceCmd)

        Try
            GOURadvanceDa.Fill(GOURadvanceDs, "Salary Advances")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim GOUResearchAdvances As Byte()
        Dim GOURadvanceTotal As Decimal = 0

        For i = 0 To GOURadvanceDs.Tables("Salary Advances").Rows.Count - 1
            With GOURadvanceDs.Tables("Salary Advances")
                GOUResearchAdvances = New UTF8Encoding(True).GetBytes(.Rows(i).Item(0).ToString + ",-" + .Rows(i).Item(1).ToString + Environment.NewLine)
                AdvanceFs.Write(GOUResearchAdvances, 0, GOUResearchAdvances.Length)
                GOURadvanceTotal = GOURadvanceTotal + .Rows(i).Item(1)
            End With
        Next

        Dim GOUResearchAdvancesDebit As Byte()
        GOUResearchAdvancesDebit = New UTF8Encoding(True).GetBytes("20000005," + GOURadvanceTotal.ToString + Environment.NewLine)
        AdvanceFs.Write(GOUResearchAdvancesDebit, 0, GOUResearchAdvancesDebit.Length)



        Dim TTIRadvanceDs As New DataSet
        Dim TTIRadvanceQuery As String = "select Account,Loan_Amount from [Salary Advances] where Branch_Name = 'TTI' and Product_Name= 'Research' and Month =" + MonthVal + " and Year = " + Year
        Dim TTIRadvanceCmd As New SqlCommand(TTIRadvanceQuery, conn)
        Dim TTIRadvanceDa As New SqlDataAdapter(TTIRadvanceCmd)

        Try
            TTIRadvanceDa.Fill(TTIRadvanceDs, "Salary Advances")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim TTIResearchAdvances As Byte()
        Dim TTIRadvanceTotal As Decimal = 0

        For i = 0 To TTIRadvanceDs.Tables("Salary Advances").Rows.Count - 1
            With TTIRadvanceDs.Tables("Salary Advances")
                TTIResearchAdvances = New UTF8Encoding(True).GetBytes(.Rows(i).Item(0).ToString + ",-" + .Rows(i).Item(1).ToString + Environment.NewLine)
                AdvanceFs.Write(TTIResearchAdvances, 0, TTIResearchAdvances.Length)
                TTIRadvanceTotal = TTIRadvanceTotal + .Rows(i).Item(1)
            End With
        Next

        Dim TTIResearchAdvancesDebit As Byte()
        TTIResearchAdvancesDebit = New UTF8Encoding(True).GetBytes("20000005," + TTIRadvanceTotal.ToString + Environment.NewLine)
        AdvanceFs.Write(TTIResearchAdvancesDebit, 0, TTIResearchAdvancesDebit.Length)



        Dim GOUSadvanceDs As New DataSet
        Dim GOUSadvanceQuery As String = "select Account,Loan_Amount from [Salary Advances] where Branch_Name = 'GOU' and Product_Name= 'Support' and Month = " + MonthVal + " and Year =" + Year
        Dim GOUSadvanceCmd As New SqlCommand(GOUSadvanceQuery, conn)
        Dim GOUSadvanceDa As New SqlDataAdapter(GOUSadvanceCmd)

        Try
            GOUSadvanceDa.Fill(GOUSadvanceDs, "Salary Advances")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim GOUSesearchAdvances As Byte()
        Dim GOUSadvanceTotal As Decimal = 0

        For i = 0 To GOUSadvanceDs.Tables("Salary Advances").Rows.Count - 1
            With GOUSadvanceDs.Tables("Salary Advances")
                GOUSesearchAdvances = New UTF8Encoding(True).GetBytes(.Rows(i).Item(0).ToString + ",-" + .Rows(i).Item(1).ToString + Environment.NewLine)
                AdvanceFs.Write(GOUSesearchAdvances, 0, GOUSesearchAdvances.Length)
                GOUSadvanceTotal = GOUSadvanceTotal + .Rows(i).Item(1)
            End With
        Next

        Dim GOUSesearchAdvancesDebit As Byte()
        GOUSesearchAdvancesDebit = New UTF8Encoding(True).GetBytes("24000005," + GOUSadvanceTotal.ToString + Environment.NewLine)
        AdvanceFs.Write(GOUSesearchAdvancesDebit, 0, GOUSesearchAdvancesDebit.Length)



        Dim EPRCRadvanceDs As New DataSet
        Dim EPRCRadvanceQuery As String = "select Account,Loan_Amount from [Salary Advances] where Branch_Name = 'EPRC' and Product_Name= 'Research' and Month = " + MonthVal + " and Year =" + Year
        Dim EPRCRadvanceCmd As New SqlCommand(EPRCRadvanceQuery, conn)
        Dim EPRCRadvanceDa As New SqlDataAdapter(EPRCRadvanceCmd)

        Try
            EPRCRadvanceDa.Fill(EPRCRadvanceDs, "Salary Advances")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim EPRCResearchAdvances As Byte()
        Dim EPRCRadvanceTotal As Decimal = 0

        For i = 0 To EPRCRadvanceDs.Tables("Salary Advances").Rows.Count - 1
            With EPRCRadvanceDs.Tables("Salary Advances")
                EPRCResearchAdvances = New UTF8Encoding(True).GetBytes(.Rows(i).Item(0).ToString + ",-" + .Rows(i).Item(1).ToString + Environment.NewLine)
                AdvanceFs.Write(EPRCResearchAdvances, 0, EPRCResearchAdvances.Length)
                EPRCRadvanceTotal = EPRCRadvanceTotal + .Rows(i).Item(1)
            End With
        Next

        Dim EPRCResearchAdvancesDebit As Byte()
        EPRCResearchAdvancesDebit = New UTF8Encoding(True).GetBytes("24000005," + EPRCRadvanceTotal.ToString + Environment.NewLine)
        AdvanceFs.Write(EPRCResearchAdvancesDebit, 0, EPRCResearchAdvancesDebit.Length)




        AdvanceFs.Close()

        Application.Exit()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 0

        ComboBox1.Text = "March"
        ComboBox2.Text = "2017"
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class
