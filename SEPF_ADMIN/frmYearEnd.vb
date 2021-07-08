Option Strict Off
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc

Imports Microsoft.Office.Interop.Excel 'Before you add this reference to your project,
' you need to install Microsoft Office and find last version of this file.

Imports Microsoft.Office.Interop
Imports System.IO

'rpadath 11/3/2019
'For year End process



Public Class frmYearEnd
    'Connection used in the frm
    Dim cn As SqlConnection
    'Transaction Object
    Dim objTrans As SqlTransaction

    Dim sPath As String

    Private Sub cmdCreateTable_Click(sender As Object, e As EventArgs) Handles cmdCreateTable.Click

        'Status label
        Label1.Text = "Creating Actuarial Tables  - Started "
        Label1.Refresh()

        'Message
        If MsgBox("You are about to Run the Actuarial create Table process. This process takes approximately 15 minutes. Do you wish to continue ", MsgBoxStyle.YesNo, "Create Actuarial Tables") = MsgBoxResult.Yes Then
            ' do the process
            Try

                Label1.Text = "Started -------"
                Label1.Refresh()

                'Proc to create Acturial tables
                If CreateActuarialTables() Then

                    Label1.Text = " -- Actuarial Tables created -------"
                    Label1.Refresh()

                End If

            Catch ex As Exception

                MsgBox(ex.Message)

            End Try

        Else
            'abort
        End If

    End Sub

    Private Function CreateActuarialTables() As Boolean
        'create actuarial tables

        'Begin trans
        objTrans = cn.BeginTransaction()

        Try
            'create participant tables
            If GetParticipant() Then
                'commit
                objTrans.Commit()
                CreateActuarialTables = True

            End If


        Catch es As SqlException
            'objTrans.Rollback()
            CreateActuarialTables = False
            MsgBox(es.Message)

        Catch ex As Exception
            '***** rollback the transaction ************************************************'
            objTrans.Rollback()
            CreateActuarialTables = False
            MsgBox(ex.Message)

        Finally
            'If cn.State <> ConnectionState.Closed Then cn.Close()
        End Try

        'objTrans.Commit()
        'Step 1


    End Function

    Private Sub frmYearEnd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If ConnectDataBase() = True Then
            ' MsgBox("Connected")
        End If

        sPath = "\\Tile\SEPF\Actuarial_Files\Cheiron_" & CStr(DatePart(DateInterval.Year, Date.Now) - 1) & "\Sepf-net\"

        txtPath.Text = sPath
    End Sub

    Private Function ConnectDataBase() As Boolean

        'call this to connect

        ConnectDataBase = False

        Try

            ' Open a database connection.
            'Dim strConnection As String =
            '   "Data Source=brickyard;Initial Catalog=SEPF-NET;" _
            '   & "Integrated Security=True;MultipleActiveResultSets=True"

            'Dim strConnectionAWS As String =
            '   "Data Source=mytestsqldb1.cxe9r3dczmpq.us-east-1.rds.amazonaws.com,1433;Initial Catalog=TestDB;" _
            '   & "User id=testdbmaster1;Password=testdbpassword1"

            'Dim strConnectionAWS1 As String =
            '   "Data Source=bacsqlserver.cxe9r3dczmpq.us-east-1.rds.amazonaws.com,1433;Initial Catalog=SEPF-NET;" _
            '   & "User id=bacSQLAdmin;Password=!8Baltimore65"

            'Dim strConnectionAWS2 As String =
            '   "Data Source=bacsqlserver.cxe9r3dczmpq.us-east-1.rds.amazonaws.com,1433;Initial Catalog=SEPF-NET;" _
            '   & "User id=sepfuser;Password=passwordsepf"

            Dim strConnectionAWS3 As String =
                 "Data Source=bacsqlprod.cxe9r3dczmpq.us-east-1.rds.amazonaws.com,1433;Initial Catalog=SEPF-NET;" _
                 & "User id=bacsqladmin;Password=sepfAdmin#;MultipleActiveResultSets=True"

            cn = New SqlConnection(strConnectionAWS3)
            cn.Open()
            ConnectDataBase = True
            Label1.Text = "Connected to Database"
            Label1.Refresh()

            'cn = New SqlConnection(Mine.ConnString.SEPFConnection)
            'cn.Open()
            'ConnectDataBase = True
            'Label1.Text = "Connected to Database"
            'Label1.Refresh()

        Catch ex As Exception

            ConnectDataBase = False

            Label1.Text = "Connection to Database -  Fail "
            Label1.Refresh()
        End Try

    End Function

    Private Function GetParticipant() As Boolean
        Dim sql As String
        Dim cmdDelete1 As SqlCommand
        Dim cmdDelete2 As SqlCommand
        Dim retValue As Integer

        Label1.Text = "Started -- "
        Label1.Refresh()

        Try

            'delete contents of tables first

            sql = "delete from tempActuarial"
            cmdDelete1 = New SqlCommand(sql, cn, objTrans)
            retValue = cmdDelete1.ExecuteNonQuery()

            sql = "delete from keyemployees"
            cmdDelete2 = New SqlCommand(sql, cn, objTrans)
            retValue = cmdDelete2.ExecuteNonQuery()

            'start filling dtata
            sql = "SELECT * FROM tempActuarial"

            ' Create Data Adapter
            Dim da As New SqlDataAdapter
            da.SelectCommand = New SqlCommand(sql, cn, objTrans)

            ' Create and fill Dataset
            Dim ds As New DataSet
            da.Fill(ds, "tempActuarial")

            ' Get the Data Table
            Dim dt As System.Data.DataTable = ds.Tables("tempActuarial")

            Dim builder As SqlCommandBuilder = New SqlCommandBuilder(da)

            builder.GetInsertCommand()

            da.TableMappings.Add("Table", "tempActuarial")

            ' Dim partCMD As SqlCommand = New SqlCommand("select * from participant where Part_id =  3195", cn, objTrans)
            '566370413

            ' Dim partCMD As SqlCommand = New SqlCommand("select * from participant where Part_ssn =  '566370413'", cn, objTrans)

            Dim partCMD As SqlCommand = New SqlCommand("select * from participant", cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read

                Label1.Text = " ---------  Processing " & myReader("Part_ID").ToString & " ------------------------"

                Label1.Refresh()

                Select Case myReader("Part_Status").ToString

                    Case "A"
                        'Call
                        ActiveParticipant(myReader, dt)
                    Case "I"
                        InActiveParticipant(myReader, dt)
                    Case "P"
                        PensionParticipant(myReader, dt)
                    Case "D"
                        If IsDBNull(myReader("Part_Retirement_Date")) Then
                            If Year(myReader("Part_DOD")) > Year(DateAdd(DateInterval.Year, -1, Date.Now)) And
                            IsDBNull(myReader("Part_Termination_Date")) Or Year(myReader("Part_Termination_Date")) > Year(DateAdd(DateInterval.Year, -1, Date.Now)) Then

                                ActiveParticipant(myReader, dt)

                            ElseIf Year(myReader("Part_DOD")) = Year(DateAdd(DateInterval.Year, -1, Date.Now)) _
                                And Year(myReader("Part_Termination_Date")) > Year(DateAdd(DateInterval.Year, -1, Date.Now)) Then

                                ActiveParticipant(myReader, dt)
                            Else
                                KeyEmployee(myReader)
                            End If

                        Else
                            If Year(myReader("Part_Retirement_Date")) <= Year(DateAdd(DateInterval.Year, -1, Date.Now)) _
                               And Year(myReader("Part_DOD")) = Year(DateAdd(DateInterval.Year, -1, Date.Now)) Then

                                PensionParticipant(myReader, dt)

                            End If
                        End If

                    Case "S"

                        If IsDBNull(myReader("Part_Retirement_Date")) Then
                        ElseIf IsDBNull(myReader("Part_Retirement_Date")) < Year(DateAdd(DateInterval.Year, -1, Date.Now)) _
                            And Year(myReader("Part_DOD")) = Year(DateAdd(DateInterval.Year, -1, Date.Now)) Then

                            PensionParticipant(myReader, dt)

                        End If

                End Select

            Loop

            '--added
            da.Update(ds)
            myReader.Close()
            GetParticipant = True

        Catch ex As Exception

            GetParticipant = False
            MsgBox(ex.Message)

            Throw ex

        End Try

        GetParticipant = True

        Label1.Text = "-------------- Complete ---------------"
        Label1.Refresh()

    End Function

    Private Function ActiveParticipant(rdReader As SqlDataReader, dsNew As System.Data.DataTable) As Boolean

        'active participant

        Try

            Dim iDate As String = "12/31/" + Convert.ToString(DatePart(DateInterval.Year, DateAdd(DateInterval.Year, -1, Date.Now)))
            Dim oDate As DateTime = Convert.ToDateTime(iDate)

            Dim newRow As DataRow = dsNew.NewRow

            newRow("ParticipantID") = rdReader!Part_ID
            newRow("contribution1990WI") = GetContribInterest(rdReader!Part_ID, 1)
            newRow("contribution1991WI") = GetContribInterest(rdReader!Part_ID, 2)
            newRow("contribution1990NI") = GetContribNoInterest(rdReader!Part_ID, 1991, "<")
            newRow("contribution1991NI") = GetContribNoInterest(rdReader!Part_ID, 1990, ">")
            newRow("ActuarialStatus") = "A"
            newRow("ServiceCreditYear") = ServiceCredit(rdReader!Part_ID, "Y", oDate)
            newRow("ServiceCreditMonth") = ServiceCredit(rdReader!Part_ID, "M", oDate)

            dsNew.Rows.Add(newRow)

            KeyEmployee(rdReader)

            ActiveParticipant = True

        Catch ex As Exception

            MsgBox(ex.Message)

            ActiveParticipant = False

            Throw ex
        End Try



    End Function

    Private Function InActiveParticipant(rdReader As SqlDataReader, dsNew As System.Data.DataTable) As Boolean

        Dim dContribInterest As Double
        Dim dContribNoInterest As Double

        Dim dContrib1 As Double
        Dim dContrib2 As Double

        Dim iDate As String = "12/31/" + Convert.ToString(DatePart(DateInterval.Year, DateAdd(DateInterval.Year, -1, Date.Now)))
        Dim oDate As DateTime = Convert.ToDateTime(iDate)

        If Year(rdReader("Part_Termination_Date")) > Year(oDate) Then
            ActiveParticipant(rdReader, dsNew)
            Exit Function
        End If

        Dim newRow As DataRow = dsNew.NewRow

        Try

            newRow("ParticipantID") = rdReader!Part_ID
            newRow("ActuarialStatus") = "T"
            newRow("ServiceCreditYear") = 0
            newRow("ServiceCreditMonth") = 0

            dContribInterest = GetContribInterest(rdReader!Part_ID, 4)

            dContribNoInterest = GetContribNoInterest(rdReader!Part_ID, 1957, ">")

            dContribNoInterest = 2 * dContribNoInterest


            'If IsDate(GetPartDateVested(rdReader!Part_ID)) = False Then

            If Year(GetPartDateVested(rdReader!Part_ID)) < 1900 Then
                dContrib2 = dContribInterest

            ElseIf dContribInterest > dContribNoInterest Then
                dContrib2 = dContribInterest
            Else
                dContrib2 = dContribNoInterest
            End If

            If rdReader!Part_Lump_Sum_Payment > dContrib2 Then
                newRow("ContributionsRemaining") = 0
            Else
                newRow("ContributionsRemaining") = dContrib2 - rdReader!Part_Lump_Sum_Payment
            End If

            newRow("ContributionsWithInterest") = dContrib2

            If rdReader!Part_Ineligible_Service_Months > 0 Then
                newRow("ContributionsWithInterest") = rdReader!Part_Lump_Sum_Payment
            End If

            If newRow("ContributionsRemaining") <> 0 And
            newRow("ContributionsRemaining") <> dContrib2 Then
                newRow("ContributionsRemaining") = 0
            End If

            If Year(DateAdd(DateInterval.Year, -1, rdReader!Part_Termination_Date)) = Year(oDate) Then

                If rdReader!Part_Lump_Sum_Payment <> 0 And rdReader!Part_Lump_Sum_Payment <> dContrib2 Then
                    newRow("ContributionsWithInterest") = rdReader!Part_Lump_Sum_Payment
                End If
            End If

            'dsNew.Tables("tempActuarial").Rows.Add(newRow)
            dsNew.Rows.Add(newRow)

            KeyEmployee(rdReader)


            InActiveParticipant = True

        Catch ex As Exception
            MsgBox(ex.Message)
            InActiveParticipant = False
            Throw ex

        End Try

        InActiveParticipant = True
    End Function

    Private Function KeyEmployee(rdReader As SqlDataReader) As Boolean

        Dim Year4 As Integer
        Dim sSql As String
        Dim dAverageMonthSal As Decimal

        'Exit Function

        Try

            Dim iDate As String = "12/31/" + Convert.ToString(DatePart(DateInterval.Year, DateAdd(DateInterval.Year, -1, Date.Now)))
            Dim oDate As DateTime = Convert.ToDateTime(iDate)

            Year4 = DatePart(DateInterval.Year, DateAdd(DateInterval.Year, -4, oDate))

            '-exit if they are Not officer Or board member
            If rdReader("Part_OE_Flag") = "O" Or rdReader("Part_OE_Flag") = "B" Then
            Else
                Exit Function
            End If

            ' Or
            'rdReader("Part_Retirement_Date") < DateAdd(DateInterval.Year, -4, oDate) Then

            'check termination date
            If IsDate(rdReader("Part_Termination_Date")) = False Then

            Else
                ' if terminated or retired
                If rdReader("Part_Termination_Date") < DateAdd(DateInterval.Year, -4, oDate) Then
                    Exit Function
                Else

                    If IsDate(rdReader("Part_Retirement_Date")) = False Then
                        Exit Function
                    ElseIf rdReader("Part_Retirement_Date") < DateAdd(DateInterval.Year, -4, oDate) Then
                        Exit Function
                    End If

                End If

            End If

            sSql = "Select * from V_ParticipantSalaryByYear where Part_ID = " & rdReader("Part_ID") & " and Salary > 160000.00 and Year >= " & Year4

            Dim cmdSalaray As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim mySalary As SqlDataReader = cmdSalaray.ExecuteReader()

            'start filling dtata
            sSql = "SELECT * FROM KeyEmployees"

            ' Create Data Adapter
            Dim da As New SqlDataAdapter
            da.SelectCommand = New SqlCommand(sSql, cn, objTrans)

            ' Create and fill Dataset
            Dim ds As New DataSet
            da.Fill(ds, "KeyEmployees")

            Dim builder As SqlCommandBuilder = New SqlCommandBuilder(da)

            builder.GetInsertCommand()

            da.TableMappings.Add("Table", "KeyEmployees")

            ' Get the Data Table
            Dim dt As System.Data.DataTable = ds.Tables("KeyEmployees")

            Dim newRow As DataRow = dt.NewRow

            newRow("Key_SSN") = rdReader!Part_SSN
            newRow("Key_Fund") = rdReader!Part_Fund
            newRow("Key_LastName") = rdReader!Part_LastName
            newRow("Key_FirstName") = rdReader!Part_FirstName
            newRow("Key_MiddleInitial") = rdReader!Part_MInitial
            newRow("Key_PensionDistribution") = PartBenefits(rdReader!Part_ID) + SurvivorBenefits(rdReader!Part_ID)
            'newRow("Key_PensionDistribution") = SurvivorBenefits(rdReader!Part_ID)

            Select Case rdReader!Part_Status
                Case "A"
                    newRow("Key_Status") = "A"
                Case "P"
                    newRow("Key_Status") = "C"
                Case "I"
                    'If IsNothing(GetPartDateVested(rdReader!Part_ID)) Then
                    If Year(GetPartDateVested(rdReader!Part_ID)) < 1900 Then
                        newRow("Key_Status") = "T"
                    Else
                        newRow("Key_Status") = "B"
                    End If
                Case "D"
                    newRow("Key_Status") = "D"
            End Select

            'rpadath
            '07072021

            If IsDate(rdReader("Part_Retirement_Date")) Then

                If IsDate(rdReader("Part_Rehire_Date")) = False Then

                    If rdReader("Part_Retirement_Date") < DateAdd(DateInterval.Year, 1, oDate) Then

                        newRow("Key_ActuarialBenefit") = rdReader("Part_Monthly_Benefit_Amount_AO")
                    End If

                Else

                    If rdReader("Part_Retirement_Date") > rdReader("Part_Rehire_Date") Then

                        If rdReader("Part_Retirement_Date") < DateAdd(DateInterval.Year, 1, oDate) Then

                            newRow("Key_ActuarialBenefit") = rdReader("Part_Monthly_Benefit_Amount_AO")
                        End If

                    End If

                End If

            Else

                dAverageMonthSal = ComputeMonthlyAverage(rdReader("Part_ID"))
                newRow("Key_ActuarialBenefit") = ComputePensionPercent(rdReader, dAverageMonthSal, oDate)
            End If


            If IsNumeric(newRow("Key_PensionDistribution")) Then
                '
            Else

                If Not IsDate(rdReader("Part_Last_Check_Date")) And rdReader("Part_Retirement_Date") = Year(Now.Date) Then

                    newRow("Key_DistributionEndDate") = "12/31/" + Convert.ToString(Year(DateAdd(DateInterval.Year, -1, Date.Now)))
                Else

                    newRow("Key_DistributionEndDate") = Month(rdReader("Part_Last_Check_Date")) + "/01/" + Year(DateAdd(DateInterval.Year, -1, Date.Now))

                End If

            End If
            'dsNew.Tables("tempActuarial").Rows.Add(newRow)
            dt.Rows.Add(newRow)

            da.Update(ds)

            mySalary.Close()

        Catch ex As Exception

            MsgBox(ex.Message)
            KeyEmployee = False
            Throw ex
        End Try

        KeyEmployee = True

    End Function

    Private Function GetContribInterest(PartID As Long, ret As Integer) As Double
        Dim sSql As String
        'Dim ret As Integer
        Dim ContribInterest As Double
        Dim lngCurYear As Long

        Try

            lngCurYear = DatePart(DateInterval.Year, Date.Now)

            sSql = "Select dbo.ContributionsWithInterest " & "(" & PartID & "," & ret & ")"

            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read
                ' Console.WriteLine("{0}", myReader.GetString(2))
                ContribInterest = myReader(0)
            Loop
            myReader.Close()

            GetContribInterest = ContribInterest

        Catch ex As Exception

            MsgBox(ex.Message)

            Throw ex

        End Try

    End Function

    Private Function GetContribNoInterest(PartID As Long, Yr As Integer, Op As String) As Double
        Dim sSql As String
        Dim ret As Integer
        Dim ContribNoInterest As Double
        Dim lngCurYear As Long

        Try

            lngCurYear = DatePart(DateInterval.Year, Date.Now)

            sSql = "Select isnull(sum(Psal_Pension_Contribution),0) from ParticipantSalary Where Part_ID =  " & PartID & " and year(psal_ContributionDate) " & Op & " " & Yr _
                & " and year(psal_ContributionDate) < " & lngCurYear

            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read
                ' Console.WriteLine("{0}", myReader.GetString(2))
                ContribNoInterest = myReader(0)
            Loop
            myReader.Close()

            GetContribNoInterest = ContribNoInterest

        Catch ex As Exception

            MsgBox(ex.Message)

            Throw ex

        End Try

    End Function

    Private Function GetPartDateVested(PartID As Long) As Date

        Dim sSql As String

        Dim dGetPartDateVested As Date
        Dim xDate As String


        Try

            sSql = "Select dbo.fn_Part_DateVested " & "(" & PartID & ")"

            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read
                ' Console.WriteLine("{0}", myReader.GetString(2))
                If IsDBNull(myReader(0)) Then
                    dGetPartDateVested = Nothing
                    'GetPartDateVested = Nothing
                    'xDate = dGetPartDateVested.ToString("d")
                    'Exit Function
                Else
                    dGetPartDateVested = myReader(0)
                End If

            Loop
            myReader.Close()

            GetPartDateVested = dGetPartDateVested

        Catch ex As Exception

            MsgBox(ex.Message)

            Throw ex

        End Try

    End Function

    Private Function GetLastBenefitCheckDate(PartID As Long) As Date
        Dim sSql As String

        Try
            sSql = "Select dbo.MaxBenefitDate " & "(" & PartID & ")"

            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReaderX As SqlDataReader = partCMD.ExecuteReader()

            Do While myReaderX.Read
                ' Console.WriteLine("{0}", myReader.GetString(2))
                GetLastBenefitCheckDate = myReaderX(0)
            Loop
            myReaderX.Close()

        Catch ex As Exception
            MsgBox(ex.Message)

            Throw ex

        End Try

    End Function

    Private Function PensionParticipant(rdReader As SqlDataReader, dsNew As System.Data.DataTable) As Boolean

        Try

            Dim iDate As String = "12/31/" + Convert.ToString(DatePart(DateInterval.Year, DateAdd(DateInterval.Year, -1, Date.Now)))
            Dim oDate As DateTime = Convert.ToDateTime(iDate)

            If Year(rdReader("Part_Award_Date")) > Year(oDate) Then
                ActiveParticipant(rdReader, dsNew)
                Exit Function
            End If

            Dim dLastBenDate As Date
            Dim iLastbenYr As Integer

            dLastBenDate = GetLastBenefitCheckDate(rdReader("Part_ID"))

            iLastbenYr = Year(dLastBenDate)

            If Year(dLastBenDate) >= Year(oDate) Then

                Dim newRow As DataRow = dsNew.NewRow

                newRow("ParticipantID") = rdReader!Part_ID
                newRow("ActuarialStatus") = "P"
                newRow("ContributionsWithInterest") = GetContribInterest(rdReader!Part_ID, 4)

                'dsNew.Tables("tempActuarial").Rows.Add(newRow)
                dsNew.Rows.Add(newRow)

            End If


            KeyEmployee(rdReader)

        Catch ex As Exception

            MsgBox(ex.Message)

            PensionParticipant = False

            Throw ex

        End Try

        PensionParticipant = True

    End Function

    Private Function ServiceCredit(PartID As Long, Ret As String, InDate As Date) As Decimal

        Dim sSql As String
        Dim dServiceCredit As Double
        Dim lngCurYear As Long

        Try

            lngCurYear = DatePart(DateInterval.Year, Date.Now)

            'sSql = "Select dbo.ServiceCredits ( " & PartID & ",'" & InDate & "',' " & Ret & "')"

            sSql = "Select dbo.fn_Part_ServiceCredits ( " & PartID & ",'" & Ret & "',' " & InDate & "')"


            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read
                ' Console.WriteLine("{0}", myReader.GetString(2))
                dServiceCredit = myReader(0)
            Loop
            myReader.Close()

            ServiceCredit = dServiceCredit

        Catch ex As Exception

            MsgBox(ex.Message)
            Throw ex

        End Try


    End Function

    Private Function PartBenefits(PartID As Long) As Decimal

        Dim sSql As String
        Dim dPartBenefit As Double
        Dim lngCurYear As Long

        PartBenefits = 0

        Try

            lngCurYear = DatePart(DateInterval.Year, Date.Now)

            sSql = "Select * from v_ParticipantBenefitByYear where Part_ID  = " & PartID


            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read
                If myReader("year") > Year(DateAdd(DateInterval.Year, -5, Now.Date)) And myReader("year") < Year(Now.Date) Then

                    dPartBenefit = dPartBenefit + myReader("benefit")

                End If
            Loop
            myReader.Close()

            PartBenefits = dPartBenefit

        Catch ex As Exception

            MsgBox(ex.Message)
            Throw ex

        End Try



    End Function

    Private Function SurvivorBenefits(PartID As Long) As Decimal

        Dim sSql As String
        Dim dPensionContrib As Double
        Dim lngCurYear As Long

        SurvivorBenefits = 0

        Try

            lngCurYear = DatePart(DateInterval.Year, Date.Now)

            sSql = "Select * from v_SurvivorBenefitByYear where Part_ID  = " & PartID


            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read

                If myReader("year") > Year(DateAdd(DateInterval.Year, -5, Now.Date)) And myReader("year") < Year(Now.Date) Then

                    dPensionContrib = dPensionContrib + myReader("benefit")

                End If
            Loop
            myReader.Close()

            SurvivorBenefits = dPensionContrib

        Catch ex As Exception

            MsgBox(ex.Message)

            Throw ex


        End Try



    End Function


    Private Function ComputeMonthlyAverage(PartID As Long) As Decimal

        Dim sSql As String
        Dim dComputeMonthlyAverage As Double
        Dim lngCurYear As Long

        Try

            lngCurYear = DatePart(DateInterval.Year, Date.Now)

            sSql = "dbo.proGet36MonthSalary ( " & PartID & ")"

            sSql = "dbo.proGet36MonthSalary " & PartID


            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read
                If IsDBNull(myReader("yr")) Then

                    dComputeMonthlyAverage = myReader("salary") / 36

                End If
            Loop
            myReader.Close()

            ComputeMonthlyAverage = dComputeMonthlyAverage

        Catch ex As Exception

            MsgBox(ex.Message)

            Throw ex

        End Try


    End Function


    Private Function ComputePensionPercent(rdReader As SqlDataReader, AverageMonthSal As Decimal, actDate As Date) As Decimal

        Dim intTotCredit As Integer
        Dim dPercent As Decimal
        Dim iPercent As Integer

        Try

            If Not IsDBNull(rdReader("Part_Past_Service_Credits")) Then

                intTotCredit = rdReader("Part_Past_Service_Credits") * 12

            End If

            intTotCredit = intTotCredit + ServiceCredit(rdReader("Part_ID"), "F", actDate)



            If (rdReader("Part_Date_Hired")) > DateSerial(2009, 12, 31) Then


                If intTotCredit < 241 Then
                    dPercent = (intTotCredit * 0.03) / 12
                End If


                If intTotCredit > 240 And intTotCredit < 361 Then
                    dPercent = (((intTotCredit - 240) * 0.02) / 12 + 0.6)
                End If

            Else

                If intTotCredit < 241 Then
                    dPercent = (intTotCredit * 0.035) / 12
                End If


                If intTotCredit > 240 And intTotCredit < 361 Then
                    dPercent = ((intTotCredit - 240) * 0.01) / 12 + 0.7
                End If

            End If


            If intTotCredit > 360 Then
                dPercent = 0.8
            End If


            iPercent = dPercent * 100


            ComputePensionPercent = AverageMonthSal * 12 * dPercent / 12

        Catch ex As Exception

            MsgBox(ex.Message)

            Throw ex

        End Try




    End Function


    Private Function GetRate(SalYear As Integer) As Decimal

        Dim sSql As String
        Dim dPartBenefit As Double
        Dim lngCurYear As Long

        Try

            lngCurYear = DatePart(DateInterval.Year, Date.Now)

            sSql = "dbo.proGetRateSelect ( " & SalYear & ")"


            Dim partCMD As SqlCommand = New SqlCommand(sSql, cn, objTrans)

            Dim myReader As SqlDataReader = partCMD.ExecuteReader()

            Do While myReader.Read

                GetRate = myReader("Rate_Amount") / 100

                If SalYear < 1998 Then

                    GetRate = 0.05

                End If

            Loop
            myReader.Close()

        Catch ex As Exception

            MsgBox(ex.Message)

            Throw ex

        End Try


    End Function


    'Private Sub Button1_Click(sender As Object, e As EventArgs)

    '    ' Dim objTrans1 As SqlTransaction

    '    'objTrans1 = cn.BeginTransaction()


    '    'start filling dtata
    '    'Dim Sql As String = "SELECT * FROM tempActuarial"

    '    Dim Sql As String = "SELECT * FROM tbltest1"

    '    ' Create Data Adapter
    '    Dim da As New SqlDataAdapter
    '    'da.SelectCommand = New SqlCommand(Sql, cn, objTrans)

    '    da.SelectCommand = New SqlCommand(Sql, cn)

    '    ' Create and fill Dataset
    '    Dim ds As New DataSet
    '    ' da.Fill(ds, "tempActuarial")
    '    da.Fill(ds, "tblTest1")

    '    ' Get the Data Table
    '    'Dim dt As DataTable = ds.Tables("tempActuarial")

    '    Dim dt As DataTable = ds.Tables("tblTest1")

    '    'Dim iDate As String = "12/31/" + Convert.ToString(DatePart(DateInterval.Year, DateAdd(DateInterval.Year, -1, Date.Now)))
    '    'Dim oDate As DateTime = Convert.ToDateTime(iDate)

    '    Dim newRow As DataRow = dt.NewRow()

    '    ' Dim newRow1 As DataRow = dt.NewRow()

    '    'newRow("ParticipantID") = -2
    '    'newRow("contribution1990WI") = 0
    '    'newRow("contribution1991WI") = 0
    '    'newRow("contribution1990NI") = 0
    '    'newRow("contribution1991NI") = 0
    '    'newRow("ActuarialStatus") = "A"
    '    'newRow("ServiceCreditYear") = 0
    '    'newRow("ServiceCreditMonth") = 0

    '    'test
    '    newRow("ID") = -2
    '    newRow("Name") = "test"

    '    'dsNew.Tables("tempActuarial").Rows.Add(newRow)
    '    dt.Rows.Add(newRow)

    '    MessageBox.Show(da.UpdateCommand.ToString)

    '    da.Update(ds, "tblTest")

    '    ''--sample
    '    'Dim table1 As DataTable = New DataTable("patients")
    '    'table1.Columns.Add("name")
    '    'table1.Columns.Add("id")
    '    'table1.Rows.Add("sam", 1)
    '    'table1.Rows.Add("mark", 2)


    'End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs)



    '    Dim Sql As String = "select ID,Name from tbltest1 where ID  =  -1"

    '    ' Create Data Adapter
    '    Dim da As New SqlDataAdapter

    '    da.SelectCommand = New SqlCommand(Sql, cn)


    '    ' Create and fill Dataset
    '    Dim ds As New DataSet

    '    ' Dim dt As DataTable = ds.Tables("tblTest1")

    '    'da.Fill(ds, "tblTest1")
    '    da.MissingSchemaAction = MissingSchemaAction.AddWithKey

    '    da.FillSchema(ds, SchemaType.Mapped, "tbltest1")

    '    ' da.Fill(ds, "tblTest1")

    '    Dim dt As DataTable = ds.Tables("tblTest1")

    '    dt.AcceptChanges()

    '    Dim newRow As DataRow = dt.NewRow()

    '    'test
    '    newRow("ID") = -2
    '    newRow("Name") = "test"

    '    dt.Rows.Add(newRow)



    '    MessageBox.Show(da.UpdateCommand.ToString)



    '    'da.Update()
    'End Sub


    ''Private Sub testUpdate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    ''    Dim strConnection As String =
    ''           "Data Source=brickyard;Initial Catalog=SEPF-NET;" _
    ''           & "Integrated Security=True;MultipleActiveResultSets=True"

    ''    Dim cn As New SqlConnection("Data Source=brickyard;Initial Catalog=SEPF-NET;Integrated Security=True")
    ''    Dim da As New SqlDataAdapter
    ''    Dim ds As New DataSet
    ''    Try
    ''        cn.Open()
    ''        da.SelectCommand = New SqlCommand("SELECT * FROM [tbltest1]", cn)
    ''        da.Fill(ds)
    ''        Dim dt As New DataTable
    ''        dt = ds.Tables("tbltest1")
    ''        Dim dr As DataRow = dt.NewRow()
    ''        dr.Item("ID") = -2
    ''        dr.Item("Name") = "Test"

    ''        dt.Rows.Add(dr)
    ''        da.Update(ds)
    ''        MsgBox("Record Successfully Inserted")
    ''    Catch ex As Exception
    ''        MsgBox(ex.Message)
    ''    End Try
    ''End Sub

    'Private Sub Button1_Click_1(sender As Object, e As EventArgs)

    '    Dim strConnection As String =
    '          "Data Source=brickyard;Initial Catalog=SEPF-NET;" _
    '          & "Integrated Security=True;MultipleActiveResultSets=True"

    '    Dim cn As New SqlConnection("Data Source=brickyard;Initial Catalog=SEPF-NET;Integrated Security=True")
    '    Dim da As New SqlDataAdapter
    '    Dim ds As New DataSet
    '    Try
    '        cn.Open()
    '        da.SelectCommand = New SqlCommand("SELECT * FROM tbltest1", cn)


    '        Dim builder As SqlCommandBuilder = New SqlCommandBuilder(da)

    '        builder.GetInsertCommand()


    '        da.Fill(ds, "tblTest1")


    '        'da.MissingSchemaAction = MissingSchemaAction.AddWithKey

    '        'da.FillSchema(ds, SchemaType.Mapped, "tbltest1")

    '        da.TableMappings.Add("Table", "tbltest1")


    '        Dim dt As New DataTable
    '        dt = ds.Tables("tbltest1")
    '        Dim dr As DataRow = dt.NewRow()
    '        dr.Item("ID") = -4
    '        dr.Item("Name") = "Test4"

    '        dt.Rows.Add(dr)
    '        da.Update(ds)
    '        MsgBox("Record Successfully Inserted")
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    'End Sub

    Private Sub ExportToExcel(sSql As String, sLoc As String)

        'Initialize the objects before use
        Dim dataAdapter As New SqlClient.SqlDataAdapter()
        Dim dataSet As New DataSet
        Dim command As New SqlClient.SqlCommand
        Dim datatableMain As New System.Data.DataTable()
        Dim fileName As String
        Dim finalPath As String


        'Dim connection As New SqlClient.SqlConnection

        ''Assign your connection string to connection object
        'connection.ConnectionString = "Data Source=.;_
        'Initial Catalog=pubs;Integrated Security=True"
        command.Connection = cn
        command.CommandType = CommandType.Text
        'You can use any command select
        command.CommandText = sSql
        dataAdapter.SelectCommand = command

        'Dim f As FolderBrowserDialog = New FolderBrowserDialog

        Label1.Text = "Exporting " & sLoc
        Label1.Refresh()

        Try

            'If f.ShowDialog() = DialogResult.OK Then

            'This section help you if your language is not English.
            System.Threading.Thread.CurrentThread.CurrentCulture =
                System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
            Dim oExcel As Excel.Application
            Dim oBook As Excel.Workbook
            Dim oSheet As Excel.Worksheet
            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add(Type.Missing)
            oSheet = oBook.Worksheets(1)

            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            'Fill data to datatable
            'connection.Open()
            dataAdapter.Fill(datatableMain)
            'connection.Close()


            'Export the Columns to excel file
            For Each dc In datatableMain.Columns
                colIndex = colIndex + 1
                oSheet.Cells(1, colIndex) = dc.ColumnName
            Next

            'Export the rows to excel file
            For Each dr In datatableMain.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In datatableMain.Columns
                    colIndex = colIndex + 1
                    oSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next


            'Check if the folder exists
            If Not Directory.Exists(txtPath.Text) Then
                Directory.CreateDirectory(txtPath.Text)
            Else

            End If

            'Set final path
            fileName = sLoc + ".xls"

            'fileName = sLoc + ".xlsx"
            'Dim finalPath = f.SelectedPath + fileName

            finalPath = txtPath.Text + fileName

            'txtPath.Text = finalPath
            oSheet.Columns.AutoFit()
            'Save file in final path

            'suppress any message
            oExcel.DisplayAlerts = False

            'oBook.SaveAs(finalPath, XlFileFormat.xlWorkbookNormal, Type.Missing,
            'Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
            'Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)

            oBook.SaveAs(finalPath, XlFileFormat.xlWorkbookNormal, Type.Missing,
                Type.Missing, Type.Missing, True, XlSaveAsAccessMode.xlExclusive,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)

            'Release the objects
            ReleaseObject(oSheet)
            oBook.Close(False, Type.Missing, Type.Missing)
            ReleaseObject(oBook)
            oExcel.Quit()
            ReleaseObject(oExcel)
            'Some time Office application does not quit after automation: 
            'so i am calling GC.Collect method.
            GC.Collect()

            'MessageBox.Show("Export done successfully!")

            Label1.Text = "Exporting Success " & sLoc
            Label1.Refresh()

            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK)
            Label1.Text = "Exporting Fail " & sLoc
            Label1.Refresh()
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdExport.Click

        'sPath = "\\Tile\SEPF\Actuarial_Files\Cheiron_" & CStr(DatePart(DateInterval.Year, Date.Now) - 1) & "\"

        'sPath = "\\Tile\SEPF\Actuarial_Files\Sepf-net\Cheiron_" & CStr(DatePart(DateInterval.Year, Date.Now) - 1) & "\"


        'ExportToExcel("Select * from " & chkPen.Tag)



        Dim ctrl As Control
        Dim ctrlType As System.Type


        For Each ctrl In Me.Controls

            ctrlType = ctrl.GetType

            'MsgBox(ctrl.Name)
            'MsgBox(ctrlType.ToString)

            If ctrlType.ToString = "System.Windows.Forms.CheckBox" Then

                Dim chkbx As System.Windows.Forms.CheckBox = CType(ctrl, System.Windows.Forms.CheckBox)

                If chkbx.checked = True Then


                    If ctrl.Tag <> "" Then
                        ExportToExcel("Select * from " & ctrl.Tag, ctrl.Text)
                    End If

                End If

            End If


            'If (ctrl.GetType() Is GetType(CheckBox)) Then
            '    Dim chkbx As CheckBox = CType(ctrl, CheckBox)
            '    chkbx.Checked = True
            '    MsgBox(ctrl.Name)
            'Else
            '    MsgBox(ctrl.Name)
            '    MsgBox(ctrl.GetType().ToString)
            'End If




            'If (ctrl.GetType() Is GetType(ComboBox)) Then
            '    Dim cbobx As ComboBox = CType(ctrl, ComboBox)
            '    cbobx.SelectedIndex = -1
            'End If
            'If (ctrl.GetType() Is GetType(DateTimePicker)) Then
            '    Dim dtp As DateTimePicker = CType(ctrl, DateTimePicker)
            '    dtp.Value = Now()
            'End If

            'If Recurse Then
            '    If (ctrl.GetType() Is GetType(Panel)) Then
            '        Dim pnl As Panel = CType(ctrl, Panel)
            '        ClearAllControls(pnl, Recurse)
            '    End If
            '    If ctrl.GetType() Is GetType(GroupBox) Then
            '        Dim grbx As GroupBox = CType(ctrl, GroupBox)
            '        ClearAllControls(grbx, Recurse)
            '    End If
            'End If
        Next
    End Sub


    Private Sub ExportToExcelOrg(sSql As String)
        'Initialize the objects before use
        Dim dataAdapter As New SqlClient.SqlDataAdapter()
        Dim dataSet As New DataSet
        Dim command As New SqlClient.SqlCommand
        Dim datatableMain As New System.Data.DataTable()
        'Dim connection As New SqlClient.SqlConnection

        ''Assign your connection string to connection object
        'connection.ConnectionString = "Data Source=.;_
        'Initial Catalog=pubs;Integrated Security=True"
        command.Connection = cn
        command.CommandType = CommandType.Text
        'You can use any command select
        command.CommandText = sSql
        dataAdapter.SelectCommand = command


        Dim f As FolderBrowserDialog = New FolderBrowserDialog
        Try

            If f.ShowDialog() = DialogResult.OK Then

                'This section help you if your language is not English.
                System.Threading.Thread.CurrentThread.CurrentCulture =
                System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oSheet As Excel.Worksheet
                oExcel = CreateObject("Excel.Application")
                oBook = oExcel.Workbooks.Add(Type.Missing)
                oSheet = oBook.Worksheets(1)

                Dim dc As System.Data.DataColumn
                Dim dr As System.Data.DataRow
                Dim colIndex As Integer = 0
                Dim rowIndex As Integer = 0

                'Fill data to datatable
                'connection.Open()
                dataAdapter.Fill(datatableMain)
                'connection.Close()


                'Export the Columns to excel file
                For Each dc In datatableMain.Columns
                    colIndex = colIndex + 1
                    oSheet.Cells(1, colIndex) = dc.ColumnName
                Next

                'Export the rows to excel file
                For Each dr In datatableMain.Rows
                    rowIndex = rowIndex + 1
                    colIndex = 0
                    For Each dc In datatableMain.Columns
                        colIndex = colIndex + 1
                        oSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                    Next
                Next

                'Set final path
                Dim fileName As String = "\Pensioners" + ".xls"
                'Dim finalPath = f.SelectedPath + fileName

                Dim finalPath = sPath + fileName

                txtPath.Text = finalPath
                oSheet.Columns.AutoFit()
                'Save file in final path
                oBook.SaveAs(finalPath, XlFileFormat.xlWorkbookNormal, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)

                'Release the objects
                ReleaseObject(oSheet)
                oBook.Close(False, Type.Missing, Type.Missing)
                ReleaseObject(oBook)
                oExcel.Quit()
                ReleaseObject(oExcel)
                'Some time Office application does not quit after automation: 
                'so i am calling GC.Collect method.
                GC.Collect()

                MessageBox.Show("Export done successfully!")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowse.Click
        Dim f As FolderBrowserDialog = New FolderBrowserDialog

        f.SelectedPath = sPath

        If f.ShowDialog() = DialogResult.OK Then
            txtPath.Text = f.SelectedPath
        End If
    End Sub

    Private Sub chkInActVest_CheckedChanged(sender As Object, e As EventArgs) Handles chkInActVest.CheckedChanged

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        Me.Close()
    End Sub
End Class
