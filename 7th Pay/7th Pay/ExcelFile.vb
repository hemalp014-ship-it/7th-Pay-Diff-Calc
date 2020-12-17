Imports Excel = Microsoft.Office.Interop.Excel
Imports _7th_Pay.Form1
Imports System.Drawing.Drawing2D

Public Class ExcelFile

    Dim Quited As Boolean = False
    Dim xlapp As Excel.Application
    Dim xlworkbook As Excel.Workbook
    Dim xlworksheet As Excel.Worksheet

    Private Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr,
              ByRef lpdwProcessId As Integer) As Integer

    Dim filepath As String = ""

    Sub New(ByVal file As String)
        InitializeComponent()

        xlapp = New Excel.Application
        filepath = file

        If xlapp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!") : Return
        Else
            Quited = False
        End If

        Application.DoEvents()
    End Sub

    Private Sub ShowProcess(ByVal text As String)
        If Me.CrtResult IsNot Nothing Then Me.CrtResult.Text = text : Form1.CurrentStatus.Text = text
        Application.DoEvents()
    End Sub

    Shadows Sub Load()
        If Not Quited Then

            Application.DoEvents()

            xlworkbook = xlapp.Workbooks.Open(filepath)
            xlworksheet = xlworkbook.Sheets("7 PAY-2016")
            If xlworksheet Is Nothing Then
                MessageBox.Show("Failed to Open required worksheet!!") : Return
            End If

            For i = 0 To 19
                Form1.PayMaster.Add(New GradePayData)

                Select Case i
                    Case 0, 1, 2
                        Form1.PayMaster(i).SetMin(Form1.mins(0))
                        Form1.PayMaster(i).SetMax(Form1.maxes(0))
                    Case 3, 4, 5, 6, 7
                        Form1.PayMaster(i).SetMin(Form1.mins(1))
                        Form1.PayMaster(i).SetMax(Form1.maxes(1))
                    Case 8, 9, 10, 11
                        Form1.PayMaster(i).SetMin(Form1.mins(2))
                        Form1.PayMaster(i).SetMax(Form1.maxes(2))
                    Case 12, 13, 14
                        Form1.PayMaster(i).SetMin(Form1.mins(3))
                        Form1.PayMaster(i).SetMax(Form1.maxes(3))
                    Case 15, 16, 17
                        Form1.PayMaster(i).SetMin(Form1.mins(4))
                        Form1.PayMaster(i).SetMax(Form1.maxes(4))
                    Case 18
                        Form1.PayMaster(i).SetMin(Form1.mins(5))
                        Form1.PayMaster(i).SetMax(Form1.maxes(5))
                    Case 19
                        Form1.PayMaster(i).SetMin(Form1.mins(6))
                        Form1.PayMaster(i).SetMax(Form1.maxes(6))
                End Select
            Next
            For i = 5 To 47
                Dim chartRange As Excel.Range = xlworksheet.Range("m" & i, "af" & i)
                Dim cnt = 0

                If i = 5 Then
                    For Each item As Excel.Range In chartRange
                        If Not item.Text.Equals("") Then
                            Form1.PayMaster(cnt).SetGradePay(item.Text)
                            Debug.WriteLine(item.Text)
                        End If
                        cnt += 1
                    Next
                    i = 7
                Else
                    For Each item As Excel.Range In chartRange
                        If Not item.Text.Equals("") Then
                            Form1.PayMaster(cnt).AddData(item.Text)
                        End If
                        cnt += 1
                    Next
                End If
            Next
            Quit()
        End If

    End Sub

    Sub BorderCell(Range As Excel.Range)
        Range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin,
                                 Excel.XlColorIndex.xlColorIndexAutomatic,
                                 Excel.XlColorIndex.xlColorIndexAutomatic)
    End Sub

    Sub CenterCell(range As Excel.Range, Optional VAlign As Short = Excel.XlVAlign.xlVAlignCenter,
                     Optional HAlign As Short = Excel.XlHAlign.xlHAlignCenter)
        range.VerticalAlignment = 2
        range.HorizontalAlignment = 3
    End Sub

    Sub Save()
        Me.Show()

        If Not Quited Then
            xlworkbook = xlapp.Workbooks.Add()
            xlworksheet = xlworkbook.Sheets("Sheet1")
            xlworksheet.Name = "7 PAY-2016"

            ShowProcess("Applying Titles")
            SetCredit() 'APPLY CREDIT
            SetTitle() 'APPLY TITLE
            TTable()  'APPLY TOP DETAILS
            BTable() 'APPLY Bottom DETAILS

            ShowProcess("Applying Page Setup")
            With xlworksheet.PageSetup
                .PaperSize = Excel.XlPaperSize.xlPaperA4
                .LeftMargin = 0.45 * 72
                .RightMargin = 0.3 * 72
            End With

            xlworkbook.SaveAs(filepath)
            xlworkbook.Close()
            Quit()

            If MsgBox("Excel file created!" + Chr(10) + "Do you want to open the file now?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Process.Start(filepath)
            End If
        End If
    End Sub

    Private Sub SetTitle()
        Dim chartrange = xlworksheet.Range("A2", "J2")
        chartrange.Merge()
        chartrange.Cells(1, 1) = "સાતમા પગારપંચ પ્રમાણે પગારની ગણતરી નો કોઠો"

        CenterCell(chartrange, Excel.XlVAlign.xlVAlignBottom)
    End Sub

    ''' <summary>
    ''' To Show Credit
    ''' </summary>
    Private Sub SetCredit()
        Dim chartrange As Excel.Range = xlworksheet.Range("A1", "J1")
        chartrange.Merge()
        chartrange.Cells(1, 1) = "CREATED BY :- G.B.PATEL(NADIAD)(MO.NO.9033424445) & HEMAL G. PATEL(NADIAD)"

        CenterCell(chartrange)
    End Sub

    ''' <summary>
    ''' To Set The Top Table
    ''' </summary>
    Private Sub TTable()

        ShowProcess("Applying Details")
        xlworksheet.Range("A3").Cells(1, 1) = "OFFICE :-"
        Dim chartrange = xlworksheet.Range("B3", "J3")
        chartrange.Merge()
        chartrange(1, 1) = Form1.TxtOffice.Text

        'LEFT DETAILS
        xlworksheet.Range("A4").Cells(1, 1) = "PLACE :-"
        chartrange = xlworksheet.Range("B4", "E4")
        chartrange.Merge()
        chartrange(1, 1) = Form1.TxtPlace.Text

        xlworksheet.Range("A5").Cells(1, 1) = "NAME :-"
        chartrange = xlworksheet.Range("B5", "E5")
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.TxtName.Text

        xlworksheet.Range("A6").Cells(1, 1) = "DESI :-"
        chartrange = xlworksheet.Range("B6", "E6")
        chartrange.Merge()
        chartrange(1, 1) = Form1.TxtDesignation.Text

        xlworksheet.Range("A3", "j6").RowHeight = 16

        xlworksheet.Range("A7").Cells(1, 1) = 1
        chartrange = xlworksheet.Range("B7", "D7")
        chartrange.Merge()
        chartrange(1, 1) = "PAY AS ON 01-07-2016"
        xlworksheet.Range("E7").Cells(1, 1) = Form1.TxtPay7.Text

        xlworksheet.Range("A8").Cells(1, 1) = 2
        chartrange = xlworksheet.Range("B8", "D8")
        chartrange.Merge()
        chartrange(1, 1) = "PAY AS ON 01-01-2016"
        xlworksheet.Range("E8").Cells(1, 1) = Form1.TxtPay6.Text

        xlworksheet.Range("A9").Cells(1, 1) = 3
        chartrange = xlworksheet.Range("B9", "D9")
        chartrange.Merge()
        chartrange(1, 1) = "GRADE PAY"
        xlworksheet.Range("E9").Cells(1, 1) = Form1.CboGradePay.Text

        xlworksheet.Range("A10").Cells(1, 1) = 4
        chartrange = xlworksheet.Range("B10", "D10")
        chartrange.Merge()
        chartrange(1, 1) = "NEW H.R.A %"
        xlworksheet.Range("E9").Cells(1, 1) = Form1.TxtHRA7.Text

        'RIGHT DETAILS
        chartrange = xlworksheet.Range("F4", "F5")
        chartrange.Merge()
        chartrange(1, 1) = "D.A %"

        CenterCell(chartrange)

        chartrange = xlworksheet.Range("G4", "I4")
        chartrange.Merge()
        chartrange(1, 1) = "01-01-2016 TO"

        chartrange = xlworksheet.Range("G5", "I5")
        chartrange.Merge()
        chartrange(1, 1) = "01-07-2016 TO"

        xlworksheet.Range("J4").Cells(1, 1) = Form1.TxtDANew1.Text & " %"
        xlworksheet.Range("J5").Cells(1, 1) = Form1.TxtDANew2.Text & " %"

        chartrange = xlworksheet.Range("F6", "I6")
        chartrange.VerticalAlignment = 2
        chartrange.HorizontalAlignment = 3
        chartrange.Merge()

        chartrange(1, 1) = "OLD H.R.A %"
        xlworksheet.Range("J6").Cells(1, 1) = Form1.TxtHRA6.Text & " %"

        chartrange = xlworksheet.Range("F7", "I7")
        chartrange.Merge()

        CenterCell(chartrange)

        chartrange(1, 1) = "OLD D.A. %"
        xlworksheet.Range("J7").Cells(1, 1) = Form1.TxtDAOld.Text & " %"

        chartrange = xlworksheet.Range("F8", "I8")
        chartrange.Merge()

        CenterCell(chartrange)

        chartrange(1, 1) = "OTHER"
        xlworksheet.Range("J8").Cells(1, 1) = Form1.TxtOther.Text

        CenterCell(chartrange)

        chartrange = xlworksheet.Range("F9", "I10")
        chartrange.Merge()
        chartrange(1, 1) = "TOTAL PAY (OLD)"

        CenterCell(chartrange)

        chartrange = xlworksheet.Range("J9", "J10")
        chartrange.Merge()
        chartrange(1, 1) = Form1.TxtPayOld.Text

        CenterCell(chartrange)

        ShowProcess("Applying Touch-up")
        For Each CELL As Excel.Range In xlworksheet.Range("A1", "J10")
            BorderCell(CELL)
        Next

        ShowProcess("Applying Headings of Tabel for Monthwise Data")
        'TABLE HEADIND
        chartrange = xlworksheet.Range("A12", "A13")
        chartrange.Cells(1, 1) = "MONTH"
        chartrange.Merge()
        CenterCell(chartrange)

        chartrange = xlworksheet.Range("B12", "E12")
        chartrange.Merge()
        chartrange(1, 1) = "AS PAR 7 PAY"

        CenterCell(chartrange)

        chartrange = xlworksheet.Range("F12", "I12")
        chartrange.Merge()
        chartrange(1, 1) = "AS PAR 6 PAY"

        CenterCell(chartrange)
        chartrange = xlworksheet.Range("J12", "J13")
        chartrange.Cells(1, 1) = "DIFF."

        chartrange.Merge()
        CenterCell(chartrange)


        ShowProcess("Applying Monthwise Data")
        'MONTH
        chartrange = xlworksheet.Range("A13", "A21")
        For I = 2 To 9
            chartrange.Cells(I, 1) = Form1.Months(I - 2).Text
        Next

        With Form1
            '7TH PAY
            SetTable("B"c, { .BPay7, .DA7, .HRA7, .TPay7})

            '6TH PAY
            SetTable("F"c, { .BPay6, .DA6, .HRA6, .TPay6})
        End With


        ShowProcess("Applying Difference")
        'DIFFERENCE
        chartrange = xlworksheet.Range("J14", "J21")
        For I = 1 To 8
            chartrange.Cells(I, 1) = Form1.Diff(I - 1).Text
        Next


        ShowProcess("Applying Total")
        'TOTAL
        xlworksheet.Range("A22").Cells(1, 1) = "TOTAL"

        chartrange = xlworksheet.Range("b22", "J22")
        chartrange.Cells(1, 1) = Form1.Total(0).Text
        chartrange.Cells(1, 2) = Form1.Total(1).Text
        chartrange.Cells(1, 3) = Form1.Total(2).Text
        chartrange.Cells(1, 4) = Form1.Total(3).Text
        chartrange.Cells(1, 5) = Form1.Total(4).Text
        chartrange.Cells(1, 6) = Form1.Total(6).Text
        chartrange.Cells(1, 7) = Form1.Total(7).Text
        chartrange.Cells(1, 8) = Form1.Total(8).Text
        chartrange.Cells(1, 9) = Form1.Total(9).Text


        ShowProcess("Applying Touch-up")
        For Each CELL As Excel.Range In xlworksheet.Range("A12", "J22")
            BorderCell(CELL)
        Next
    End Sub

    ''' <summary>
    ''' To Set The Bottom Table
    ''' </summary>
    Private Sub BTable()

        ShowProcess("Applying Difference Tabel")
        LLTable() 'APPLY LEFT PART

        ShowProcess("Applying Installment Data")
        LRTable() 'APPLY RIGHT PART
    End Sub

    ''' <summary>
    ''' To Set The Lower Left Table
    ''' </summary>
    Private Sub LLTable()
        Dim I = 24
        Dim chartrange = xlworksheet.Range("B" & I, "G" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "MONTH AUGUST-2016 PAY"

        CenterCell(chartrange)

        chartrange = xlworksheet.Range("B25", "C25")
        chartrange.Merge()

        I = 25
        chartrange = xlworksheet.Range("f" & I, "g" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "PAY 7th"

        CenterCell(chartrange)

        I = 25
        chartrange = xlworksheet.Range("D" & I, "E" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "PAY 6th"

        CenterCell(chartrange)

        I = 26
        chartrange = xlworksheet.Range("B" & I, "C" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "BASIC PAY"
        CenterCell(chartrange)

        chartrange = xlworksheet.Range("D" & I, "E" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.DA6(6).Text

        chartrange = xlworksheet.Range("F" & I, "G" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.DA7(6).Text

        I = 27
        chartrange = xlworksheet.Range("B" & I, "C" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "D.A."

        CenterCell(chartrange)

        chartrange = xlworksheet.Range("D" & I, "E" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.DA6(6).Text

        chartrange = xlworksheet.Range("F" & I, "G" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.DA7(6).Text

        I = 28
        chartrange = xlworksheet.Range("B" & I, "C" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "H.R.A."

        CenterCell(chartrange)
        chartrange = xlworksheet.Range("D" & I, "E" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.HRA6(6).Text

        chartrange = xlworksheet.Range("F" & I, "G" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.HRA7(6).Text

        I = 29
        chartrange = xlworksheet.Range("B" & I, "C" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "TOTAL PAY"

        CenterCell(chartrange)

        chartrange = xlworksheet.Range("D" & I, "E" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.TPay6(6).Text

        chartrange = xlworksheet.Range("F" & I, "G" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = Form1.TPay7(6).Text


        ShowProcess("Applying Touch-up")
        For Each CELL As Excel.Range In xlworksheet.Range("B24", "G29")
            BorderCell(CELL)
        Next

    End Sub
    ''' <summary>
    ''' To Set The Lower Right Table
    ''' </summary>
    Private Sub LRTable()
        Dim I = 24
        Dim chartrange = xlworksheet.Range("I" & I, "J" & I)
        chartrange.Merge()
        chartrange.Cells(1, 1) = "PAYABLE DIFFERENCE"

        I = 25
        chartrange = xlworksheet.Range("I" & I)
        chartrange.Cells(1, 1) = CInt(Form1.Diff(0).Text) + CInt(Form1.Diff(1).Text)

        chartrange = xlworksheet.Range("J" & I)
        chartrange.Cells(1, 1) = "1st"

        I = 26
        chartrange = xlworksheet.Range("I" & I)
        chartrange.Cells(1, 1) = CInt(Form1.Diff(2).Text) + CInt(Form1.Diff(3).Text)

        chartrange = xlworksheet.Range("J" & I)
        chartrange.Cells(1, 1) = "2nd"

        I = 27
        chartrange = xlworksheet.Range("I" & I)
        chartrange.Cells(1, 1) = CInt(Form1.Diff(4).Text) + CInt(Form1.Diff(5).Text) + CInt(Form1.Diff(6).Text)

        chartrange = xlworksheet.Range("J" & I)
        chartrange.Cells(1, 1) = "3rd"

        ShowProcess("Applying Touch-up")
        For Each CELL As Excel.Range In xlworksheet.Range("I24", "J27")
            BorderCell(CELL)
        Next
    End Sub

    'Private Function InceraseChar(RefChar As Char) As Char
    '    Return (Chr(Asc(RefChar) + 1))
    'End Function
    Private Sub InceraseChar(ByRef RefChar As Char)
        RefChar = (Chr(Asc(RefChar) + 1))
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="StartPoint">Would we less or equal to char W</param>
    ''' <param name="Labels">A squence of Labels(BPay, DA, HRA, TPay</param>
    Private Sub SetTable(StartPoint As Char, Labels()() As Label)
        Dim chartrange = xlworksheet.Range("B13")
        chartrange.Cells(1, 1) = "BASIC PAY"
        chartrange.ColumnWidth = 9.3

        CenterCell(chartrange)

        chartrange = xlworksheet.Range(StartPoint + "14", StartPoint + "B21")
        For I = 1 To 8
            chartrange.Cells(I, 1) = Labels(0)(I - 1).Text
        Next

        InceraseChar(StartPoint)

        chartrange = xlworksheet.Range(StartPoint + "13")
        chartrange.Cells(1, 1) = "D.A."
        CenterCell(chartrange)

        chartrange = xlworksheet.Range(StartPoint + "14", StartPoint + "21")
        For I = 1 To 8
            chartrange.Cells(I, 1) = Labels(1)(I - 1).Text
        Next

        InceraseChar(StartPoint)

        chartrange = xlworksheet.Range(StartPoint + "13")
        chartrange.Cells(1, 1) = "H.R.A."
        CenterCell(chartrange)
        chartrange = xlworksheet.Range(StartPoint + "14", StartPoint + "21")
        For I = 1 To 8
            chartrange.Cells(I, 1) = Labels(2)(I - 1).Text
        Next

        InceraseChar(StartPoint)

        chartrange = xlworksheet.Range(StartPoint + "13")
        chartrange.ColumnWidth = 9.71
        chartrange.Cells(1, 1) = "TOTAL PAY"
        CenterCell(chartrange)

        chartrange = xlworksheet.Range(StartPoint + "14", StartPoint + "21")
        For I = 1 To 8
            chartrange.Cells(I, 1) = Labels(3)(I - 1).Text
        Next
    End Sub

    Friend Sub Quit()
        'If Not xlworkbook Is Nothing Then xlworkbook.Saved = True

        ReleaseObject(xlworksheet)
        ReleaseObject(xlworkbook)

        Dim xlHWND As Integer = xlapp.Hwnd

        Dim ProcIdXL As Integer = 0

        GetWindowThreadProcessId(xlHWND, ProcIdXL)

        Dim xproc As Process = Process.GetProcessById(ProcIdXL)

        xlapp.Quit()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)

        xlapp = Nothing

        If Not xproc.HasExited Then xproc.Kill()
        Quited = True

        Form1.CurrentStatus.Text = "Ready"
        Me.Close()
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing

        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class