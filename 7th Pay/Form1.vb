Imports System
Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Drawing.Drawing2D

Public Class Form1

#Region "Public"

    Public mins() As Integer = {4440, 5200, 9300, 15600, 37400, 67000, 80000}
    Public maxes() As Integer = {7440, 20200, 34800, 39100, 67000, 79000, 80000}
    Public PayMaster As New List(Of GradePayData)

    Dim NumberTranslator As New સંખ્યાનુવાદક({"૦", "૧", "૨", "૩", "૪", "૫", "૬", "૭", "૮", "૯"})

#End Region

#Region "Friend"
    Friend Total As Label()

    Friend BPay7 As Label() : Friend BPay6 As Label()
    Friend DA7 As Label() : Friend DA6 As Label()
    Friend HRA7 As Label() : Friend HRA6 As Label()
    Friend TPay7 As Label() : Friend TPay6 As Label()

    Friend Months As Label()

    Friend Diff As Label()
    Friend GradePays As Label()
#End Region

    Dim Labels As New List(Of Label)

    Dim DAOld As Integer = 0
    Dim DANew1 As Integer = 0
    Dim DANew2 As Integer = 0

    Dim HRAOld As Integer = 0
    Dim HRANew As Integer = 0

    Dim Other As Integer = 0

    Dim Pay7th As Integer
    Dim Pay6th As Integer

    Dim GradePay As Integer

    Dim InputList As Type() = {GetType(TextBox), GetType(ComboBox)}
    Dim Inputs As New List(Of Object)

    Dim filepath As String

    Dim cbochanged As Boolean = False

    Enum Setting
        Office
        Place
        Name
        Designation
    End Enum
    Class GradePayData
        Friend min As Integer
        Friend max As Integer
        Friend GradePay As Integer
        Friend Data As New List(Of Integer)

        Sub SetMin(ByVal value As Integer)
            Me.min = value
        End Sub
        Sub SetMax(ByVal value As Integer)
            Me.max = value
        End Sub
        Sub SetGradePay(ByVal value As Integer)
            Me.GradePay = value
        End Sub
        Sub AddData(ByVal value As Integer)
            Data.Add(value)
        End Sub
        Sub Load(ByVal str As String)
            Dim items() As String = str.Split(",")

            Dim item = items(0).Split("-")
            min = item(0) : max = item(1)

            GradePay = items(1)
            For Each item1 In items(2).Split(":")
                Data.Add(item1)
            Next
        End Sub
        Overloads Function ToString()
            Dim strdata As String = Nothing

            strdata += min & "-" & max & "," & GradePay & ","
            For Each item In Data
                strdata += item & ":"
            Next
            Return strdata.Substring(0, strdata.Length - 1)
        End Function
    End Class
    Function Encrypt(ByVal data As String) As String
        Dim bytes() As Byte = Encoding.ASCII.GetBytes(data)

        For i = 0 To bytes.Length - 1
            If bytes(i) >= Byte.MaxValue - (60 - 1) Then
                bytes(i) = Byte.MaxValue - bytes(i)
            Else
                bytes(i) += 60
            End If
        Next

        Return Encoding.ASCII.GetString(bytes)
    End Function
    ''' <summary>
    ''' 60-bit decryption
    ''' </summary>
    ''' <param name="data"></param>
    ''' <returns></returns>
    Function Decrypt(ByVal data As String) As String
        Dim bytes() As Byte = Encoding.ASCII.GetBytes(data)

        For i = 0 To bytes.Length - 1
            If bytes(i) <= 60 - 1 Then
                bytes(i) = Byte.MaxValue - bytes(i)
            Else
                bytes(i) -= 60
            End If
        Next

        Return Encoding.ASCII.GetString(bytes)
    End Function
    Private Function FindLabels(ByVal Nums() As Integer) As Label()
        Dim labs As New List(Of Label)
        For Each item In Nums
            labs.Add(FindLabel("Label" & item))
        Next
        Return labs.ToArray
    End Function
    Private Function FindLabel(ByVal Name As String) As Label
        Dim labs As New Label With {.Name = "missing"}

        For Each item1 In Labels
            If TypeOf item1 Is Label Then
                Dim lab As Label = item1
                If lab.Name = Name Then labs = lab
            End If
        Next
        If labs.Name = "missing" Then
            'Debug.WriteLine("not found : " & Name)
        End If
        Return labs
    End Function
    Private Function JoinArray(ByVal arr1 As Integer(), ByVal arr2 As Integer()) As Integer()
        Dim arr As New List(Of Integer)
        For Each item In arr1
            arr.Add(item)
        Next
        For Each item In arr2
            arr.Add(item)
        Next
        Return arr.ToArray
    End Function
    Function RoundIt(ByVal value As Single) As Integer
        Dim val As Integer = 0

        If CInt(value / 10) * 10 - CInt(value) / 10 * 10 = 0 Then
            val = CInt(value / 10) * 10
        ElseIf CInt(value / 10) - CInt(value) / 10 > 1 Then
            val = CInt(value / 10) + 10
        Else
            val = CInt(value / 10) * 10 + 10
        End If
        Return val
    End Function
    Private Sub FindPanels(ByRef Item As Control)
        For Each _item As Control In Item.Controls
            If TypeOf _item Is TableLayoutPanel Then
                FindElement(_item)
                FindPanels(_item)
            End If
        Next
    End Sub
    Sub FindElement(ByRef Item As Control)
        For Each _item In Item.Controls
            If TypeOf _item Is Label Then
                Labels.Add(_item)
            ElseIf InputList.Contains(_item.GetType) Then
                Inputs.Add(_item)
            End If
        Next
    End Sub
    Sub SetGradient(ctrl As Object, Optional opacity As Single = 1)
        With ctrl
            .ForeColor = Color.Black
            Dim img As Image = New Bitmap(CInt(.Width), .Height)
            Dim g = Graphics.FromImage(img)
            Dim path As New GraphicsPath()
            Dim Brush = New LinearGradientBrush(CType(.Bounds, Rectangle), Color.White, Color.Black, 90)

            Dim y As New ColorBlend(3) With {
                        .Colors = {Color.FromArgb(255 * opacity, Color.White),
                                Color.FromArgb(255 * opacity, 100, 100, 100),
                                Color.FromArgb(255 * opacity, Color.White)},
                        .Positions = {0, 0.75, 1}
                    }

            Brush.InterpolationColors = y
            g.FillRectangle(Brush, .Bounds)

            .BackgroundImage = img
            .Invalidate()
        End With
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim foundIt As Boolean = False
        Dim PastItems As New List(Of Control)
        Dim cnt As Integer = 0

        SetGradient(Menu)
        SetGradient(ExportMenuItem)
        ExitMenuItem.BackgroundImage = ExportMenuItem.BackgroundImage
        SetGradient(Me)

        'SetBody()

        'For Each item In Me.Controls
        '    If TypeOf item Is Panel Then
        '        FindPanels(item)
        '    End If
        '    Application.DoEvents()
        'Next

        FindElement(Body)

        Debug.WriteLine("Labels found : " & Labels.Count)
        Debug.WriteLine("Inputs found : " & Inputs.Count)

        For Each item In Inputs
            Select Case item.GetType
                Case GetType(TextBox)
                    AddHandler CType(item, TextBox).TextChanged, AddressOf Txt_TextChanged
                Case GetType(ComboBox)

                    If item.DropDownStyle = ComboBoxStyle.DropDownList Then
                        AddHandler CType(item, ComboBox).SelectedIndexChanged, AddressOf Cbo_SelectedIndexChanged
                    Else
                        AddHandler CType(item, ComboBox).TextChanged, AddressOf Txt_TextChanged
                    End If
            End Select
            AddHandler CType(item, Control).KeyDown, AddressOf _KeyDown
            Application.DoEvents()
        Next

        Months = FindLabels(GetSeries(26, 1, 33))

        Total = FindLabels({71, 72, 73, 74, 78, 128, 77, 76, 75, 123})

        BPay7 = FindLabels(GetSeries(39, 4, 67))
        BPay6 = FindLabels(JoinArray({113}, GetSeries(106, -4, 82)))

        DA7 = FindLabels(GetSeries(40, 4, 68))
        DA6 = FindLabels(GetSeries(109, -4, 81))

        HRA7 = FindLabels(GetSeries(41, 4, 69))
        HRA6 = FindLabels(GetSeries(108, -4, 80))

        TPay7 = FindLabels(GetSeries(42, 4, 70))
        TPay6 = FindLabels(GetSeries(107, -4, 79))

        Diff = FindLabels(GetSeries(115, 1, 122))

        GradePays = FindLabels(JoinArray({15, 124, 127}, GetSeries(133, -1, 129)))

        My.Settings.Reload()

        Dim data As String = ""

        If My.Settings.Data.Equals("") Then

fromfile:   If My.Computer.FileSystem.FileExists(Application.StartupPath + "\data.txt") Then
                My.Settings.Data = My.Computer.FileSystem.ReadAllText(Application.StartupPath + "\data.txt")
                My.Settings.Save()
            Else
                MsgBox("Contact us to get the data file!")
            End If
        Else
            If Not My.Computer.FileSystem.FileExists(Application.StartupPath + "\data.txt") Then
                My.Computer.FileSystem.WriteAllText(Application.StartupPath + "\data.txt", My.Settings.Data, False)
            End If
        End If

        data = Decrypt(My.Settings.Data)

        Try
            For Each item In data.Split(";")
                PayMaster.Add(New GradePayData())
                PayMaster.Last.Load(item)
            Next
        Catch ex As Exception
            If TypeOf ex Is InvalidCastException Or TypeOf ex Is FormatException Then
                GoTo fromfile
            Else
                MsgBox(ex.Message & vbCrLf & ex.InnerException.ToString)
            End If
        End Try

        If Body.Top - Body.Height > 0 Then
            Body.AutoScrollMargin = New Size(Body.AutoScrollMargin.Width, Body.Top - Body.Height)
        End If
    End Sub
    Private Sub AddItem(ByVal data() As Integer)
        For Each item In data
            If Not CboGradePay.Items.Contains(item) Then
                CboGradePay.Items.Add(item)
            End If
        Next
    End Sub
    Private Function GetSeries(ByVal A As Integer, ByVal d As Integer, ByVal last As Integer) As Integer()
        Dim int As New List(Of Integer)
        For i = A To last Step d
            int.Add(i)
        Next
        Return int.ToArray()
    End Function
    Private Sub PlaceInMiddle(ByVal ctrl As Control, parent As Control)
        If ctrl.Width > parent.Width Then
        Else
            ctrl.Left = parent.Width / 2 - ctrl.Width / 2
        End If
        If ctrl.Height > parent.Height Then
        Else
            ctrl.Top = parent.Height / 2 - ctrl.Height / 2
        End If
    End Sub
    Private Sub Form1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'Dim h As Integer = Me.ClientSize.Height - MenuStrip1.Height - 10
        'If Panel1.Height > h Or Panel1.Height <= 0 Then Panel1.Height = h

        'Dim w As Integer = Me.ClientSize.Width - 10
        'If Panel1.Width > w Or Panel1.Width <= 0 Then Panel1.Width = w

        PlaceInMiddle(Format, Body)
        Menu.BringToFront()
    End Sub
    Private Sub BtnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SaveData()

        GetAndSet(TxtDANew2, DANew2)
        GetAndSet(TxtDANew1, DANew1)
        GetAndSet(TxtDAOld, DAOld)
        GetAndSet(TxtHRA7, HRANew)
        GetAndSet(TxtHRA6, HRAOld)
        GetAndSet(TxtOther, Other)
        GetAndSet(TxtPay6, Pay6th)
        GetAndSet(TxtPay7, Pay7th)

        Integer.TryParse(CboGradePay.Text, GradePay)

        Dim Pagar As Integer = (GradePay + Pay6th) * 2.57

        BPay7(0).Text = Pay7th
        BPay7(1).Text = Pay7th
        BPay7(2).Text = Pay7th
        BPay7(3).Text = Pay7th
        BPay7(4).Text = Pay7th
        BPay7(5).Text = Pay7th

        For Each item In PayMaster
            If item.GradePay = Me.GradePay And Pay6th > item.min And Pay6th < item.max Then
                For Each num In item.Data
                    If num > Pay7th Then Pagar = num : Exit For
                Next
            End If
            Application.DoEvents()
        Next

        BPay7(6).Text = Pagar
        BPay7(7).Text = Pagar

        For i = 0 To 5
            BPay6(i).Text = Pay6th
            GradePays(i).Text = GradePay
            Application.DoEvents()
        Next

        TxtPayOld.Text = Pay6th + GradePay

        BPay6(6).Text = Pay6th + RoundIt((Pay6th + GradePay) * 0.03)
        GradePays(6).Text = GradePay

        For i = 0 To 5
            DA7(i).Text = CInt(Integer.Parse(BPay7(i).Text) * DANew1 / 100)
            Application.DoEvents()
        Next
        For i = 6 To 7
            DA7(i).Text = CInt(Integer.Parse(BPay7(i).Text) * DANew2 / 100)
            Application.DoEvents()
        Next
        For i = 0 To 7
            DA6(i).Text = CInt((Integer.Parse(BPay6(i).Text) + GradePays(i).Text) * DAOld / 100)
            Application.DoEvents()
        Next
        For i = 0 To 7
            HRA6(i).Text = CInt((Integer.Parse(BPay6(i).Text) + GradePays(i).Text) * HRAOld / 100)
            Application.DoEvents()
        Next
        For i = 0 To 7
            HRA7(i).Text = CInt((BPay7(i).Text) * HRANew / 100)
            Application.DoEvents()
        Next
        For i = 0 To 7
            TPay6(i).Text = BPay6(i).Text + Integer.Parse(GradePays(i).Text) + Integer.Parse(DA6(i).Text) + HRA6(i).Text
            Application.DoEvents()
        Next
        For i = 0 To 7
            TPay7(i).Text = BPay7(i).Text + Integer.Parse(DA7(i).Text) + HRA7(i).Text
            Application.DoEvents()
        Next
        For i = 0 To 6
            Diff(i).Text = TPay7(i).Text - TPay6(i).Text
            Application.DoEvents()
        Next
        For i = 0 To 9
            Select Case i
                Case 0
                    Total(i).Text = TotalOf(BPay7)
                Case 1
                    Total(i).Text = TotalOf(DA7)
                Case 2
                    Total(i).Text = TotalOf(HRA7)
                Case 3
                    Total(i).Text = TotalOf(TPay7)
                Case 4
                    Total(i).Text = TotalOf(BPay6)
                Case 5
                    Total(i).Text = TotalOf(GradePays)
                Case 6
                    Total(i).Text = TotalOf(DA6)
                Case 7
                    Total(i).Text = TotalOf(HRA6)
                Case 8
                    Total(i).Text = TotalOf(TPay6)
                Case 9
                    Total(i).Text = TotalOf(Diff)
            End Select
            Application.DoEvents()
        Next

        'Panel1.SuspendLayout()
        'For i = 183 To 0 Step -3
        '    If Panel1.Top <= MenuStrip1.Height Then Exit For
        '    Panel1.Height += 2 * 3
        '    Panel1.Top -= 1 * 3
        '    Application.DoEvents()
        'Next
        'Panel1.ResumeLayout()

        'Panel1.Height = 781
        'PlaceInMiddle(Panel1)
        Body.AutoScroll = True
    End Sub
    Function TotalOf(ByVal t() As Label) As Integer
        Dim sum As Integer = 0
        For Each item In t
            sum = sum + item.Text
        Next
        Return sum
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tb">Text box to operate</param>
    ''' <param name="value">Integer variable to get the data</param>
    ''' <remarks>DON'T CALL THIS SUB ROUTINE FROM FROM tb'S TEXT CHANGED EVENT</remarks>
    Sub GetAndSet(ByVal tb As TextBox, ByRef value As Integer)
        Integer.TryParse(NumberTranslator.GetNumber(tb.Text), value) : tb.Text = value
    End Sub
    Private Sub SaveData()
        Dim data As String = Nothing
        For Each item In PayMaster
            data += item.ToString() & ";"
        Next

        data = data.Substring(0, data.Length - 1)

        My.Settings.Data = Encrypt(data)
        My.Settings.Save()

        My.Computer.FileSystem.WriteAllText(Application.StartupPath + "\data.txt", Encrypt(data), False, Encoding.ASCII)
    End Sub
    Private Sub Cbo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles CboGradePay.SelectedIndexChanged
        cbochanged = True
        SetButton()
    End Sub
    Sub SetButton()
        BtnSubmit.Enabled = cbochanged And Not TxtPay7.Text.Equals("") And Not TxtPay6.Text.Equals("") And Not TxtDAOld.Text.Equals("") And Not TxtDANew1.Text.Equals("") And Not TxtDANew2.Text.Equals("")

        'Dim data As Boolean = True
        'For Each item In Inputs : data = data And Not item.text.Equals("") : Next
        'ExportMenuItem.Enabled = data

        If BtnSubmit.Enabled Then
            BtnSubmit.BackColor = Color.DarkBlue
            BtnSubmit.ForeColor = Color.White
            ExportMenuItem.Enabled = True
        End If
    End Sub
    Private Sub Txt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles TxtPay7.TextChanged, TxtPay6.TextChanged, TxtDAOld.TextChanged, TxtDANew1.TextChanged, TxtDANew2.TextChanged
        SetButton()
    End Sub
    Sub SetTextBox(TB As Control)
        If TB.Text Is Nothing Then
            TB.Text = "____"
        Else
            If TB.Text.Equals("") Then TB.Text = "____"
        End If
    End Sub
    Private Sub ExportMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportMenuItem.Click
        SetTextBox(TxtName)
        SetTextBox(TxtOffice)
        SetTextBox(TxtDesignation)
        SetTextBox(TxtPlace)

        Using SaveFile As New SaveFileDialog
            SaveFile.Filter = "Excel File|*.xlsx|Excel Old File|*.xls"
            SaveFile.FileName = TxtName.Text

            If SaveFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                If My.Computer.FileSystem.FileExists(SaveFile.FileName) Then
                    Try
retry:                  My.Computer.FileSystem.DeleteFile(SaveFile.FileName)
                    Catch ex As IOException
                        If ex.Message.Contains("used by another process") Then
                            If MsgBox("The File '" + SaveFile.FileName + "' is used by Another Application" + Chr(13) + "Please, Close That Application and then click 'Retry' Button", vbRetryCancel, "Found An Error!") = vbRetry Then
                                GoTo retry
                            Else
                                MsgBox("Process is aborted by User!")
                                Exit Sub
                            End If
                        End If
                    End Try
                End If

                Dim file As New ExcelFile(SaveFile.FileName)
                file.Save()
            End If
        End Using
    End Sub
    Private Sub FileMenu_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileMenu.MouseEnter
        If Not ExportMenuItem.Enabled Then CurrentStatus.Text = "Please, Alteast fill all marked(*) fields to Export As Excel"
    End Sub
    Private Sub FileMenu_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileMenu.MouseLeave
        CurrentStatus.Text = "Ready"
    End Sub
    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If BtnSubmit.Bounds.IntersectsWith(New Rectangle(e.Location, New Size(10, 10))) Then
            CurrentStatus.Text = "Please Fill all nessecory fields to Unlock Submit Button"
        Else
            CurrentStatus.Text = "Ready"
        End If
    End Sub
    Private Sub _KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If BtnSubmit.Enabled Then BtnSubmit.PerformClick()

            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub ExitMenuItem_Click(sender As Object, e As EventArgs) Handles ExitMenuItem.Click
        Me.Close()
    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        SetBody()
    End Sub
    Sub SetBody()
        Body.Top = Menu.Top + Menu.Height
        Body.Left = 0
        Body.Height = ClientSize.Height - Body.Top - Status.Height
        Body.Width = ClientSize.Width
    End Sub

End Class
