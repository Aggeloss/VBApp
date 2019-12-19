Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1

    Public dataOfAllStudents As DataOfStudents = New DataOfStudents()
    Public Shared filepath As String = ""
    Public Shared startLineX As Integer
    Public Shared endLineX As Integer
    Public Shared startLineY As Integer
    Public Shared endLineY As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '============ Find the avg value of each student for the given values ============= 

        'Me.ShowGraph(dataOfAllStudents.getDataList())

        '==================================================================================

        '============ Separate each of the given values to be grouped into triads for each admeasurement =============

        'Me.ShowGroupGraph(dataOfAllStudents.getDataList())

        '=============================================================================================================
        'Me.DisplayGraph()

        'MsgBox(getStartLineX() & " " & getEndLineX() & " " & getStartLineY() & " " & getEndLineY())

        If (getFilePath() = "") Then
            MsgBox("You must open a file to continue")
        Else
            Dim line1 As String

            For Each item In dataOfAllStudents.getVariables()
                line1 = item & " : "
                Form3.RichTextBox2.AppendText(line1 & Environment.NewLine)
            Next

            Form3.DataGridView5.Visible = False
            Form3.Label9.Visible = False
            Form3.Label13.Visible = False
            Form3.Label11.Visible = False
            Form3.ComboBox5.Visible = False

            Form3.DataGridView1.ReadOnly = True
            Form3.DataGridView2.ReadOnly = True
            Form3.DataGridView3.ReadOnly = True

            Form3.Button9.Enabled = False
            Form3.Button6.Enabled = False
            'Form3.Button4.Enabled = False

            Form3.Show()
            Me.Hide()

        End If
    End Sub

    Public Sub ShowGraph(list As List(Of Integer))

        Dim totalValue As Integer = 0
        Dim avg As New List(Of Double)
        Dim a As Integer = 0
        Dim totalSteps = (getEndLineY() * getEndLineX())

        For i = (getStartLineY() * getStartLineX()) To (totalSteps - 1) Step 1
            a += 1
            If (a < getEndLineX()) Then
                'Console.WriteLine(list(i))
                totalValue += list(i)
            Else
                'Console.WriteLine(totalValue)
                avg.Add(totalValue / getEndLineX())
                totalValue = 0
                a = 0
            End If
        Next

        'For Each item As Integer In list
        'Console.WriteLine(item)
        'a += 1
        'If (a < getEndLineX()) Then
        'totalValue += item
        'Else
        'Console.WriteLine(totalValue)
        'avg.Add(totalValue / getEndLineX())
        'totalValue = 0
        'a = 0
        'End If
        'Next

        a = 1

        For Each item As Double In avg
            'If (dataOfAllStudents.getAgeList()(a) >= 24 And dataOfAllStudents.getAgeList()(a) <= 64) Then
            'Form3.Chart2.Series("Data").Points.AddXY(dataOfAllStudents.getDataOfParticipants(a), item)
            'Form3.Chart2.Series("Data").Points.AddXY(a, item)
            'End If
            a += 1
        Next
    End Sub

    Public Sub ConsoleData(list As List(Of Integer))

        Dim counter As Integer = 0

        For Each item In list
            counter += 1
        Next

        Console.WriteLine("The amount Of total values is: " & counter & ", The starting position is: " & getStartLineX() & ", The ending position is: " & getEndLineX())

    End Sub

    Public Sub ShowGroupGraph(list As List(Of Integer))

        'Dim valuesOfAllStudents As Integer() = New Integer(list.Count) {}

        Dim groupTotalValue As Integer = 0
        Dim avgValues As New List(Of Double)
        Dim b As Integer = 0
        Dim c As Integer = 0
        Dim flag As Integer = 0
        Dim totalSteps = list.Count - 1
        Dim lastStepFlag As Integer = (totalSteps / dataOfAllStudents.getDataOfFile()(2)) - 1

        For i = getStartLineY() To totalSteps Step 1
            b += 1
            If (b < getEndLineY() - getStartLineY()) Then
                'If (dataOfAllStudents.getAgeList()(c) >= 24 And dataOfAllStudents.getAgeList()(c) <= 64) Then
                groupTotalValue += list(i)
                If (b <> getEndLineY() - getStartLineY() - 1) Then
                    i += (list.Count / getEndLineY()) - 1
                End If
                If (flag = lastStepFlag And b = 1) Then
                    avgValues.Add(groupTotalValue / (getEndLineY() - getStartLineY() - 1))
                End If
                c += 1
                'Else
                'Continue For
                'End If
            Else
                avgValues.Add(groupTotalValue / (getEndLineY() - getStartLineY() - 1))
                groupTotalValue = 0
                i = flag
                flag += 1
                totalSteps -= getEndLineY() - getStartLineY()
                b = 0
            End If
        Next

        b = 0

        For Each item As Double In avgValues
            b += 1
            'Form3.Chart2.Series("Data").Points.AddXY(b, item)
        Next
    End Sub

    Public Function getData() As List(Of Integer)

        Dim FileNum As Integer = FreeFile()
        Dim TempS As String = ""
        Dim TempL As String

        Dim totalLines As Integer = 0
        Dim totalValues As Integer = 0

        Dim list As New List(Of String)

        Try
            FileOpen(FileNum, getFilePath(), OpenMode.Input)
        Catch ex As Exception
            MsgBox("Something went wrong with the filepath" & getFilePath())
        End Try

        Do Until EOF(FileNum)

            totalLines += 1

            TempL = LineInput(FileNum)

            'Console.WriteLine(TempL)

            Dim split As String() = TempL.Split(" ")
            Dim s As String

            If (totalLines = 1) Then
                totalValues = split.Length - 1
            End If

            Dim i As Integer = 0

            For Each s In split
                If s.Trim() <> "" Then
                    list.Add(s)
                End If
                i += 1
            Next s
            list.Add("/")
        Loop

        Dim counter As Integer = 0
        Dim flag As Boolean = True
        Dim firstTime As Boolean = True

        For Each item In list
            If (item = "/" And flag = True) Then
                flag = False
                firstTime = False
            End If
            If (flag = True And firstTime = True) Then
                counter += 1
            End If
        Next

        Dim dataOfFile As New List(Of Integer)

        dataOfFile.Add(2)
        dataOfFile.Add(totalLines)
        dataOfFile.Add(counter)
        dataOfFile.Add(totalLines - 1)

        setEndLineX(counter)
        setEndLineY(totalLines)

        dataOfAllStudents.setDataOfFile(dataOfFile)

        FileClose(FileNum)

        dataOfAllStudents.calculateValues(list)



        Return dataOfAllStudents.getDataList()

    End Function

    Public Function getAllVariablesReadingData()
        getData()
        Return dataOfAllStudents.getVariables()
    End Function

    Public Function getAllVariables()
        Return dataOfAllStudents.getVariables()
    End Function

    Public Function getAllData()

        For Each item In dataOfAllStudents.getAllData()
            Console.WriteLine(item)
        Next

        Return dataOfAllStudents.getAllData()
    End Function

    Private Sub DisplayGraph()
        'Here we can get our Scatter Graph and display it.
        'xlApp.Visible = True
        'xlWorkBook.Activate()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("You can change X and Y by doing so, if X = 0 - 10 can modify and replace it with, X = 5 - 8.")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        RichTextBox1.ReadOnly = True

        Dim extension As String
        Dim line1 As String
        Dim line2 As String
        Dim line3 As String
        Dim line4 As String

        OpenFileDialog1.Filter = "text files|*.txt"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.Title = "Open Text File"

        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            'filepath = OpenFileDialog1.FileName
            setFilePath(OpenFileDialog1.FileName)
            extension = Path.GetExtension(filepath)
            If (extension = ".txt") Then

                Me.getData()

                Dim data As Integer() = New Integer(4) {}

                Dim i As Integer = 0
                'Console.WriteLine(dataOfAllStudents.getDataOfFile().Count)
                For Each item In dataOfAllStudents.getDataOfFile()
                    data(i) = item
                    i += 1
                Next

                MsgBox("The chosen filepath is: " & getFilePath())
                line1 = "The data starts at line " & data(0)
                line2 = "The data ends at line " & data(1)
                line3 = "There are, " & data(2) & " assignments on X"
                line4 = "There are, " & data(3) & " assignments on Y"
                'line5 = "X = " & 0 & " - " & data(2)
                'line6 = "Y = " & 0 & " - " & data(3)
                RichTextBox1.AppendText(line1 & Environment.NewLine)
                RichTextBox1.AppendText(line2 & Environment.NewLine)
                RichTextBox1.AppendText(line3 & Environment.NewLine)
                RichTextBox1.AppendText(line4 & Environment.NewLine)
                'RichTextBox1.AppendText("" & Environment.NewLine)
                'RichTextBox1.AppendText(line5 & Environment.NewLine)
                'RichTextBox1.AppendText(line6 & Environment.NewLine)
            Else
                MsgBox("You can only choose text files")
                setFilePath("")
            End If
        End If
    End Sub

    Public Sub setFilePath(fp As String)
        filepath = fp
    End Sub

    Public Shared Function getFilePath() As String
        Return filepath
    End Function

    Public Sub setStartLineX(sX As Integer)
        startLineX = sX
    End Sub

    Public Shared Function getStartLineX() As Integer
        Return startLineX
    End Function

    Public Sub setEndLineX(eX As Integer)
        endLineX = eX
    End Sub

    Public Shared Function getEndLineX() As Integer
        Return endLineX
    End Function

    Public Sub setStartLineY(sY As Integer)
        startLineY = sY
    End Sub

    Public Shared Function getStartLineY() As Integer
        Return startLineY
    End Function

    Public Sub setEndLineY(eY As Integer)
        endLineY = eY
    End Sub

    Public Shared Function getEndLineY() As Integer
        Return endLineY
    End Function
End Class