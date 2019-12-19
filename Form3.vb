Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form3

    '============= Two Basic Arrays ===============

    Dim calcListX As Integer(,)
    Dim calcListY As Integer(,)

    '==============================================

    Dim firstForm As Form1 = New Form1()
    Dim secondForm As Form2 = New Form2()

    Dim table As New DataTable("DataGridView1")
    Dim tableAge As New DataTable("DataGridView2")
    Dim tableData As New DataTable("DataGridView3")
    Dim tableCalc As New DataTable("DataGridView4")
    Dim tableCalcCells As New DataTable("DataGridView5")

    '================= Main Data ==================

    Dim valuesOfColumns As List(Of Integer)
    Dim valuesOfRows As List(Of Integer)
    Dim valuesOfCells As List(Of Integer)

    Dim checkClicked As Boolean = False
    Dim checkContinued As Boolean = False
    Dim checkEdited As Boolean = False

    Dim allData As List(Of List(Of Integer))

    '==============================================

    Dim valuesOfFileInRichTextBox2 As String()
    Dim searchOptions As String()

    Dim selectedOption As String
    Dim newListWithVariables As List(Of String)

    Dim valuesOfEachColumn As List(Of String)

    Dim index As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        table.Columns.Add("", Type.GetType("System.String"))
        table.Columns.Add("LB", Type.GetType("System.String"))
        table.Columns.Add("RB", Type.GetType("System.String"))
        table.Columns.Add("Aud", Type.GetType("System.String"))
        table.Columns.Add("Vis", Type.GetType("System.String"))
        table.Columns.Add("Ling", Type.GetType("System.String"))
        table.Columns.Add("Kin", Type.GetType("System.String"))
        table.Columns.Add("Inter", Type.GetType("System.String"))
        table.Columns.Add("Intra", Type.GetType("System.String"))
        table.Columns.Add("Match", Type.GetType("System.String"))
        table.Columns.Add("Analo", Type.GetType("System.String"))
        table.Columns.Add("Sub", Type.GetType("System.String"))
        table.Columns.Add("Intersec", Type.GetType("System.String"))
        table.Columns.Add("Con", Type.GetType("System.String"))
        table.Columns.Add("Recon", Type.GetType("System.String"))
        table.Columns.Add("Rules", Type.GetType("System.String"))

        table.Rows.Add("Left Brain", "1", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        table.Rows.Add("Right Brain", "", "1", "", "", "", "", "", "", "", "", "", "", "", "", "")
        table.Rows.Add("Auditory", "", "", "1", "", "", "", "", "", "", "", "", "", "", "", "")
        table.Rows.Add("Visual", "", "", "", "1", "", "", "", "", "", "", "", "", "", "", "")
        table.Rows.Add("Linguistic", "", "", "", "", "1", "", "", "", "", "", "", "", "", "", "")
        table.Rows.Add("Kin aesthetic", "", "", "", "", "", "1", "", "", "", "", "", "", "", "", "")
        table.Rows.Add("Interpersonal", "", "", "", "", "", "", "1", "", "", "", "", "", "", "", "")
        table.Rows.Add("Intrapersonal", "", "", "", "", "", "", "", "1", "", "", "", "", "", "", "")
        table.Rows.Add("Matching", "", "", "", "", "", "", "", "", "1", "", "", "", "", "", "")
        table.Rows.Add("Analogies", "", "", "", "", "", "", "", "", "", "1", "", "", "", "", "")
        table.Rows.Add("Subsets", "", "", "", "", "", "", "", "", "", "", "1", "", "", "", "")
        table.Rows.Add("Intersections", "", "", "", "", "", "", "", "", "", "", "", "1", "", "", "")
        table.Rows.Add("Construction", "", "", "", "", "", "", "", "", "", "", "", "", "1", "", "")
        table.Rows.Add("Reconstruction", "", "", "", "", "", "", "", "", "", "", "", "", "", "1", "")
        table.Rows.Add("Rules", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "1")

        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

        DataGridView1.DataSource = table
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        tableAge.Columns.Add("", Type.GetType("System.String"))
        tableAge.Columns.Add("leftbrain", Type.GetType("System.String"))
        tableAge.Columns.Add("rightbrain", Type.GetType("System.String"))
        tableAge.Columns.Add("aud", Type.GetType("System.String"))
        tableAge.Columns.Add("vis", Type.GetType("System.String"))
        tableAge.Columns.Add("ling", Type.GetType("System.String"))
        tableAge.Columns.Add("kin", Type.GetType("System.String"))
        tableAge.Columns.Add("interp", Type.GetType("System.String"))
        tableAge.Columns.Add("intra", Type.GetType("System.String"))
        tableAge.Columns.Add("match", Type.GetType("System.String"))
        tableAge.Columns.Add("analo", Type.GetType("System.String"))
        tableAge.Columns.Add("sub", Type.GetType("System.String"))
        tableAge.Columns.Add("inters", Type.GetType("System.String"))
        tableAge.Columns.Add("const", Type.GetType("System.String"))
        tableAge.Columns.Add("recon", Type.GetType("System.String"))
        tableAge.Columns.Add("rules", Type.GetType("System.String"))
        tableAge.Columns.Add("math", Type.GetType("System.String"))

        tableAge.Rows.Add("age 24-34", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        tableAge.Rows.Add("age 34-44", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        tableAge.Rows.Add("age 44-54", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        tableAge.Rows.Add("age 54-64", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

        DataGridView2.DataSource = tableAge
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlWorkBook = xlApp.Workbooks.Add()
        xlWorkSheet = xlWorkBook.Sheets(1)
        xlWorkSheet.Cells(3, 7) = "ScatterPlots Diagram"

        Try
            Dim myChart As Excel.ChartObject
            Dim sr As Excel.Series

            Dim x As String
            Dim i As Integer
            Dim j As Integer
            Dim listOfXValues As New List(Of Integer)
            Dim listOfYValues As New List(Of Integer)

            MsgBox("Excel file exported with name ScExcel.xlsx.")

            For i = 0 To DataGridView1.RowCount - 2
                For j = 0 To DataGridView1.ColumnCount - 1
                    x = DataGridView1.Rows(i).Cells(j).Value
                    If IsNumeric(x) Then
                        listOfXValues.Add(i)
                        listOfYValues.Add(j)
                    End If
                Next
            Next

            Dim xValues As Integer() = New Integer(listOfXValues.Count) {}
            Dim yValues As Integer() = New Integer(listOfYValues.Count) {}

            i = 0

            For Each item As Integer In listOfXValues
                xValues(i) = item
                Console.WriteLine(item)
                i += 1
            Next

            i = 0

            For Each item As Integer In listOfYValues
                yValues(i) = item
                i += 1
            Next

            myChart = xlWorkSheet.ChartObjects.Add(95, 50, 500, 200)
            sr = myChart.Chart.SeriesCollection.NewSeries
            sr.XValues = xValues
            sr.Values = yValues

            With myChart
                .Chart.ChartType = Excel.XlChartType.xlXYScatter
                .Chart.ApplyLayout(4)
                .Chart.ChartStyle = 2
            End With

            xlWorkSheet.SaveAs("C:\Users\ANGELOS\Desktop\ScExcel.xlsx")
        Catch ex As Exception
            MsgBox("Something went wrong while exporting ScatterPlots graph.")
        End Try

        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

        'Select Case ComboBox1.SelectedItem.ToString()
        'Case "Show Graph"
        'Me.Chart2.Series("Data").Points.Clear()
        'firstForm.ShowGraph(firstForm.getData())
        'Case "Show Group Graph"
        'Me.Chart2.Series("Data").Points.Clear()
        'firstForm.ShowGroupGraph(firstForm.getData())
        'Case "Console Data"
        'firstForm.ConsoleData(firstForm.getData())
        'End Select
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If checkClicked Then
            MsgBox("You have already added the values")
        Else
            newListWithVariables = New List(Of String)

            Dim variables As List(Of String)
            Dim data As List(Of String)
            Dim chunkCollection As Integer = 0

            variables = firstForm.getAllVariablesReadingData()

            searchOptions = New String(variables.Count - 1) {}

            Dim counterSearch As Integer = 0

            For Each item In variables
                searchOptions(counterSearch) = item
                counterSearch += 1
            Next

            selectedOption = searchOptions(0)

            ComboBox4.DataSource = variables

            data = firstForm.getAllData()

            chunkCollection = variables.Count

            Console.WriteLine("Sizeee : " & variables.Count)

            For Each item In variables
                newListWithVariables.Add(item)
                If tableData.Columns.Contains(item) Then
                    MsgBox("You have already added the values")
                    Exit For
                Else
                    tableData.Columns.Add(item, Type.GetType("System.String"))
                End If
            Next

            Dim row As String() = New String(chunkCollection - 1) {}
            Dim counter As Integer = 0

            For Each item In data
                row(counter) = item
                If (counter = chunkCollection - 1) Then
                    tableData.Rows.Add(row)
                    row = New String(chunkCollection - 1) {}
                    counter = 0
                    Continue For
                End If
                counter += 1
            Next

            DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

            DataGridView3.DataSource = tableData

            checkClicked = True
        End If

    End Sub

    Private Sub DataGridView3_clicked(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick

        Dim variables As List(Of String)
        variables = firstForm.getAllVariables()

        index = e.RowIndex
        Dim selectedRow As DataGridViewRow

        selectedRow = DataGridView3.Rows(index)

        Dim line1 As String
        Dim counter As Integer = 0

        RichTextBox2.Text = ""

        For Each item In variables
            line1 = item & " : " & selectedRow.Cells(counter).Value.ToString()
            RichTextBox2.AppendText(line1 & Environment.NewLine)
            counter += 1
        Next
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Dim variables As List(Of String) = firstForm.getAllVariables()
        valuesOfFileInRichTextBox2 = New String(variables.Count - 1) {}
        Dim i As Integer

        Dim context As String()

        For i = 0 To variables.Count - 1 Step 1
            context = RichTextBox2.Lines(i).ToString().Split
            valuesOfFileInRichTextBox2(i) = context(2)
        Next

        For Each item In valuesOfFileInRichTextBox2
            Console.WriteLine(item)
        Next

        Dim counterNull As Integer = 0
        Dim counter As Integer = 0

        For Each item In valuesOfFileInRichTextBox2
            counter += 1
            If item = "" Then
                counterNull += 1
            End If
            If counterNull = valuesOfFileInRichTextBox2.Length() Then
                MsgBox("There is no data to add")
            ElseIf counter = valuesOfFileInRichTextBox2.Length() Then
                tableData.Rows.Add(valuesOfFileInRichTextBox2)
            End If
        Next

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim variables As List(Of String) = firstForm.getAllVariables()
        valuesOfFileInRichTextBox2 = New String(variables.Count - 1) {}
        Dim i As Integer

        Dim context As String()

        For i = 0 To variables.Count - 1 Step 1
            context = RichTextBox2.Lines(i).ToString().Split
            valuesOfFileInRichTextBox2(i) = context(2)
        Next

        Dim selectedRow As DataGridViewRow

        selectedRow = DataGridView3.Rows(index)

        Dim counterNull As Integer = 0
        Dim counter As Integer = 0

        For Each item In valuesOfFileInRichTextBox2
            counter += 1
            If item = "" Then
                counterNull += 1
            End If
            If counterNull = valuesOfFileInRichTextBox2.Length() Then
                MsgBox("There is no data to update")
            ElseIf counter = valuesOfFileInRichTextBox2.Length() Then
                For i = 0 To variables.Count - 1 Step 1
                    selectedRow.Cells(i).Value = valuesOfFileInRichTextBox2(i)
                Next
            End If
        Next

        'tableData.Rows.Update(valuesOfFileInRichTextBox2(0))

        'For Each item In valuesOfFileInRichTextBox2
        'Console.WriteLine(item)
        'Next

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'Dim selectedRow As DataGridViewRow
        'selectedRow = DataGridView3.Rows(index)
        'DataGridView3.Rows.Remove(selectedRow)

        For Each row As DataGridViewRow In DataGridView3.SelectedRows
            DataGridView3.Rows.Remove(row)
        Next

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        RichTextBox2.Text = ""

        Dim variables As List(Of String) = firstForm.getAllVariables()

        Dim line1 As String
        Console.WriteLine(variables.Count)
        If variables.Count <= 1 Then
            For Each item In firstForm.getAllVariablesReadingData()
                line1 = item & " : "
                RichTextBox2.AppendText(line1 & Environment.NewLine)
            Next
        Else
            For Each item In firstForm.getAllVariables()
                line1 = item & " : "
                RichTextBox2.AppendText(line1 & Environment.NewLine)
            Next
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim variables As List(Of String) = firstForm.getAllVariables()

        Dim DV As New DataView(tableData)

        DV.RowFilter = String.Format(selectedOption & " Like '%{0}%'", TextBox1.Text)
        DataGridView3.DataSource = DV

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        selectedOption = searchOptions(ComboBox4.SelectedIndex)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Dim variables As List(Of String) = firstForm.getAllVariables()

        Dim checkIfThereIsVariableNamedAge As Boolean = False

        If variables.Count <= 1 Then
            MsgBox("You should first add values")
        Else
            For Each item In firstForm.getAllVariables()
                If "age".ToUpper() = item.ToString.ToUpper() Then
                    checkIfThereIsVariableNamedAge = True
                End If
            Next
        End If

        If checkIfThereIsVariableNamedAge = True Then

            Dim DV As New DataView(tableData)

            Select Case ComboBox3.SelectedItem.ToString()
                Case "remove filter"
                    DV.RowFilter = String.Format("")
                    DataGridView3.DataSource = DV
                Case "< 14"
                    DV.RowFilter = String.Format("Age < 14")
                    DataGridView3.DataSource = DV
                Case ">= 14 and < 24"
                    DV.RowFilter = String.Format("Age >= 14 And Age < 24")
                    DataGridView3.DataSource = DV
                Case ">= 24 and < 34"
                    DV.RowFilter = String.Format("Age >= 24 And Age < 34")
                    DataGridView3.DataSource = DV
                Case ">= 34"
                    DV.RowFilter = String.Format("Age >= 34")
                    DataGridView3.DataSource = DV
            End Select
        ElseIf checkIfThereIsVariableNamedAge = False And Not variables.Count <= 1 Then
            MsgBox("This operation demands a variable named Age or age")
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        valuesOfEachColumn = New List(Of String)

        Dim variables As List(Of String) = firstForm.getAllVariables()
        Dim counter As Integer = 0
        Dim i As Integer

        Dim tempList As List(Of String) = New List(Of String)
        Dim rowInsert As String()
        Dim j As Integer = 0
        Dim h As Integer = 0

        If variables.Count <= 1 Then
            MsgBox("You should first add values")
        ElseIf valuesOfColumns Is Nothing And valuesOfRows Is Nothing Or checkContinued And Not checkEdited Then

            checkContinued = True

            tableCalc.Rows.Clear()
            tableCalc.Columns.Clear()

            valuesOfColumns = New List(Of Integer)
            valuesOfRows = New List(Of Integer)

            newListWithVariables = New List(Of String)

            For Each item In variables
                newListWithVariables.Add(item)
            Next

            If DataGridView3.SelectedRows.Count < 1 Then
                For Each row As DataGridViewRow In DataGridView3.Rows
                    counter += 1
                Next

                Dim firstTime As Boolean = True

                For Each row As DataGridViewRow In DataGridView3.Rows
                    For i = 0 To variables.Count - 1 Step 1
                        If IsNumeric(row.Cells(i).Value) Then
                            valuesOfEachColumn.Add(row.Cells(i).Value)
                            valuesOfRows.Add(row.Cells(i).Value)
                        Else
                            If firstTime Then
                                newListWithVariables.RemoveAt(i)
                            End If
                            Continue For
                        End If
                    Next
                    firstTime = False
                Next
            Else
                counter += 1
                For Each row As DataGridViewRow In DataGridView3.SelectedRows
                    counter += 1
                Next

                Dim firstTime As Boolean = True

                For Each row As DataGridViewRow In DataGridView3.SelectedRows
                    For i = 0 To variables.Count - 1 Step 1
                        If IsNumeric(row.Cells(i).Value) Then
                            valuesOfEachColumn.Add(row.Cells(i).Value)
                            valuesOfRows.Add(row.Cells(i).Value)
                        Else
                            If firstTime Then
                                newListWithVariables.RemoveAt(i)
                            End If
                            Continue For
                        End If
                    Next
                    firstTime = False
                Next
            End If

            For Each item In newListWithVariables
                Console.WriteLine(item)
            Next
            rowInsert = New String(counter - 1) {}

            For i = 0 To newListWithVariables.Count - 1 Step 1
                h = i
                While j < counter
                    If j = 0 Then
                        rowInsert(j) = newListWithVariables(i)
                    Else
                        rowInsert(j) = valuesOfEachColumn(h)
                        Console.WriteLine(rowInsert(j))
                        h += newListWithVariables.Count
                    End If
                    j += 1
                End While
                j = 0
                For a = 1 To rowInsert.Length - 1 Step 1
                    If IsNumeric(rowInsert(a)) Then
                        valuesOfColumns.Add(rowInsert(a))
                    Else
                        Continue For
                    End If
                Next
            Next
        Else
            'tableCalc.Rows.Clear()
            'tableCalc.Columns.Clear()

            'counter = 0

            'For Each row As DataGridViewRow In DataGridView3.SelectedRows
            'counter += 1
            'Next

            'For Each row As DataGridViewRow In DataGridView3.SelectedRows
            'For i = 0 To variables.Count - 1 Step 1
            'If IsNumeric(row.Cells(i).Value) Then
            'valuesOfEachColumn.Add(row.Cells(i).Value)
            'valuesOfRows.Add(row.Cells(i).Value)
            'Else
            'Continue For
            'End If
            '   Next
            '  Next
            '
            'rowInsert = New String(counter - 1) {}

            'For i = 0 To newListWithVariables.Count - 1 Step 1
            'h = i
            'While j < counter
            'If j = 0 Then
            'rowInsert(j) = newListWithVariables(i)
            'Else
            'rowInsert(j) = valuesOfEachColumn(h)
            'h += newListWithVariables.Count
            'End If
            '   j += 1
            '  End While
            ' j = 0
            'For a = 1 To rowInsert.Length - 1 Step 1
            'If IsNumeric(rowInsert(a)) Then
            'valuesOfColumns.Add(rowInsert(a))
            'Else
            'Continue For
            'End If
            '   Next
            '  Next
        End If

        allData = New List(Of List(Of Integer))

        Dim secondCounter As Integer = 0

        If Not valuesOfColumns Is Nothing Then
            allData.Add(getValuesOfColumns())
            secondCounter += 1
        End If

        If Not valuesOfRows Is Nothing Then
            allData.Add(getValuesOfRows())
            secondCounter += 1
        End If

        If Not valuesOfCells Is Nothing Then
            allData.Add(getValuesOfCells())
            secondCounter += 1

            DataGridView5.Visible = True
            Label9.Visible = True
            Label13.Visible = True
            Label11.Visible = True
            ComboBox5.Visible = True
        End If

        If secondCounter >= 2 Then
            Console.WriteLine(allData.Count)
            'TabControl1.TabPages(1).Visible = True
            TabControl1.SelectedIndex = 1 'It selects second tab
            If secondForm.getIndexOfRemovedRow().Count >= 1 Then
                tableCalc.Rows.Clear()
                tableCalc.Columns.Clear()
            End If
            For Each index In secondForm.getIndexOfRemovedRow()
                Console.WriteLine("Sizeee: " & newListWithVariables.Count & " and the index is: " & index & " deleted item: " & newListWithVariables(index))
                newListWithVariables.RemoveAt(index)
            Next
            If Not getValuesOfCells() Is Nothing Then
                If getValuesOfCells().Count > 1 Then
                    tableCalcCells.Rows.Clear()
                    tableCalcCells.Columns.Clear()
                    Dim count As Integer = 0

                    For Each item In getValuesOfCells()
                        tableCalcCells.Columns.Add(count, Type.GetType("System.String"))
                        count += 1
                    Next

                    Dim stringArray As String() = New String(count - 1) {}
                    Dim x As Integer = 0

                    For Each item In getValuesOfCells()
                        stringArray(x) = item
                        x += 1
                    Next

                    tableCalcCells.Rows.Add(stringArray)

                    DataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

                    DataGridView5.DataSource = tableCalcCells
                End If
            End If
        Else
            MsgBox("Something went wrong. Please try again.", MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        checkEdited = True

        Dim tableData2 As New DataTable("Form2.DataGridView1")

        valuesOfEachColumn = New List(Of String)

        Dim variables As List(Of String) = firstForm.getAllVariables()
        Dim counter As Integer = 0
        Dim i As Integer

        Dim tempList As List(Of String) = New List(Of String)
        Dim rowInsert As String() = New String() {}
        Dim j As Integer = 0
        Dim h As Integer = 0

        valuesOfColumns = New List(Of Integer)
        valuesOfRows = New List(Of Integer)

        newListWithVariables = New List(Of String)

        For Each item In variables
            newListWithVariables.Add(item)
        Next

        Console.WriteLine(variables.Count)
        If variables.Count <= 1 Then
            MsgBox("You should first add values")
        ElseIf DataGridView3.SelectedRows.Count < 1 Then
            For Each row As DataGridViewRow In DataGridView3.Rows
                tableData2.Columns.Add(counter, Type.GetType("System.String"))
                counter += 1
            Next

            Dim firstTime As Boolean = True

            For Each row As DataGridViewRow In DataGridView3.Rows
                For i = 0 To variables.Count - 1 Step 1
                    If IsNumeric(row.Cells(i).Value) Then
                        valuesOfEachColumn.Add(row.Cells(i).Value)
                        valuesOfRows.Add(row.Cells(i).Value)
                    Else
                        If firstTime Then
                            newListWithVariables.RemoveAt(i)
                        End If
                        Continue For
                    End If
                Next
                firstTime = False
            Next

            tableData2.Columns.Add(counter + 1, Type.GetType("System.String")) 'cause of variables

            rowInsert = New String(counter - 1) {}

            For i = 0 To newListWithVariables.Count - 1 Step 1
                h = i
                While j < counter
                    If j = 0 Then
                        rowInsert(j) = newListWithVariables(i)
                    Else
                        rowInsert(j) = valuesOfEachColumn(h)
                        h += newListWithVariables.Count
                    End If
                    j += 1
                End While
                j = 0
                For a = 1 To rowInsert.Length - 1 Step 1
                    If IsNumeric(rowInsert(a)) Then
                        valuesOfColumns.Add(rowInsert(a))
                    Else
                        Continue For
                    End If
                Next
                tableData2.Rows.Add(rowInsert)
            Next

            Form2.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

            Form2.DataGridView1.DataSource = tableData2
            Form2.Show()
        Else

            counter = 0

            tableData2.Columns.Add(counter, Type.GetType("System.String")) 'cause of variables

            For Each row As DataGridViewRow In DataGridView3.SelectedRows
                counter += 1
                tableData2.Columns.Add(counter, Type.GetType("System.String"))
            Next

            newListWithVariables = New List(Of String)

            For Each item In variables
                newListWithVariables.Add(item)
            Next

            Dim firstTime As Boolean = True

            For Each row As DataGridViewRow In DataGridView3.SelectedRows
                For i = 0 To variables.Count - 1 Step 1
                    If IsNumeric(row.Cells(i).Value) Then
                        valuesOfEachColumn.Add(row.Cells(i).Value)
                        valuesOfRows.Add(row.Cells(i).Value)
                        Console.WriteLine(row.Cells(i).Value)
                    Else
                        If firstTime Then
                            newListWithVariables.RemoveAt(i)
                        End If
                        Continue For
                    End If
                Next
                firstTime = False
            Next
            Console.WriteLine("============================================")
            rowInsert = New String(counter) {}

            h = 0
            newListWithVariables.Reverse()
            valuesOfEachColumn.Reverse()
            For Each item In valuesOfEachColumn
                Console.WriteLine(item)
            Next
            For i = newListWithVariables.Count - 1 To 0 Step -1
                h = i
                While j < counter + 1
                    If j = 0 Then
                        rowInsert(j) = newListWithVariables(i)
                    Else
                        rowInsert(j) = valuesOfEachColumn(h)
                        'Console.WriteLine(rowInsert(j))
                        h += newListWithVariables.Count
                    End If
                    j += 1
                End While
                j = 0
                For a = 1 To rowInsert.Length - 1 Step 1
                    If IsNumeric(rowInsert(a)) Then
                        valuesOfColumns.Add(rowInsert(a))
                    Else
                        Continue For
                    End If
                Next
                tableData2.Rows.Add(rowInsert)
            Next

            newListWithVariables.Reverse()

            Form2.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

            Form2.DataGridView1.DataSource = tableData2
            Form2.Show()
        End If
    End Sub

    Private Sub GetSelectedCellsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetSelectedCellsToolStripMenuItem.Click

        Dim listOfValues As List(Of Integer) = New List(Of Integer)
        valuesOfCells = New List(Of Integer)

        For Each cell As DataGridViewCell In DataGridView3.SelectedCells
            If Not IsNumeric(cell.Value) Then
                Dim result As Integer = MessageBox.Show("You have chosen a value with type of string. Do you want to continue disregarding this value ?", "Different Type Of Values Message", MessageBoxButtons.YesNoCancel)
                If result = DialogResult.Cancel Then
                    Exit For
                ElseIf result = DialogResult.No Then
                    Exit For
                ElseIf result = DialogResult.Yes Then
                    Continue For
                End If
            Else
                listOfValues.Add(cell.Value)
            End If
        Next

        listOfValues.Reverse()

        For Each item In listOfValues
            valuesOfCells.Add(item)
        Next
    End Sub

    Public Function getValuesOfColumns() As List(Of Integer)

        Dim temp As List(Of Integer) = secondForm.getValuesOfColumns()

        If temp.Count > 1 Then
            Console.WriteLine("Columns")
            valuesOfColumns = temp
        End If

        'Console.WriteLine("Size #form2 : " & temp.Count)

        Return valuesOfColumns

    End Function

    Public Function getValuesOfRows() As List(Of Integer)

        Dim temp As List(Of Integer) = secondForm.getValuesOfColumns()

        If temp.Count > 1 Then
            Console.WriteLine("Rows")
            valuesOfRows = temp
        End If

        Return valuesOfRows

    End Function

    Public Function getValuesOfCells() As List(Of Integer)

        Return valuesOfCells

    End Function

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Select Case ComboBox2.SelectedItem.ToString()
            Case "X axis"
                If Not getValuesOfColumns() Is Nothing And Not getValuesOfRows() Is Nothing Then

                    Me.Cursor = Cursors.WaitCursor

                    calcListY = Nothing

                    tableCalc.Rows.Clear()
                    tableCalc.Columns.Clear()

                    Dim counter As Integer = 0
                    Dim temp As String() = New String(newListWithVariables.Count - 1) {}
                    Dim i As Integer

                    For i = 0 To newListWithVariables.Count - 1 Step 1
                        tableCalc.Columns.Add(newListWithVariables(i))
                    Next
                    Dim position As Integer
                    Dim j As Integer = 0

                    calcListX = New Integer((getValuesOfColumns().Count / newListWithVariables.Count) - 1, newListWithVariables.Count - 1) {}

                    For i = 0 To (getValuesOfColumns().Count / newListWithVariables.Count) - 1 Step 1
                        position = i
                        While j < newListWithVariables.Count
                            temp(counter) = getValuesOfColumns()(position)
                            calcListX(i, j) = temp(counter)
                            'Console.WriteLine("Inserted: " & temp(counter) & ", at position => " & position)
                            counter += 1
                            position += ((getValuesOfColumns().Count) / newListWithVariables.Count)
                            j += 1
                        End While
                        j = 0
                        If counter = newListWithVariables.Count Then
                            tableCalc.Rows.Add(temp)
                            temp = New String(newListWithVariables.Count - 1) {}
                            counter = 0
                        End If
                    Next

                    DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

                    DataGridView4.DataSource = tableCalc

                    Me.Cursor = Cursors.Default
                Else
                    MsgBox("You should first add values")
                End If
            Case "Y axis"
                If Not getValuesOfColumns() Is Nothing And Not getValuesOfRows() Is Nothing Then

                    Me.Cursor = Cursors.WaitCursor

                    calcListX = Nothing

                    tableCalc.Rows.Clear()
                    tableCalc.Columns.Clear()

                    Dim variables As List(Of String) = firstForm.getAllVariables()
                    Dim counter As Integer = 0
                    Dim i As Integer

                    tableCalc.Columns.Add(0, Type.GetType("System.String")) 'cause of variables

                    For Each row As DataGridViewRow In DataGridView3.Rows
                        counter += 1
                        tableCalc.Columns.Add(counter, Type.GetType("System.String"))
                    Next

                    Dim variablesCounter As Integer = 0

                    Dim rowInsert As String() = New String(counter) {}
                    Dim j As Integer = 0
                    Dim h As Integer = 0

                    Dim listOfValuesOfColumns As List(Of Integer) = getValuesOfColumns()

                    For Each item In newListWithVariables
                        Console.WriteLine(item)
                    Next

                    calcListY = New Integer(newListWithVariables.Count - 1, (listOfValuesOfColumns.Count / newListWithVariables.Count) - 1) {}

                    Dim z As Integer = 0

                    For i = 0 To newListWithVariables.Count - 1 Step 1
                        While j < (listOfValuesOfColumns.Count / newListWithVariables.Count) + 1
                            If j = 0 Then
                                rowInsert(j) = newListWithVariables(i)
                            Else
                                rowInsert(j) = listOfValuesOfColumns(h)
                                calcListY(i, z) = rowInsert(j)
                                z += 1
                                h += 1
                            End If
                            j += 1
                        End While
                        z = 0
                        j = 0
                        tableCalc.Rows.Add(rowInsert)
                    Next

                    DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

                    DataGridView4.DataSource = tableCalc

                    Me.Cursor = Cursors.Default
                Else
                    MsgBox("You should first add values")
                End If
        End Select
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        Dim average As List(Of Integer) = New List(Of Integer)
        Dim total As Integer = 0
        Dim i As Integer
        Dim j As Integer

        If Not calcListX Is Nothing Then
            For i = 0 To calcListX.GetLength(0) - 1 Step 1
                For j = 0 To calcListX.GetLength(1) - 1 Step 1
                    total += calcListX(i, j)
                Next
                average.Add(total / calcListX.GetLength(1))
                total = 0
            Next
        ElseIf Not calcListY Is Nothing Then
            For i = 0 To calcListY.GetLength(0) - 1 Step 1
                For j = 0 To calcListY.GetLength(1) - 1 Step 1
                    total += calcListY(i, j)
                Next
                average.Add(total / calcListY.GetLength(1))
                total = 0
            Next
        End If

        Dim a As Integer = 1

        For Each item In average
            Form4.Chart1.Series("Average").Points.AddXY(a, item)
            a += 1
        Next

        Form4.Show()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        Dim average As List(Of Integer) = New List(Of Integer)
        Dim total As Integer = 0
        Dim i As Integer
        Dim j As Integer

        Dim deviation As List(Of Integer) = New List(Of Integer)

        If Not calcListX Is Nothing Then
            For i = 0 To calcListX.GetLength(0) - 1 Step 1
                For j = 0 To calcListX.GetLength(1) - 1 Step 1
                    total += calcListX(i, j)
                Next
                average.Add(total / calcListX.GetLength(1))
                total = 0
            Next
            For i = 0 To calcListX.GetLength(0) - 1 Step 1
                For j = 0 To calcListX.GetLength(1) - 1 Step 1
                    total += ((calcListX(i, j) - average(i)) ^ 2)
                Next
                deviation.Add(Math.Sqrt(total / (calcListX.GetLength(1) - 1)))
                total = 0
            Next
        ElseIf Not calcListY Is Nothing Then
            For i = 0 To calcListY.GetLength(0) - 1 Step 1
                For j = 0 To calcListY.GetLength(1) - 1 Step 1
                    total += calcListY(i, j)
                Next
                average.Add(total / calcListY.GetLength(1))
                total = 0
            Next
            For i = 0 To calcListY.GetLength(0) - 1 Step 1
                For j = 0 To calcListY.GetLength(1) - 1 Step 1
                    total += ((calcListY(i, j) - average(i)) ^ 2)
                Next
                deviation.Add(Math.Sqrt(total / calcListY.GetLength(1)))
                total = 0
            Next
        End If

        Dim a As Integer = 1

        For Each item In deviation
            Form5.Chart1.Series("Deviation").Points.AddXY(a, item)
            a += 1
        Next

        Form5.Show()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        Dim average As List(Of Integer) = New List(Of Integer)
        Dim total As Integer = 0
        Dim i As Integer
        Dim j As Integer

        Dim fluctuation As List(Of Integer) = New List(Of Integer)

        If Not calcListX Is Nothing Then
            For i = 0 To calcListX.GetLength(0) - 1 Step 1
                For j = 0 To calcListX.GetLength(1) - 1 Step 1
                    total += calcListX(i, j)
                Next
                average.Add(total / calcListX.GetLength(1))
                total = 0
            Next
            For i = 0 To calcListX.GetLength(0) - 1 Step 1
                For j = 0 To calcListX.GetLength(1) - 1 Step 1
                    total += ((calcListX(i, j) - average(i)) ^ 2)
                Next
                fluctuation.Add(total / calcListX.GetLength(1))
                total = 0
            Next
        ElseIf Not calcListY Is Nothing Then
            For i = 0 To calcListY.GetLength(0) - 1 Step 1
                For j = 0 To calcListY.GetLength(1) - 1 Step 1
                    total += calcListY(i, j)
                Next
                average.Add(total / calcListY.GetLength(1))
                total = 0
            Next
            For i = 0 To calcListY.GetLength(0) - 1 Step 1
                For j = 0 To calcListY.GetLength(1) - 1 Step 1
                    total += ((calcListY(i, j) - average(i)) ^ 2)
                Next
                fluctuation.Add(total / calcListY.GetLength(1))
                total = 0
            Next
        End If

        Dim a As Integer = 1

        For Each item In fluctuation
            Form6.Chart1.Series("Fluctuation").Points.AddXY(a, item)
            a += 1
        Next

        Form6.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedItem.ToString()
            Case "ScatterPlot Chart"
                Dim xlApp As New Excel.Application
                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet

                xlWorkBook = xlApp.Workbooks.Add()
                xlWorkSheet = xlWorkBook.Sheets(1)
                xlWorkSheet.Cells(3, 7) = "ScatterPlots Diagram"

                Try
                    Dim myChart As Excel.ChartObject
                    Dim sr As Excel.Series

                    MsgBox("Excel file exported with name ScExcel_ScatterPlot_Chart.xlsx.")

                    Dim average As List(Of Integer) = New List(Of Integer)
                    Dim total As Integer = 0
                    Dim i As Integer
                    Dim j As Integer

                    If Not calcListX Is Nothing Then
                        For i = 0 To calcListX.GetLength(0) - 1 Step 1
                            For j = 0 To calcListX.GetLength(1) - 1 Step 1
                                total += calcListX(i, j)
                            Next
                            average.Add(total / calcListX.GetLength(1))
                            total = 0
                        Next
                    ElseIf Not calcListY Is Nothing Then
                        For i = 0 To calcListY.GetLength(0) - 1 Step 1
                            For j = 0 To calcListY.GetLength(1) - 1 Step 1
                                total += calcListY(i, j)
                            Next
                            average.Add(total / calcListY.GetLength(1))
                            total = 0
                        Next
                    End If

                    Dim xValues As Integer() = New Integer(average.Count - 1) {}
                    Dim yValues As Integer() = New Integer(average.Count - 1) {}

                    i = 0

                    For Each item As Integer In average
                        yValues(i) = item
                        'Console.WriteLine(item)
                        i += 1
                    Next

                    i = 0
                    Dim a As Integer = 0

                    For Each item As Integer In average
                        a += 10
                        xValues(i) = a
                        i += 1
                    Next

                    For Each item In xValues
                        Console.WriteLine(item)
                    Next

                    myChart = xlWorkSheet.ChartObjects.Add(95, 50, 500, 200)
                    sr = myChart.Chart.SeriesCollection.NewSeries
                    sr.XValues = xValues
                    sr.Values = yValues

                    With myChart
                        .Chart.ChartType = Excel.XlChartType.xlXYScatter

                        .Chart.ApplyLayout(4)
                        .Chart.ChartStyle = 2
                    End With

                    xlWorkSheet.SaveAs("C:\Users\ANGELOS\Desktop\ScExcel_ScatterPlot_Chart.xlsx")
                Catch ex As Exception
                    MsgBox("Something went wrong while exporting ScatterPlots graph.")
                End Try

                xlWorkBook.Close()
                xlApp.Quit()
                releaseObject(xlApp)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkSheet)
            Case "Line Chart"
                Dim xlApp As New Excel.Application
                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet

                xlWorkBook = xlApp.Workbooks.Add()
                xlWorkSheet = xlWorkBook.Sheets(1)
                xlWorkSheet.Cells(4, 7) = "Line Diagram"

                Try
                    Dim myChart As Excel.ChartObject
                    Dim sr As Excel.Series

                    MsgBox("Excel file exported with name ScExcel_Line_Chart.xlsx")

                    Dim average As List(Of Integer) = New List(Of Integer)
                    Dim total As Integer = 0
                    Dim i As Integer
                    Dim j As Integer

                    If Not calcListX Is Nothing Then
                        For i = 0 To calcListX.GetLength(0) - 1 Step 1
                            For j = 0 To calcListX.GetLength(1) - 1 Step 1
                                total += calcListX(i, j)
                            Next
                            average.Add(total / calcListX.GetLength(1))
                            total = 0
                        Next
                    ElseIf Not calcListY Is Nothing Then
                        For i = 0 To calcListY.GetLength(0) - 1 Step 1
                            For j = 0 To calcListY.GetLength(1) - 1 Step 1
                                total += calcListY(i, j)
                            Next
                            average.Add(total / calcListY.GetLength(1))
                            total = 0
                        Next
                    End If

                    Dim xValues As Integer() = New Integer(average.Count - 1) {}
                    Dim yValues As Integer() = New Integer(average.Count - 1) {}

                    i = 0

                    For Each item As Integer In average
                        yValues(i) = item
                        'Console.WriteLine(item)
                        i += 1
                    Next

                    i = 0
                    Dim a As Integer = 0

                    For Each item As Integer In average
                        a += 10
                        xValues(i) = a
                        i += 1
                    Next

                    For Each item In xValues
                        Console.WriteLine(item)
                    Next

                    myChart = xlWorkSheet.ChartObjects.Add(95, 50, 500, 200)
                    sr = myChart.Chart.SeriesCollection.NewSeries
                    sr.XValues = xValues
                    sr.Values = yValues

                    With myChart
                        .Chart.ChartType = Excel.XlChartType.xlLine

                        .Chart.ApplyLayout(4)
                        .Chart.ChartStyle = 2
                    End With

                    xlWorkSheet.SaveAs("C:\Users\ANGELOS\Desktop\ScExcel_Line_Chart.xlsx")
                Catch ex As Exception
                    MsgBox("Something went wrong while exporting Line graph.")
                End Try

                xlWorkBook.Close()
                xlApp.Quit()
                releaseObject(xlApp)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkSheet)
            Case "Pie Chart"
                Dim xlApp As New Excel.Application
                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet

                xlWorkBook = xlApp.Workbooks.Add()
                xlWorkSheet = xlWorkBook.Sheets(1)
                xlWorkSheet.Cells(4, 7) = "Pie Diagram"

                Try
                    Dim myChart As Excel.ChartObject
                    Dim sr As Excel.Series

                    MsgBox("Excel file exported with name ScExcel_Pie_Chart.xlsx")

                    Dim average As List(Of Integer) = New List(Of Integer)
                    Dim total As Integer = 0
                    Dim i As Integer
                    Dim j As Integer

                    If Not calcListX Is Nothing Then
                        For i = 0 To calcListX.GetLength(0) - 1 Step 1
                            For j = 0 To calcListX.GetLength(1) - 1 Step 1
                                total += calcListX(i, j)
                            Next
                            average.Add(total / calcListX.GetLength(1))
                            total = 0
                        Next
                    ElseIf Not calcListY Is Nothing Then
                        For i = 0 To calcListY.GetLength(0) - 1 Step 1
                            For j = 0 To calcListY.GetLength(1) - 1 Step 1
                                total += calcListY(i, j)
                            Next
                            average.Add(total / calcListY.GetLength(1))
                            total = 0
                        Next
                    End If

                    Dim xValues As Integer() = New Integer(average.Count - 1) {}
                    Dim yValues As Integer() = New Integer(average.Count - 1) {}

                    i = 0

                    For Each item As Integer In average
                        yValues(i) = item
                        'Console.WriteLine(item)
                        i += 1
                    Next

                    i = 0
                    Dim a As Integer = 0

                    For Each item As Integer In average
                        a += 10
                        xValues(i) = a
                        i += 1
                    Next

                    For Each item In xValues
                        Console.WriteLine(item)
                    Next

                    myChart = xlWorkSheet.ChartObjects.Add(95, 50, 500, 200)
                    sr = myChart.Chart.SeriesCollection.NewSeries
                    sr.XValues = xValues
                    sr.Values = yValues

                    With myChart
                        .Chart.ChartType = Excel.XlChartType.xlPie

                        .Chart.ApplyLayout(4)
                        .Chart.ChartStyle = 2
                    End With

                    xlWorkSheet.SaveAs("C:\Users\ANGELOS\Desktop\ScExcel_Pie_Chart.xlsx")
                Catch ex As Exception
                    MsgBox("Something went wrong while exporting Pie graph.")
                End Try

                xlWorkBook.Close()
                xlApp.Quit()
                releaseObject(xlApp)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkSheet)
        End Select
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Dim allValuesOfCells As List(Of Integer) = getValuesOfCells()

        Dim addValue As Integer = 0
        Dim subValue As Integer = 0
        Dim mulValue As Long = 0
        Dim divValue As Integer = 0
        Dim avgValue As Integer = 0
        Dim devValue As Integer = 0
        Dim fluValue As Integer = 0

        Dim counter As Integer = 0

        If Not allValuesOfCells Is Nothing Then

            subValue = allValuesOfCells(0)
            mulValue = allValuesOfCells(0)
            divValue = allValuesOfCells(0)

            For Each item In allValuesOfCells
                counter += 1
                addValue += item
            Next

            Dim i As Integer = 0

            For i = 1 To allValuesOfCells.Count - 1 Step 1
                subValue -= allValuesOfCells(i)
                mulValue *= allValuesOfCells(i)
                divValue /= allValuesOfCells(i)
            Next

            avgValue = addValue / allValuesOfCells.Count

            Dim newAllValuesOfCells As List(Of Integer) = New List(Of Integer)
            Dim newAllValuesOfCells_1 As List(Of Integer) = New List(Of Integer)

            For Each item In allValuesOfCells
                newAllValuesOfCells.Add((item - avgValue) ^ 2)
                newAllValuesOfCells_1.Add((item - avgValue) ^ 2)
            Next

            For Each item In newAllValuesOfCells
                devValue += item
            Next

            For Each item In newAllValuesOfCells_1
                fluValue += item
            Next

            devValue = Math.Sqrt(devValue / newAllValuesOfCells.Count)
            fluValue = fluValue / newAllValuesOfCells_1.Count

            Select Case ComboBox5.SelectedItem.ToString()
                Case "Add"
                    MsgBox(addValue)
                Case "Substract"
                    MsgBox(subValue)
                Case "Multiply"
                    MsgBox(mulValue)
                Case "Divide"
                    MsgBox(divValue)
                Case "Calculate Average Value"
                    MsgBox(avgValue)
                Case "Calculate Deviation"
                    MsgBox(devValue)
                Case "Calculate Fluctuation"
                    MsgBox(fluValue)
            End Select
        End If
    End Sub
End Class