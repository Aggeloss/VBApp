Public Class Form2

    Shared valuesOfColumns As List(Of Integer) = New List(Of Integer)
    Shared indexOfRemovedRow As List(Of Integer) = New List(Of Integer)

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        valuesOfColumns = New List(Of Integer)
        indexOfRemovedRow = New List(Of Integer)

        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            indexOfRemovedRow.Add(row.Index)
            DataGridView1.Rows.Remove(row)
            'For Each item As DataGridViewCell In row.Cells
            'If IsNumeric(item) Then
            'valuesOfColumns.Add(item.Value)
            'Else
            'Continue For
            'End If
            'Next
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For Each row As DataGridViewRow In DataGridView1.Rows
            For Each cell As DataGridViewCell In row.Cells
                If IsNumeric(cell.Value) Then
                    valuesOfColumns.Add(cell.Value)
                Else
                    Continue For
                End If
            Next
        Next
        Me.Close()
    End Sub

    Public Shared Function getValuesOfColumns() As List(Of Integer)
        Return valuesOfColumns
    End Function

    Public Shared Function getIndexOfRemovedRow() As List(Of Integer)
        Return indexOfRemovedRow
    End Function
End Class