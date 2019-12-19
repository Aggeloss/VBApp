Imports Microsoft.VisualBasic

Public Class DataOfStudents

    Private variables As New List(Of String)
    Private dataOfParticipants As New List(Of String)
    Private dataOfFile As New List(Of Integer)
    Private dataList As New List(Of Integer)
    Private allData As New List(Of String)

    Public Sub calculateValues(list As List(Of String))
        Dim flag As Boolean = False
        Dim i As Integer = 0

        Dim a As Integer = 0

        Dim firstTime As Boolean = True

        For Each item In list
            If item.Trim() <> "" Then
                If (flag = False And firstTime = True And item <> "/") Then
                    variables.Add(item)
                ElseIf (flag = False And item <> "|") Then
                    dataOfParticipants.Add(item)
                    'timeOfExamList.Add(Convert.ToInt32(item))
                    'ElseIf (flag = False And i = 1) Then
                    'ageList.Add(Convert.ToInt32(item))
                    'ElseIf (flag = False And i = 2) Then
                    'sexList.Add(item)
                Else
                    If (item = "|") Then
                        flag = True
                    End If

                    If (flag = True And item <> "|" And item <> "/") Then
                        'Console.WriteLine(item)
                        'Console.WriteLine(item)
                        dataList.Add(Convert.ToInt32(item))
                    End If
                End If
            End If
            If (item = "/") Then
                firstTime = False
            End If
            If (item = "/") Then
                flag = False
            End If
            If (firstTime = False And item <> "/" And item <> "|") Then
                allData.Add(item)
            End If
        Next item

        Me.setDataOfParticipants(dataOfParticipants)
        'Me.setDataOfFile(dataOfFile)
        Me.setVariables(variables)
        Me.setDataList(dataList)
    End Sub

    Public Sub setDataOfParticipants(dop As List(Of String))
        dataOfParticipants = dop
    End Sub

    Public Function getDataOfParticipants() As List(Of String)
        Return dataOfParticipants
    End Function

    Public Sub setDataOfFile(dof As List(Of Integer))
        dataOfFile = dof
    End Sub

    Public Function getDataOfFile() As List(Of Integer)
        Return dataOfFile
    End Function

    Public Sub setVariables(var As List(Of String))
        variables = var
    End Sub

    Public Function getVariables() As List(Of String)
        Return variables
    End Function

    Public Sub setDataList(dl As List(Of Integer))
        dataList = dl
    End Sub

    Public Function getDataList() As List(Of Integer)
        Return dataList
    End Function

    Public Sub setAllData(ad As List(Of String))
        allData = ad
    End Sub

    Public Function getAllData() As List(Of String)
        Return allData
    End Function
End Class
