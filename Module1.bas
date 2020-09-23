Attribute VB_Name = "Module1"
'***********************************************
'*             This is not my code             *
'*     It's someone's listview order code      *
'*                 Dunno who's!                *
'***********************************************
Option Explicit
Public strSettingFile As String
Public strHeaders As String
Public strServer As String
Public intPort As String
Public intPortProx As String
Public strProxy As String
Public Const sortAlphanumeric = 0
Public Const sortNumeric = 1
Public Const sortDate = 2
Public Const sortAscending = 3
Public Const sortDescending = 4
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function
Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function
Public Sub WriteSet(strKey As String, strValue As String) '
    Call WriteINI("SiteDetective", strKey, strValue, strSettingFile)
End Sub
Public Function ReadSet(strKey As String) As String
    ReadSet = ReadINI("SiteDetective", strKey, strSettingFile)
End Function





Function SortColumn(ByVal ListViewControl As MSComctlLib.ListView, ColumnIndex As Integer, SortType As Integer, SortOrder As Integer) As Boolean
    Dim X As Integer, y As Integer
    On Error GoTo ErrHandler
    


    Select Case SortType
        
        '*** Alphanumeric sort
        Case sortAlphanumeric


        DoSort ListViewControl, SortOrder, ColumnIndex - 1
            
            '*** Numeric Sort
            Case sortNumeric
            Dim strMax As String, strNew As String
            
            'Find the longest (whole) number string
            '     length in the column


            If ColumnIndex > 1 Then


                For X = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1)) <> 0 Then 'ignores 0 length strings


                        If Len(CStr(Int(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1)))) > Len(strMax) Then
                            strMax = CStr(Int(ListViewControl.ListItems(X).SubItems(ColumnIndex - 1)))
                        End If
                    End If
                Next
            Else


                For X = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(X)) <> 0 Then


                        If Len(CStr(Int(ListViewControl.ListItems(X)))) > Len(strMax) Then
                            strMax = CStr(Int(ListViewControl.ListItems(X)))
                        End If
                    End If
                Next
            End If
            
            'hide the control - speeds up the sort
            ListViewControl.Visible = False
            


            If ColumnIndex > 1 Then


                For X = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1)) = 0 Then
                        ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1) = "0" 'make 0 length strings = To "0"
                    ElseIf Len(CStr(Int(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1)))) < Len(strMax) Then
                        'prefix all numbers with 0's as required
                        '
                        strNew = ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1)


                        For y = 1 To Len(strMax) - Len(CStr(Int(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1))))
                            strNew = "0" & strNew
                        Next
                        ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1) = strNew
                    End If
                Next
            Else


                For X = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(X).Text) = 0 Then
                        ListViewControl.ListItems(X).Text = "0" 'make 0 length strings = To "0"
                    ElseIf Len(CStr(Int(ListViewControl.ListItems(X)))) < Len(strMax) Then
                        'prefix all numbers with 0's as required
                        '
                        strNew = ListViewControl.ListItems(X).Text


                        For y = 1 To Len(strMax) - Len(CStr(Int(ListViewControl.ListItems(X))))
                            strNew = "0" & strNew
                        Next
                        ListViewControl.ListItems(X).Text = strNew
                    End If
                Next
            End If
            


            DoSort ListViewControl, SortOrder, ColumnIndex - 1
                


                If ColumnIndex > 1 Then
                    'Remove preceding 0's


                    For X = 1 To ListViewControl.ListItems.Count
                        ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1) = CDbl(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1))
                        If ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1) = 0 Then ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1) = ""
                    Next
                Else
                    'Remove preceding 0's


                    For X = 1 To ListViewControl.ListItems.Count
                        ListViewControl.ListItems(X).Text = CDbl(ListViewControl.ListItems(X).Text)
                        If ListViewControl.ListItems(X).Text = 0 Then ListViewControl.ListItems(X).Text = ""
                    Next
                End If
                ListViewControl.Visible = True
                
                '*** Date Sort
                Case sortDate
                ListViewControl.Visible = False


                If ColumnIndex > 1 Then
                    'Convert dates to format that can be sor
                    '     ted alphanumerically


                    For X = 1 To ListViewControl.ListItems.Count
                        ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1), "YYYY MM DD hh:mm:ss")
                    Next


                    DoSort ListViewControl, SortOrder, ColumnIndex - 1
                        'Convert dates back to General Date form
                        '     at


                        For X = 1 To ListViewControl.ListItems.Count
                            ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(X).ListSubItems(ColumnIndex - 1), "General Date")
                        Next
                    Else
                        'Convert dates to format that can be sor
                        '     ted alphanumerically


                        For X = 1 To ListViewControl.ListItems.Count
                            ListViewControl.ListItems(X).Text = Format(ListViewControl.ListItems(X).Text, "YYYY MM DD hh:mm:ss")
                        Next


                        DoSort ListViewControl, SortOrder, ColumnIndex - 1
                            'Convert dates back to General Date form
                            '     at


                            For X = 1 To ListViewControl.ListItems.Count
                                ListViewControl.ListItems(X).Text = Format(ListViewControl.ListItems(X).Text, "General Date")
                            Next
                            
                        End If
                        
                        ListViewControl.Visible = True
                    End Select
                SortColumn = True
                
Exit_Function:
                Exit Function
                
ErrHandler:
                MsgBox Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "ListView Sort module Error"
                SortColumn = False
                Resume Exit_Function
            End Function


Private Sub DoSort(ByVal ListViewControl As MSComctlLib.ListView, SortOrder As Integer, SortKey As Integer)


    If SortOrder = sortAscending Then
        ListViewControl.SortOrder = lvwAscending
    ElseIf SortOrder = sortDescending Then
        ListViewControl.SortOrder = lvwDescending
    End If
    ListViewControl.SortKey = SortKey
    ListViewControl.Sorted = True
End Sub


