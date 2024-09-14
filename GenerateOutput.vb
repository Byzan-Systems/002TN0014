Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System
Imports System.Data
Imports Microsoft.Office.Interop

Module GenrateOutput

    Dim objLogCls As New ClsErrLog
    Dim objGetSetINI As ClsShared
    Dim objBaseClass As ClsBase
    Dim objValidationClass As ClsValidation
    Dim SumOfAmount As Double = 0
   
    Public Function GenerateOutPutFile(ByRef dtEpay As DataTable, ByRef dtAdvice As DataTable, ByRef dtEpayUnsucess As DataTable, ByVal strFileName As String) As Boolean
        'Dim gstrA2Afile As String = String.Empty
        'Dim strMethodCalForEpay As Boolean = False
        'Try
        '    objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
        '    objValidationClass = New ClsValidation(strFileName, objBaseClass.gstrIniPath)
        '    FileCounter = objBaseClass.GetINISettings("General", "File Counter", My.Application.Info.DirectoryPath & "\settings.ini")
        '    FileCounter = FileCounter + 1

        '    If Len(FileCounter) < 3 Then
        '        FileCounter = FileCounter.PadLeft(4, "0").Trim()
        '        FileCounter = FileCounter.Substring(FileCounter.Length - 3, 3)
        '    End If

        '    strFileName = (objValidationClass.IsJustAlpha(Path.GetFileNameWithoutExtension(gstrInputFile), 10, "N"))


        '    gstrOutputFile_EPAY = strFileName & "_" & DateTime.Now.ToString("ddMMyyyyHHmmss") & ".xls "
        '    gstrOutputFile_EpayText = Path.GetFileNameWithoutExtension(gstrOutputFile_EPAY) & ".txt"
        '    gstrOutputFile_ADV = Path.GetFileNameWithoutExtension(gstrOutputFile_EPAY) & "_ADV.txt"

        '    'If dtEpayUnsucess.Rows.Count = 0 Then

        '    If strEpayOptFile_Format.ToString().Trim().ToUpper() = "TXT" Then
        '        If Generate_OutPut_Epay_Text(dtEpay, strFileName) = False And Generate_Output_ADV(dtAdvice, strFileName) = False Then
        '            GenerateOutPutFile = False
        '        Else
        '            GenerateOutPutFile = True
        '            Call objBaseClass.SetINISettings("General", "File Counter", Val(FileCounter), My.Application.Info.DirectoryPath & "\settings.ini")
        '        End If
        '    Else
        '        If Generate_Output_EPAY(dtEpay, strFileName) = False And Generate_Output_ADV(dtAdvice, strFileName) = False Then
        '            GenerateOutPutFile = False
        '        Else
        '            GenerateOutPutFile = True

        '            Call objBaseClass.SetINISettings("General", "File Counter", Val(FileCounter), My.Application.Info.DirectoryPath & "\settings.ini")
        '        End If
        '    End If
        '    ' End If
        'Catch ex As Exception
        '    GenerateOutPutFile = False
        '    objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "GenerateOutPutFile")
        'End Try


    End Function

    Public Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then

                ' Check_Comma = Chr(34) & strTemp & Chr(34) & ","
                Check_Comma = strTemp
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Payment", "Check_Comma")

        End Try
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function

    Function RemoveCharacter(ByVal stringToCleanUp As String)
        Dim characterToRemove As String = ""
        characterToRemove = Chr(34) + "=~^!#$%&'()*+,-@`/\:{}[]"

        Dim firstThree As Char() = characterToRemove.Take(30).ToArray()
        For index = 0 To firstThree.Length - 1
            stringToCleanUp = stringToCleanUp.ToString.Replace(firstThree(index), "")
        Next
        Return stringToCleanUp
    End Function
End Module
