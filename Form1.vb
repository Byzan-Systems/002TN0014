Imports System.IO
Imports System.Data
Imports System.Security.Cryptography
Imports System.Text
Public Class Form1
    Dim objBaseClass As ClsBase
    Dim objFileValidate As ClsValidation
    Dim objGetSetINI As ClsShared
    Dim StrEncrpt As String = String.Empty

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try

            Timer1.Interval = 1000
            Timer1.Enabled = False

            Conversion_Process()

            Timer1.Enabled = True

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form Load", "Timer1_Tick")
        End Try
    End Sub
    Private Sub Generate_SettingFile()

        Dim strConverterCaption As String = ""
        Dim strSettingsFilePath As String = My.Application.Info.DirectoryPath & "\settings.ini"

        Try
            objGetSetINI = New ClsShared

            '-Genereate Settings.ini File-
            If Not File.Exists(strSettingsFilePath) Then

                '-General Section-
                Call objGetSetINI.SetINISettings("General", "Date", Format(Now, "dd/MM/yyyy"), strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Audit Log", My.Application.Info.DirectoryPath & "\Audit", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Error Log", My.Application.Info.DirectoryPath & "\Error", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Input Folder", My.Application.Info.DirectoryPath & "\INPUT", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Output Folder", My.Application.Info.DirectoryPath & "\Output", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Response Folder", My.Application.Info.DirectoryPath & "\Response", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Res_Output Folder", My.Application.Info.DirectoryPath & "\Res_Output", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "Archived FolderSuc", My.Application.Info.DirectoryPath & "\Archive\Success", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderUnSuc", My.Application.Info.DirectoryPath & "\Archive\UnSuccess", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "Master", My.Application.Info.DirectoryPath & "\Master\Client_Master.csv", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "Converter Caption", "DecryptionFiles_Application", strSettingsFilePath)
                '  Call objGetSetINI.SetINISettings("General", "Process Output File Ignoring Invalid Transactions", "N", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "File Counter", "0", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "==", "==========================================", strSettingsFilePath) 'Separator

                Call objGetSetINI.SetINISettings("General", "Decrypt Key", My.Application.Info.DirectoryPath & "\Yes.key", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "==", "==========================================", strSettingsFilePath) 'Separator

                '-Client Details Section-
                Call objGetSetINI.SetINISettings("Client Details", "Input Date Format", "dd/MM/yyyy", strSettingsFilePath)
                '   Call objGetSetINI.SetINISettings("Client Details", "Utility Code", "NACH00000000022441", strSettingsFilePath)     '''''''''Commented by swati dtd 2021-04-19
                '  Call objGetSetINI.SetINISettings("Client Details", "Client short name", "HAFEDH", strSettingsFilePath)    '''''''''Commented by swati dtd 2021-04-19
                Call objGetSetINI.SetINISettings("Client Details", "YBL Nomenclature Format", "NACH_CR", strSettingsFilePath)

                'Response
                Call objGetSetINI.SetINISettings("Client Details", "Bank Scheme code", "YESB", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Bank User Name", "EKHARID", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Response File Nomenclature Format", "ACH-CR", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("Client Details", "==", "====================================", strSettingsFilePath) 'Separator
            End If

            '-Get Converter Caption from Settings-
            If File.Exists(strSettingsFilePath) Then
                strConverterCaption = objGetSetINI.GetINISettings("General", "Converter Caption", strSettingsFilePath)
                If strConverterCaption <> "" Then
                    Text = strConverterCaption.ToString() & " - Version " & Mid(Application.ProductVersion.ToString(), 1, 3)
                Else
                    MsgBox("Either settings.ini file does not contains the key as [ Converter Caption ] or the key value is blank" & vbCrLf & "Please refer to " & strSettingsFilePath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End
                End If
            End If

        Catch ex As Exception
            MsgBox("Error" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error while Generating Settings File")
            End

        Finally
            If Not objGetSetINI Is Nothing Then
                objGetSetINI.Dispose()
                objGetSetINI = Nothing
            End If

        End Try

    End Sub
    Private Sub Conversion_Process()
        Dim objfolderAll As DirectoryInfo
        Try
            If objBaseClass Is Nothing Then
                objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            End If

            '-Get Settings-
            If GetAllSettings() = True Then
                MsgBox("Either file path is invalid or any key value is left blank in settings.ini file", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error In Settings")
                Exit Sub
            End If


            '-Process Input-
            objfolderAll = New DirectoryInfo(strInputFolderPath)
            If objfolderAll.GetFiles.Length = 0 Then
                objfolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started for INPUT Files")

                For Each file As FileInfo In objfolderAll.GetFiles("*")
                    objBaseClass.isCompleteFileAvailable(file.FullName)
                    If Mid(file.FullName, file.FullName.Length - 3, 4).ToString().ToUpper() <> ".txt".ToUpper And Mid(file.FullName, file.FullName.Length - 3, 4).ToString().ToUpper() <> ".BAK" Then
                        objBaseClass.LogEntry("Invalid File Format", False)
                    Else
                        objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("INPUT File [ " & file.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)
                        '   strFileType = "IMP"
                        Process_Each(file.FullName, "INPUT")

                        objfolderAll.Refresh()
                    End If
                Next
            End If

            ' For Response 

            objfolderAll = Nothing

            objfolderAll = New DirectoryInfo(strResponseFolderPath)

            If objfolderAll.GetFiles.Length = 0 Then
                objfolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started for RESPONSE Files")

                For Each objFileOne As FileInfo In objfolderAll.GetFiles()
                    objBaseClass.isCompleteFileAvailable(objFileOne.FullName)
                    If Mid(objFileOne.FullName, objFileOne.FullName.Length - 3, 4).ToString().ToUpper() = ".txt".ToUpper And Mid(objFileOne.FullName, objFileOne.FullName.Length - 3, 4).ToString().ToUpper() <> ".BAK" Then
                        objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("RESPONSE File [ " & objFileOne.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)
                        '  strFileType = "RES"
                        Process_Each(objFileOne.FullName, "RES")

                        objfolderAll.Refresh()

                    End If
                Next
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Conversion_Process")

        Finally
            If Not objBaseClass Is Nothing Then
                objBaseClass.Dispose()
                objBaseClass = Nothing
            End If
        End Try
    End Sub
    Private Sub Process_Each(ByVal strInputFileName As String, ByVal strFileProcess As String)
        Dim TrnProcSuc As Boolean

        Dim strFileType As String = ""
        Dim strArray() As String = Nothing
        Dim strResArray() As String = Nothing
        Dim strFileSeqNo As String = ""
        Dim strOptClientName As String = ""

        Dim strtext As String = ""
        Dim outtexxt As String = ""
        Dim strMsg As String = ""

        Dim OutputFolderPath As String = ""
        Dim FileArrLength As Integer = 0

        Dim strFileSettelmentDate As String = ""
        Dim strResFileNameExtension As String = ""

        ''''Payment file''''''''''''
        '   Dim strUtilityCode As String = ""

        Dim strUtilityCode As String = ""
        Dim strClientShortName As String = ""
        Dim DrMaster() As DataRow = Nothing

        Dim file As System.IO.StreamWriter = Nothing

        Try

            gstrInputFolder = strInputFileName.Substring(0, strInputFileName.LastIndexOf("\"))
            gstrInputFile = strInputFileName.Substring(strInputFileName.LastIndexOf("\"))
            gstrInputFile = Path.GetFileName(gstrInputFile).ToUpper.Trim()

            '-Conversion Process-

            objBaseClass.LogEntry("", False)
            objBaseClass.LogEntry("Process Started")

            DtMaster = objBaseClass.MyGetDatatable_Text(strMasterPath, ",")      '''''''''Added by swati dtd 2021-04-19

            'If Path.GetFileNameWithoutExtension(gstrInputFile).Substring(Path.GetFileNameWithoutExtension(gstrInputFile).Length - 4, 4) = "-RES" Then '" Then
            If gstrInputFile.Contains("-RES") Then '" Then
                strArray = Path.GetFileNameWithoutExtension(gstrInputFile).ToString().Trim().ToUpper.Split("_")
                FileArrLength = 6
                strFileType = "RES"
                strMsg = "RES"
            Else
                strArray = Path.GetFileNameWithoutExtension(gstrInputFile).ToString().Trim().ToUpper.Split("-")
                FileArrLength = 7
                strFileType = "Input"
                strMsg = "Input"
            End If
            If strFileProcess.ToUpper <> strFileType.ToUpper Then
                objBaseClass.LogEntry("Input File put in wrong folder path [ " & strInputFileName & "]")
                Exit Sub
            End If

            If strArray.Length > 0 Then

                If strArray.Length = FileArrLength Then

                    If strFileType = "RES" Then
                        strFileSeqNo = strArray(5)
                        'strResArray = strArray(5).Split("-")

                        '  strFileSeqNo = strArray(5).Substring(0, strArray(5).LastIndex("-"))
                        'If strResArray.Length > 0 Then
                        '    strFileSeqNo = strResArray(0)

                        '    For i = 0 To strResArray.Length - 1
                        '        strResFileNameExtension = strResFileNameExtension & "-" & strResArray(i)
                        '    Next
                        'End If

                        OutputFolderPath = strRes_OutputFolderPath
                        strFileSettelmentDate = strArray(2)
                        '  gstrOutputFile = strResFile_Nomenclature_Format & "-" & strBank & "-" & strBankUserName & "-" & DateTime.Now.ToString("ddMMyyyy") & "-" & strFileSeqNo & ".txt" ' & "-RES" & ".txt"
                        gstrOutputFile = strResFile_Nomenclature_Format & "-" & strBank & "-" & strBankUserName & "-" & strFileSettelmentDate & "-" & strFileSeqNo & ".txt" ' & "-RES" & ".txt"
                    Else
                        strFileSeqNo = strArray(5)
                        strFileSettelmentDate = strArray(4)

                        '''''''''Commented by swati dtd 2021-04-19
                        'If strClientShortName.Length >= 6 Then
                        '    strOptClientName = strClientShortName.Substring(0, 6)
                        'Else
                        '    strOptClientName = strClientShortName
                        'End If
                        OutputFolderPath = strOutputFolderPath
                        'gstrOutputFile = strYBL_Nomenclature_Format & "_" & strFileSettelmentDate & "_" & strUtilityCode & "_" & strOptClientName & "_" & strFileSeqNo & ".txt"     '''''''''Commented by swati dtd 2021-04-19
                    End If
                Else
                    TrnProcSuc = False
                    objBaseClass.LogEntry("Invalid Input File Nomenclature " & gstrInputFile, False)
                    objBaseClass.LogEntry("Please check Audit and Error Log file.Process Terminated", False)
                    objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                    objBaseClass.LogEntry("Input file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)
                    Exit Sub
                End If
            Else
                TrnProcSuc = False
                objBaseClass.LogEntry("Invalid Input File")
                objBaseClass.LogEntry("Please check Audit and Error Log file.Process Terminated", False)
                objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                objBaseClass.LogEntry("Input file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)
                Exit Sub
            End If

            objBaseClass.LogEntry("Reading " & strMsg & " File " & gstrInputFile, False)
            ''''''''File Reading
            System.Windows.Forms.Application.DoEvents()

            If strFileType = "Input" Then
                Dim fileReader As String
                fileReader = My.Computer.FileSystem.ReadAllText(gstrInputFolder & "\" & gstrInputFile)
                strtext = fileReader
            End If

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If strFileType = "RES" Then
                ''''''''''reading respons file

                strtext = ResponseFile_ReadProc(gstrInputFolder & "\" & gstrInputFile)
                outtexxt = strtext
            Else
                outtexxt = Decrypt(strtext, strDecryptKey)
            End If
            'outtexxt = Decrypt(strtext, strDecryptKey)
            If outtexxt <> Nothing Then
                If strFileType = "Input" Then
                    strUtilityCode = outtexxt.Substring(156, 18).ToString.Trim().ToUpper  '''''Utility code 
                    DrMaster = DtMaster.Select("[Column_1] ='" & strUtilityCode & "'")  '''''Added 2021-04-10
                    If DrMaster.Length > 0 Then
                        strClientShortName = DrMaster(0)("Column_2").ToString().Trim()

                        gstrOutputFile = strYBL_Nomenclature_Format & "_" & strFileSettelmentDate & "_" & strUtilityCode & "_" & strClientShortName & "_" & strFileSeqNo & ".txt"
                    Else
                        TrnProcSuc = False
                        objBaseClass.LogEntry("The Utility code [" & strUtilityCode & "] in input file [" & gstrInputFile & "] does not match with master file [" & Path.GetFileName(strMasterPath) & " ] utility code .")
                        objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                        objBaseClass.LogEntry("Input file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)
                        Exit Sub
                    End If
                End If

                objBaseClass.LogEntry("Output File Generation Process Started")
                If System.IO.File.Exists(OutputFolderPath & "\" & gstrOutputFile) = True Then
                    System.IO.File.Delete(OutputFolderPath & "\" & gstrOutputFile)
                End If

                '   Dim file As System.IO.StreamWriter
                'file = My.Computer.FileSystem.OpenTextFileWriter(OutputFolderPath & "\" & gstrOutputFile)''''Commented by swati dtd 2021-09-29
                File = My.Computer.FileSystem.OpenTextFileWriter(OutputFolderPath & "\" & gstrOutputFile, True, System.Text.Encoding.GetEncoding(28597)) ''''Added by swati dtd 2021-09-29

                '   file.WriteLine(outtexxt) ''''Commented by swati dtd 2021-09-29

                File.Write(outtexxt) ''''Added by swati dtd 2021-09-29
                File.Close()

                TrnProcSuc = True
                If strFileType = "Input" Then
                    objBaseClass.LogEntry(strMsg & " Output Files is Decrpted Successfully [" & gstrOutputFile & "]", False)
                Else
                    objBaseClass.LogEntry(strMsg & " Output Files is Generated Successfully [" & gstrOutputFile & "]", False)
                End If

                objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderSuc & "\" & gstrInputFile)
                objBaseClass.LogEntry(strMsg & " file [" + Path.GetFileName(strInputFileName) + "] Is Moved to " + strArchivedFolderSuc)

            Else
                TrnProcSuc = False
                If strFileType = "Input" Then
                    objBaseClass.LogEntry(strMsg & "Output File Decrption process failed due to Error", True)
                Else
                    objBaseClass.LogEntry(strMsg & "Output File Generation process failed due to Error", True)
                End If

                objBaseClass.LogEntry("Please check Audit and Error Log file.Process Terminated", False)
                objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                objBaseClass.LogEntry(strMsg & " file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)
                '   Exit Sub
            End If

            If TrnProcSuc <> False Then
                objBaseClass.LogEntry("Process Completed Successfully", False)
                objBaseClass.LogEntry("-------------------------------------------------------------------------------------", False)

            Else
                objBaseClass.LogEntry("Process Terminated", False)
                objBaseClass.LogEntry("-------------------------------------------------------------------------------------", False)
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "DecryptionFiles_Application", "Process_Each")

        Finally

            If Not objFileValidate Is Nothing Then
                objBaseClass.ObjectDispose(objFileValidate.DtInput)
                objBaseClass.ObjectDispose(objFileValidate.DtUnSucInput)

                objFileValidate.Dispose()
                objFileValidate = Nothing
            End If


            If Not File Is Nothing Then
                File.Close()
                File.Dispose()

            End If
        End Try
    End Sub
    Public Function ResponseFile_ReadProc(ByVal FilePath As String) As String
        Dim strTempLine As String
        Dim strTempLineModified As String = ""
        Dim objStrmReader As StreamReader = Nothing

        Dim objStrmWriter As StreamWriter = Nothing
        Dim strOutPutLine As String = ""
        Dim strLineNo As Integer = 0

        Dim strReplaceNewValue As String = ""
        Dim strOldValue As String = ""
        Try
            '239,250

            objStrmReader = New StreamReader(FilePath)
            While Not objStrmReader.EndOfStream

                strTempLine = objStrmReader.ReadLine()
                strReplaceNewValue = ""
                strLineNo += 1
                strTempLineModified = ""
                strOldValue = ""

                'If objStrmReader.EndOfStream = True Then
                '    If strTempLine = "" Then
                '        Exit While
                '    End If
                'End If
                If strTempLine <> "" Then

                    If strLineNo = 1 Then

                        strOldValue = strTempLine.Substring(239 - 1, 11).ToString().Trim()  ''Start position 239

                        If strOldValue.Length < 11 Then
                            strReplaceNewValue = Pad_Length(strOldValue.PadLeft(11, "0"), 11)
                            strTempLineModified = ReplaceAt(strTempLine, 239 - 1, 11, strReplaceNewValue)
                        Else
                            strTempLineModified = strTempLine
                        End If
                    Else

                        strOldValue = strTempLine.Substring(265 - 1, 15).ToString().Trim() ''Start position 265
                        If strOldValue = "" Then
                            strReplaceNewValue = Pad_Length(strOldValue.PadLeft(15, "0"), 15)
                            strTempLineModified = ReplaceAt(strTempLine, 265 - 1, 15, strReplaceNewValue)
                        Else
                            strTempLineModified = strTempLine
                        End If
                    End If

                    If objStrmReader.EndOfStream = True Then
                        strOutPutLine = strOutPutLine & strTempLineModified ' & vbNewLine
                    Else
                        strOutPutLine = strOutPutLine & strTempLineModified & vbLf '& vbNewLine
                    End If
                End If
            End While

            Return strOutPutLine
        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "ResponseFile_ReadProc")
        Finally
            If Not objStrmReader Is Nothing Then
                objStrmReader.Close()
                objStrmReader.Dispose()

            End If

            If Not objStrmWriter Is Nothing Then
                objStrmWriter.Close()
                objStrmWriter.Dispose()
            End If
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
    Public Function ReplaceAt(ByVal str As String, ByVal index As Integer, ByVal length As Integer, ByVal replace As String) As String


        Return str.Remove(index, Math.Min(length, str.Length - index)).Insert(index, replace)


    End Function


    Public Function Decrypt(ByVal textToDecrypt As String, ByVal FilePath As String) As String
        Try
            Dim rijndaelCipher As RijndaelManaged = New RijndaelManaged()
            rijndaelCipher.Mode = CipherMode.CBC
            rijndaelCipher.Padding = PaddingMode.PKCS7
            rijndaelCipher.KeySize = &H80
            rijndaelCipher.BlockSize = &H80

            Dim encryptedData As Byte() = Convert.FromBase64String(textToDecrypt)
            Dim pwdBytes As Byte() = GetFileBytes(FilePath)
            Dim keyBytes As Byte() = New Byte(15) {}
            Dim len As Integer = pwdBytes.Length

            If len > keyBytes.Length Then
                len = keyBytes.Length
            End If

            Array.Copy(pwdBytes, keyBytes, len)
            rijndaelCipher.Key = keyBytes
            rijndaelCipher.IV = keyBytes
            Dim plainText As Byte() = rijndaelCipher.CreateDecryptor().TransformFinalBlock(encryptedData, 0, encryptedData.Length)
            Return Encoding.UTF8.GetString(plainText)
        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Decrypt")
        Finally

        End Try

    End Function

    Public Function GetFileBytes(ByVal filePath As String) As Byte()
        Dim buffer As Byte()
        Dim fileStream As FileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read)

        Try
            Dim length As Integer = CInt(fileStream.Length)
            buffer = New Byte(length - 1) {}
            Dim count As Integer
            Dim sum As Integer = 0

            While (CSharpImpl.__Assign(count, fileStream.Read(buffer, sum, length - sum))) > 0
                sum += count
            End While

        Finally
            fileStream.Close()
        End Try

        Return buffer
    End Function

    Private Class CSharpImpl
        <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
        Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Timer1.Interval = 100
            Timer1.Enabled = True

            Generate_SettingFile()

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form Load", "form1_Load")
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub
    Private Function GetAllSettings() As Boolean
        Try
            GetAllSettings = False

            If Not File.Exists(My.Application.Info.DirectoryPath & "\settings.ini") Then
                GetAllSettings = True
                MsgBox("Either settings.ini file does not exists or invalid file path" & vbCrLf & My.Application.Info.DirectoryPath & "\settings.ini", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            End If

            '-Audit Folder Path-
            If strAuditFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strAuditFolderPath) Then
                    Directory.CreateDirectory(strAuditFolderPath)
                    If Not Directory.Exists(strAuditFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Audit Log folder. Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", True)
                        End If
                        MsgBox("Invalid path for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                        Exit Function
                    End If
                End If
            End If

            '-Error Folder Path-
            If strErrorFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Error Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strErrorFolderPath) Then
                    Directory.CreateDirectory(strErrorFolderPath)
                    If Not Directory.Exists(strErrorFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Error Log folder. Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Error Log folder." & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Input Folder Path-
            If strInputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Input Folder " & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strInputFolderPath) Then
                    Directory.CreateDirectory(strInputFolderPath)
                    If Not Directory.Exists(strInputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Input Folder. Please check settings.ini file, the key as [ Input Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Input Folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "settings Error")
                    End If
                End If
            End If

            '-Archived Success Path-
            If strArchivedFolderSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archived Success folder" & vbCrLf & "Please check settings.ini file, the key as [ Archived Success Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderSuc) Then
                    Directory.CreateDirectory(strArchivedFolderSuc)
                    If Not Directory.Exists(strArchivedFolderSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archived Success Please check [ settings.ini ] file, the key as [ Archived Success Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archived Success Folder." & vbCrLf & "Please check settings.ini file, the key as [ Archived Success Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Archived Unsuccess Path-
            If strArchivedFolderUnSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archived Unsuccess folder" & vbCrLf & "Please check settings.ini file, the key as [ Archived Unsuccess Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderUnSuc) Then
                    Directory.CreateDirectory(strArchivedFolderUnSuc)
                    If Not Directory.Exists(strArchivedFolderUnSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archived Unsuccess Folder. Please check [ settings.ini ] file, the key as [ Archived Unsuccess Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archived Unsuccess Folder." & vbCrLf & "Please check settings.ini file, the key as [ Archived Unsuccess Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Output Folder Path-
            If strOutputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Output Folder" & vbCrLf & "Please check settings.ini file, the key as [ Output  Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strOutputFolderPath) Then
                    Directory.CreateDirectory(strOutputFolderPath)
                    If Not Directory.Exists(strOutputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for  Output Folder. Please check [ settings.ini ] file, the key as [  Output Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for  Output Folder." & vbCrLf & "Please check settings.ini file, the key as [  Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If


            '-Response Folder Path-
            If strResponseFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Response folder" & vbCrLf & "Please check settings.ini file, the key as [ Response Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strResponseFolderPath) Then
                    Directory.CreateDirectory(strResponseFolderPath)
                    If Not Directory.Exists(strResponseFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Output Folder. Please check [ settings.ini ] file, the key as [ Response Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Response Folder." & vbCrLf & "Please check settings.ini file, the key as [ Response Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Output Response Folder Path-
            If strRes_OutputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Response Output folder" & vbCrLf & "Please check settings.ini file, the key as [  Response Output Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strRes_OutputFolderPath) Then
                    Directory.CreateDirectory(strRes_OutputFolderPath)
                    If Not Directory.Exists(strRes_OutputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Response Output Folder. Please check [ settings.ini ] file, the key as [  Response Output Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for  Response Output Folder." & vbCrLf & "Please check settings.ini file, the key as [ Response Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

        Catch ex As Exception
            GetAllSettings = True
            'MsgBox("Error - " & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error While Getting Log Path from Settings.ini File")
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "GetAllSettings")

        End Try

    End Function


End Class
