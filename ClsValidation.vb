Imports System
Imports System.IO

Public Class ClsValidation

    Implements IDisposable

    Private ObjBaseClass As ClsBase         ''need to be dispose 

    Private DtValidation As DataTable       ''need to be dispose
    Private DtHaderValidation As DataTable
    Private DtSpCharValidation As DataTable       ''need to be dispose
    Private DtInputMapping As DataTable
    Private DtOutPutMapping As DataTable
    ''---

    Public DtInput As DataTable             ''need to be dispose
    Public DtInputHeader As DataTable
    Public DtUnSucInput As DataTable        ''need to be dispose
    Private DtTemp As DataTable             ''need to be dispose
    Private DtTempHeader As DataTable


    Public DtInputResp As DataTable                     ''need to be dispose
    Public DtUnSucResp As DataTable                ''need to be dispose

    Private StrFilePath As String
    'Private ValidationPath As String
    Private ValidationPath_B As String
    Private ValidationPath_P As String
    Private SpCharValidationPath As String
    ''---

    Public StrSettingPath As String

    Public ErrorMessage As String
    Private ValidationPath As String
    Public DtInputAdvice As DataTable
    Public DtUnSucInputAdvice As DataTable
    Public DtUnSucEmail_ADV As DataTable

    Private DtStatus As DataTable       ''need to be dispose
    Public HeaderUploadDate As String       ''need to be dispose


    Public Sub New(ByVal _strFilePath As String, ByVal _SettINIPath As String)

        StrFilePath = _strFilePath
        StrSettingPath = _SettINIPath

        Try
            ObjBaseClass = New ClsBase(_SettINIPath)

            ValidationPath = ObjBaseClass.GetINISettings("General", "Validation", _SettINIPath)
            DtInput = New DataTable("DtInput")
            '  DefineColumn(DtInput)
            DtUnSucInput = New DataTable("DtInput")
            '  DefineColumn(DtUnSucInput)

            'DtInputAdvice = New DataTable("RBIInput")
            'DefineColumnForAdvice(DtInputAdvice)
            'DtUnSucInputAdvice = New DataTable("RBIUnSucInput")
            'DefineColumnForAdvice(DtUnSucInputAdvice)

            'DtInputResp = New DataTable("InputRes")
            'DefineColumnForRevResponse(DtInputResp)
            'DtUnSucResp = New DataTable("UnSucResInput")
            'DefineColumnForRevResponse(DtUnSucResp)

        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Constructor")
        End Try

    End Sub

    'Private Sub DefineColumn(ByRef DtInput As DataTable)

    '    DtInput.Columns.Add(New DataColumn("Detail Indicator"))   '0
    '    DtInput.Columns.Add(New DataColumn("Message Type"))  '1

    '    DtInput.Columns.Add(New DataColumn("Debit Account Number"))  '2
    '    DtInput.Columns.Add(New DataColumn("Remitter Name"))  '3
    '    DtInput.Columns.Add(New DataColumn("Address Line 1"))  '4
    '    DtInput.Columns.Add(New DataColumn("Address Line 2"))  '5
    '    DtInput.Columns.Add(New DataColumn("Address Line 3"))  '6

    '    DtInput.Columns.Add(New DataColumn("Beneficiary Bank IFS Code"))  '7
    '    DtInput.Columns.Add(New DataColumn("Beneficiary Bank Account No"))  '8
    '    DtInput.Columns.Add(New DataColumn("Beneficiary Name"))  '9

    '    DtInput.Columns.Add(New DataColumn("Bene Add Line 1"))  '10
    '    DtInput.Columns.Add(New DataColumn("Bene Add Line 2"))  '11
    '    DtInput.Columns.Add(New DataColumn("Bene Add Line 3"))  '12
    '    DtInput.Columns.Add(New DataColumn("Bene Add Line 4"))  '13

    '    DtInput.Columns.Add(New DataColumn("Transaction Ref No"))  '14
    '    DtInput.Columns.Add(New DataColumn("Upload Date"))  '15
    '    DtInput.Columns.Add(New DataColumn("Amount"))  '16 ' 
    '    DtInput.Columns.Add(New DataColumn("Sender To Receiver Info"))  '17

    '    DtInput.Columns.Add(New DataColumn("Add Info 1"))  '18
    '    DtInput.Columns.Add(New DataColumn("Add Info 2"))  '19
    '    DtInput.Columns.Add(New DataColumn("Add Info 3"))  '20
    '    DtInput.Columns.Add(New DataColumn("Add Info 4"))  '21

    '    DtInput.Columns.Add(New DataColumn("TXN_NO"))   ''22  ', System.Type.GetType("System.Int32")
    '    DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   ''23
    '    DtInput.Columns.Add(New DataColumn("Reason"))   ''24

    'End Sub
    'Private Sub DefineColumnForAdvice(ByRef DtInput As DataTable)

    '    DtInput.Columns.Add(New DataColumn("Transaction ref no")) ''0
    '    DtInput.Columns.Add(New DataColumn("Beneficiary Name"))   ''1
    '    DtInput.Columns.Add(New DataColumn("Beneficiary Code"))    ''2
    '    DtInput.Columns.Add(New DataColumn("Bene address"))    ''3
    '    DtInput.Columns.Add(New DataColumn("Bene bank name")) ''4
    '    DtInput.Columns.Add(New DataColumn("Bene Branch Name"))  ''5
    '    DtInput.Columns.Add(New DataColumn("Net Amount"))   ''6
    '    DtInput.Columns.Add(New DataColumn("Payment details1"))   ''7
    '    DtInput.Columns.Add(New DataColumn("Payment details2"))   ''8
    '    DtInput.Columns.Add(New DataColumn("Payment details3"))   ''9
    '    DtInput.Columns.Add(New DataColumn("Payment details4"))   ''10
    '    DtInput.Columns.Add(New DataColumn("Payment details5"))   ''11
    '    DtInput.Columns.Add(New DataColumn("Payment details6"))   ''12
    '    DtInput.Columns.Add(New DataColumn("Payment details7"))   ''13
    '    DtInput.Columns.Add(New DataColumn("Bene email id"))   ''14
    '    ' DtInput.Columns.Add(New DataColumn("Phone"))   ''15

    '    DtInput.Columns.Add(New DataColumn("TXN_NO", System.Type.GetType("System.Int32")))     '15
    '    DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))    '16
    '    DtInput.Columns.Add(New DataColumn("Reason"))    '17
    '    DtInput.Columns.Add(New DataColumn("EmailLog"))    '18
    'End Sub

    Public Function CheckValidateFile(ByVal gstrInputFile As String) As Boolean

        Try
            If Not File.Exists(gstrInputFile) Then
                Call ObjBaseClass.Handle_Error(New ApplicationException("Input file path is incorrect or not file found. [" & StrFilePath & "]"), "ClsValidation", -123, "CheckValidateFile")
                CheckValidateFile = False
                Exit Function
            End If

            '  CheckValidateFile = Validate()


        Catch ex As Exception
            CheckValidateFile = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "CheckValidateFile")
        End Try

    End Function

    Private Function GetInArrayByComma(ByVal pStrValue As String) As String()

        Try

            Dim Tmpstr As String = ""
            Dim Index_S, Index_E, TmpIndex As Integer


            Index_E = InStr(pStrValue, Chr(34))

            If Index_E > 0 Then

                Index_S = 0
                Tmpstr = ""
                While True

                    Index_E = InStr(Index_S + 1, pStrValue, Chr(34))

                    If Index_E > 0 Then

                        Tmpstr += pStrValue.Substring(Index_S, Index_E - Index_S - 1).Replace(",", "|")
                        Index_S = Index_E
                        Index_E = InStr(Index_E + 1, pStrValue, Chr(34))
                        Tmpstr += pStrValue.Substring(Index_S, (Index_E - Index_S) - 1)
                        Index_S = Index_E

                    Else
                        Tmpstr += pStrValue.Substring(Index_S, pStrValue.Length - Index_S).Replace(",", "|")
                        GetInArrayByComma = Tmpstr.Split("|")
                        Exit While
                    End If

                End While

            Else
                GetInArrayByComma = pStrValue.Split(",")

            End If

        Catch ex As Exception

        End Try

    End Function

    Public Function RemoveBlankRow(ByRef _DtTemp As DataTable)
        'To Remove Blank Row Exists in DataTable
        Dim blnRowBlank As Boolean
        Dim delIndexStr As String = ""
        Dim DelIndex() As String
        Try

            For i As Integer = 0 To _DtTemp.Rows.Count - 1
                blnRowBlank = True
                Dim vRow As DataRow = _DtTemp.Rows(i)
                For intCol As Int32 = 0 To _DtTemp.Columns.Count - 1
                    If vRow.Item(intCol).ToString().Trim() <> "" Then
                        blnRowBlank = False
                        Exit For
                    End If
                Next

                If blnRowBlank = True Then
                    'DtTemp1.Rows(i).Delete()
                    delIndexStr = delIndexStr & i & ","
                End If

            Next

            If delIndexStr <> "" Then
                delIndexStr = Left(delIndexStr, delIndexStr.Length - 1)
                DelIndex = delIndexStr.Split(",")
                For j As Integer = 0 To DelIndex.Length - 1

                    If DelIndex(j) <> "" Then

                        If j = 0 Then
                            _DtTemp.Rows(DelIndex(j)).Delete()
                        Else
                            _DtTemp.Rows(DelIndex(j) - j).Delete()
                        End If
                        _DtTemp.AcceptChanges()

                    End If
                Next


            End If

            '------------------End Here

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveBlankRow")

        End Try

    End Function

    Private Sub ClearArray(ByRef pArr() As String)
        Try
            For I As Int16 = 0 To pArr.Length - 1
                pArr(I) = ""
            Next

        Catch ex As Exception

        End Try

    End Sub
    Private Sub ClearArraySplit(ByRef pArr() As String, ByVal inputPos As Integer)
        Try

            For I As Int16 = 0 To pArr.Length - 1
                If inputPos <> 0 And inputPos <> 10 And inputPos <> 11 Then
                    pArr(I) = ""
                End If

            Next

        Catch ex As Exception

        End Try

    End Sub

    Private Function GetSubstring(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring = "~Error~"
                Else
                    'GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    If Len(Mid(pStrValue, pStartPos + 1, Len(pStrValue))) < (pEndPos - pStartPos) Then
                        GetSubstring = Mid(pStrValue, pStartPos + 1, pEndPos - pStartPos)
                    Else
                        GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    End If
                End If
            End If

        Catch ex As Exception
            GetSubstring = "~Error~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")
        End Try

    End Function
    
    Private Function GetValidateDate(ByRef pStrDate As String) As Boolean

        Try
            ''Commented and Added by Lakshmi dtd 08-05-12
            'strInputDateFormat = strInputDateFormat.ToUpper()


            strInputDateFormat = strInputDateFormat.ToUpper()
            ''-

            Dim TmpstrInputDateFormat() As String
            Dim TempStrDateValue() As String = pStrDate.Split(" ")

            If InStr(TempStrDateValue(0), "/") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("/")
                TmpstrInputDateFormat = strInputDateFormat.Split("/")
            ElseIf InStr(TempStrDateValue(0), "-") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("-")
                If strInputDateFormat.Contains("-") Then
                    TmpstrInputDateFormat = strInputDateFormat.Split("-")
                Else
                    TmpstrInputDateFormat = strInputDateFormat.Split("/")
                End If

                '   TmpstrInputDateFormat = strInputDateFormat.Split("-")
            End If

            Dim HsUserDate As New Hashtable
            Dim HsSystemDate As New Hashtable
            Dim StrFinalDate As String

            If TempStrDateValue.Length = 3 Then
                For IntStr As Integer = 0 To TempStrDateValue.Length - 1
                    HsUserDate.Add(GetShort(TmpstrInputDateFormat(IntStr)), TempStrDateValue(IntStr))
                Next
                Dim SysDate() As String
                Dim dtSys As String = System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern.ToUpper()
                If InStr(dtSys, "/") > 0 Then
                    SysDate = dtSys.Split("/")
                ElseIf InStr(dtSys, "-") > 0 Then
                    SysDate = dtSys.Split("-")
                End If

                StrFinalDate = ""
                For IntStr As Integer = 0 To SysDate.Length - 1
                    If StrFinalDate = "" Then
                        StrFinalDate += HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    Else
                        StrFinalDate += "/" & HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    End If
                Next

                Try
                    ''pStrDate = Format(CDate(StrFinalDate), "dd/MM/yyyy")
                    pStrDate = CDate(StrFinalDate)
                    'InputDate = CDate(StrFinalDate)
                    GetValidateDate = True

                Catch ex As Exception
                    GetValidateDate = False

                End Try
            Else
                GetValidateDate = False
            End If

        Catch ex As Exception
            GetValidateDate = False

        End Try
    End Function
    Private Function GetShort(ByVal pStr As String) As String

        pStr = pStr.ToUpper

        If InStr(pStr, "D") > 0 Then
            GetShort = "D"
        ElseIf InStr(pStr, "M") > 0 Then
            GetShort = "M"
        ElseIf InStr(pStr, "Y") > 0 Then
            GetShort = "Y"
        End If

    End Function

    Private Sub AddRowsToDataTable(ByVal pNotValid As Boolean, ByVal Data() As String)
        Try
            If Data Is Nothing Then Exit Sub

            If pNotValid = True Then
                DtUnSucInput.Rows.Add(Data)
            Else
                DtInput.Rows.Add(Data)
            End If


        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "AddRowsToDataTable")
        End Try
    End Sub

    Private Function GetValueFormArray(ByRef pArray() As Object, ByVal pPosition As Int16) As String

        Try
            If pArray.Length >= pPosition Then
                GetValueFormArray = pArray(pPosition - 1).ToString()
            Else
                GetValueFormArray = "~ERROR~"
            End If

        Catch ex As Exception

            GetValueFormArray = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetValueFormArray")

        End Try

    End Function

    Public Function IsJustAlpha(ByVal sText As String, ByVal num As Integer, ByVal ReplaceSpace As String, Optional ByVal ShowMsgYN As String = "") As String
        Try
            Dim SpecialCharReplace() As DataRow = Nothing
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""


            'If sText <> "" Then
            For n = 1 To iTextLen
                sChar = Mid(sText, n, 1)
                If ChkText(sChar, num) Then
                    IsJustAlpha = IsJustAlpha + sChar
                Else

                    If ShowMsgYN = "Y" Then
                        IsJustAlpha = "Y"
                        Exit Function
                    Else
                        If ReplaceSpace = "Y" Then
                            IsJustAlpha = IsJustAlpha + " "
                        End If

                    End If

                End If
            Next
            'End If

            If Not IsJustAlpha Is Nothing Then
                Return IsJustAlpha
            Else
                Return ""
            End If


        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "IsJustAlpha")
        End Try
    End Function
    

    Private Function ChkText(ByVal sChr As String, ByVal num As Integer) As Boolean

        Try
            Select Case num
                Case 1
                    '- name field 
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]"
                    'ChkText = True
                Case 2
                    '- amount field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[.]" 'Or sChr Like "[,]"
                    'ChkText = True
                Case 3
                    '- alhpa numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[,]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[ ]" Or sChr Like "[.]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[:]"
                    'ChkText = True
                Case 4
                    '- address field
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 5
                    '- number field
                    ChkText = sChr Like "[0-9]"
                    'ChkText = True
                Case 6
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]"
                    'ChkText = True
                Case 7
                    '- Date field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[:]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[-]" Or sChr Like "[.]"
                    'ChkText = True
                Case 8
                    '- alhpa numeric field & All Characters on Keyboard
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[_]" Or sChr Like "[&]" Or sChr Like "[$]" Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[\]" Or sChr Like "[[]" Or sChr Like "[]]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[']" Or sChr Like "[ ]" Or sChr Like "[;]" Or sChr Like "[#]" Or sChr Like "[%]" Or sChr Like "[^]" Or sChr Like "[*]" Or sChr Like "[=]" Or sChr Like "[|]"
                    'ChkText = True
                Case 9
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]"
                Case 10
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[-]" Or sChr Like "[ ]" Or sChr Like "[_]"

                Case 11
                    '- alhpa numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[,]" Or sChr Like "[ ]" Or sChr Like "[.]"
                Case 12
                    '- address field
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[|]" Or sChr Like "[!]" Or sChr Like "[#]" Or sChr Like "[@]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[%]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 13
                    '- name field 
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]"
                Case 14
                    '- Bene ID
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[_]" Or sChr Like "[-]" Or sChr Like "[/]"
                Case 15
                    '- PayDate
                    ChkText = sChr Like "[0-9]" Or sChr Like "[/]" Or sChr Like "[|]" Or sChr Like "[~]"
                Case 16  ''''If amount in (-) minus
                    '- amount field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[.]" Or sChr Like "[-]"
                    'ChkText = True
                Case 17
                    '- Beneficiary Bank Account No field  "0 - 9", "a-z", "A-Z", ", . / ( ) :"
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[,]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[:]"
                    'ChkText = True
                Case Else
                    ChkText = False
            End Select

            Return ChkText

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "ChkText")
        End Try
    End Function

    Public Function RemoveJunk(ByVal sText As String) As String
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsAlpha(sChar) Then
                        RemoveJunk = RemoveJunk + sChar
                    End If
                Next
            End If

        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "RemoveJunk")

        End Try
    End Function

    Private Function IsAlpha(ByVal sChr As String) As Boolean
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters

        IsAlpha = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" _
        Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[;]" Or sChr Like "[:]" _
        Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[?]" Or sChr Like "[/]" _
        Or sChr Like "[']" Or sChr Like "[""]" Or sChr Like "[|]" Or sChr Like "[\]" _
        Or sChr Like "[{]" Or sChr Like "[[]" Or sChr Like "[}]" Or sChr Like "[]]" _
        Or sChr Like "[+]" Or sChr Like "[=]" Or sChr Like "[_]" Or sChr Like "[-]" _
        Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[*]" Or sChr Like "[&]" _
        Or sChr Like "[^]" Or sChr Like "[%]" Or sChr Like "[$]" Or sChr Like "[#]" _
        Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[`]" Or sChr Like "[~]" _
        Or sChr Like "[ ]" 'commented dtd 03-06-2011

    End Function


    Public Function SpCharValidation(ByVal StringValue As String, ByRef _dtSpChar As DataTable) As String

        ''Added by Jaiwant dtd  03-Dec-2010
        Dim ArrSpChar(0) As String
        Dim intSpCharRow As Integer
        ''---
        ClearArray(ArrSpChar)
        Array.Resize(ArrSpChar, _dtSpChar.Select.Length)
        intSpCharRow = 0

        For Each SVRow As DataRow In _dtSpChar.Rows
            ArrSpChar(intSpCharRow) = SVRow(0).ToString
            intSpCharRow += 1
        Next

        ''Added By Jaiwant dtd  03-Dec-2010 ''For All Special Characters
        Dim StrOriginalValue As String = ""
        Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&", " "}

        Try
            ''To remove special chars from array which need to ignore.
            For iIChar As Int16 = 0 To ArrSpChar.Length - 1
                For iSChar As Int16 = 0 To arrSpecialChar.Length - 1
                    If ArrSpChar(iIChar) = arrSpecialChar(iSChar) Then
                        arrSpecialChar(iSChar) = Nothing
                    End If
                Next
            Next
            SpCharValidation = ""
            Dim i As Integer
            For i = 0 To arrSpecialChar.Length - 1
                If InStr(StringValue, arrSpecialChar(i), CompareMethod.Binary) <> 0 Then
                    SpCharValidation = SpCharValidation & arrSpecialChar(i)
                End If
            Next

            Return SpCharValidation

        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "SpCharValidation")

        End Try
    End Function
    Public Function RemoveSplChar(ByVal sText As String, ByVal intType As Integer) As String
        ''-To remove Junk Characters-
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsSplChar(sChar, intType) = True Then
                        RemoveSplChar = RemoveSplChar & sChar
                    Else
                        RemoveSplChar = RemoveSplChar & " "
                    End If
                Next
            Else
                RemoveSplChar = ""
            End If

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveSplChar")

        End Try

    End Function

    Private Function IsSplChar(ByVal strChar As String, ByVal intType As Integer) As Boolean


        Select Case intType

            Case 1

                IsSplChar = strChar Like "[0-9]"

            Case 2

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]"

            Case 3

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" Or strChar Like "[/]" _
                            Or strChar Like "[:]" Or strChar Like "[-]" Or strChar Like "[?]" Or strChar Like "[+]" _
                            Or strChar Like "[(]" Or strChar Like "[)]" Or strChar Like "[.]" Or strChar Like "[,]"
            Case 4

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[/]" Or strChar Like "[-]"

            Case 5

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[.]"

            Case 6
                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" Or strChar Like "[/]" _
                            Or strChar Like "[:]" Or strChar Like "[(]" Or strChar Like "[)]" Or strChar Like "[.]" Or strChar Like "[,]"

            Case 7
                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" _
                             Or strChar Like "[.]" Or strChar Like "[_]" Or strChar Like "[@]"
            Case 8

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" Or strChar Like "[.]"
        End Select
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function

#Region " IDisposable Support "

    Public Sub Dispose() Implements IDisposable.Dispose

        If Not ObjBaseClass Is Nothing Then ObjBaseClass.Dispose()
        If Not DtValidation Is Nothing Then DtValidation.Dispose()
        ''Added by Jaiwant dtd  03-Dec-2010
        If Not DtSpCharValidation Is Nothing Then DtSpCharValidation.Dispose()
        ''----
        If Not DtInput Is Nothing Then DtInput.Dispose()
        If Not DtUnSucInput Is Nothing Then DtUnSucInput.Dispose()
        If Not DtTemp Is Nothing Then DtTemp.Dispose()

        ObjBaseClass = Nothing
        DtValidation = Nothing
        ''Added by Jaiwant dtd  03-Dec-2010
        DtSpCharValidation = Nothing
        ''----
        DtInput = Nothing
        DtUnSucInput = Nothing
        DtTemp = Nothing

        GC.SuppressFinalize(Me)
    End Sub
    Private Function GetSubstring1(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring1 = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring1 = "~Error~"
                Else
                    'GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    If Len(Mid(pStrValue, pStartPos + 1, Len(pStrValue))) < (pEndPos - pStartPos) Then
                        GetSubstring1 = Mid(pStrValue, pStartPos + 1, pEndPos - pStartPos)
                    Else
                        GetSubstring1 = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    End If
                End If
            End If

        Catch ex As Exception
            GetSubstring1 = "~Error~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")
        End Try

    End Function

    'Private Function Validate() As Boolean

    '    Validate = False
    '    Dim DrValidOutputColumn() As DataRow = Nothing
    '    Dim StrDataRow(24) As String
    '    Dim ArrDataRow As Object
    '    Dim InputLineNumber As Int32 = 0
    '    Dim Amount As Double = 0

    '    Dim TXN_NO As Integer
    '    Dim SUBTXN_NO As Integer = 1

    '    Dim intPosField As Integer = 3
    '    Dim HardCode As Integer = 2
    '    Dim MandatoryPos As Integer = 4
    '    Dim LengthPosMax As Integer = 5
    '    Dim CharType As Integer = 6
    '    Dim ReplaceSpace As Integer = 7
    '    Dim IFSCCOde As String = ""

    '    Dim inputPos() As String = Nothing
    '    '  Dim SrNo As Integer

    '    Try


    '        DtValidation = ObjBaseClass.GetDataTable_ExcelSheet(strValidationPath, "Epay")
    '        DrValidOutputColumn = DtValidation.Select("[SRNO] <> 0  ", "[SRNO]")

    '        DtTemp = New DataTable()
    '        DtTemp = ObjBaseClass.GetDataTable_ExcelSheetInput(gstrInputFolder & "\" & gstrInputFile, "", "")

    '        RemoveBlankRow(DtTemp)

    '        If DtTemp.Rows(0)(0).ToString().Trim().ToUpper() <> "Transcation Type".ToUpper Then
    '            ObjBaseClass.LogEntry("Header missing in Input File  .")
    '            Validate = False
    '            Exit Function
    '        End If
    '        For i = 0 To DtTemp.Columns.Count - 1
    '            DtTemp.Columns(i).ColumnName = DtTemp.Rows(0)(i)
    '        Next

    '        If DtTemp.Rows.Count > 0 Then
    '            TXN_NO = 0

    '            For Each dtRow As DataRow In DtTemp.Rows
    '                InputLineNumber += 1
    '                If InputLineNumber > 1 Then

    '                    ClearArray(StrDataRow)
    '                    ArrDataRow = dtRow.ItemArray()

    '                    For intIndex As Int32 = 0 To DrValidOutputColumn.Length - 1
    '                        If Val(DrValidOutputColumn(intIndex)(intPosField).ToString().Trim()) <> 0 Then
    '                            inputPos = DrValidOutputColumn(intIndex)(intPosField).ToString().Split(",")
    '                            For index = 0 To inputPos.Length - 1
    '                                StrDataRow(intIndex) = StrDataRow(intIndex).Trim() & GetValueFormArray(ArrDataRow, inputPos(index)).Trim()
    '                            Next
    '                        Else
    '                            If StrDataRow(intIndex) = "~Error~" Then
    '                                StrDataRow(24) = "Input Line " & InputLineNumber & "  " & DrValidOutputColumn(intIndex)(2).ToString().Trim() & " Error in Input Position |"
    '                            Else
    '                                StrDataRow(intIndex) = ""
    '                            End If
    '                        End If

    '                        '  HardCode Value
    '                        If DrValidOutputColumn(intIndex)(HardCode).ToString().Trim() <> "" Then
    '                            StrDataRow(intIndex) = DrValidOutputColumn(intIndex)(HardCode).ToString()
    '                        End If

    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Debit Account Number".Trim().ToUpper() Then
    '                            If strDebitAccNo <> "" Then
    '                                StrDataRow(intIndex) = IsJustAlpha(strDebitAccNo.ToString().Trim().Replace("&", "And"), Val(DrValidOutputColumn(intIndex)(CharType).ToString().Trim()), DrValidOutputColumn(intIndex)(ReplaceSpace).ToString().Trim())
    '                                If StrDataRow(intIndex).Length() <> 15 Then
    '                                    StrDataRow(24) = StrDataRow(24) & "Input Line " & InputLineNumber & " Debit Account Number [" & StrDataRow(intIndex) & "] should be 15 Digit|"
    '                                End If
    '                            End If
    '                        End If
    '                        ''''Replace & with and
    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper = "Transaction Ref No".ToUpper Or DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper = "Debit Account Number".ToUpper Or DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper = "Beneficiary Bank Account No".ToUpper Or DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper = "Beneficiary Bank IFS Code".ToUpper Then
    '                        Else
    '                            StrDataRow(intIndex) = StrDataRow(intIndex).Replace("&", "And")
    '                        End If
    '                        ''''''''''''''''''Added by swati dtd 2021-02-25
    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Beneficiary Bank Account No".Trim().ToUpper() Then
    '                            If StrDataRow(intIndex).ToString() <> "" Then
    '                                If StrDataRow(intIndex).ToString().Contains("-") Then
    '                                    StrDataRow(intIndex) = StrDataRow(intIndex).ToString().Replace("-", "")
    '                                Else
    '                                    StrDataRow(intIndex) = StrDataRow(intIndex).ToString()
    '                                End If
    '                            End If
    '                        End If
    '                        ''''''''''''''''''''''''''''''''''''''''''''''''''
    '                        'character validation
    '                        If Val(DrValidOutputColumn(intIndex)(CharType).ToString().Trim()) > 0 Then
    '                            If IsJustAlpha(StrDataRow(intIndex).Trim(), Val(DrValidOutputColumn(intIndex)(CharType).ToString().Trim()), DrValidOutputColumn(intIndex)(ReplaceSpace).ToString().Trim(), DrValidOutputColumn(intIndex)(8).ToString().Trim()) = "Y" Then
    '                                StrDataRow(24) = StrDataRow(24) & "Input Line :" & InputLineNumber & " column :" & DrValidOutputColumn(intIndex)(1).ToString().Trim() & " [" & StrDataRow(intIndex) & "]  This field contains special characters  |"
    '                                Continue For
    '                            Else
    '                                StrDataRow(intIndex) = IsJustAlpha(StrDataRow(intIndex).Trim(), Val(DrValidOutputColumn(intIndex)(CharType).ToString().Trim()), DrValidOutputColumn(intIndex)(ReplaceSpace).ToString().Trim(), DrValidOutputColumn(intIndex)(8).ToString().Trim())
    '                            End If
    '                        End If

    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Message Type".Trim().ToUpper() Then
    '                            IFSCCOde = IsJustAlpha(dtRow(24), Val(DrValidOutputColumn(7)(CharType).ToString().Trim()), "N").Replace(" ", "")
    '                            If IFSCCOde.StartsWith("YESB") Then
    '                                If IFSCCOde.Length > 5 Then
    '                                    Dim strCheckAlph As String
    '                                    strCheckAlph = IFSCCOde.Substring(5, 3)
    '                                    If IsNumeric(strCheckAlph) Then
    '                                        StrDataRow(intIndex) = "A"
    '                                    Else
    '                                        StrDataRow(intIndex) = "N06"
    '                                    End If
    '                                End If

    '                            ElseIf (StrDataRow(intIndex).ToString.ToUpper() = "I") Then
    '                                StrDataRow(intIndex) = "N06"
    '                            ElseIf (StrDataRow(intIndex).ToString.ToUpper() = "R") Then
    '                                StrDataRow(intIndex) = "R41"
    '                            ElseIf (StrDataRow(intIndex).ToString.ToUpper() = "N") Then
    '                                StrDataRow(intIndex) = "N06"
    '                            Else
    '                                StrDataRow(24) = StrDataRow(24) & "Input Line " & InputLineNumber & " Please check Transaction Type [" & StrDataRow(intIndex) & "]" & "|"
    '                            End If
    '                        End If

    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Upload Date".Trim().ToUpper() Then
    '                            Dim str As String = StrDataRow(intIndex).ToString.Trim()
    '                            If str.ToString().Trim() <> "" Then
    '                                If GetValidateDate(str) = True Then
    '                                    StrDataRow(intIndex) = Format(CDate(str), "dd\/MM\/yyyy")
    '                                    If DateDiff("d", CDate(str), Now.Date()) > 0 Then
    '                                        StrDataRow(24) = StrDataRow(24) & "Input Line :" & InputLineNumber & " column Name " & DrValidOutputColumn(intIndex)(1).ToString().Trim() & "[" & StrDataRow(intIndex) & "]Please Check transaction Is -Past dated  |"
    '                                    End If
    '                                Else
    '                                    StrDataRow(24) = StrDataRow(24) & "Input Line " & InputLineNumber & " column Name " & DrValidOutputColumn(intIndex)(1).ToString().Trim() & "[" & StrDataRow(intIndex) & "] Is Not Valid Date Format|"
    '                                End If
    '                            End If
    '                        End If

    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Amount".Trim().ToUpper() Then
    '                            If StrDataRow(intIndex).ToString().Trim() <> "" Then
    '                                If StrDataRow(intIndex) = "0" Or StrDataRow(intIndex) = "0.00" Then
    '                                    StrDataRow(24) = StrDataRow(24) & "Input Line " & InputLineNumber & " column Name " & DrValidOutputColumn(intIndex)(1).ToString().Trim() & "[" & StrDataRow(intIndex) & "]  Amount is Zero ,For Transaction Ref No[" & StrDataRow(14) & "]|"
    '                                ElseIf (Val(StrDataRow(intIndex))) < 0 Then
    '                                    StrDataRow(24) = StrDataRow(24) & "Input Line " & InputLineNumber & " column Name " & DrValidOutputColumn(intIndex)(1).ToString().Trim() & "[" & StrDataRow(intIndex) & "]  Amount is Negative,For Transaction Ref No[" & StrDataRow(14) & "]|"
    '                                Else
    '                                    StrDataRow(intIndex) = Val(StrDataRow(intIndex)).ToString(".00").Trim()
    '                                End If
    '                            End If
    '                        End If

    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Beneficiary Bank IFS Code".Trim().ToUpper() Then
    '                            If StrDataRow(intIndex).Length() <> 11 Then
    '                                StrDataRow(24) = StrDataRow(24) & "Input Line : " & InputLineNumber & " IFSC Code [" & StrDataRow(intIndex) & "] should be 11 Digit|"
    '                            End If
    '                        End If

    '                        ''''''''''''''''''''''Commented by swati dtd 2021-02-25
    '                        'If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Beneficiary Bank Account No".Trim().ToUpper() Then
    '                        '    StrDataRow(intIndex) = StrDataRow(intIndex).ToString()
    '                        'End If
    '                        '''''''''''''''''''''''''''''
    '                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "Sender To Receiver Info".Trim().ToUpper() Then
    '                            If StrDataRow(intIndex).ToString() <> "" Then
    '                                StrDataRow(intIndex) = StrDataRow(intIndex).ToString()
    '                            Else
    '                                StrDataRow(intIndex) = "Reliance Gen Ins"
    '                            End If
    '                        End If

    '                        '  ------------End Here
    '                        '  --------------Check mandatory 

    '                        If DrValidOutputColumn(intIndex)(MandatoryPos).ToString().Trim() = "M" And StrDataRow(intIndex).Trim = "" Then
    '                            StrDataRow(24) = StrDataRow(24) & "Input Line " & InputLineNumber & "  " & DrValidOutputColumn(intIndex)(1).ToString().Trim() & " This is Mandatory Field & it is Blank |"
    '                        End If


    '                        If strEpayOptFile_Format.ToString().Trim().ToUpper() = "TXT" Then
    '                            ' ''--------------Padding LenthWise

    '                            If DrValidOutputColumn(intIndex)(1).ToString.Trim().ToUpper() = "Amount".ToUpper Then
    '                                StrDataRow(intIndex) = Pad_Length(StrDataRow(intIndex).ToString().PadLeft(DrValidOutputColumn(intIndex)(LengthPosMax), "0"), DrValidOutputColumn(intIndex)(LengthPosMax))
    '                            Else
    '                                StrDataRow(intIndex) = Pad_Length(StrDataRow(intIndex).ToString().PadRight(DrValidOutputColumn(intIndex)(LengthPosMax), " "), DrValidOutputColumn(intIndex)(LengthPosMax))
    '                            End If
    '                            '''''
    '                        ElseIf (strEpayOptFile_Format.ToString().Trim().ToUpper() = "XLS" Or strEpayOptFile_Format.ToString().Trim().ToUpper() = "XLSX") Then
    '                            If StrDataRow(intIndex).Length > Val(DrValidOutputColumn(intIndex)(LengthPosMax).ToString()) Then
    '                                StrDataRow(intIndex) = Left(StrDataRow(intIndex).PadRight(Val(DrValidOutputColumn(intIndex)(LengthPosMax).ToString()), ""), Val(DrValidOutputColumn(intIndex)(LengthPosMax).ToString())).Trim()
    '                            End If
    '                        End If
    '                    Next

    '                    TXN_NO += 1
    '                    StrDataRow(22) = TXN_NO
    '                    StrDataRow(23) = SUBTXN_NO
    '                    If StrDataRow(24).ToString().Trim() = "" Then
    '                        DtInput.Rows.Add(StrDataRow)
    '                    Else
    '                        DtUnSucInput.Rows.Add(StrDataRow)
    '                    End If
    '                End If
    '            Next

    '        Else
    '            Call ObjBaseClass.Handle_Error(New ApplicationException("Validation is not maintained properly in " & Path.GetFileName(strValidationPath) & " validation file. It must be atleast 4 columns defination."), "ClsValidation", -123, "Validate")

    '        End If

    '        Validate_Advice()

    '        Validate = True

    '    Catch ex As Exception

    '        Validate = False
    '        ErrorMessage = ex.Message
    '        Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Validate")

    '    Finally
    '        DrValidOutputColumn = Nothing

    '        ObjBaseClass.ObjectDispose(DtTemp)
    '        ObjBaseClass.ObjectDispose(DtValidation)

    '    End Try

    'End Function


    Function RupeesToWord(ByVal MyNumber)
        Dim Temp
        Dim Rupees, Paisa As String
        Dim DecimalPlace, iCount
        Dim Hundreds, Words As String
        Dim place(9) As String
        place(0) = " Thousand "
        place(2) = " Lakh "
        place(4) = " Crore "
        place(6) = " Arab "
        place(8) = " Kharab "
        On Error Resume Next
        ' Convert MyNumber to a string, trimming extra spaces.
        MyNumber = Trim(Str(MyNumber))

        ' Find decimal place.
        DecimalPlace = InStr(MyNumber, ".")

        ' If we find decimal place...
        If DecimalPlace > 0 Then
            ' Convert Paisa
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Paisa = " and " & ConvertTens(Temp) & " Paisa"

            ' Strip off paisa from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        '===============================================================
        Dim TM As String  ' If MyNumber between Rs.1 To 99 Only.
        TM = Right(MyNumber, 2)

        If Len(MyNumber) > 0 And Len(MyNumber) <= 2 Then
            If Len(TM) = 1 Then
                Words = ConvertDigit(TM)
                'RupeesToWord = "Rupees " & Words & Paisa & " Only"
                RupeesToWord = Words & Paisa

                Exit Function

            Else
                If Len(TM) = 2 Then
                    Words = ConvertTens(TM)
                    'RupeesToWord = "Rupees " & Words & Paisa & " Only"
                    RupeesToWord = Words & Paisa
                    Exit Function

                End If
            End If
        End If
        '===============================================================


        ' Convert last 3 digits of MyNumber to ruppees in word.
        Hundreds = ConvertHundreds(Right(MyNumber, 3))
        ' Strip off last three digits
        MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        iCount = 0
        Do While MyNumber <> ""
            'Strip last two digits
            Temp = Right(MyNumber, 2)
            If Len(MyNumber) = 1 Then


                If Trim(Words) = "Thousand" Or _
                Trim(Words) = "Lakh  Thousand" Or _
                Trim(Words) = "Lakh" Or _
                Trim(Words) = "Crore" Or _
                Trim(Words) = "Crore  Lakh  Thousand" Or _
                Trim(Words) = "Arab  Crore  Lakh  Thousand" Or _
                Trim(Words) = "Arab" Or _
                Trim(Words) = "Kharab  Arab  Crore  Lakh  Thousand" Or _
                Trim(Words) = "Kharab" Then

                    Words = ConvertDigit(Temp) & place(iCount)
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                Else

                    Words = ConvertDigit(Temp) & place(iCount) & Words
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                End If
            Else

                If Trim(Words) = "Thousand" Or _
                   Trim(Words) = "Lakh  Thousand" Or _
                   Trim(Words) = "Lakh" Or _
                   Trim(Words) = "Crore" Or _
                   Trim(Words) = "Crore  Lakh  Thousand" Or _
                   Trim(Words) = "Arab  Crore  Lakh  Thousand" Or _
                   Trim(Words) = "Arab" Then


                    Words = ConvertTens(Temp) & place(iCount)


                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else

                    '=================================================================
                    ' if only Lakh, Crore, Arab, Kharab

                    If Trim(ConvertTens(Temp) & place(iCount)) = "Lakh" Or _
                       Trim(ConvertTens(Temp) & place(iCount)) = "Crore" Or _
                       Trim(ConvertTens(Temp) & place(iCount)) = "Arab" Then

                        Words = Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    Else
                        Words = ConvertTens(Temp) & place(iCount) & Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    End If

                End If
            End If

            iCount = iCount + 2
        Loop

        'RupeesToWord = "Rupees " & Words & Hundreds & Paisa & " Only"
        RupeesToWord = Words & Hundreds & Paisa
    End Function

    Private Function ConvertDigit(ByVal MyDigit)
        Select Case Val(MyDigit)
            Case 1 : ConvertDigit = "One"
            Case 2 : ConvertDigit = "Two"
            Case 3 : ConvertDigit = "Three"
            Case 4 : ConvertDigit = "Four"
            Case 5 : ConvertDigit = "Five"
            Case 6 : ConvertDigit = "Six"
            Case 7 : ConvertDigit = "Seven"
            Case 8 : ConvertDigit = "Eight"
            Case 9 : ConvertDigit = "Nine"
            Case Else : ConvertDigit = ""
        End Select
    End Function

    Private Function ConvertTens(ByVal MyTens)
        Dim Result As String

        ' Is value between 10 and 19?
        If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
                Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
        End If

        ConvertTens = Result
    End Function

    Private Function ConvertHundreds(ByVal MyNumber)
        Dim Result As String

        ' Exit if there is nothing to convert.
        If Val(MyNumber) = 0 Then Exit Function

        ' Append leading zeros to number.
        MyNumber = Right("000" & MyNumber, 3)

        ' Do we have a hundreds place digit to convert?
        If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred And "
        End If

        ' Do we have a tens place digit to convert?
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
        Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
        End If

        ConvertHundreds = Trim(Result)
    End Function

    Public Function IsJustAlpha1(ByVal sText As String, ByVal num As Integer, ByVal ReplaceSpace As String, Optional ByVal ShowMsgYN As String = "") As String
        Try
            Dim SpecialCharReplace() As DataRow = Nothing
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""


            'If sText <> "" Then
            For n = 1 To iTextLen
                sChar = Mid(sText, n, 1)
                If ChkText(sChar, num) Then
                    IsJustAlpha1 = IsJustAlpha1 + sChar
                Else

                    If ShowMsgYN = "Y" Then
                        IsJustAlpha1 = "Y"
                        Exit Function
                    Else
                        If ReplaceSpace = "Y" Then
                            IsJustAlpha1 = IsJustAlpha1 + " "
                        End If

                    End If

                End If
            Next
            'End If

            If Not IsJustAlpha1 Is Nothing Then
                Return IsJustAlpha1
            Else
                Return ""
            End If


        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "IsJustAlpha")
        End Try
    End Function
#End Region

End Class
