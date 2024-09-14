Option Explicit On

Module ModGen


    Public blnErrorLog As Boolean
    Public strAuditFolderPath As String
    Public strErrorFolderPath As String

    Public strInputFolderPath As String
    Public gstrInputFile As String
    Public gstrInputFolder As String

    Public gstrOutputFile As String

    Public strOutputFolderPath As String
   
    ''Res
    Public gstrResOutputfile As String
    Public gstrResponseInputFolder As String
    Public gstrResponseInputFile As String

    Public strResponseFolderPath As String             ' Response folder path
    Public strRes_OutputFolderPath As String            ' RevResponse folder path

    ''Archive
    Public strArchivedFolderSuc As String
    Public strArchivedFolderUnSuc As String
    ''''''''''''''''''

    Public strProceed As String
    Public strInvalidTrans As String
    Public FileCounter As String
    Public strEpayOptFile_Format As String

    Public strValidationPath As String

    '-Client Details-
    Public strClientCode As String
    Public strClientName As String
    Public strInputDateFormat As String
    ' Public strUtilityCode As String  ''''Comment by swati dt 2021-04-19
    ' Public strClientShortName As String
    Public strYBL_Nomenclature_Format As String
    Public strResFile_Nomenclature_Format As String

    Public strBank As String
    Public strBankUserName As String

    Public strDecryptKey As String

    Public DtMaster As DataTable
    Public strMasterPath As String
End Module


