Option Infer On
Option Strict On

Imports System.IO
Imports EaseCore
Imports System.Text.RegularExpressions
Imports EaseCore.DAL

''' <summary>
''' Holds the application configuration for the customer/client
''' </summary>
''' <remarks></remarks>
Public Class AppConfig
    Private Shared ReadOnly BASE_ERRORNUMBER As Integer = EaseCore.ErrorNumbers.EASECLASS_APPCONFIG
    Private Const LOG_APPNAME As String = "EASECLASS"

#Region "Enums, Fields, Properties, Structures"

    Public Structure stLanguageList
        Public LanguageID As Integer
        <VBFixedString(40)> Public LanguageDescX As String
    End Structure

    Public Structure stEASESys1
        Public ID As UInt32
        <VBFixedString(80)> Public DescX As String
        <VBFixedString(80)> Public ForeignDescX As String
        Public WL As Int16
        Public L As Int16
        Public Type As Int16
        Public DP As Int16
    End Structure

    Public Structure stWordRecOLD          'oldversion of data
        <VBFixedString(50)> Public w As String
        Public wl As Int16
        Public l As Int16
        Public type As Int16
        Public dp As Int16
    End Structure

    Public Structure stEASESys          'holds easesys configuration
        <VBFixedString(50)> Public Version As String
        <VBFixedString(40)> Public EASEConfig As String
        Public DBKey As Int16
    End Structure

    ''' <summary>
    ''' Holds the words from easesys1.dat - USED only for backward compatibility.
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure stWordRec
        <VBFixedString(50)> Public w As String
        Public wl As Int16
        Public l As Int16
        Public type As Int16
        Public dp As Int16
    End Structure

    Public wrd As stWordRec

    Public Structure stMTMRec
        <VBFixedString(8)> Public Code As String
        <VBFixedString(40)> Public Desc As String
        <VBFixedString(6)> Public TMU As String
    End Structure

    Public MTM As stMTMRec

    Public Structure stWordUpdate           'used to update the easesys1.dat
        Public WordPos As Int16
        <VBFixedString(80)> Public NewWord As String
        <VBFixedString(80)> Public OldWord As String
        <VBFixedString(80)> Public OldWordDesc As String
        Public l As Int16
        Public type As Int16
        Public dp As Int16
    End Structure

    Public WPRefType As Int16
    Public VerifyFiles As Boolean

    Public Property EncodeConnectionString() As Boolean
        Get
            Return gObjAppConfig.EncodeConnectionString
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.EncodeConnectionString = Value
        End Set
    End Property

    Private _sharedSubHeaderUsageWord As String

    Public ReadOnly Property SharedSubHeaderUsageWord As String
        Get
            If (IsNothing(_sharedSubHeaderUsageWord) OrElse _sharedSubHeaderUsageWord = "") Then
                Dim strTemp2 As String = Ec.AppConfig.GetWrd(2954)
                If strTemp2.Trim = "" Then strTemp2 = "Usage"
                _sharedSubHeaderUsageWord = SharedSubHeaderWord.Trim & " " & strTemp2.Trim
            End If
            Return _sharedSubHeaderUsageWord
        End Get
    End Property

    Private _sharedSubHeaderWord As String

    Public ReadOnly Property SharedSubHeaderWord As String
        Get
            If (IsNothing(_sharedSubHeaderWord) OrElse _sharedSubHeaderWord = "") Then
                _sharedSubHeaderWord = GetWrd(7197)
                If _sharedSubHeaderWord.Trim = "" Then _sharedSubHeaderWord = "Shared Sub-Header"
            End If
            Return _sharedSubHeaderWord
        End Get
    End Property

    Private _operationWord As String

    Public ReadOnly Property OperationWord As String
        Get
            If (IsNothing(_operationWord) OrElse _operationWord = "") Then
                If Ec.AppConfig.AstonMartin Then
                    _operationWord = Ec.AppConfig.GetWrd(1428) 'station
                Else
                    _operationWord = Ec.AppConfig.GetWrd(4018) 'operation
                End If
            End If
            Return _operationWord
        End Get
    End Property

    Private _sharedExistingSubHeaderWord As String

    Public ReadOnly Property SharedExistingSubHeaderWord As String
        Get
            If (IsNothing(_sharedExistingSubHeaderWord) OrElse _sharedExistingSubHeaderWord = "") Then
                _sharedExistingSubHeaderWord = Ec.AppConfig.GetWrd(7198)
                If _sharedExistingSubHeaderWord.Trim = "" Then _sharedExistingSubHeaderWord = "Shared existing Sub-Header"
            End If
            Return _sharedExistingSubHeaderWord
        End Get
    End Property

    Private _messystem As Boolean?

    Public ReadOnly Property MESSystem As Boolean
        Get
            If (_messystem Is Nothing) Then
                _messystem = IsEASEConfigValue(13, "3")
            End If
            Return _messystem.Value
        End Get
    End Property

    Private _standardwork As Boolean?

    Public ReadOnly Property StandardWork As Boolean
        Get
            If (_standardwork Is Nothing) Then
                _standardwork = IsEASEConfigValue(14, "1")
            End If
            Return _standardwork.Value
        End Get
    End Property

    Private _LPA As Boolean?

    ReadOnly Property LPA As Boolean
        Get
            If _LPA Is Nothing Then
                _LPA = IsEASEConfigValue(15, "7")
            End If
            Return _LPA.Value
        End Get
    End Property

    Public Property SplashMessage As String

    Private _operatortraining As Boolean?

    ReadOnly Property OperatorTraining As Boolean
        Get
            If _operatortraining Is Nothing Then
                _operatortraining = IsEASEConfigValue(16, "3")
                If Not _operatortraining Then
                    If Ec.AppConfig.AstonMartin Then
                        _operatortraining = True
                    End If
                End If
            End If
            Return _operatortraining.Value
        End Get
    End Property

    Public Function IsEASEConfigValue(byval position As integer, byval value as String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strEASEConfig As String = Trim(gObjAppConfig.EASEConfig)
        dim blnRtnValue = strEASEConfig.Length >= position AndAlso Mid(strEASEConfig, position, 1) = value

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function
    
    Public Property ApplicationIsAmlViewEase As Boolean

#End Region

#Region "Constructors, Initialization, and Load"

#End Region

#Region "Event Handlers"


#End Region

#Region "Main Methods"


#End Region

#Region "Utility Methods"

    Public Function BrowseToDirectory(Optional ByVal strInitialDir As String = "",
                                      Optional ByVal blnUserEASConfigDir As Boolean = False) As String
#If TRACE Then
        Dim startTicks As Long = Log.FILE_DIR_IO(
            String.Format("Enter strInitialDir:({0}) blnUserEASConfigDir:({1})", strInitialDir, blnUserEASConfigDir),
            LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim strDir As String = ""

        If blnUserEASConfigDir AndAlso strInitialDir.Trim = "" Then
            strInitialDir = EASEConfigDirectory
        End If
        If strInitialDir.Length > 0 Then
            If System.IO.Directory.Exists(strInitialDir) Then strDir = strInitialDir.Trim
        End If

        Using folderBrowseDialog As New System.Windows.Forms.FolderBrowserDialog()
            If Trim(strDir) = "" Then
                folderBrowseDialog.RootFolder = Environment.SpecialFolder.MyComputer
            Else
                folderBrowseDialog.RootFolder = Environment.SpecialFolder.MyComputer
                folderBrowseDialog.SelectedPath = strDir
            End If

            If folderBrowseDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                strDir = EaseCore.Extensions.Strings.AddSlash(folderBrowseDialog.SelectedPath)
            Else
                strDir = strDir
            End If
        End Using

#If TRACE Then
        Log.FILE_DIR_IO(String.Format("Exit ({0})", strDir), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strDir
    End Function

#End Region

#Region "Protected Methods"


#End Region

#Region "Private Methods"


#End Region


    Public Function GetWrd(ByVal intWord As Int32, ByVal defaultValue As String) As String


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_LOW(
            String.Format("Enter intWord;({0}) defaultValue:({1})", intWord, defaultValue), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = GetWrd(intWord)

        If String.IsNullOrEmpty(returnValue) Then
            returnValue = defaultValue
        End If


#If TRACE Then
        Log.EASESYS_IO_LOW(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    ''' <summary>
    ''' Gets the Word Description for the passed record position from easesys1
    ''' </summary>
    ''' <param name="intWord">Record  position</param>
    ''' <param name="intWL">WL Field value for the record position (BYREF)</param>
    ''' <param name="intL">L Field value for the record position (BYREF)</param>
    ''' <param name="intType">Type Field value for the record position (BYREF)</param>
    ''' <param name="intDP">DP Field value for the record position (BYREF)</param>
    ''' <returns>Description field for the record position</returns>
    ''' <remarks></remarks>
    Public Function GetWrd(ByVal intWord As Int32,
                           Optional ByRef intWL As Int16 = 0,
                           Optional ByRef intL As Int16 = 0,
                           Optional ByRef intType As Int16 = 0,
                           Optional ByRef intDP As Int16 = 0,
                           Optional ByVal blnGetEnglishWord As Boolean = False,
                           Optional ByVal blnGetForeignWord As Boolean = False,
                           Optional ByVal blnRemoveAmbersand As Boolean = False) As String


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_LOW(
                    String.Format("Enter intWord:({0}) blnGetEnglishWord:({1}) blnGetForeignWord:({2}) blnRemoveAmbersand:({3})",
                                  intWord, blnGetEnglishWord, blnGetForeignWord, blnRemoveAmbersand), LOG_APPNAME,
                    BASE_ERRORNUMBER + 0)
#End If

        '**************************************************************************
        'GetWrd function returns value only for Client Server Applications
        'Web Applications use local Getwrd function instead
        '**************************************************************************

        'blnGetEnglishWord, blnGetForeignWord used only in updating the new words file, not used anywhere else

        Dim str_RtnValue As String = ""

        If Not gBlnEASESys1Loaded Then GoTo ExitThisFunction 'make sure the easesys1 object is loaded

        Dim str_Search As String = CStr(intWord)
        Dim obj_Row As DataRow
        Try
            'blnGetForeignWord = True
            obj_Row = gObjEasesys1.Rows.Find(str_Search)
            If Not (obj_Row Is Nothing) Then

                If blnGetEnglishWord Then GoTo UseEnglish
                If blnGetForeignWord Then GoTo UseForeignWord
                'If DBConfig.LanguageID  > 0 then
                If DBConfig.LanguageID > 0 Then 'If DBConfig.UseForeignLanguage Then         'gUseEnglishLanguage
                    UseForeignWord:
                    str_RtnValue = CStr(obj_Row("intldesc"))

                    If blnGetForeignWord = False AndAlso Trim(str_RtnValue) = "" Then
                        'default for most of the getwrds
                        GoTo UseEnglish
                    End If
                Else
                    UseEnglish:
                    str_RtnValue = CStr(obj_Row("desc"))
                End If
                GetOtherFields:
                intWL = Extensions.Data.GetDataRowValue (Of Int16)(obj_Row("wl"), 0)
                intL = Extensions.Data.GetDataRowValue (Of Int16)(obj_Row("l"), 0)
                intType = Extensions.Data.GetDataRowValue (Of Int16)(obj_Row("type"), 0)
                intDP = Extensions.Data.GetDataRowValue (Of Int16)(obj_Row("dp"), 0)
            End If
        Catch ex As Exception
            Log.Error("Position: " & intWord & "(" & ex.Message & ")", LOG_APPNAME)
            Call GenerateException("GetWrd: Position:" & intWord, ex)
        Finally
            obj_Row = Nothing
        End Try

        If blnRemoveAmbersand Then
            str_RtnValue = str_RtnValue.Replace("&", "")
        End If

        ExitThisFunction:


#If TRACE Then
        Log.EASESYS_IO_LOW(
            String.Format("Exit ({0}) intWL:({1}) intL:({2}) intType:({3}) intDP:({4})", str_RtnValue, intWL, intL, intType,
                          intDP), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return str_RtnValue
    End Function

    ''' <summary>
    ''' Unloads the object which holds the easesys records
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UnLoadEaseSys1()

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        gBlnEASESys1Loaded = False

        gObjEasesys1.Dispose()
        gObjEasesys1 = Nothing

#If TRACE Then
        Log.EASESYS_IO("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Returns database Type for the database name 
    ''' If you pass "Access", this function will return 0 
    ''' 'Database type (0-access, 1-Microsoft Oracle .NET Provider, 2 - Sql Server, 3- Oracle .NET Provider, 4- Oracle Managed .NET Provider)
    ''' </summary>
    ''' <param name="strDBType">Database type (Access, Oracle, SQL Server, Oracle MSFT)</param>
    ''' <returns>Database type as integer</returns>
    ''' <remarks></remarks>
    Public Function GetDBType(ByVal strDBType As String) As Int16

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter strDBType:({0})", strDBType), LOG_APPNAME,
                                                     BASE_ERRORNUMBER + 0)
#End If

        'Get the Database type based on passed params

        Dim intRtnValue As Int16 = 0

        'Database type (0-access, 1-Microsoft Oracle .NET Provider, 2 - Sql Server, 3- Oracle .NET Provider)
        If strDBType = "Access" Then
            intRtnValue = 0
        ElseIf strDBType = "Oracle" Then
            intRtnValue = 1
        ElseIf strDBType = "SQL Server" Then
            intRtnValue = 2
        ElseIf strDBType = "OracleNet" Then
            intRtnValue = 3
            'JVG - EVC-784 - Trap & Retry Oracle database connection drop start
        ElseIf strDBType = "Oracle Managed Provider" Then
            intRtnValue = 4
            'JVG - EVC-784 - Trap & Retry Oracle database connection drop end
        End If


#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", intRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return intRtnValue
    End Function

    ''' <summary>
    ''' Returns the Database type name for the type (integer field)
    ''' 'Database type (0-access, 1-Microsoft Oracle .NET Provider, 2 - Sql Server, 3- Oracle .NET Provider, 4- Oracle Managed .NET Provider)
    ''' </summary>
    ''' <param name="intDBType">Database type (as integer)</param>
    ''' <returns>Database Type name (as string)</returns>
    ''' <remarks></remarks>
    Public Function GetDatabaseType(ByVal intDBType As Integer) As String

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter intDBType:({0})", intDBType), LOG_APPNAME,
                                                     BASE_ERRORNUMBER + 0)
#End If

        'Database type (0-access, 1-Microsoft Oracle .NET Provider, 2 - Sql Server, 3- Oracle .NET Provider, 4- Oracle Managed .NET Provider)
        Dim strRtnValue As String = ""
        If intDBType = 0 Then
            strRtnValue = "Access"
        ElseIf intDBType = 1 Then 'Oracle - MSFT Provider
            strRtnValue = "Oracle"
        ElseIf intDBType = 2 Then 'SQL Server
            strRtnValue = "SQL Server"
        ElseIf intDBType = 3 Then 'Oracle - Oracle Provider
            strRtnValue = "Oracle Managed Data Access" '"OracleNet"
            'JVG - EVC-784 - Trap & Retry Oracle database connection drop
        ElseIf intDBType = 4 Then 'Oracle - Oracle Managed Provider
            strRtnValue = "Oracle Managed Provider"
        End If


#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function


    ''' <summary>
    ''' Check the database already exists in easesys3.dat
    ''' </summary>
    ''' <param name="strConnectionString">Connection String</param>
    ''' <param name="blnEditDS">Are you trying to Edit an existing datasource or not</param>
    ''' <returns>True/False. If the matching database key found, returns true.</returns>
    ''' <remarks></remarks>
    Public Function CheckForDuplicateEntry(ByVal strConnectionString As String,
                                           Optional ByVal blnEditDS As Boolean = False,
                                           Optional ByRef blnIncomplete As Boolean = False) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(
            String.Format("Enter strConnectionString:({0}) blnEditDS:({1}) ", strConnectionString, blnEditDS), LOG_APPNAME,
            BASE_ERRORNUMBER + 0)
#End If
        'Check for the Database Key existence in the easesys3.dat and returns true if the dbkey field matches
        Dim blnRtnValue As Boolean = False
        Dim EASEDBs() As stDataSource
        Dim intMax As Int16 = 0
        If blnEditDS Then intMax = 1 'cover to editing an existing data source

        Dim intCounter = 0
        blnIncomplete = False
        Try
            ReDim EASEDBs(0)
            EASEDBs = ReadEaseSys3(True)
            'Call ReadEaseSys3(strEaseSys3Filename, EASEDBs)
            strConnectionString = strConnectionString.Trim.ToLower
            For intI = 1 To UBound(EASEDBs)
                If strConnectionString = EASEDBs(intI).ConnectionString.Trim.ToLower Then
                    intCounter += 1

                    If EASEDBs(intI).Update7Flag <> 1 Then
                        blnIncomplete = True
                    End If
                End If
            Next

            If intCounter > intMax Then
                blnRtnValue = True
            End If

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("CheckForDuplicateEntry", ex)
        Finally
            EASEDBs = Nothing
        End Try

        ExitThisFunction:

#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0}) blnIncomplete:({1})", blnRtnValue, blnIncomplete), LOG_APPNAME,
                            BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Public Function ReadEaseSys() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'MsgBox("ReadEaseSys")

        Dim strFileName As String = ""

        strFileName = GetEASESysDatabaseName()
        If Not File.Exists(strFileName) Then
            Call GenerateException("LoadEaseSys: File not found. (Location: " & strFileName & ")", New Exception)
            GoTo ExitThisSub
        End If

        Dim result = False
        Dim objEASESys As stEASESys
        Try
            'MsgBox("B4 ReadEASESysRecords")
            objEASESys = ReadEASESysRecords()

            gObjAppConfig.EASEConfig = objEASESys.EASEConfig
            gObjAppConfig.DBKey = objEASESys.DBKey
            gObjAppConfig.Version = objEASESys.Version

            gObjAppConfig.TimeUnit = 60 'use minutes as default. This value will be changed after reading readclientfile sub

            'MsgBox("B4 ConfigureApplicationSettings")
            ConfigureApplicationSettings(gObjAppConfig.EASEConfig)
            result = True

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("ReadEaseSys", ex)
        Finally
            objEASESys = Nothing
            'objStreamReader = Nothing
        End Try

        ExitThisSub:

        'gObjAppConfig.ExtraMinuteDecimal = intExtraMinuteDecimal
        'gObjAppConfig.SharedDirectory = strSharedDirectory
        'gObjAppConfig.EncodeConnectionString = blnEncodeCS


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result
    End Function

    ''' <summary>
    ''' Set the application configuration params based on easeconfig from easesys
    ''' </summary>
    ''' <param name="strEASEConfig"></param>
    ''' <remarks></remarks>
    Private Sub ConfigureApplicationSettings(ByVal strEASEConfig As String)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO(String.Format("Enter strEASEConfig:({0})", strEASEConfig), LOG_APPNAME,
                                                BASE_ERRORNUMBER + 0)
#End If

        'EASEConfig = 
        '    1 - modf
        '    2 - modf2
        '    3 - capp
        '    4 - LB
        '    5 - material
        '    6 - expertease
        '    7 - machine
        '    8 - UAS
        '    9 - CAPPonly
        '   10 - single User
        '   11 - MDM
        '   12 - MCR
        '   13 - MES
        '   14 - Standard Work
        '   15 - LPA
        '   16 - Operator Training
        '   17 - ToolManagement

        If Trim(strEASEConfig) = "" Then Exit Sub

        Dim intPos As Int16 = 0
        Try

            If strEASEConfig.Trim.Length < 12 Then
                GoTo ExitThisSub
            End If
            gObjAppConfig.Modf = EaseCore.Extensions.Strings.ToByte(Mid(strEASEConfig, 1, 1))
            gObjAppConfig.Modf2 = EaseCore.Extensions.Strings.ToByte(Mid(strEASEConfig, 2, 1))

            gObjAppConfig.Material = False
            If Ec.AppConfig.Modf2 = 3 Then
                If EaseCore.Extensions.Strings.ToInt16(Mid(strEASEConfig, 5, 1)) = 1 Then gObjAppConfig.Material = True
            End If

            gObjAppConfig.CAPP = False
            gObjAppConfig.ControlPlan = False
            If EaseCore.Extensions.Strings.ToInt16(Mid(strEASEConfig, 3, 1)) = 1 Then
                gObjAppConfig.CAPP = True
                gObjAppConfig.ControlPlan = True
            End If

            gObjAppConfig.LineBalance = False
            If Mid(strEASEConfig, 4, 1) = "1" Then gObjAppConfig.LineBalance = True

            gObjAppConfig.ExpertEASE = False
            intPos = 6
            If strEASEConfig.Length >= intPos AndAlso Mid(strEASEConfig, intPos, 1) = "7" Then gObjAppConfig.ExpertEASE = True

            gObjAppConfig.Machine = False
            intPos = 7
            If strEASEConfig.Length >= intPos AndAlso Mid(strEASEConfig, intPos, 1) = "1" And gObjAppConfig.CAPP Then _
                gObjAppConfig.Machine = True

            gObjAppConfig.UASUser = False
            intPos = 8
            If strEASEConfig.Length >= intPos AndAlso Mid(strEASEConfig, intPos, 1) = "1" Then gObjAppConfig.UASUser = True

            gObjAppConfig.CAPPOnly = False
            intPos = 9
            If strEASEConfig.Length >= intPos AndAlso Mid(strEASEConfig, intPos, 1) = "1" Then gObjAppConfig.CAPPOnly = True

            gObjAppConfig.SingleUserLicense = False
            intPos = 10
            If strEASEConfig.Length >= intPos AndAlso Mid(strEASEConfig, intPos, 1) = "1" Then _
                gObjAppConfig.SingleUserLicense = True

            gObjAppConfig.MDM = False
            intPos = 11
            If strEASEConfig.Length >= intPos AndAlso Mid(strEASEConfig, intPos, 1) = "1" Then gObjAppConfig.MDM = True

            gObjAppConfig.MCR = False
            intPos = 12
            If _
                strEASEConfig.Length >= intPos AndAlso
                (Mid(strEASEConfig, intPos, 1) = "1" Or Mid(strEASEConfig, intPos, 1) = "7") Then gObjAppConfig.MCR = True

            '13th position is for 3 (MES System, look for the function in the sub)
            '14 - Standard Work
            '15 - LPA 
            '16 - Training

            'R.J 
            '17 - Tool Management
            gObjAppConfig.ToolManagement = False
            intPos = 17
            If strEASEConfig.Length >= intPos AndAlso Mid(strEASEConfig, intPos, 1) = "1" Then _
                gObjAppConfig.ToolManagement = True

            ExitThisSub:
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("ConfigureApplicationSettings", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Returns the number of available database count
    ''' </summary>
    ''' <param name="blnAllDatabase">true, to get all database count, including incomplete database</param>
    ''' <returns>Number of database count.</returns>
    ''' <remarks></remarks>
    Public Function GetADatabaseCount(ByVal blnAllDatabase As Boolean) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter blnAllDatabase:({0})", blnAllDatabase), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If

        Dim intRtnValue = 0
        Dim EASEDBs() As stDataSource

        Try
            ReDim EASEDBs(0)
            EASEDBs = ReadEaseSys3(blnAllDatabase)

            intRtnValue = UBound(EASEDBs)
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("GetADatabaseCount", ex)
        Finally
            EASEDBs = Nothing
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", intRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return intRtnValue
    End Function

    ''' <summary>
    ''' Reads the available datasources and find the matching default database
    ''' returns the database details as object.
    ''' </summary>
    ''' <param name="intDBKey">Database key to search for.</param>
    ''' <param name="intDBCount">Number of available database count. (BYREF)</param>
    ''' <param name="blnDBFound">Match found for the database passed parameter. (BYREF)</param>
    ''' <returns>Object holds the database details for the passed parameter dbkey.</returns>
    ''' <remarks></remarks>
    Public Function GetADatabase(ByVal intDBKey As Integer, Optional ByRef intDBCount As Int16 = 0,
                                 Optional ByRef blnDBFound As Boolean = False) As stDataSource

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter intDBKey:({0})", intDBKey), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If

        'Get the Database info based on the DBKey passed

        Dim objDB As New stDataSource
        Dim EASEDBs() As stDataSource

        blnDBFound = False
        Try
            ReDim EASEDBs(0)

            EASEDBs = ReadEaseSys3(False, intDBKey)
            If UBound(EASEDBs) > 0 Then
                blnDBFound = True
            End If

            If blnDBFound Then objDB = EASEDBs(1)

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("GetADatabase", ex)
        Finally
            EASEDBs = Nothing
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0}) intDBCount:({1}) blnDBFound:({2})", objDB, intDBCount, blnDBFound),
                           LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objDB
    End Function

    ''' <summary>
    ''' Returns the customer name from easesys1
    ''' </summary>
    ''' <returns>Customer name</returns>
    ''' <remarks></remarks>
    Private Function GetCustomerName() As String

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'get the Customer name from easesys1
        Dim strPassword As String = "", strTemp As String = ""

        strTemp = GetWrd(9).Trim & GetWrd(8).Trim & GetWrd(7).Trim
        strTemp = GeneralFunctions.Decrypt(strTemp).Trim()

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", strTemp.Trim), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strTemp.Trim
    End Function

    Public Sub SetCustomerName(Optional ByVal blnNavyCustomFlag As Boolean = False)

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW(String.Format("Enter blnNavyCustomFlag:({0})", blnNavyCustomFlag),
                                                       LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        If blnNavyCustomFlag Then
            gObjAppConfig.CustomerName = "NAVY"
        Else
            gObjAppConfig.CustomerName = GetCustomerName()
        End If

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Encrypts the connection String to non-human readable form.
    ''' </summary>
    ''' <param name="strCS">Connection String</param>
    ''' <returns>Encrypted Connection string.</returns>
    ''' <remarks></remarks>
    Public Function EncryptConnectionString(ByVal strCS As String) As String


#If TRACE Then
        Dim startTicks As Long = Log.SECURITY_LOW(String.Format("Enter strCS:({0})", strCS), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'encode the connection string
        strCS = Trim(strCS)
        If InStr(strCS.ToLower, ".mdb") > 0 Then GoTo ExitThisFunction
        Dim strTemp = ""
        Dim intJ = 1
        For intI = 1 To Len(strCS) 'To 1 Step -1
            strTemp = strTemp & Chr(Asc(Mid(strCS, intI, 1)) + intJ)
            intJ += 1
            If intJ Mod 4 = 0 Then intJ = 1
        Next

        'now reverse it
        strCS = ""
        For intI = Len(strTemp) To 1 Step - 1
            strCS &= Mid(strTemp, intI, 1)
        Next
        ExitThisFunction:


#If TRACE Then
        Log.SECURITY_LOW(String.Format("Exit ({0})", strCS), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strCS.Trim
    End Function

    ''' <summary>
    ''' Decrypt the connection to human readable format.
    ''' </summary>
    ''' <param name="strCS">Encrypted connection string.</param>
    ''' <returns>Human readable connection string.</returns>
    ''' <remarks></remarks>
    Public Function DecodeConnectionString(ByVal strCS As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.SECURITY_LOW(String.Format("Enter strCS:({0})", strCS), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'decode the connection string
        Dim strTemp As String = ""
        strCS = Trim(strCS)
        If InStr(strCS.ToLower, ".mdb") > 0 Then GoTo ExitThisFunction

        'first reverse the string
        For intI = Len(strCS) To 1 Step - 1
            strTemp &= Mid(strCS, intI, 1)
        Next
        'strTemp = strCS
        Dim intJ = 1
        strCS = ""
        For intI = 1 To Len(strTemp)
            strCS = strCS & Chr(Asc(Mid(strTemp, intI, 1)) - intJ)

            intJ += 1
            If intJ Mod 4 = 0 Then intJ = 1
        Next

        ExitThisFunction:


#If TRACE Then
        Log.SECURITY_LOW(String.Format("Exit ({0})", strCS), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strCS.Trim
    End Function


    Private _minuteText As String
    Private _secondText As String
    Private _hoursText As String

    ''' <summary>
    ''' Returns Time unit in Sting
    ''' </summary>
    ''' <param name="intTimeUnit">Time Unit in Integer (60-mins, 3600-secs, else hrs)</param>
    ''' <returns>Time Unit in string format</returns>
    ''' <remarks></remarks>
    Public Function GetTimeUnitText(ByVal intTimeUnit As Int16) As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW(String.Format("Enter intTimeUnit:({0})", intTimeUnit), LOG_APPNAME,
                                                 BASE_ERRORNUMBER + 0)
#End If

        Dim strResult As String

        Select Case intTimeUnit
            Case 60
                If (String.IsNullOrEmpty(_minuteText)) Then
                    _minuteText = GetWrd(3778) '"mins"
                End If
                strResult = _minuteText
            Case 3600
                If (String.IsNullOrEmpty(_secondText)) Then
                    _secondText = GetWrd(3779) '"secs"
                End If
                strResult = _secondText
            Case Else
                If (String.IsNullOrEmpty(_hoursText)) Then
                    _hoursText = GetWrd(3777) '"hrs"
                End If
                strResult = _hoursText
        End Select
#If TRACE Then
        Log.UTILITY_LOW(String.Format("Exit ({0})", strResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strResult
    End Function

    ''' <summary>
    ''' Dispose the EASEClass library from memory
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose()

#If TRACE Then
        Dim startTicks As Long = Log.Trace29("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Ec = Nothing
        gObjAppConfig = Nothing
        gObjQF = Nothing
        gObjCurrentDB = Nothing

#If TRACE Then
        Log.Trace29("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Check the EASE configuration files (easesys.mdb) attributes and 
    ''' if the files are read-only, change it read-write.
    '''     </summary>
    ''' <param name="strMsg">Error message, if any (BYREF)</param>
    ''' <returns>True/False, result of this process</returns>
    ''' <remarks></remarks>
    Public Function CheckEASEConfigFiles(Optional ByRef strMsg As String = "") As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.FILE_DIR_IO("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False

        dim strFileName = GetEASESysDatabaseName()
        If Not SetFileAttributes(strFileName, strMsg) Then GoTo ExitThisFunction

        blnRtnValue = True

        ExitThisFunction:

#If TRACE Then
        Log.FILE_DIR_IO(String.Format("Exit ({0}) strMsg:({1})", blnRtnValue, strMsg), LOG_APPNAME, BASE_ERRORNUMBER + 0,
                        startTicks)
#End If
        Return blnRtnValue
    End Function

    ''' <summary>
    ''' Set the file attribute to Normal (Read-Write)
    ''' </summary>
    ''' <param name="strFilename">filename with file path</param>
    ''' <param name="strMsg">Error message (BYREF)</param>
    ''' <returns>True/False, result of this process.</returns>
    ''' <remarks></remarks>
    Public Function SetFileAttributes(ByVal strFilename As String, Optional ByRef strMsg As String = "") As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.FILE_DIR_IO_LOW(String.Format("Enter strFilename:({0})", strFilename), LOG_APPNAME,
                                                     BASE_ERRORNUMBER + 0)
#End If


        Dim dDirectoryInfo As DirectoryInfo = New DirectoryInfo(strFilename)
        Dim DAttributes As FileAttributes = dDirectoryInfo.Attributes()
        Dim blnRtnValue As Boolean = False

        Try

            If Not File.Exists(strFilename) Then
                strMsg = "EASE Configuration File (" & strFilename & ") is missing."
                GoTo ExitThisFunction
            End If

            If DAttributes = FileAttributes.ReadOnly Then
                File.SetAttributes(strFilename, FileAttributes.Normal)
            End If
            blnRtnValue = True

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            blnRtnValue = False
            strMsg = "EASE configuration file (" & Trim(strFilename) & ") is read-only. " & vbCrLf & vbCrLf &
                     "Unable to set the file SetFileAttributes."
        End Try
        ExitThisFunction:


#If TRACE Then
        Log.FILE_DIR_IO_LOW(String.Format("Exit ({0}) strMsg:({1})", blnRtnValue, strMsg), LOG_APPNAME, BASE_ERRORNUMBER + 0,
                            startTicks)
#End If
        Return blnRtnValue
    End Function

    ''' <summary>
    ''' Set the Customer name flag for (Harley - Milwaukee, York, Kansas, Cummins)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ConfigureCustomerNameFlags()

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strCustomerName As String = gObjAppConfig.CustomerName
        Ec.AppConfig.HarleyMil = False
        Ec.AppConfig.HarleyYork = False
        Ec.AppConfig.Cummins = False
        Ec.AppConfig.EASE = False

        strCustomerName = UCase(strCustomerName)
        If InStr(strCustomerName, "HARLEY-DAVIDSON") > 0 Then
            Ec.AppConfig.HarleyMil = True
            If InStr(strCustomerName, "SOFTAIL") > 0 Or
               InStr(strCustomerName, "YORK") > 0 Or
               InStr(strCustomerName, "KANSAS") > 0 Then

                Ec.AppConfig.HarleyYork = True
            End If
        End If

        If InStr(strCustomerName, "CUMMINS") > 0 Then
            Ec.AppConfig.Cummins = True
        End If

        If InStr(strCustomerName, "ARMADA") > 0 Then
            Ec.AppConfig.Armada = True
        End If

        If InStr(strCustomerName, "DIESEL") > 0 Then
            Ec.AppConfig.DieselExchange = True
        End If

        If InStr(strCustomerName, "EASE") > 0 Then
            Ec.AppConfig.EASE = True
        End If


#If TRACE Then
        Log.SYSTEM_CONFIG_LOW("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function GetWordDescriptionFromDatabase(ByVal intWordID As Int32) As String
#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter intWordID:({0})", intWordID), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If


        Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
        Dim strSQL As String = "select descx from easesys1 where id=" & intWordID
        Dim returnValue = ""
        Using connection = New EaseCore.DAL.Connection(strCS)
            returnValue = connection.ExecuteScalar (Of String)(strSQL, "")
        End Using

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Sub ReadEaseSys2(Optional ByVal intID As Int32 = 0)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter intID:({0})", intID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        'open the EASESys2 table
        Dim strSQL As String = "", strCS As String = ""
        Dim strSys2TableName As String

        If EASEClass7.DBConfig.Metric = True Then
            strSys2TableName = "Sys2Metric"
        Else
            strSys2TableName = "Sys2Imperial"
        End If

        Dim obj_TmpEasesys2 As New DataTable(strSys2TableName), obj_Row As DataRow
        Dim intTempFileNo As Int16 = 0

        Dim PrimaryKeyColumns(0) As DataColumn
        Dim strCond As String = " where "

        Try

            strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            strSQL = "select * from " & strSys2TableName

            If intID > 0 Then
                strCond = " and "
                strSQL &= strCond & " id=" & intID
            End If
            strSQL &= " order by id"

            Using connection = New EaseCore.DAL.Connection(strCS)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then

                        obj_TmpEasesys2.Columns.Add("no", Type.GetType("System.Int16"))
                        obj_TmpEasesys2.Columns.Add("code", Type.GetType("System.String"))
                        obj_TmpEasesys2.Columns.Add("desc", Type.GetType("System.String"))
                        obj_TmpEasesys2.Columns.Add("tmu", Type.GetType("System.Int16"))

                        For Each reader As DataRow In table.Rows
                            Dim str_Tmp = Trim(Extensions.Data.GetDataRowValue(reader("code"), ""))
                            If Trim(str_Tmp) = "" Then GoTo SkipThisRecord

                            obj_Row = obj_TmpEasesys2.NewRow
                            obj_Row("no") = reader("id")
                            obj_Row("code") = str_Tmp
                            str_Tmp = Trim(Extensions.Data.GetDataRowValue(reader("descx"), ""))
                            obj_Row("desc") = str_Tmp

                            Dim intTemp = Extensions.Data.GetDataRowValue(reader("tmu"), 0)
                            obj_Row("tmu") = intTemp

                            obj_TmpEasesys2.Rows.Add(obj_Row)
                            SkipThisRecord:
                        Next
                    End If
                End Using
            End Using

            PrimaryKeyColumns(0) = obj_TmpEasesys2.Columns("no")
            obj_TmpEasesys2.PrimaryKey = PrimaryKeyColumns

            gObjEasesys2 = obj_TmpEasesys2  'Set the Global Variable
            gBlnEASESys2Loaded = True
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("ReadEaseSys2", ex)
            gBlnEASESys1Loaded = False
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Gets the mtmrec for the passed record position from easesys2.dat
    ''' </summary>
    ''' <param name="intmtm">Record  position</param>
    ''' <returns>Description field for the record position</returns>
    ''' <remarks></remarks>
    Public Function GetMTM(ByVal intMTM As Int32) As stMTMRec


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter intMTM:({0})", intMTM), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If

        '**************************************************************************
        'Getmtm function returns value only for Client Server Applications
        'Web Applications use local Getmtm function instead
        '**************************************************************************


        'Get the data from easesys2.dat (which is already open)

        Dim str_RtnValue As String = ""
        Dim objMTM As stMTMRec
        objMTM.Code = ""
        objMTM.Desc = ""
        objMTM.TMU = "0"


        If Not gBlnEASESys2Loaded Then GoTo ExitThisFunction 'make sure the easesys2 object is loaded

        Dim str_Search As String = CStr(intMTM)
        Dim obj_Row As DataRow

        Try

            obj_Row = gObjEasesys2.Rows.Find(str_Search)
            If Not (obj_Row Is Nothing) Then
                objMTM.Desc = Extensions.Data.GetDataRowValue(obj_Row("desc"), "")
                objMTM.Code = Extensions.Data.GetDataRowValue(obj_Row("code"), "")
                objMTM.TMU = Extensions.Data.GetDataRowValue(obj_Row("tmu"), "")
            End If
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("LoadEaseSys2", ex)
        Finally
            obj_Row = Nothing
        End Try

        ExitThisFunction:


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", objMTM), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objMTM
    End Function

    ''' <summary>
    ''' Unloads the object which holds the easesys2.dat records
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UnLoadEaseSys2()

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        gBlnEASESys2Loaded = False

        gObjEasesys2.Dispose()
        gObjEasesys2 = Nothing

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function GetNumberOfLicenses() As Int16

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'GetUserLicense, GetApplicationLicense,getlicensecount,GetUserLicense, GetApplicationLicense

        Dim intLen As Int16 = 0, intWL As Int16 = 0, intType As Int16 = 0, intLicense As Int16 = 0
        Ec.AppConfig.GetWrd(6, intWL, intLen, intType)
        If intLen > 0 And intType > 0 Then
            intLicense = intLen*intType
        End If
        If intLicense = 0 Then intLicense = 5


#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", intLicense), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return intLicense
    End Function

    Public Function GetNumberOfPlants() As Int16

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intPlantsCount As Int16 = 0
        Try
            GetWrd(7, , , intPlantsCount)
        Catch ex As Exception
            Call GenerateException("GetNumberOfPlants", ex)
        End Try
        If intPlantsCount <= 0 Then intPlantsCount = 1


#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", intPlantsCount), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return intPlantsCount
    End Function


#Region "***  EASESys Configuration ** Access EASESys.MDB **"

    Public Sub UpdateEaseSys1_RouteHeaderFields(ByRef objEaseSys1NEW As DataTable)


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intDP As Int16 = 0, intType As Int16 = 0
        Dim strWord As String = ""
        Dim strReplaceWith As String = ""

        Try
            'Header Field Type: 0-Plan Header, 1-Operation Header, 2-Element Header, 3-Plan Extra fields, 4-Operation Extra fields
            '                 : 5-Tooling records, 6-Machines, 7-Work Center, 8- Custom ODBC Link,9-Material User Defined Fields
            '                 : 10- Plan Display Fields (used only for reporting and display purpose)
            '                 : 11- Operation Display Fields (used only for reporting and display purpose)
            '                 : 12 - Shared Operation header fields, 13- MCR fields, 14 - MDM Editor - Shared Documents
            '                 : 15 - Location Document fields,  16- FMEA- Doc Hdr- General , 17 -FMEA header; 
            '                 : 18-  Control plan header ,  19-LPA header, 20- StdWorkHeader, 21- StdwrkSignoff header
            '                 : 22-MES Header , 23- FMEA-DocHdr-FMEA , 24- FMEA DocHdr MQV, 25- FMEA DocHdr Control plan

            intHeaderType = 0         'plan Header
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)

            intWL = 0
            intType = 0
            For intK = 1 To UBound(objHeaderFields)

                Dim intKeyX = 23 + objHeaderFields(intK).FieldSeq 'intK
                strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                strWord = objHeaderFields(intK).UserDefName.Trim

                'if there is no change, don't update the word
                If strWord.Trim = strFind.Trim Then GoTo SkipUpdtingWordsInMemory

                Dim intL = strWord.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
                '-1 to update all fields

                SkipUpdtingWordsInMemory:

                'update similiar words for each field
                Select Case intKeyX
                    Case 24, 25
                        strReplaceWith = objHeaderFields(intK).UserDefName.Trim    ' updating similiar wordsused in 

                        'if same words, don't update the file
                        If strFind.Trim = strReplaceWith.Trim Then GoTo SkipUpdatingField1
                        UpdateSimiliarWords(objEaseSys1NEW, intKeyX, strFind, strReplaceWith)
                        SkipUpdatingField1:
                        If intKeyX = 24 Then
                            'update partnumber abbreviation 1
                            intKeyX = 4206
                            strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                            strReplaceWith = objHeaderFields(intK).Abbrv1

                            'if same words, don't update the file
                            If strFind.Trim = strReplaceWith.Trim Then GoTo SkipUpdatingField2

                            UpdateSimiliarWords(objEaseSys1NEW, intKeyX, strFind, strReplaceWith)
                            SkipUpdatingField2:
                        End If
                End Select
            Next intK
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_RouteHeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_OperationHeaderFields(ByRef objEaseSys1NEW As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intDP As Int16 = 0, intType As Int16 = 0
        Dim strWord As String = ""
        Dim strReplaceWith As String = ""
        Try
            intHeaderType = 1         'Operation Header
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)

            intWL = 0
            intType = 0
            For intK = 1 To UBound(objHeaderFields)
                Dim intKeyX = 48 + objHeaderFields(intK).FieldSeq 'intK
                strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX) ' updating similiar wordsused in 

                strWord = objHeaderFields(intK).UserDefName.Trim

                'if there is no change, don't update the word
                If strWord.Trim = strFind.Trim Then GoTo SkipUpdtingWordsInMemory

                Dim intL = strWord.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
                '-1 to update all fields
                SkipUpdtingWordsInMemory:
                'update similiar words for each field
                Select Case intKeyX
                    Case 49, 51
                        strReplaceWith = objHeaderFields(intK).UserDefName.Trim

                        'if same words, don't update the file
                        If strFind.Trim = strReplaceWith.Trim Then GoTo SkipUpdatingField1
                        'strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX) : strReplaceWith = objHeaderFields(intK).UserDefName.Trim
                        UpdateSimiliarWords(objEaseSys1NEW, intKeyX, strFind, strReplaceWith)
                        SkipUpdatingField1:
                        If intKeyX = 49 Then
                            'update Operation - Abbreviation 1
                            intKeyX = 4218
                            strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                            strReplaceWith = objHeaderFields(intK).Abbrv1

                            If strFind.Trim = strReplaceWith.Trim Then GoTo SkipUpdatingField2 _
                            'if same words, don't update the file
                            UpdateSimiliarWords(objEaseSys1NEW, intKeyX, strFind, strReplaceWith)
                            SkipUpdatingField2:

                            'update Op - Abbreviation 2
                            intKeyX = 4370 '4018
                            strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                            strReplaceWith = objHeaderFields(intK).Abbrv2

                            If strFind.Trim = strReplaceWith.Trim Then GoTo SkipUpdatingField3 _
                            'if same words, don't update the file
                            UpdateSimiliarWords(objEaseSys1NEW, intKeyX, strFind, strReplaceWith)

                            SkipUpdatingField3:
                        End If
                End Select
            Next
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_OperationHeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_ElementHeaderFields(ByRef objEaseSys1NEW As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intDP As Int16 = 0, intType As Int16 = 0
        Dim strReplaceWith As String = ""
        Try
            intHeaderType = 2         'Element Header
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            intWL = 0
            intType = 0
            For intK = 1 To UBound(objHeaderFields)
                Dim intKeyX = 2201 + intK
                Dim strWord = objHeaderFields(intK).UserDefName.Trim
                Dim intL = strWord.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
                '-1 to update all fields
            Next
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_ElementHeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_SubHeaderFields(ByRef objEaseSys1NEW As DataTable)


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim intWL As Int16 = 0, intL As Int16 = 0, intDP As Int16 = 0, intType As Int16 = 0
        Dim intKeyX As Int16 = 0, strWord As String = "", intK As Int16 = 0
        Dim strReplaceWith As String = "", blnRtnValue As Boolean = False
        Try
            intKeyX = 3698  '3698 is already updated (in ClientEditor-> frmCustomeWords)
            strWord = GetWrd(intKeyX, intWL, intL, intType, intDP, True)

            'get the word from Client table 
            strFind = Ec.DBConfig.ReadFromClientTableRecord(132)
            strReplaceWith = strWord
            If Trim(strFind) <> Trim(strReplaceWith) Then
                UpdateSimiliarWords(objEaseSys1NEW, intKeyX, strFind, strReplaceWith)
            End If

            'Update the Database
            blnRtnValue = UpdateEaseSys1Table(objEaseSys1NEW, GetEASESysDatabaseName())

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_SubHeaderFields", ex)
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_SharedOperation(ByRef objEaseSys1NEW As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0
        Dim strReplaceWith As String = ""
        Try
            intHeaderType = 12         'Shared Operation
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            intWL = 0
            intType = 0
            For intK = 1 To UBound(objHeaderFields)
                Dim intKeyX = 4234 + intK
                Dim strWord = objHeaderFields(intK).UserDefName.Trim
                Dim intL = strWord.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
                '-1 to update all fields
            Next

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_SharedOperation", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_MDMSharedDocsHeaderFields(ByRef objEaseSys1NEW As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0
        Dim strReplaceWith As String = ""
        Try
            intHeaderType = 14         'MDM Editor shared docs header fields
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            intWL = 0
            intType = 0
            For intK = 1 To UBound(objHeaderFields)
                Dim intKeyX = 5301 + intK
                Dim strWord = objHeaderFields(intK).UserDefName.Trim
                Dim intL = strWord.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
                '-1 to update all fields
            Next

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_MDMSharedDocsHeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_MDMLocationDocsHeaderFields(ByRef objEaseSys1NEW As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0
        Dim strReplaceWith As String = ""
        Try
            intHeaderType = 15         'MDM Editor Location docs header fields
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            intWL = 0
            intType = 0
            For intK = 1 To UBound(objHeaderFields)
                Dim intKeyX = 5409 + intK
                Dim strWord = objHeaderFields(intK).UserDefName.Trim
                Dim intL = strWord.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
                '-1 to update all fields
            Next

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_MDMLocationDocsHeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_MCRHeaderFields(ByRef objEaseSys1NEW As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0, strFind As String = ""
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0
        Dim strReplaceWith As String = ""
        Try
            intHeaderType = 13         'MCR
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            intWL = 0
            intType = 0
            For intK = 1 To UBound(objHeaderFields)
                Dim intKeyX = 0
                Select Case intK
                    Case Is <= 13
                        intKeyX = 1399 + intK
                    Case 14
                        intKeyX = 1609
                    Case 15
                        intKeyX = 1610
                End Select
                Dim strWord = objHeaderFields(intK).UserDefName.Trim
                Dim intL = strWord.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
                '-1 to update all fields
            Next

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_MCRHeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_StationCheckListHdr_HeaderFields(ByRef objEaseSys1NEW As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0
        Try
            intHeaderType = 26         'Station CheckList Hdr
            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            intWL = 0
            intType = 0
            Dim intKeyX = 5610
            For intK = 1 To UBound(objHeaderFields)

                Dim strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                Dim strReplaceWith = objHeaderFields(intK).UserDefName.Trim
                If strFind = strReplaceWith Then GoTo SkipUpdatingWordInMemory
                Dim intL = strReplaceWith.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strReplaceWith, strReplaceWith, intL, intWL, intType, - 1) _
                '-1 to update all fields
                SkipUpdatingWordInMemory:
                intKeyX += 1
            Next

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_MCRHeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub
    '*** This sub is shared among FMEA CP HDR, FMEA Hdr and Control Plan Header fields for FMEA and also for LPA Header fields*********'
    Public Sub UpdateEaseSys1_FMEA_LPA_HdrFields(ByRef objEaseSys1NEW As DataTable, ByVal strParam As String)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter strParam:({0})", strParam), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If
        Dim intHeaderType As Int16 = 0
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0
        Try
            Dim intKeyX = 0
            Select Case strParam
                'Pr - new design 10 Oct (MQV)
                'Case "FMEA-CPHEADER"
                '    intHeaderType = 16
                '    intKeyX = 5601   

                ' intKEYX --> Index of user defined name in word editor for corr type
                Case "FMEA-DOCHDR-GENERAL"
                    intHeaderType = 16
                    intKeyX = 6601
                Case "FMEA-DOCHDR-FMEA"
                    intHeaderType = 23
                    intKeyX = 6614
                Case "FMEA-DOCHDR-MQV"
                    intHeaderType = 24
                    'intKeyX = 6628
                    intKeyX = 6630                ''Design change- remove BIS Data Range; warranty data range
                Case "FMEA-DOCHDR-CP" 'Doc header for control plan
                    intHeaderType = 25
                    intKeyX = 6640
                Case "FMEA-HEADER"
                    intHeaderType = 17
                    intKeyX = 5626
                Case "FMEA-CTRLPLANHEADER"
                    intHeaderType = 18
                    intKeyX = 5648
                Case "LPA-HEADER"
                    intHeaderType = 19
                    intKeyX = 5665
            End Select

            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            For intK = 1 To UBound(objHeaderFields)
                Dim strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                Dim strReplaceWith = objHeaderFields(intK).UserDefName.Trim
                If strFind = strReplaceWith Then GoTo SkipUpdatingWordInMemory
                Dim intL = strReplaceWith.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strReplaceWith, strReplaceWith, intL, intWL, intType, - 1) _
                '-1 to update all fields
                SkipUpdatingWordInMemory:
                intKeyX += 1
            Next
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_FMEA_HeaderFields", ex)
        Finally
            objHeaderFields = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_StdWorkHeaderFields(ByRef objEaseSys1NEW As DataTable, ByVal strParam As String)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter strParam:({0})", strParam), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If

        Dim intHeaderType As Int16 = 0
        Dim objHeaderFields(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0
        Try
            Dim intKeyX = 0
            Select Case strParam
                Case "STDWORK-HEADER"
                    intHeaderType = 20
                    intKeyX = 5689          ' Index of user defined name for Std work header in word file starts @ index 5689 
                Case "STDWORK-SIGNOFFHEADER"
                    intHeaderType = 21
                    intKeyX = 4535 _
                    'Index of user defined name for StdWork- Signoff header in word file starts @ index 4535 
            End Select

            objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            For intK = 1 To UBound(objHeaderFields)
                '*****26 July 2013- To implement stdwork changes that are applicable only for cummims-CMI as per new requirement 
                'Cummins-CMI has 13 header fields; other customers have only 5; 
                'Fields 6 to 13 are included later & r not in consecutive places with those initial 5 fields; Cover it 

                If strParam = "STDWORK-HEADER" And intHeaderType = 20 And intK = 6 Then intKeyX = 7150

                Dim strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                Dim strReplaceWith = objHeaderFields(intK).UserDefName.Trim
                If strFind = strReplaceWith Then GoTo SkipUpdatingWordInMemory
                Dim intL = strReplaceWith.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strReplaceWith, strReplaceWith, intL, intWL, intType, - 1) _
                '-1 to update all fields
                SkipUpdatingWordInMemory:
                intKeyX += 1
            Next
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_StdWorkHeaderFields", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_StdWork_Template_HeaderFields(ByRef objEaseSys1NEW As DataTable, ByVal intTemplateID As Int16,
                                                            ByVal intPlantID As Int16)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(
            String.Format("Enter intTemplateID:({0}) intPlantID:({1})", intTemplateID, intPlantID), LOG_APPNAME,
            BASE_ERRORNUMBER + 0)
#End If


        Dim objStdWrkTemplateHF(0) As EASEClass7.DBConfig.stHeaderFields
        Dim intWL As Int16 = 0, intType As Int16 = 0

        Try
            Dim intKeyX = 0
            If intTemplateID = 1 Then
                intKeyX = 4544          'User defined name STD WORK -ASM template hdr  in word file starts @ index 4544 
            ElseIf intTemplateID = 2 Then
                intKeyX = 4581            'User defined name STD WORK -MACH template hdr  in word file starts @ index 4581
                '*****26 July 2013-Cummins-CMI-Applicable only for cummins - Introduced 2 new templates
            ElseIf intTemplateID = 3 Then
                intKeyX = 3891          'User defined name STD WORK -CMI Assembly template hdr  in word file starts @ index 3891
            ElseIf intTemplateID = 4 Then
                intKeyX = 3896 _
                'User defined name STD WORK -CMI Machining template hdr  in word file starts @ index 3896
            End If


            'objStdWrkTemplateHF = Ec.MDM.GetHeaderFields_StdWork_Template(intTemplateID, intPlantID)
            objStdWrkTemplateHF = Ec.StandardWork.GetHeaderFields_StdWork_Template(intTemplateID, intPlantID)
            For intK = 1 To UBound(objStdWrkTemplateHF)
                Dim strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                Dim strReplaceWith = objStdWrkTemplateHF(intK).UserDefName.Trim
                If strFind = strReplaceWith Then GoTo SkipUpdatingWordInMemory
                Dim intL = strReplaceWith.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strReplaceWith, strReplaceWith, intL, intWL, intType, - 1) _
                '-1 to update all fields
                SkipUpdatingWordInMemory:
                intKeyX += 1
            Next
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_StdWork_Template_HeaderFields", ex)
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub
    '***** Updates the word file in memory with the default words fetched from headers table******'
    Public Sub UpdateEaseSys1_MESHeaderFields(ByRef objEaseSys1NEW As DataTable)


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intWL As Int16 = 0, intType As Int16 = 0

        Try
            Dim intHeaderType = 22
            Dim intKeyX = 6221          ' Index of user defined name for System header in word file starts @ index 6221

            Dim objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            For intK = 1 To UBound(objHeaderFields)
                Dim strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                Dim strReplaceWith = objHeaderFields(intK).UserDefName.Trim
                If strFind = strReplaceWith Then GoTo SkipUpdatingWordInMemory
                Dim intL = strReplaceWith.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strReplaceWith, strReplaceWith, intL, intWL, intType, - 1) _
                '-1 to update all fields
                SkipUpdatingWordInMemory:
                intKeyX += 1
            Next
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_MESHeaderFields", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEaseSys1_MDMImageSheet_HeaderFields(ByRef objEaseSys1NEW As DataTable)


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intWL As Int16 = 0, intType As Int16 = 0

        Try
            Dim intHeaderType = 28
            Dim intKeyX = 6718          ' Index of user defined name for System header in word file starts @ index 6718

            Dim objHeaderFields = Ec.DBConfig.GetHeaderFields(intHeaderType)
            For intK = 1 To UBound(objHeaderFields)
                Dim strFind = GetWordfromMemory(objEaseSys1NEW, intKeyX)
                Dim strReplaceWith = objHeaderFields(intK).UserDefName.Trim
                If strFind = strReplaceWith Then GoTo SkipUpdatingWordInMemory
                Dim intL = strReplaceWith.Trim.Length
                UpdateWordInMemory(objEaseSys1NEW, intKeyX, strReplaceWith, strReplaceWith, intL, intWL, intType, - 1) _
                '-1 to update all fields
                SkipUpdatingWordInMemory:
                intKeyX += 1
            Next
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1_MDMImageSheetHeaderFields", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function UpdateEaseSys1Words(ByVal strNewWordFileName As String,
                                        Optional ByVal blnUpdateWordFileInMemory As Boolean = False,
                                        Optional ByVal strLastUpdatedDateTime As String = "") As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(
            String.Format("Enter strNewWordFileName:({0}) blnUpdateWordFileInMemory:({1}) strLastUpdatedDateTime:({2})",
                          strNewWordFileName, blnUpdateWordFileInMemory, strLastUpdatedDateTime),
            LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If


        'This sub is used to update the systemupdate.new file. Updates the word file in word object and updates the tables

        'if blnUpdateWordFileInMemory=true, the update the word file in memory (instead of systemupdate.new)

        'update the following sections/words in the easesys1 table
        '1. Customer name
        '2. License info (viewease)
        '3. Plan Header Fields
        '4. Operation Header Fields
        '5. Element Header fields
        '6. Sub Header caption
        '7. Plan header screen caption
        '8. MCR Fields
        '9. Update Shared Operation Header fields
        '10. Update MDM, location document header fields

        Dim blnRtnValue As Boolean = False
        Dim objEaseSys1NEW As New DataTable("easesys1")
        Dim strLastUpdatedExistingFile As String = "", strLastUpdatedNewFile As String = ""
        Dim strTemp As String = "", intKeyX = 0, strWord As String = ""
        Dim lngTemp1 As Long = 0, lngTemp2 As Long = 0, strErrLocation As String = "Start"
        Dim intWL As Int16 = 0, intL As Int16 = 0, intDP As Int16 = 0, intType As Int16 = 0
        Dim intWordsCount As Int32 = 0, intWordsCountNEW As Int32 = 0
        Dim strDefFileName As String = "", blnCheckForNewWords As Boolean = False
        Try


            If blnUpdateWordFileInMemory Then
                strErrLocation = "GetEASESysDatabaseName"

                strNewWordFileName = GetEASESysDatabaseName()
                objEaseSys1NEW = gObjEasesys1
                If objEaseSys1NEW.Rows.Count = 0 Then GoTo ExitThisFunction
                GoTo UpdateHeaderFields
            Else
                strErrLocation = "Load Words from SystemUpdate.NEW"

                'default option, while updating systemupdat.new
                'get the list of words from the new words file
                objEaseSys1NEW = Ec.AppConfig.GetEasesys1Words(, True, strNewWordFileName)
                If objEaseSys1NEW.Rows.Count = 0 Then GoTo ExitThisFunction
                blnCheckForNewWords = True
            End If


            strErrLocation = "Check Last updated Date Time Stamp"
            strLastUpdatedExistingFile = GetWrd(16, , , , , True)       'get the last updated date and time from the word file
            strLastUpdatedNewFile = GetWordfromMemory(objEaseSys1NEW, 16)

            'the new file is missing the last updated date and time, may be corrupt, do nothing
            If strLastUpdatedNewFile.Trim = "" Then GoTo ExitThisFunction

            'Backup the existing easesys.mdb. IF any error occurred, we can use the file to recover, brilliant huh!.
            strErrLocation = "Backup existing easesys.mdb"
            strDefFileName = GetEASESysDatabaseName()
            strTemp = Ec.GeneralFunc.GetDirectoryPath(strDefFileName) & "easesys-" & Ec.GeneralFunc.GetRandomFileName(True) &
                      ".mdb"
            Ec.GeneralFunc.FileCopy(strDefFileName, strTemp)


            '1. Customer name
            strErrLocation = "Update: Customer Name"
            intKeyX = 4         'Customer Name (OLD)
            strWord = GetWrd(intKeyX, intWL, intL, intType, intDP, True)
            UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, "", intL, intWL, intType, - 1)      '-1 to update all fields

            intKeyX = 7         'Customer Name (NEW)
            strWord = GetWrd(intKeyX, intWL, intL, intType, intDP, True)
            UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, "", intL, intWL, intType, - 1)      '-1 to update all fields

            intKeyX = 8         'Customer Name (NEW)
            strWord = GetWrd(intKeyX, intWL, intL, intType, intDP, True)
            UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, "", intL, intWL, intType, - 1)      '-1 to update all fields
            intKeyX = 9         'Customer Name (NEW)
            strWord = GetWrd(intKeyX, intWL, intL, intType, intDP, True)
            UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, "", intL, intWL, intType, - 1)      '-1 to update all fields

            '2. License info (viewease)
            intKeyX = 6     'View License details
            strWord = GetWrd(intKeyX, intWL, intL, intType, intDP, True)
            UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
            '-1 to update all fields

            UpdateHeaderFields:
            '3. Plan Header Fields
            strErrLocation = "UpdateEaseSys1_RouteHeaderFields"
            UpdateEaseSys1_RouteHeaderFields(objEaseSys1NEW)

            '4. Operation Header Fields
            strErrLocation = "UpdateEaseSys1_OperationHeaderFields"
            UpdateEaseSys1_OperationHeaderFields(objEaseSys1NEW)

            '5. Element Header fields
            strErrLocation = "UpdateEaseSys1_ElementHeaderFields"
            UpdateEaseSys1_ElementHeaderFields(objEaseSys1NEW)

            '6. Sub Header caption
            strErrLocation = "UpdateEaseSys1_SubHeaderFields"
            UpdateEaseSys1_SubHeaderFields(objEaseSys1NEW) _
            '** Also updates the database *** , the update flag is reset ** intermediate save *****

            '7. Plan header screen caption
            strErrLocation = "Update Plan Header Caption"
            intKeyX = 23     'Route Header Screen Caption
            strWord = GetWrd(intKeyX, intWL, intL, intType, intDP, True)
            UpdateWordInMemory(objEaseSys1NEW, intKeyX, strWord, strWord, intL, intWL, intType, - 1) _
            '-1 to update all fields

            '8. MCR Fields
            If Ec.AppConfig.MCR Then
                strErrLocation = "UpdateEaseSys1_MCRHeaderFields"
                UpdateEaseSys1_MCRHeaderFields(objEaseSys1NEW)
            End If


            '9. Update Shared Operation Header fields
            strErrLocation = "UpdateEaseSys1_SharedOperation"
            UpdateEaseSys1_SharedOperation(objEaseSys1NEW)


            If Ec.AppConfig.MDM Then
                '10. Update MDM location document, Shared document header fields
                strErrLocation = "UpdateEaseSys1_MDMLocationDocsHeaderFields"
                UpdateEaseSys1_MDMLocationDocsHeaderFields(objEaseSys1NEW)

                strErrLocation = "UpdateEaseSys1_MDMSharedDocsHeaderFields"
                UpdateEaseSys1_MDMSharedDocsHeaderFields(objEaseSys1NEW)
            End If

            ''====================================================================================
            'If Ec.AppConfig.MDM Then
            '    strErrLocation = "UpdateEaseSys1_FMEA_LPA_HdrFields"
            '    UpdateEaseSys1_FMEA_LPA_HdrFields(objEaseSys1NEW, "FMEA-DOCHDR-GENERAL")
            '    UpdateEaseSys1_FMEA_LPA_HdrFields(objEaseSys1NEW, "FMEA-DOCHDR-FMEA")
            '    UpdateEaseSys1_FMEA_LPA_HdrFields(objEaseSys1NEW, "FMEA-DOCHDR-MQV")
            '    UpdateEaseSys1_FMEA_LPA_HdrFields(objEaseSys1NEW, "FMEA-DOCHDR-CP")
            '    UpdateEaseSys1_FMEA_LPA_HdrFields(objEaseSys1NEW, "FMEA-HEADER")
            '    UpdateEaseSys1_FMEA_LPA_HdrFields(objEaseSys1NEW, "FMEA-CTRLPLANHEADER")
            '    UpdateEaseSys1_FMEA_LPA_HdrFields(objEaseSys1NEW, "LPA-HEADER")
            'End If

            '' Include the check for Std Work If Ec.Appconfig.StandardWork then
            'strErrLocation = "UpdateEaseSys1_StdWorkHeaderFields"
            'UpdateEaseSys1_StdWorkHeaderFields(objEaseSys1NEW, "STDWORK-HEADER")

            'strErrLocation = "UpdateEaseSys1_StdWork_Template_HeaderFields"
            'UpdateEaseSys1_StdWork_Template_HeaderFields(objEaseSys1NEW, 1, 1)

            'If Ec.AppConfig.MESSystem Then
            '    strErrLocation = "UpdateEaseSys1_MESHeaderFields"
            '    UpdateEaseSys1_MESHeaderFields(objEaseSys1NEW)
            'End If
            '====================================================================================


            If blnCheckForNewWords Then
                strErrLocation = "Add New Words"
                ''check for new words in the systemupdate.new
                intWordsCount = GetTotalNumberOfWordsInEaseSys1()       'get the number of records from existing easesys1 

                intWordsCountNEW = GetTotalNumberOfWordsInEaseSys1(strNewWordFileName) _
                'get the number of records from existing easesys1 
                If intWordsCountNEW > intWordsCount Then
                    'new words found
                    intL = 0
                    intWL = 0
                    intType = 0
                    For intK = intWordsCount + 1 To intWordsCountNEW
                        intKeyX = intK

                        'update the "updateflag" field in the table to insert new records
                        UpdateWordInMemory(objEaseSys1NEW, intKeyX, "", "", intL, intWL, intType, - 99)      '-99 - New Record
                    Next
                End If
            End If

            'update the last updated date and time stamp
            strErrLocation = "UpdateLastUpdatedDateTimeInEASESys1"
            UpdateLastUpdatedDateTimeInEASESys1(objEaseSys1NEW, strLastUpdatedDateTime)

            'Update the Database
            strErrLocation = "UpdateEaseSys1Table"
            blnRtnValue = UpdateEaseSys1Table(objEaseSys1NEW, GetEASESysDatabaseName())

            ExitThisFunction:
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateEaseSys1Words: " & strErrLocation, ex)
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Private Sub UpdateLastUpdatedDateTimeInEASESys1(ByRef objEaseSys1NEW As DataTable, ByVal strLastUpdatedDateTime As String)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter strLastUpdatedDateTime:({0})", strLastUpdatedDateTime),
                                                    LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'update the easesys1 table with last update date and time
        Dim intKeyX As Int16 = 0
        intKeyX = 16
        If strLastUpdatedDateTime.Trim = "" Then strLastUpdatedDateTime = GetCurrentDateTimeForWordUpdate()
        UpdateWordInMemory(objEaseSys1NEW, intKeyX, strLastUpdatedDateTime, strLastUpdatedDateTime, 0, 0, 0, - 1)


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function UpdateEaseSys1Table(ByRef objEaseSys1NEW As DataTable, ByVal strWordFileName As String,
                                        Optional ByVal strLastUpdatedDateTime As String = "",
                                        Optional ByVal blnUpdateAllDATA As Boolean = False) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(
            String.Format("Enter strWordFileName:({0}) strLastUpdatedDateTime:({1}) blnUpdateAllDATA:({2})", strWordFileName,
                          strLastUpdatedDateTime, blnUpdateAllDATA), LOG_APPNAME,
            BASE_ERRORNUMBER + 0)
#End If

        'Update the words in Easesys1 table

        Dim intQueryType As QueryBuilder.QueryType = QueryBuilder.QueryType.Insert, blnRtnValue As Boolean = False
        Dim dtDate As Date = EaseDate.Now

        Try

            If strLastUpdatedDateTime.Trim <> "" Then 'update the last updated date and time in easesys1 word file
                UpdateLastUpdatedDateTimeInEASESys1(objEaseSys1NEW, strLastUpdatedDateTime)
            End If

            'OK, now we got all the words in the object, update the words in the file
            dim strCS = EaseCore.DAL.Connection.BuildConnectionString(Connection.DatabaseTypes.Access, GetEASESysDatabaseName)            

            Dim sqlList = New List(Of String)()

            Using connection = New EaseCore.DAL.Connection(strCS)
                For Each objRow As DataRow In objEaseSys1NEW.Rows
                    If Not Extensions.Data.GetDataRowValue(objRow("updateflag"), False) Then GoTo SkipThisWord
                    dim intKeyX = Extensions.Data.GetDataRowValue (Of Int16)(objRow("no"), 0)
                    dim intDP = Extensions.Data.GetDataRowValue (Of Int16)(objRow("dp"), 0)

                    If intKeyX >= 4600 And intKeyX <= 4900 Then GoTo SkipThisWord _
                    'skip updating the application constants. this records must be updated only in work editor
                    intQueryType = EASEClass7.QueryBuilder.QueryType.Update 'update query

                    Dim queryBuilder As EASEClass7.QueryBuilder
                    If intQueryType = QueryBuilder.QueryType.Insert Or intDP = - 99 Then 'insert new words
                        queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(intQueryType, "easesys1")
                        queryBuilder.AddField("id", intKeyX, True)
                        queryBuilder.AddField("descx", objRow("desc"), False, False)
                        queryBuilder.AddField("foreigndescx", objRow("intldesc"), False, False)
                        queryBuilder.AddField("wl", objRow("wl"), True)
                        queryBuilder.AddField("Len", objRow("l"), True)
                        queryBuilder.AddField("type", objRow("type"), True)
                        queryBuilder.AddField("dp", 0, True)
                    Else 'update existing record (DEFAULT OPTION)
                        queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(intQueryType, "easesys1")
                        queryBuilder.AddField("descx", objRow("desc"), False, False)
                        queryBuilder.AddField("wl", objRow("wl"), True)
                        queryBuilder.AddField("Len", objRow("l"), True)
                        queryBuilder.AddField("type", objRow("type"), True)
                        queryBuilder.AddConditionField("id", intKeyX.ToString(), True)

                        'FieldExist() takes connstr corr to actualDB(orcl/acc/sql) where as we need conn str corr to easesys.mdb. so pass it explicitly. Change the DataSource as needed
                        If DAL.Schema.FieldExist(connection, "easesys1", "LastUpdated") Then
                            queryBuilder.AddField("LastUpdated", dtDate, False) _
                            'Fill date - needed while export/import of easesys.mdb to/from excel
                        End If
                    End If

                    sqlList.Add(queryBuilder.GenerateQuery)
                    SkipThisWord:
                Next
                If (sqlList.Count > 0) Then
                    connection.ExecuteNonQuery(sqlList)
                End If
            End Using

            'ok no errors occurred, now reset the updateflag in the global variable
            objEaseSys1NEW.AcceptChanges()
            For Each objRow As DataRow In objEaseSys1NEW.Rows
                If Not Extensions.Data.GetDataRowValue(objRow("updateflag"), False) Then GoTo SkipThisWord1

                objRow.BeginEdit()
                objRow("updateflag") = False
                objRow.EndEdit()
                SkipThisWord1:
            Next
            blnRtnValue = True
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Private Function GetWordsList_PartDesc() As Int32() 'get the list of words to update for Part Description

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(1) As Int32
        ArrWords(1) = 25
        'ArrWords(2) = 4613

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_OperationNumber() As Int32() 'get the list of words to update for Part Description

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(5) As Int32
        ArrWords(1) = 206
        ArrWords(2) = 49
        ArrWords(3) = 1899
        ArrWords(4) = 3942
        ArrWords(5) = 6351

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit Count:({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_Plant() As Int32() 'get the list of words to update for Part Description

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(0) As Int32
        Dim intK = 0

        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 1403  'The word is Plant, change it to 'Test Plant', the replace function updates to 'Test Test Plant', so ignore the first record, already covered.Km : 4/9/2014

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 216
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 218
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 488
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 1403
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1419
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1441
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1614
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1615

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2724
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3062
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3063
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3081
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3083
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3084
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3543
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3902
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7258
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7259 '4239
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4325
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4327
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 4663      application constant, don't change it
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5434
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5435
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5436
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5438
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5440
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5441
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5443
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5444

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5713
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5998
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 6018
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 6824
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7097
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7114
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 9005

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_SubHeader() As Int32()

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(0) As Int32
        Dim intK = 0

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 37
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 184
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 252
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 475
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 485

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1442
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1443

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1874
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1875
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1876
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1898
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1900
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1942
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1943
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1985
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1986
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1987
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1995
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1996

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2006
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2007
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2008
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2009
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2010
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2018
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2039
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2042
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2045
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2046
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2049
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2059
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2060
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2061
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2066
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2073
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2075
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2115
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2119

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2144
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2145
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2146

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2733
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2734
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2793
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2807
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2808
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2809
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2848
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2983

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3528
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3566
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3698
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3871
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3882
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3883

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4026
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4137
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4140
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4316
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4319
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4320
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4353
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4373
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4390
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4421
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4458
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4494
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4497
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4504
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4505
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4506
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4507
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4509
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4511
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4512
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4513
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4514

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4586
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4648
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 6355
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 6385
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7197

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 8719
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 9014

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit Count:({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_OperationWorkCenter() As Int32()

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(23) As Int32

        ArrWords(1) = 51
        ArrWords(2) = 214
        ArrWords(3) = 1810
        ArrWords(4) = 1870
        ArrWords(5) = 1884
        ArrWords(6) = 2053
        ArrWords(7) = 2065
        ArrWords(8) = 2148
        ArrWords(9) = 2176
        ArrWords(10) = 2184

        ArrWords(11) = 3670

        ArrWords(12) = 4278
        ArrWords(13) = 4361
        ArrWords(14) = 4362
        ArrWords(15) = 4363
        ArrWords(16) = 4364
        'ArrWords(1) = 4664      'application constants; restricted
        ArrWords(17) = 5446
        ArrWords(18) = 5450
        ArrWords(19) = 5451
        ArrWords(20) = 5452
        ArrWords(21) = 5469
        ArrWords(22) = 5471
        ArrWords(23) = 5475

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit Count:({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_OperationAbbrev2() As Int32() 'get the list of words to update for OP

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(0) As Int32, intK = 0

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 442
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1954
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1977

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1991
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1993
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1994
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2130
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2132
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2173
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2179
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2180
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2186
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5519
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7216
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 9011
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 9031
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 9032
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 9033


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_OperationAbbrev1() As Int32() 'get the list of words to update for Operation#


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(0) As Int32, intK = 0

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 49
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 115
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 116
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 206
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 212
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 248

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 253
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 277
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 318
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 322
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 344

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 356

        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 365
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 405
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 407
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 411
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 417
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 446
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 454
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 458
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 460
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 467
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 477
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 479

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1102

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1442

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1806
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1812
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1814
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1819
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1821
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1831
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1833
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1836
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1858
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1863
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1885
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1887
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1889
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1893
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1895
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1897
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1899
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1938
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1946
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1947
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1948

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1950

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1952
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1953
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1955
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1956
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1961
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1965
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1971
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1974
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1976
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1988


        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2001
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2002
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2003
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2004
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2041
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2044
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2048
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2051
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2052
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2056
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2059
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2060
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2062
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2066
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2070
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2071
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2077
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2078
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2082
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2083
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2084
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2085
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2086
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2087
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2111
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2114
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2115
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2117
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2118
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2119
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2120
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2121
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2122
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2123
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2124
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2125
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2126
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2128
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2130
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2133
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2140
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2148

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2174
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2175
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2188

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2697
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2731
        ' intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 2733
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 2734
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 2737  GetWordsList_OperationAbbrev2
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2738

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3169
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3170
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3173
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3174
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3188
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3191
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3192
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3193
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3197
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3198
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3204
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3205
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3207
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3210
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3252
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3253

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3520
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3584
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3598
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3676
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3685
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3691
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3692
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3839
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3846
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3848
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3875
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3878
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3879
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3881
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3886
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3934
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3936
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3939
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3942
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3944
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3947
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3948
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3954
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3960

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4019
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4020
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4203
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4207
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4208
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4209
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4215
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4218
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4233
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4234
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4234
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7261 '4241
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7263 '4243
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7264 '4244
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7265 '4245
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7266 '4246
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7267 '4247
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7268 '4248
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7269 '4249
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4250
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4252
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4253
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4254
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4255
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4257
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4259
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4260
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4305
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4306
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4313
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4314
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4319
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4331
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4332
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4373
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4374

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4375
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4376
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4385
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4394
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4395
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4418
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4420
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4446
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4460
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4461
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4468
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4478
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4497
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4502
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4565
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4573
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4588
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4589
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4594
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4595
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4596
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4598
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4599
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 4627  application constants, restricted, km
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 4628  application constants, restricted, km
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5026
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 6250
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 6306
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7218
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7220
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 7221

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_Activity() As Int32()

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrWords(14) As Int32

        ArrWords(1) = 369
        ArrWords(2) = 378
        ArrWords(3) = 384
        ArrWords(4) = 391

        'ArrWords(1) = 501
        ArrWords(5) = 510
        ArrWords(6) = 1861
        ArrWords(7) = 2844
        ArrWords(8) = 2845
        ArrWords(9) = 4427
        ArrWords(10) = 4623
        ArrWords(11) = 4624
        ArrWords(12) = 4638
        ArrWords(13) = 5703
        ArrWords(14) = 5730


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_PartNO() As Int32() 'get the list of words to update for Part Number 

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Full word (Plan number)
        Dim ArrWords(3) As Int32

        ArrWords(1) = 24
        ArrWords(2) = 2178
        ArrWords(3) = 3171
        'ArrWords(2) = 4608         'restricted - KM **


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit Count:({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordsList_PartNOAbbrev1() As Int32() 'get the list of words to update for Part# (Abbrev1)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Abbreviation (Ex: Plan#)
        Dim ArrWords(0) As Int32
        Dim intK = 0

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 205
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 211
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 247 - SHARE PLAN
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 254
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 322
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 397
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 398
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 400
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 406
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 417
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 434
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 436
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 442
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 448
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 460
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 467
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 470
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 471
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 477
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 479
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 857

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1811
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1813
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1815
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1818
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1820
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1822
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1886
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1892
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1894
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1896
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1932
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1935
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1936
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1959
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1960
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1962
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 1966


        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2017
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2040
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 2042
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2043
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2047
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2050
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2055
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2069
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2071
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2076
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2079
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2080
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2081
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2082
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2083
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2084
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2090
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2096
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2100
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2111
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2113
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2114
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2116
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2117
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2118
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2119
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2120
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2121
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2122
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2124
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2125
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2128
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2129
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2133
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2143
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2146
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2147
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2148
        'intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 2149

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2154
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2155
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2156
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2158
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2159

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2161
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2162
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2164

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2166
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2170
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2174
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2178

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2737
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2738
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2754
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 2755

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3061
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3164
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3171
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3172
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3173
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3181
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3183
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3184
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3185
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3186
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3187
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3190
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3194
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3196
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3199
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3200
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3201
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3202
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3203

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3208
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3235
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3511
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3512
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3516
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3517
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3518
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3519
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3523
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3524
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3525
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3527
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3530
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3533
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3534
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3537
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3543
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3551
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3554
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3555
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3556
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3567
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3568
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3675
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3678
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3679
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3847

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3848
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3851
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3852
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3854
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3876
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3938
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3950
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3951
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3953
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3959
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3960
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3961
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3964
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3965
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3966
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3968
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3969
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3970
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3974
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3980
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3982
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3983
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3984
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3985
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3988
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3989
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 3991

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4017
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4202
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4206

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4284
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4285

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4293
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4294
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4300
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4301
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4303
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4304
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4305
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4310
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4311
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4323
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4350
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4359
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4387
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4350
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4387
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4417
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4419
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4422

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4431
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4435
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4436
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4441
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4442
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4446
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4450
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4451
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4468
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4472
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4477
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4480

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4481
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4483
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4494
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4496
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4498
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4500
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4514

        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4521
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4522
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4523
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 4524


        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5439
        intK += 1
        ReDim Preserve ArrWords(intK)
        ArrWords(intK) = 5443
        ' intK += 1 : ReDim Preserve ArrWords(intK) : ArrWords(intK) = 5713


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit Count:({0})", UBound(ArrWords)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrWords
    End Function

    Private Function GetWordfromMemory(ByVal objEaseSys1 As DataTable, ByVal intID As Integer) As String

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter objEaseSys1:({0}) intID:({1})", objEaseSys1, intID),
                                                    LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = ""

        Try
            Dim strSearchFor = CStr(intID)
            Dim objRow = objEaseSys1.Rows.Find(strSearchFor)
            If (Not IsNothing(objRow)) Then
                returnValue = Trim(Extensions.Data.GetDataRowValue(objRow("desc"), ""))
            End If
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("GetWordfromMemory", ex)
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Sub UpdateWordInMemory(ByRef objEaseSys1 As DataTable, ByVal intID As Integer,
                                  ByVal strDesc As String, ByVal strIntlDesc As String, ByVal intL As Integer,
                                  ByVal intWL As Int16, ByVal intType As Int16, ByVal intDP As Int16)


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(
                    String.Format(
                        "Enter intID:({0}) strDesc:({1}) strIntlDesc:({2}) intL:({3}) intWL:({4}) intType:({5}) intDP:({6}) ",
                        intID, strDesc, strIntlDesc, intL, intWL, intType, intDP), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If


        'used in AppConfig and Config class only

        Dim objRow As DataRow, strSearchFor As String = ""
        Try
            strSearchFor = CStr(intID)
            objRow = objEaseSys1.Rows.Find(strSearchFor)
            If Not (objRow Is Nothing) Then
                objEaseSys1.AcceptChanges()
                objRow.BeginEdit()

                If Trim(strDesc) <> "" Then objRow("desc") = strDesc.Trim
                If Trim(strIntlDesc) <> "" Then objRow("intldesc") = strIntlDesc.Trim

                If intL > 0 Then objRow("l") = intL
                If intWL > 0 Then objRow("wl") = intWL
                If intType > 0 Then objRow("type") = intType
                If intDP > 0 Then objRow("dp") = intDP
                objRow("updateflag") = True         'update the database
                objRow.EndEdit()


                'Ec.GeneralFunc.WriteToEASELog(strSearchFor.Trim & "-" & strDesc.Trim)
            End If
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("GetWordfromMemory", ex)
        Finally
            objRow = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", objEaseSys1), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateSimiliarWords(ByRef objEaseSys1 As DataTable, ByVal intKeyX As Integer, ByVal strFind As String,
                                   ByVal strReplace As String)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(
            String.Format("Enter intKeyX:({0}) strFind:({1}) strReplace:({2})", intKeyX, strFind, strReplace), LOG_APPNAME,
            BASE_ERRORNUMBER + 0)
#End If


        'used in AppConfig and Config class only

        Dim strWord As String = "", intID As Int32 = 0
        Dim ArrWords(0) As Int32, strFinalWord As String = ""
        Dim chrPound As Char = "#"c

        Try
            If strFind.Trim = "" Then GoTo ExitThisSub 'find shouldn't be blank , defensive coding
            If strReplace.Trim = "" Then GoTo ExitThisSub 'replacing value should never be blank, defensive coding

            Select Case intKeyX
                Case 24 'part number
                    ArrWords = GetWordsList_PartNO()
                Case 4206 'partno abbrev1
                    ArrWords = GetWordsList_PartNOAbbrev1()
                Case 25 'Part Desc          'need in Config program, don't comment it
                    ArrWords = GetWordsList_PartDesc()
                Case 49 'Operation Number
                    ArrWords = GetWordsList_OperationNumber()
                Case 4218 'Operation
                    ArrWords = GetWordsList_OperationAbbrev1()
                Case 4370 'OP-- 4370/check for OP /not again for operation
                    ArrWords = GetWordsList_OperationAbbrev2()
                Case 51 'Work Center
                    ArrWords = GetWordsList_OperationWorkCenter()
                Case 3698 'subheader
                    ArrWords = GetWordsList_SubHeader()
                Case 1403
                    ArrWords = GetWordsList_Plant()
                Case 501 'R.J
                    ArrWords = GetWordsList_Activity()
            End Select
            If UBound(ArrWords) > 0 Then
                For intK = 1 To UBound(ArrWords)
                    intID = ArrWords(intK)
                    If intID <= 0 Then GoTo SkipUpdating 'skip unused record position

                    strWord = GetWordfromMemory(objEaseSys1, intID)

                    If strFind = "Op" Then
                        If InStr(strWord, chrPound & chrPound) > 0 Then
                            strWord = Regex.Replace(strWord, chrPound, String.Empty)
                        End If

                        Dim intPosPound = - 1

                        If InStr(strWord, strFind & " ") > 0 Then
                            intPosPound = InStr(strWord, strFind & " ") - 1
                        ElseIf InStr(strWord, strFind & "/") > 0 Then
                            intPosPound = InStr(strWord, strFind & "/") - 1
                        ElseIf (InStr(strWord, strFind) = strWord.Length - 1) Then
                            intPosPound = strWord.Length - strFind.Length
                        End If

                        If intPosPound > - 1 Then
                            strFinalWord = strWord.Substring(0, intPosPound) & strReplace &
                                           strWord.Substring(intPosPound + strFind.Length)
                        Else
                            strFinalWord = strWord
                        End If
                    Else
                        strFinalWord = Ec.GeneralFunc.ReplaceString(strWord, strFind, strReplace)
                    End If

                    UpdateWordInMemory(objEaseSys1, intID, strFinalWord, "", 0, 0, 0, - 1)
                    'If strWord.Trim = strFinalWord.Trim Then
                    '    'same nothing changes
                    '    Ec.GeneralFunc.WriteToEASELog("ID: " & intID & "   : Word: " & strWord & ", FIND:" & strFind & " -- " & strReplace)
                    'End If

                    SkipUpdating:
                Next
            End If
            ExitThisSub:
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("UpdateSimiliarWords", ex)
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", objEaseSys1), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Writes all database records in object EaseSys3 into easesys3.dat
    ''' Any existing records in easesys3.dat will be deleted.
    ''' </summary>
    ''' <param name="EaseSys3">Object holds all database details.</param>
    ''' <param name="blnResult">Result of the process (Success/Failure)</param>
    ''' <remarks></remarks>
    Public Sub WriteEaseSys3AllRecords(ByVal EaseSys3() As EASEClass7.stDataSource,
                                       Optional ByRef blnResult As Boolean = False,
                                       Optional ByVal strLOCALEASESysDir As String = "")


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(
            String.Format("Enter EaseSys3:({0}) strLOCALEASESysDir:({1})", EaseSys3, strLOCALEASESysDir), LOG_APPNAME,
            BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string

            '**EASESys1 consolidation- KM **
            If strLOCALEASESysDir.Trim <> "" Then 'IS USED ONLY IN WORD EDITOR (EASE INTERNAL USE ONLY)
                strCS = EaseCore.DAL.Connection.BuildConnectionString(0, strLOCALEASESysDir)
            End If

            Dim strConn1 As String = "",
                strSpare1 As String = "",
                strSpare2 As String = "",
                strSpare3 As String = "",
                strTemp2 As String = ""
            Dim blnUseSpare As Boolean = False

            Using connection = New EaseCore.DAL.Connection(strCS)
                'delete all records
                strSQL = "delete from easesys3"
                connection.ExecuteNonQuery(strSQL)

                For intK = 1 To UBound(EaseSys3)
                    strCS = EncryptConnectionString(EaseSys3(intK).ConnectionString)

                    If strCS.Trim.Length > 255 Then
                        strTemp2 = strCS.Trim
                        strConn1 = Microsoft.VisualBasic.Left(strTemp2, 255)
                        strTemp2 = Replace(strTemp2, strConn1, "")   '255 - remaining

                        If strTemp2.Trim.Length > 0 Then
                            blnUseSpare = True
                            strSpare1 = Microsoft.VisualBasic.Left(strTemp2, 40)
                            strTemp2 = Replace(strTemp2, strSpare1, "") '290-remaining  '45
                        End If

                        If strTemp2.Trim.Length > 0 Then
                            blnUseSpare = True
                            strSpare2 = Microsoft.VisualBasic.Left(strTemp2, 40)
                            strTemp2 = Replace(strTemp2, strSpare2, "") '321-remaining '40
                        End If

                        If strTemp2.Trim.Length > 0 Then
                            blnUseSpare = True
                            strSpare3 = Microsoft.VisualBasic.Left(strTemp2, 40)
                            strTemp2 = Replace(strTemp2, strSpare3, "") '361-remaining
                        End If
                    End If

                    If blnUseSpare Then
                        strCS = strConn1
                        EaseSys3(intK).Spare1 = strSpare1
                        EaseSys3(intK).Spare2 = strSpare2
                        EaseSys3(intK).Spare3 = strSpare3
                    End If

                    strSQL = "insert into easesys3(dbkey,connectstring,sharedeasedirectory,databasetype,refdocpath,cadpath," &
                             "videopath,descx,update7flag,spare1,spare2,spare3) values(" &
                             EaseSys3(intK).Key & ",'" & EaseCore.Extensions.Strings.ReplaceSingleQuote(strCS) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).SharedDirectory) & "'," &
                             EaseSys3(intK).DBType & ",'" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).DocPath) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).CadPath) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).VideoPath) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).DataSourceName) & "'," &
                             EaseSys3(intK).Update7Flag & "," &
                             "'" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).Spare1) & "', " &
                             "'" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).Spare2) & "', " &
                             "'" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3(intK).Spare3) & "') "
                    connection.ExecuteNonQuery(strSQL)
                Next
            End Using
            blnResult = True
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("WriteEaseSys3AllRecords", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function GetMaxDBKeyForEASESys3() As Integer

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = 0
        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            Dim strSQL = "select max(dbkey) from easesys3 "

            Using connection = New EaseCore.DAL.Connection(strCS)
                returnValue = connection.ExecuteScalar (Of Integer)(strSQL, 0)
            End Using
            returnValue += 1

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("GetMaxDBKeyForEASESys3", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Sub UpdateDataSourceUpdate7Flag(ByVal intDBKey As Int16, ByVal intKeyFlag As Int16)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter intDBKey:({0}) intKeyFlag:({1})", intDBKey, intKeyFlag),
                                                    LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Try

            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            Using connection = New EaseCore.DAL.Connection(strCS)
                Dim strSQL As String = "update easesys3 set update7flag=" & intKeyFlag & " where dbkey=" & intDBKey
                connection.ExecuteNonQuery(strSQL)
            End Using

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("UpdateDataSourceupdate7flag", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub WriteEASESys3(ByVal EaseSys3 As EASEClass7.stDataSource, ByRef intDBKey As Integer,
                             Optional ByRef blnResult As Boolean = False)


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter EaseSys3:({0}) ", EaseSys3), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If

        'Note: BYRef : intDBKey
        'update/add single record into easesys3

        Dim strSQL As String = ""
        Dim strCS As String = "", blnNewRecord As Boolean = False
        blnResult = False
        Dim strConn1 As String = "",
            strSpare1 As String = "",
            strSpare2 As String = "",
            strSpare3 As String = "",
            strTemp2 As String = ""
        Dim blnUseSpare As Boolean = False
        Try

            If intDBKey = - 1 Then 'new record
                intDBKey = GetMaxDBKeyForEASESys3()
                blnNewRecord = True
            End If
            strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string

            Using connection = New EaseCore.DAL.Connection(strCS)
                strCS = EncryptConnectionString(EaseSys3.ConnectionString)

                If strCS.Trim.Length > 255 Then
                    strTemp2 = strCS.Trim
                    strConn1 = Microsoft.VisualBasic.Left(strTemp2, 255)
                    strTemp2 = Replace(strTemp2, strConn1, "")   '255 - remaining

                    If strTemp2.Trim.Length > 0 Then
                        blnUseSpare = True
                        strSpare1 = Microsoft.VisualBasic.Left(strTemp2, 40)
                        strTemp2 = Replace(strTemp2, strSpare1, "") '290-remaining  '45
                    End If

                    If strTemp2.Trim.Length > 0 Then
                        blnUseSpare = True
                        strSpare2 = Microsoft.VisualBasic.Left(strTemp2, 40)
                        strTemp2 = Replace(strTemp2, strSpare2, "") '321-remaining '40
                    End If

                    If strTemp2.Trim.Length > 0 Then
                        blnUseSpare = True
                        strSpare3 = Microsoft.VisualBasic.Left(strTemp2, 40)
                        strTemp2 = Replace(strTemp2, strSpare3, "") '361-remaining
                    End If
                End If

                If blnUseSpare Then
                    strCS = strConn1
                    EaseSys3.Spare1 = strSpare1
                    EaseSys3.Spare2 = strSpare2
                    EaseSys3.Spare3 = strSpare3
                End If

                If blnNewRecord Then
                    'JVG - EVC-784 - Trap & Retry Oracle database connection drop
                    strSQL = "insert into easesys3(dbkey,connectstring,sharedeasedirectory,databasetype,refdocpath,cadpath," &
                             "videopath,descx,update7flag,spare1,spare2,spare3) values(" &
                             intDBKey & ",'" & EaseCore.Extensions.Strings.ReplaceSingleQuote(strCS) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.SharedDirectory) & "'," & EaseSys3.DBType &
                             ",'" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.DocPath) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.CadPath) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.VideoPath) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.DataSourceName) & "'," &
                             EaseSys3.Update7Flag & ",'" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.Spare1) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.Spare2) & "','" &
                             EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.Spare3) & "') "
                Else
                    'JVG - EVC-784 - Trap & Retry Oracle database connection drop
                    strSQL = "update easesys3 set connectstring='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(strCS) &
                             "'," &
                             " sharedeasedirectory='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.SharedDirectory) &
                             "'," &
                             " databasetype=" & EaseSys3.DBType & "," &
                             " refdocpath='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.DocPath) & "'," &
                             " cadpath='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.CadPath) & "'," &
                             " videopath='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.VideoPath) & "'," &
                             " descx='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.DataSourceName) & "'," &
                             " spare1='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.Spare1) & "'," &
                             " spare2='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.Spare2) & "'," &
                             " spare3='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(EaseSys3.Spare3) & "'" &
                             " where dbkey=" & intDBKey

                End If

                connection.ExecuteNonQuery(strSQL)
            End Using
            blnResult = True
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("UpdateEASESys3", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit intDBKey:({0}) blnResult:({1})", intDBKey, blnResult), LOG_APPNAME,
                           BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Private Function GetTotalNumberOfWordsInEaseSys1(Optional ByVal strNEWWordFileName As String = "") As Int32

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter strNEWWordFileName:({0})", strNEWWordFileName),
                                                    LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'strCS is used to get the number of words from the systemUpdateNEW MDB file
        Dim returnValue = 0

        Try

            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            If Trim(strNEWWordFileName) <> "" Then
                strCS = EaseCore.DAL.Connection.BuildConnectionString(0, strNEWWordFileName)            'strNewWordFileName
            End If

            Dim strSQL = "select count(1) from easesys1"
            Using connection = New EaseCore.DAL.Connection(gStrConnectionString)
                returnValue = connection.ExecuteScalar (Of Integer)(strSQL, 0)
            End Using

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("GetTotalNumberOfWordsInEaseSys1", ex)
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Friend Function GetEasesys1Words(Optional ByVal intID As Integer = 0,
                                     Optional ByVal blnRetrieveAllRecords As Boolean = False,
                                     Optional ByVal strDBFileName As String = "",
                                     Optional ByVal blnUseEASEDatabase As Boolean = False,
                                     Optional ByVal intLanguage As Integer = 1) As DataTable


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(
                    String.Format(
                        "Enter intID:({0}) blnRetrieveAllRecords:({1}) strDBFileName:({2}) blnUseEASEDatabase:({3}) intLanguage:({4})",
                        intID, blnRetrieveAllRecords, strDBFileName, blnUseEASEDatabase, intLanguage), LOG_APPNAME,
                    BASE_ERRORNUMBER + 0)
#End If

        'This sub gets the words from easesys1 table and holds
        'in dataTable. The data table is used to sort and search for the words
        'which is essential for EASE applications.

        'intLanguage - default language is English

        '**EASESys1 consolidation- KM **
        'blnUseEASEDatabase is used to read the word from ease database (called in UpdateLocalAccessWordsFile, LOADEASE)

        Dim strCS As String = ""
        Dim strTemp1 As String = ""
        Dim easeSys1DataTable As New DataTable("easesys1")
        Dim PrimaryKeyColumns(0) As DataColumn

        Try
            If Trim(strDBFileName) = "" Then 'default selction 
                strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            Else
                strCS = EaseCore.DAL.Connection.BuildConnectionString(0, strDBFileName)
            End If

            Dim strSQL = String.Format("select * from {0} ", If(ApplicationIsAmlViewEase, "easesys1_aml_viewease", "easesys1"))
            If intID > 0 Then
                strSQL &= " where id=" & intID
            End If
            strSQL &= " order by id"

            Using connection = New EaseCore.DAL.Connection(strCS)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        easeSys1DataTable.Columns.Add("no", Type.GetType("System.Int16"))
                        easeSys1DataTable.Columns.Add("desc", Type.GetType("System.String"))              'english
                        easeSys1DataTable.Columns.Add("intldesc", Type.GetType("System.String")) _
                        'international language


                        easeSys1DataTable.Columns.Add("wl", Type.GetType("System.Int16"))
                        easeSys1DataTable.Columns.Add("l", Type.GetType("System.Int16"))
                        easeSys1DataTable.Columns.Add("type", Type.GetType("System.Int16"))
                        easeSys1DataTable.Columns.Add("dp", Type.GetType("System.Int16"))
                        easeSys1DataTable.Columns.Add("updateflag", Type.GetType("System.Boolean"))

                        Dim strLanguageFieldName = "foreigndescx"  'default foreign language
                        If intLanguage > 1 Then _
'field name: Foreightdescx, foreigndescx2, foreigndescx3. foreigndescx4,foreigndescx5....
                            strLanguageFieldName &= intLanguage.ToString _
                            'foreigndescx2, foreigndescx3. foreigndescx4,foreigndescx5....
                        End If

                        For Each reader As DataRow In table.Rows

                            Dim str_Tmp = Trim(Extensions.Data.GetDataRowValue(reader("descx"), ""))
                            If blnRetrieveAllRecords Then GoTo SkipBlankRecordsCheck
                            If Trim(str_Tmp) = "" Then GoTo SkipThisRecord

                            SkipBlankRecordsCheck:
                            Dim obj_Row = easeSys1DataTable.NewRow
                            obj_Row("no") = reader("id")
                            obj_Row("desc") = Trim(str_Tmp)

                            strTemp1 = Trim(Extensions.Data.GetDataRowValue(reader(strLanguageFieldName), ""))
                            obj_Row("intldesc") = Trim(strTemp1)

                            'SkipForeignLanguage:
                            Dim intTemp = Extensions.Data.GetDataRowValue(reader("wl"), 0)
                            obj_Row("wl") = intTemp

                            intTemp = Extensions.Data.GetDataRowValue(reader("len"), 0)
                            obj_Row("l") = intTemp

                            intTemp = Extensions.Data.GetDataRowValue(reader("type"), 0)
                            obj_Row("type") = intTemp

                            intTemp = Extensions.Data.GetDataRowValue(reader("dp"), 0)
                            obj_Row("dp") = intTemp
                            obj_Row("updateflag") = False  'used in updating the words file (Sub: UpdateEaseSys1Words)
                            easeSys1DataTable.Rows.Add(obj_Row)
                            SkipThisRecord:
                            strTemp1 = ""
                        Next
                    End If
                End Using
            End Using
            PrimaryKeyColumns(0) = easeSys1DataTable.Columns("no")
            easeSys1DataTable.PrimaryKey = PrimaryKeyColumns
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("GetEasesys1Words", ex)
        Finally
            PrimaryKeyColumns = Nothing
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return easeSys1DataTable
    End Function

    Public Sub ReadEaseSys1(Optional ByVal intID As Integer = 0, Optional ByVal blnRetrieveAllRecords As Boolean = False,
                            Optional ByVal intLanguageID As Integer = 0, optional ByVal amlViewEase as Boolean = false)


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(
                    String.Format("Enter intID:({0}) blnRetrieveAllRecords:({1}) intLanguageID:({2})", intID,
                                  blnRetrieveAllRecords, intLanguageID), LOG_APPNAME,
                    BASE_ERRORNUMBER + 0)
#End If


        'This sub gets the words from easesys1 table and holds
        'in dataTable. The data table is used to sort and search for the words
        'which is essential for EASE applications.
        'This sub is used in Process Planning, ClientEditor and all the EASE apps.
        'The Word Editor is not using this Sub (it uses GetEASESys1Records)
        Try
            ApplicationIsAmlViewEase = amlViewEase
            gObjEasesys1 = GetEasesys1Words(intID, blnRetrieveAllRecords, , , intLanguageID)
            gBlnEASESys1Loaded = True
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("ReadEaseSys1", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function ReadEaseSys3(Optional ByVal blnAllDatabase As Boolean = False, Optional ByVal intDBKey As Integer = 0,
                                 Optional ByVal strLOCALEASESysDir As String = "") As EASEClass7.stDataSource()


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(String.Format("Enter blnAllDatabase:({0}) intDBKey:({1}) strLOCALEASESysDir:({2})",
                                                 blnAllDatabase, intDBKey, strLOCALEASESysDir), LOG_APPNAME,
                                   BASE_ERRORNUMBER + 0)
#End If


        'strLOCALEASESysDir (EASE INTERNAL USE ONLY) param is used only in EASE word editor and should NEVER BE USED ANYWHERE ELSE - KM

        Dim objDS() As EASEClass7.stDataSource
        Dim strCond As String = " where "
        ReDim objDS(0)

        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string

            If strLOCALEASESysDir.Trim <> "" Then _
'IS USED ONLY IN WORD EDITOR (EASE INTERNAL USE ONLY)  **EASESys1 consolidation- KM **
                strCS = EaseCore.DAL.Connection.BuildConnectionString(0, strLOCALEASESysDir)
            End If
            Dim strSQL = "select * from easesys3 "

            If Not blnAllDatabase Then
                strSQL &= strCond & " update7flag=1"
                strCond = " and "
            End If
            If intDBKey <> 0 Then
                strSQL &= strCond & " dbkey=" & intDBKey
                strCond = " and "
            End If
            strSQL &= " order by dbkey"

            Using connection = New EaseCore.DAL.Connection(strCS)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objDS(table.Rows.Count)
                        Dim counter = 0
                        For Each reader As DataRow In table.Rows

                            counter += 1

                            objDS(counter).CadPath = Trim(Extensions.Data.GetDataRowValue(reader("cadpath"), ""))
                            objDS(counter).ConnectionString = Trim(Extensions.Data.GetDataRowValue(reader("connectstring"), ""))
                            objDS(counter).DBType = Extensions.Data.GetDataRowValue (Of EaseCore.Dal.Connection.DatabaseTypes)(reader("databasetype"), Connection.DatabaseTypes.Access)
                            objDS(counter).Key = Extensions.Data.GetDataRowValue (Of Int16)(reader("dbkey"), 0)
                            objDS(counter).DataSourceName = Trim(Extensions.Data.GetDataRowValue(reader("descx"), ""))
                            objDS(counter).Update7Flag = Extensions.Data.GetDataRowValue (Of Int16)(reader("update7flag"), 0)
                            objDS(counter).DocPath = Trim(Extensions.Data.GetDataRowValue(reader("refdocpath"), ""))
                            objDS(counter).SharedDirectory = Trim(Extensions.Data.GetDataRowValue(reader("sharedeasedirectory"),
                                                                                                  ""))
                            If (objDS(counter).SharedDirectory <> "") Then _
                                objDS(counter).SharedDirectory =
                                    EaseCore.Extensions.Strings.AddSlash(objDS(counter).SharedDirectory)
                            objDS(counter).VideoPath = Trim(Extensions.Data.GetDataRowValue(reader("videopath"), ""))
                            objDS(counter).Spare1 = Extensions.Data.GetDataRowValue(reader("spare1"), "")
                            objDS(counter).Spare2 = Extensions.Data.GetDataRowValue(reader("spare2"), "")
                            objDS(counter).Spare3 = Extensions.Data.GetDataRowValue(reader("spare3"), "")
                            objDS(counter).ConnectionString = objDS(counter).ConnectionString.Trim & objDS(counter).Spare1.Trim &
                                                              objDS(counter).Spare2.Trim & objDS(counter).Spare3.Trim
                            objDS(counter).ConnectionString = DecodeConnectionString(objDS(counter).ConnectionString)
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("ReadEASESys3Records", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit Count:({0})", UBound(objDS)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objDS
    End Function

    Public Sub UpdateDefaultDatabaseInEaseSys(ByVal intDBKey As Int16, Optional ByVal strLOCALEASESysDir As String = "")

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(
            String.Format("Enter intDBKey:({0}) strLOCALEASESysDir:({1})", intDBKey, strLOCALEASESysDir), LOG_APPNAME,
            BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "update easesys set dbkey=" & intDBKey & " where id=1"
        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string

            '**EASESys1 consolidation- KM **
            If strLOCALEASESysDir.Trim <> "" Then 'IS USED ONLY IN WORD EDITOR (EASE INTERNAL USE ONLY)
                strCS = EaseCore.DAL.Connection.BuildConnectionString(0, strLOCALEASESysDir)
            End If
            Using connection = New EaseCore.DAL.Connection(strCS)
                connection.ExecuteNonQuery(strSQL)
            End Using
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("UpdateDefaultDatabaseInEaseSys", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Update the datasource description for the passed parameter strDBKey
    ''' </summary>
    ''' <param name="intDBKey">Database key</param>
    ''' <param name="strDataSourceDesc">Data source description.</param>
    ''' <param name="strSharedDirectory">Shared EASE directory.</param>
    ''' <returns>True/False, result of this process</returns>
    ''' <remarks></remarks>
    Public Function UpdateDataSourceDescription(ByVal intDBKey As Int16,
                                                ByVal strDataSourceDesc As String,
                                                Optional ByVal strSharedDirectory As String = "") As Boolean

#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(String.Format("Enter intDBKey:({0}) strDataSourceDesc:({1}) strSharedDirectory:({2})",
                                                 intDBKey, strDataSourceDesc, strSharedDirectory), LOG_APPNAME,
                                   BASE_ERRORNUMBER + 0)
#End If


        Dim strSQL As String = "update easesys3 set descx='" & EaseCore.Extensions.Strings.ReplaceSingleQuote(strDataSourceDesc) &
                               "' where dbkey=" & intDBKey
        Dim blnRtnValue As Boolean = False
        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            Using connection = New EaseCore.DAL.Connection(strCS)
                connection.ExecuteNonQuery(strSQL)
            End Using
            blnRtnValue = True
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("UpdateDataSourceDescription", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Function CheckEASESysDatabaseConnection() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = False
        Dim strCS = ReadEASESysDatabaseConnectionString()
        Using connection = New EaseCore.DAL.Connection(strCS)
            Dim strSQL = "select count(1) from easesys1"
            returnValue = connection.ExecuteScalar (Of Integer)(strSQL, 0) > 0
        End Using

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetEASESysDatabaseName() As String

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strDBName As String = EaseCore.Extensions.Strings.AddSlash(Ec.AppConfig.EASEConfigDirectory) & "easesys.mdb"


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", strDBName), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strDBName
    End Function

    Public Function GetEaseSys1MaxID() As Int32

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(id) from easesys1"
        Dim returnValue = 0
        Using connection = New EaseCore.DAL.Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar (Of Integer)(strSQL, 0)
        End Using
        returnValue += 1

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function ReadEASESysDatabaseConnectionString() As String

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strDB As String = GetEASESysDatabaseName()
        Dim strCS As String = EaseCore.DAL.Connection.BuildConnectionString(0, strDB)


#If TRACE Then
        Log.EASESYS_IO_LOW(String.Format("Exit ({0})", strCS), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strCS
    End Function

    Public Function GetDefaultDatabaseConnectionString() As String

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'this sub is used only in Web apps, if the ClientEditor has been changed with new data source, the other web apps
        'need to refresh their database settings instead of using the existing session vars
        Dim strCS As String = "", strSQL As String = ""
        Dim strReturnValue As String = ""
        Dim intDBKey = 0, blnDBFound As Boolean = False
        Dim objDB As New stDataSource
        Try
            strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            strSQL = "select dbkey from easesys where id=1"
            Using connection = New EaseCore.DAL.Connection(strCS)
                intDBKey = connection.ExecuteScalar (Of Integer)(strSQL, 0)
            End Using

            objDB = GetADatabase(intDBKey, , blnDBFound)
            If blnDBFound Then
                strReturnValue = objDB.ConnectionString.Trim
            End If

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("ReadEASESysRecords", ex)
        Finally
            objDB = Nothing
        End Try

#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strReturnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strReturnValue
    End Function

    Public Function ReadEASESysRecords(Optional ByVal strLOCALEASESysDir As String = "") As stEASESys


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter strLOCALEASESysDir:({0})", strLOCALEASESysDir),
                                                    LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objEASESys As New stEASESys
        Try
            objEASESys.EASEConfig = ""
            objEASESys.Version = ""
            objEASESys.DBKey = 0
            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string

            '**EASESys1 consolidation- KM **
            If strLOCALEASESysDir.Trim <> "" Then 'IS USED ONLY IN WORD EDITOR (EASE INTERNAL USE ONLY)
                strCS = EaseCore.DAL.Connection.BuildConnectionString(0, strLOCALEASESysDir)
            End If

            Dim strSQL = "select * from easesys where id=1"
            Using connection = New EaseCore.DAL.Connection(strCS)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim reader As DataRow = table.Rows(0)

                        objEASESys.EASEConfig = Trim(Extensions.Data.GetDataRowValue(reader("EASEConfig"), ""))
                        objEASESys.Version = Trim(Extensions.Data.GetDataRowValue(reader("version"), ""))
                        objEASESys.DBKey = Extensions.Data.GetDataRowValue (Of Int16)(reader("dbkey"), 0)
                    End If
                End Using
            End Using
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("ReadEASESysRecords", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objEASESys
    End Function

    Public Sub UpdatePreviousVersionRecords(ByVal objTable As DataTable)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter objTable:({0})", objTable), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False

        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()
            Using connection = New EaseCore.DAL.Connection(strCS)
                connection.BeginTransaction()
                Try
                    For Each obj_Row As DataRow In objTable.Rows
                        Dim intID = Extensions.Data.GetDataRowValue(obj_Row("no"), 0)
                        Dim strTemp = Extensions.Data.GetDataRowValue(obj_Row("desc"), "")
                        Dim strSQL = "delete from easesys1 where id=" & intID
                        connection.ExecuteNonQuery(strSQL)

                        strSQL = ""
                        If intID > 818 Then
                            strSQL = ""

                            Ec.GeneralFunc.StripAllSpecialCharacters(strTemp, True)

                        End If

                        Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(EASEClass7.QueryBuilder.QueryType.Insert,
                                                                                  "easesys1")
                        queryBuilder.AddField("id", intID, True)
                        queryBuilder.AddField("descx", strTemp, False, False)
                        queryBuilder.AddField("wl", obj_Row("wl"), True)
                        queryBuilder.AddField("Len", obj_Row("l"), True)
                        queryBuilder.AddField("type", obj_Row("type"), True)
                        queryBuilder.AddField("dp", obj_Row("dp"), True)
                        strSQL = queryBuilder.GenerateQuery
                        connection.ExecuteNonQuery(strSQL)
                    Next
                    connection.CommitTransaction()
                    blnResult = True
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Throw ex
                End Try
            End Using
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("UpdatePreviousVersionRecords", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub AddBlankRecordsinEASESys1(ByVal intFromID As Int32, ByVal intToID As Int32)


#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter intFromID:({0}) intToID:({1})", intFromID, intToID),
                                                    LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If


        Dim strSQL As String = "", blnResult As Boolean = False
        Dim dtDate As System.DateTime = EaseDate.Now

        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()
            Using connection = New EaseCore.DAL.Connection(strCS)
                connection.BeginTransaction()
                Try
                    For intK = intFromID To intToID
                        Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(EASEClass7.QueryBuilder.QueryType.Insert,
                                                                                  "easesys1")
                        queryBuilder.AddField("id", intK, True)
                        queryBuilder.AddField("descx", "", False, False)
                        queryBuilder.AddField("ForeignDescX", "", False, False)
                        queryBuilder.AddField("wl", 0, True)
                        queryBuilder.AddField("Len", 0, True)
                        queryBuilder.AddField("type", 0, True)
                        queryBuilder.AddField("dp", 0, True)
                        'FieldExist() takes connstr corr to actualDB(orcl/acc/sql) where as we need conn str corr to easesys.mdb. so pass it explicitly. Change the DataSource as needed
                        If Ec.IO.FieldExist("easesys1", "LastUpdated") Then _
' this field may not be present for existing cust. Added this flag to cover existing cust
                            queryBuilder.AddField("LastUpdated", dtDate, False) _
                            'Fill date - needed while export/import of easesys.mdb to/from excel
                        End If

                        strSQL = queryBuilder.GenerateQuery
                        connection.ExecuteNonQuery(strSQL)
                    Next

                    connection.CommitTransaction()
                    blnResult = True
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Throw ex
                End Try
            End Using
        Catch ex As Exception
            Call GenerateException("AddBlankRecords", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateEASESys1Record(ByVal intID As Int32, ByVal strDesc As String, ByVal strForeignDescX As String,
                                    ByVal intWL As Int16, ByVal intLen As Integer, ByVal intType As Int16, ByVal intDP As Int16,
                                    Optional ByVal strLastUpdatedDateTime As String = "")


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(
                    String.Format(
                        "Enter intID:({0}) strDesc:({1}) strForeignDescX:({2}) intWL:({3}) intLen:({4}) intType:({5}) intDP:({6}) strLastUpdatedDateTime:({7})",
                        intID, strDesc, strForeignDescX, intWL, intLen, intType, intDP, strLastUpdatedDateTime), LOG_APPNAME,
                    BASE_ERRORNUMBER + 0)
#End If

        'write word in the database
        WriteWordRecord(QueryBuilder.QueryType.Update, intID, strDesc, strForeignDescX, intWL, intLen, intType, intDP, True,
                        strLastUpdatedDateTime)

        'update the word in the memory (so using getwrd will return the updated word)
        UpdateWordInMemory(gObjEasesys1, intID, strDesc, strForeignDescX, 0, intWL, intType, intDP)


#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub UpdateSimilarWordsPlant(ByVal intKeyX As Int32, ByVal strFind As String, ByVal strReplaceWith As String)

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter intKeyX:({0})  strFind:({1}) strReplaceWith:({2})",
                                                                  intKeyX, strFind, strReplaceWith), LOG_APPNAME,
                                                    BASE_ERRORNUMBER + 0)
#End If

        UpdateSimiliarWords(gObjEasesys1, intKeyX, strFind, strReplaceWith)

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub


    Public Function GetCurrentDateTimeForWordUpdate() As String

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim dtDate As DateTime = EaseDate.Now, strTemp As String = ""
        strTemp = dtDate.Year & dtDate.Month & dtDate.Day & "-" & dtDate.Hour & dtDate.Minute & dtDate.Second
        'strTemp = Format(dtDate, "yyyyMMddHHmmss")         **EASESys1 consolidation- KM **


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", strTemp), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strTemp
    End Function

    Public Function GetEASEWorksActivationCode2() As String

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strRtnValue As String = ""
        Try
            Dim strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            Dim strSQL = "select descx FROM easesys1 where (id between 4895 and 4899) and dp = 99 order by id"

            Using connection = New EaseCore.DAL.Connection(strCS)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each reader As DataRow In table.Rows
                            strRtnValue &= Trim(Extensions.Data.GetDataRowValue(reader("descx"), ""))
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            Call GenerateException("GetEASEWorksActivationCode2", ex)
            gBlnEASESys1Loaded = False
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function

    Public Function WriteEASEWorksActivationCode2(ByVal strActivationCode As String, Optional ByRef strErrorMsg As String = "") _
        As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter strActivationCode:({0}) ", strActivationCode),
                                                    LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strCS As String = ""
        Dim ArrTemp() As String
        Dim intKeyX = 0, strSQL As String = ""
        Dim strTemp As String = "", blnRtnValue As Boolean = False

        Try

            ArrTemp = Ec.GeneralFunc.SplitTextToArray(strActivationCode, 60)
            If UBound(ArrTemp) = 0 Then
                strErrorMsg = "Activiation Code missing."
                GoTo ExitThisFunction
            End If
            If UBound(ArrTemp) > 6 Then
                strErrorMsg = "Invalid Activiation Code."
                GoTo ExitThisFunction
            End If

            strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string
            Using connection = New EaseCore.DAL.Connection(strCS)
                intKeyX = 4894
                For intK = 1 To UBound(ArrTemp)
                    strTemp = ArrTemp(intK)
                    intKeyX += 1

                    Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(EASEClass7.QueryBuilder.QueryType.Update,
                                                                              "easesys1")
                    queryBuilder.AddField("descx", strTemp, False, False)
                    queryBuilder.AddField("ForeignDescX", "", False, False)
                    queryBuilder.AddField("wl", 0, True)
                    queryBuilder.AddField("Len", 0, True)
                    queryBuilder.AddField("type", 0, True)
                    queryBuilder.AddField("dp", 99, True)
                    queryBuilder.AddConditionField("id", intKeyX.ToString(), True)
                    strSQL = queryBuilder.GenerateQuery
                    connection.ExecuteNonQuery(strSQL)
                Next
                blnRtnValue = True
            End Using

            ExitThisFunction:
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("WriteEASEWorksActivationCode2", ex)
        End Try


#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit ({0}) strErrorMsg:({1})", blnRtnValue, strErrorMsg), LOG_APPNAME,
                           BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Sub EditWordRecord(ByVal intQueryType As QueryBuilder.QueryType, ByVal intID As Int32,
                              ByVal strForeignDescX As String, Optional ByVal intselectedindex As Int32 = 0)


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(
                    String.Format("Enter intQueryType:({0})  intID:({1})  strForeignDescX:({2})  intselectedindex:({3}) ",
                                  intQueryType, intID, strForeignDescX, intselectedindex), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        '1-Insert Query, 2-Update Query, 3-Delete Query
        Dim strCS As String = ""
        Dim intPos As Int16 = 0
        Dim dtDate As System.DateTime = EaseDate.Now
        Try
            strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string

            Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(intQueryType, "easesys1")

            If intselectedindex = 1 Then
                queryBuilder.AddField("ForeignDescX1", strForeignDescX, False, False)

            ElseIf intselectedindex = 2 Then
                queryBuilder.AddField("ForeignDescX2", strForeignDescX, False, False)

            ElseIf intselectedindex = 3 Then
                queryBuilder.AddField("ForeignDescX3", strForeignDescX, False, False)

            ElseIf intselectedindex = 4 Then
                queryBuilder.AddField("ForeignDescX4", strForeignDescX, False, False)
            ElseIf intselectedindex = 5 Then
                queryBuilder.AddField("ForeignDescX5", strForeignDescX, False, False)

            ElseIf intselectedindex = 6 Then
                queryBuilder.AddField("ForeignDescX6", strForeignDescX, False, False)

            ElseIf intselectedindex = 7 Then
                queryBuilder.AddField("ForeignDescX7", strForeignDescX, False, False)

            ElseIf intselectedindex = 8 Then
                queryBuilder.AddField("ForeignDescX8", strForeignDescX, False, False)

            ElseIf intselectedindex = 9 Then
                queryBuilder.AddField("ForeignDescX9", strForeignDescX, False, False)

            ElseIf intselectedindex = 10 Then
                queryBuilder.AddField("ForeignDescX10", strForeignDescX, False, False)

            ElseIf intselectedindex = 11 Then
                queryBuilder.AddField("ForeignDescX11", strForeignDescX, False, False)

            ElseIf intselectedindex = 12 Then
                queryBuilder.AddField("ForeignDescX12", strForeignDescX, False, False)

            ElseIf intselectedindex = 13 Then
                queryBuilder.AddField("ForeignDescX13", strForeignDescX, False, False)

            ElseIf intselectedindex = 14 Then
                queryBuilder.AddField("ForeignDescX14", strForeignDescX, False, False)

            ElseIf intselectedindex = 15 Then
                queryBuilder.AddField("ForeignDescX15", strForeignDescX, False, False)

            Else
                queryBuilder.AddField("ForeignDescX", strForeignDescX, False, False)

            End If
            'R.J  12/11/2013


            If Ec.IO.FieldExist("easesys1", "LastUpdated", strCS) Then
                queryBuilder.AddField("LastUpdated", dtDate, False)
            End If

            If intQueryType = QueryBuilder.QueryType.Insert Then
                queryBuilder.AddField("id", intID, True)
            ElseIf intQueryType = QueryBuilder.QueryType.Update Then
                queryBuilder.AddConditionField("id", intID.ToString(), True)
            End If

            intPos = 16
            intQueryType = QueryBuilder.QueryType.Update

            queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(intQueryType, "easesys1")
            If intselectedindex = 1 Then
                queryBuilder.AddField("ForeignDescX", strForeignDescX, False, False)
            ElseIf intselectedindex = 2 Then
                queryBuilder.AddField("ForeignDescX2", strForeignDescX, False, False)

            ElseIf intselectedindex = 3 Then
                queryBuilder.AddField("ForeignDescX3", strForeignDescX, False, False)

            ElseIf intselectedindex = 4 Then
                queryBuilder.AddField("ForeignDescX4", strForeignDescX, False, False)
            ElseIf intselectedindex = 5 Then
                queryBuilder.AddField("ForeignDescX5", strForeignDescX, False, False)

            ElseIf intselectedindex = 6 Then
                queryBuilder.AddField("ForeignDescX6", strForeignDescX, False, False)

            ElseIf intselectedindex = 7 Then
                queryBuilder.AddField("ForeignDescX7", strForeignDescX, False, False)

            ElseIf intselectedindex = 8 Then
                queryBuilder.AddField("ForeignDescX8", strForeignDescX, False, False)

            ElseIf intselectedindex = 9 Then
                queryBuilder.AddField("ForeignDescX9", strForeignDescX, False, False)

            ElseIf intselectedindex = 10 Then
                queryBuilder.AddField("ForeignDescX10", strForeignDescX, False, False)

            ElseIf intselectedindex = 11 Then
                queryBuilder.AddField("ForeignDescX11", strForeignDescX, False, False)

            ElseIf intselectedindex = 12 Then
                queryBuilder.AddField("ForeignDescX12", strForeignDescX, False, False)

            ElseIf intselectedindex = 13 Then
                queryBuilder.AddField("ForeignDescX13", strForeignDescX, False, False)

            ElseIf intselectedindex = 14 Then
                queryBuilder.AddField("ForeignDescX14", strForeignDescX, False, False)

            ElseIf intselectedindex = 15 Then
                queryBuilder.AddField("ForeignDescX15", strForeignDescX, False, False)

            Else
                queryBuilder.AddField("ForeignDescX", strForeignDescX, False, False)

            End If
            'R.J  12/11/2013
            queryBuilder.AddConditionField("id", intID.ToString(), True)

            strSQL = queryBuilder.GenerateQuery
            Using connection = New EaseCore.DAL.Connection(strCS)
                connection.ExecuteNonQuery(strSQL)
            End Using

            ExitThisSub:
        Catch ex As Exception
            Call GenerateException("WriteWordRecord", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub WriteWordRecord(ByVal intQueryType As QueryBuilder.QueryType, ByVal intID As Int32, ByVal strDesc As String,
                               ByVal strForeignDescX As String,
                               ByVal intWL As Integer, ByVal intLen As Integer, ByVal intType As Int16, ByVal intDP As Int16,
                               Optional ByVal blnUpdateDateTimeStamp As Boolean = False,
                               Optional ByVal strLastUpdatedDateTime As String = "",
                               Optional ByVal blnProtectAppConstants As Boolean = True,
                               Optional ByVal intLanguage As Int32 = 1, 
                                Optional Byval amlViewEase As Boolean = False )


#If TRACE Then
        Dim startTicks As Long =
                Log.EASESYS_IO_MED(
                    String.Format(
                        "Enter intQueryType:({0}) intID:({1}) strDesc:({2}) strForeignDescX:({3}) intWL:({4}) intLen:({5}) intType:({6}) intDP:({7}) blnUpdateDateTimeStamp:({8}) strLastUpdatedDateTime:({9}) blnProtectAppConstants:({10}) intLanguage:({11}) ",
                        intQueryType, intID, strDesc, strForeignDescX, intWL, intLen, intType, intDP, blnUpdateDateTimeStamp,
                        strLastUpdatedDateTime, blnProtectAppConstants, intLanguage), LOG_APPNAME,
                    BASE_ERRORNUMBER + 0)
#End If


        'blnUpdateDateTimeStamp : is used in EASE word editor
        'Don't allow the anyone to update the Application Constants (so, we don't use the 4600-4900 range)

        Dim strSQL As String = ""
        '1-Insert Query, 2-Update Query, 3-Delete Query
        Dim strCS As String = ""
        Dim intPos As Int16 = 0
        Dim dtDate As System.DateTime = EaseDate.Now
        Try
            strCS = ReadEASESysDatabaseConnectionString()       'get the EASE config database connection string

            Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(intQueryType, If (amlViewEase, "easesys1_aml_viewease", "easesys1"))
            queryBuilder.AddField("descx", strDesc, False, False)
            If intLanguage = 2 Then
                queryBuilder.AddField("ForeignDescX2", strForeignDescX, False, False)
            ElseIf intLanguage = 3 Then
                queryBuilder.AddField("ForeignDescX3", strForeignDescX, False, False)
            ElseIf intLanguage = 4 Then
                queryBuilder.AddField("ForeignDescX4", strForeignDescX, False, False)
            ElseIf intLanguage = 5 Then
                queryBuilder.AddField("ForeignDescX5", strForeignDescX, False, False)
            ElseIf intLanguage = 6 Then
                queryBuilder.AddField("ForeignDescX6", strForeignDescX, False, False)
            ElseIf intLanguage = 7 Then
                queryBuilder.AddField("ForeignDescX7", strForeignDescX, False, False)
            ElseIf intLanguage = 8 Then
                queryBuilder.AddField("ForeignDescX8", strForeignDescX, False, False)
            ElseIf intLanguage = 9 Then
                queryBuilder.AddField("ForeignDescX9", strForeignDescX, False, False)
            ElseIf intLanguage = 10 Then
                queryBuilder.AddField("ForeignDescX10", strForeignDescX, False, False)
            ElseIf intLanguage = 11 Then
                queryBuilder.AddField("ForeignDescX11", strForeignDescX, False, False)
            ElseIf intLanguage = 12 Then
                queryBuilder.AddField("ForeignDescX12", strForeignDescX, False, False)
            ElseIf intLanguage = 13 Then
                queryBuilder.AddField("ForeignDescX13", strForeignDescX, False, False)
            ElseIf intLanguage = 14 Then
                queryBuilder.AddField("ForeignDescX14", strForeignDescX, False, False)
            ElseIf intLanguage = 15 Then
                queryBuilder.AddField("ForeignDescX15", strForeignDescX, False, False)
            ElseIf intLanguage = 16 Then
                queryBuilder.AddField("ForeignDescX16", strForeignDescX, False, False)
            Else 'default- first foreign language

                queryBuilder.AddField("ForeignDescX", strForeignDescX, False, False)
            End If

            queryBuilder.AddField("wl", intWL, True)
            queryBuilder.AddField("Len", intLen, True)
            queryBuilder.AddField("type", intType, True)
            queryBuilder.AddField("dp", intDP, True)

            'FieldExist() takes connstr corr to actualDB(orcl/acc/sql) where as we need conn str corr to easesys.mdb. so pass it explicitly. Change the DataSource as needed 
            If Ec.IO.FieldExist("easesys1", "LastUpdated", strCS) Then _
' this field may not be present for existing cust. Added this flag to cover existing cust
                queryBuilder.AddField("LastUpdated", dtDate, False) _
                'Fill date - needed while export/import of easesys.mdb to/from excel
            End If


            If intQueryType = QueryBuilder.QueryType.Insert Then
                queryBuilder.AddField("id", intID, True)
            ElseIf intQueryType = QueryBuilder.QueryType.Update Then
                queryBuilder.AddConditionField("id", intID.ToString(), True)
            End If
            strSQL = queryBuilder.GenerateQuery

            If intID >= 4600 And intID <= 4900 Then
                'protect application constants
                If blnProtectAppConstants Then GoTo ExitThisSub
            End If


            Using connection = New EaseCore.DAL.Connection(strCS)
                connection.ExecuteNonQuery(strSQL)

                If blnUpdateDateTimeStamp Then
                    intPos = 16
                    intQueryType = EASEClass7.QueryBuilder.QueryType.Update

                    'update the last updated date in the words file. this is used to update the client's words file
                    Dim strTemp = strLastUpdatedDateTime.Trim
                    If strTemp.Trim = "" Then strTemp = GetCurrentDateTimeForWordUpdate()

                    queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(intQueryType, "easesys1")
                    queryBuilder.AddField("descx", strTemp, False, False)
                    If intLanguage = 2 Then
                        queryBuilder.AddField("ForeignDescX2", "", False, False)
                    ElseIf intLanguage = 3 Then
                        queryBuilder.AddField("ForeignDescX3", "", False, False)
                    ElseIf intLanguage = 4 Then
                        queryBuilder.AddField("ForeignDescX4", "", False, False)
                    ElseIf intLanguage = 5 Then
                        queryBuilder.AddField("ForeignDescX5", "", False, False)
                    ElseIf intLanguage = 6 Then
                        queryBuilder.AddField("ForeignDescX6", "", False, False)
                    ElseIf intLanguage = 7 Then
                        queryBuilder.AddField("ForeignDescX7", "", False, False)
                    ElseIf intLanguage = 8 Then
                        queryBuilder.AddField("ForeignDescX8", "", False, False)
                    ElseIf intLanguage = 9 Then
                        queryBuilder.AddField("ForeignDescX9", "", False, False)
                    ElseIf intLanguage = 10 Then
                        queryBuilder.AddField("ForeignDescX10", "", False, False)
                    ElseIf intLanguage = 11 Then
                        queryBuilder.AddField("ForeignDescX11", "", False, False)
                    ElseIf intLanguage = 12 Then
                        queryBuilder.AddField("ForeignDescX12", "", False, False)
                    ElseIf intLanguage = 13 Then
                        queryBuilder.AddField("ForeignDescX13", "", False, False)
                    ElseIf intLanguage = 14 Then
                        queryBuilder.AddField("ForeignDescX14", "", False, False)
                    ElseIf intLanguage = 15 Then
                        queryBuilder.AddField("ForeignDescX15", "", False, False)
                    ElseIf intLanguage = 16 Then
                        queryBuilder.AddField("ForeignDescX16", "", False, False)
                    Else
                        queryBuilder.AddField("ForeignDescX", "", False, False)
                    End If
                    queryBuilder.AddField("wl", 0, True)
                    queryBuilder.AddField("Len", 0, True)
                    queryBuilder.AddField("type", 0, True)
                    queryBuilder.AddField("dp", 0, True)
                    queryBuilder.AddConditionField("id", intPos.ToString(), True)

                    strSQL = queryBuilder.GenerateQuery
                    connection.ExecuteNonQuery(strSQL)
                End If
            End Using
            ExitThisSub:
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            Call GenerateException("WriteWordRecord", ex)
        End Try

#If TRACE Then
        Log.EASESYS_IO_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Property EASESys1Object() As DataTable
        'set/returns the EASESYS1 table 
        Get
            Return gObjEasesys1
        End Get
        Set(ByVal Value As DataTable)
            gObjEasesys1 = Value
        End Set
    End Property

#End Region

#Region "***  Appconfig Properties (Get/Set)"

    ''' <summary>
    ''' Get/Set EASE configuration parameter (EX: 100000000)
    ''' </summary>
    Public Property EASEConfig() As String
        Get
            Return Trim(gObjAppConfig.EASEConfig)
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.EASEConfig = Trim(Value)
        End Set
    End Property

    ''' <summary>
    ''' Get/Set first field value from EASE configuration parameter (returns 1 if EASE Config is 100000000)
    ''' </summary>

        Public Property Modf() As Byte
        Get
            Return gObjAppConfig.Modf
        End Get
        Set(ByVal Value As Byte)
            gObjAppConfig.Modf = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Second field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property Modf2() As Byte
        Get
            Return gObjAppConfig.Modf2
        End Get
        Set(ByVal Value As Byte)
            gObjAppConfig.Modf2 = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set EASE application version.
    ''' </summary>
    Public Property Version() As String
        Get
            Return Trim(gObjAppConfig.Version)
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.Version = Trim(Value)
        End Set
    End Property

    Public Property EASEConfigDirectory() As String
        Get
            Return gObjAppConfig.EASEConfigDirectory.Trim
        End Get
        Set(ByVal value As String)
            gObjAppConfig.EASEConfigDirectory = value
        End Set
    End Property

    Public Property EASETempPath() As String
        Get
            Return gObjAppConfig.EASETempPath.Trim
        End Get
        Set(ByVal value As String)
            gObjAppConfig.EASETempPath = value
        End Set
    End Property

    Public ReadOnly Property EASETempDirectory() As String
        Get
            Return gObjAppConfig.EASETempPath.Trim
        End Get
    End Property

    ''' <summary>
    ''' Get/Set Shared EASE directory.
    ''' </summary>
    Public Property SharedDirectory() As String
        Get
            Return Trim(gObjAppConfig.SharedDirectory)
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.SharedDirectory = Trim(Value)
        End Set
    End Property

    ''' <summary>
    ''' Get/Set EASE database key.
    ''' </summary>
    Public Property DBKey() As Integer
        Get
            Return gObjAppConfig.DBKey
        End Get
        Set(ByVal Value As Integer)
            gObjAppConfig.DBKey = Value
        End Set
    End Property

    Public Property PowerpointPath() As String
        Get
            Return Trim(gObjAppConfig.PowerpointPath)
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.PowerpointPath = Trim(Value)
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Adobe Acrobat Reader path.
    ''' </summary>
    Public Property AcrobatPath() As String
        Get
            Return Trim(gObjAppConfig.AcrobatPath)
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.AcrobatPath = Trim(Value)
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Customer name.
    ''' </summary>
    Public Property CustomerName() As String
        Get
            Return Trim(gObjAppConfig.CustomerName)
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.CustomerName = Trim(Value)
        End Set
    End Property

    Public Property RetailEASE() As Boolean
        Get
            Return gObjAppConfig.RetailEASE
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.RetailEASE = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set whether the application is configured to run in 'Logistics' mode.
    ''' </summary>
    Public Property Logistics() As Boolean
        Get
            Return gObjAppConfig.Logistics
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.Logistics = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set whether the application is configured to run in 'Office' mode.
    ''' </summary>
    Public Property Office As Boolean
        Get
            Return gObjAppConfig.Office
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.Office = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set CAPP flag, Third field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property CAPP As Boolean
        Get
            Return gObjAppConfig.CAPP
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.CAPP = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set CAPP only flag, ninth field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property CAPPOnly() As Boolean
        Get
            Return gObjAppConfig.CAPPOnly
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.CAPPOnly = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Single User License flag,  tenth field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property SingleUserLicense() As Boolean
        Get
            Return gObjAppConfig.SingleUserLicense
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.SingleUserLicense = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Expert EASE Flag, sixth field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property ExpertEASE() As Boolean
        Get
            Return gObjAppConfig.ExpertEASE
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.ExpertEASE = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set MDM flag,  11th field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property MDM() As Boolean
        Get
            Return gObjAppConfig.MDM
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.MDM = Value
        End Set
    End Property
    'R.J to SET MES and Tool Management Flag
    '''' <summary>
    '''' Get/Set MES flag,  17th field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    '''' </summary>
    'Public Property MES() As Boolean
    '    Get
    '        Return gObjAppConfig.MES
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        gObjAppConfig.MES = Value
    '    End Set
    'End Property
    ''' <summary>
    ''' Get/Set ToolManagement flag,  18th field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property ToolManagement() As Boolean
        Get
            Return gObjAppConfig.ToolManagement
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.ToolManagement = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set MCR flag,  12th field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property MCR() As Boolean
        Get
            Return gObjAppConfig.MCR
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.MCR = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set UAS User flag, eighth field value from EASE configuration parameter (returns 0 if EASE Config is 100000000)
    ''' </summary>
    Public Property UASUser() As Boolean
        Get
            Return gObjAppConfig.UASUser
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.UASUser = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Control Plan flag (if CAPP is enabled, then Control Plan is also enabled)
    ''' </summary>
    Public Property ControlPlan() As Boolean
        Get
            Return gObjAppConfig.ControlPlan
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.ControlPlan = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Line Balance flag (4th position in EASE configuration)
    ''' </summary>
    Public Property LineBalance() As Boolean
        Get
            Return gObjAppConfig.LineBalance
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.LineBalance = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Machine flag (7th position in EASE configuration)
    ''' </summary>
    Public Property Machine() As Boolean
        Get
            Return gObjAppConfig.Machine
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.Machine = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Material flag (if modf2 flag is 3))
    ''' </summary>
    Public Property Material() As Boolean
        Get
            Return gObjAppConfig.Material
        End Get
        Set(ByVal Value As Boolean)
            gObjAppConfig.Material = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set MH Flag
    ''' </summary>
    Public Property MHFlag() As Int16
        Get
            Return gObjAppConfig.MHFlag
        End Get
        Set(ByVal Value As Int16)
            gObjAppConfig.MHFlag = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Time Unit (1-Hours, 60-Minutes, 3600-Seconds)
    ''' </summary>
    Public Property TimeUnit() As Int16
        Get
            Return gObjAppConfig.TimeUnit
        End Get
        Set(ByVal Value As Int16)
            gObjAppConfig.TimeUnit = Value

            'reset the time format
            SetTimeFormat(Value)
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Time format for time units (Hours, Minutes, Seconds)
    ''' </summary>
    Public Property TimeFormat() As String
        Get
            Return gObjAppConfig.TimeFormat
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.TimeFormat = Value
        End Set
    End Property

    Public ReadOnly Property CurrentCulture() As String
        Get
            Return System.Threading.Thread.CurrentThread.CurrentCulture.Name
        End Get
    End Property


    ''' <summary>
    ''' Set Time format for the time unit.
    ''' </summary>
    Public Sub SetTimeFormat(ByVal intTimeUnit As Int16)

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strFormat As String = ""
        Select Case intTimeUnit
            Case 1
                strFormat = TimeFormatHours()
            Case 60
                strFormat = TimeFormatMinutes()
            Case Else
                strFormat = TimeFormatSeconds()
        End Select

        gObjAppConfig.TimeFormat = strFormat

#If TRACE Then
        Log.UTILITY_LOW("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    ''' <summary>
    ''' Set Numeric Precision format for the time unit.
    ''' </summary>
    Public Sub SetNumericFormat()

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        gObjAppConfig.NumericFormat = "####0.0000000"

#If TRACE Then
        Log.UTILITY_LOW("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function TimeFormatHours() As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strTimeFormatHours As String = "####0.00000"

#If TRACE Then
        Log.UTILITY_LOW(String.Format("Exit ({0})", strTimeFormatHours), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strTimeFormatHours
    End Function

    Public Function TimeFormatMinutes() As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strTimeFormatMinutes As String = "#####0.0000"

#If TRACE Then
        Log.UTILITY_LOW(String.Format("Exit ({0})", strTimeFormatMinutes), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strTimeFormatMinutes
    End Function

    Public Function TimeFormatSeconds() As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strTimeFormatSeconds As String = "######0.000"


#If TRACE Then
        Log.UTILITY_LOW(String.Format("Exit ({0})", strTimeFormatSeconds), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strTimeFormatSeconds
    End Function

    Public Function GetPCSPerHourFormat() As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strFormat As String = "######0.000"
        If Ec.AppConfig.SubZero Then
            'EVC-680: Subzero: PCS/HR in single precision
            strFormat = "######0.0"
        End If

#If TRACE Then
        Log.UTILITY_LOW(String.Format("Exit ({0})", strFormat.Trim), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strFormat.Trim '"######0.000"
    End Function

    ''' <summary>
    ''' Get/Set Time Unit (Hours, Minutes, Seconds)
    ''' </summary>
    Public Property NumericFormat() As String
        'GetNumberFormat
        'GetNumericFormat
        Get
            Return gObjAppConfig.NumericFormat
        End Get
        Set(ByVal Value As String)
            gObjAppConfig.NumericFormat = Value
        End Set
    End Property
    '
    ''' <summary>
    ''' Get/Set HarleyMil flag. (Applicable to All harley plants)
    ''' </summary>
    Public Property HarleyMil() As Boolean
        Get
            Return gHarleyMil
        End Get
        Set(ByVal Value As Boolean)
            gHarleyMil = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Harley-York flag.
    ''' </summary>
    Public Property HarleyYork() As Boolean
        Get
            Return gHarleyYork
        End Get
        Set(ByVal Value As Boolean)
            gHarleyYork = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Cummins flag.
    ''' </summary>
    Public Property Cummins() As Boolean
        Get
            Return gCummins
        End Get
        Set(ByVal Value As Boolean)
            gCummins = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Cummins flag.
    ''' </summary>
    Public Property EASE() As Boolean
        Get
            Return gEaseUser
        End Get
        Set(ByVal Value As Boolean)
            gEaseUser = Value
        End Set
    End Property


    ''' <summary>
    ''' Get/Set Armada flag.
    ''' </summary>
    Public Property Armada() As Boolean
        Get
            Return gArmada
        End Get
        Set(ByVal Value As Boolean)
            gArmada = Value
        End Set
    End Property

    ''' <summary>
    ''' Get/Set Diesel Exchange flag.
    ''' </summary>
    Public Property DieselExchange() As Boolean
        Get
            Return gDieselExchange
        End Get
        Set(ByVal Value As Boolean)
            gDieselExchange = Value
        End Set
    End Property

    Public ReadOnly Property MaximumFileSizeLimit() As Long
        Get
            Return 31457280     '30 MB
        End Get
    End Property

    Public Function Hitachi() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("HITACHI")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult
    End Function

    Public Function Navy() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("NAVY")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function SubZero() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("ZERO")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function ChristieDigital() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("CHRISTIE")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function


    Private _astonMartin As Boolean?

    ReadOnly Property AstonMartin As Boolean
        Get
            If _AstonMartin Is Nothing Then
                _AstonMartin = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("ASTON")
            End If
            Return _AstonMartin.Value
        End Get
    End Property

    Private _astonMartinStAthan As boolean?

    Readonly Property AstonMartinStAthan As Boolean
        Get
            If (_astonMartinStAthan Is nothing) then
                _astonMartinStAthan = AstonMartin AndAlso Trim(UCase(Ec.AppConfig.CustomerName)).Contains("ATHAN")
            End If
            return _astonMartinStAthan.Value
        End Get
    End Property

    Private _astonMartinLagonda As boolean?

    Readonly Property AstonMartinLagonda As Boolean
        Get
            If (_astonMartinLagonda Is nothing) then
                _astonMartinLagonda = AstonMartin AndAlso Trim(UCase(Ec.AppConfig.CustomerName)).Contains("LAGONDA")
            End If
            return _astonMartinLagonda.Value
        End Get
    End Property

    Public Function GeneralDynamics() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("DYNAMICS")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function TeslaMotors() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("TESLA")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function StandardTextile() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("TEXTILE")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function OregonIronWorks() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        blnResult = Trim(UCase(Ec.AppConfig.CustomerName)).Contains("OREGON")

#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

#End Region


    Public Function GetLanguageList() As stLanguageList()

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objLanguages(0) As stLanguageList
        Dim strWord As String = "", intWord As Int16 = 0
        Dim intType As Int16 = 0, intPos As Integer = 0
        Dim intK = 0
        Try
            intPos = 0
            strWord = "English"
            intK += 1
            ReDim Preserve objLanguages(intK)
            objLanguages(intK).LanguageID = intPos
            objLanguages(intK).LanguageDescX = strWord.Trim

            intPos = 1
            strWord = Ec.AppConfig.GetWrd(4861, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 2
            intWord = 4862
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 3
            intWord = 4863
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If
            intPos = 4
            intWord = 4864
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 5
            intWord = 4865
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 6
            intWord = 4866
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 7
            intWord = 4867
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 8
            intWord = 4868
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 9
            intWord = 4869
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 10
            intWord = 4870
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 11
            intWord = 4871
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 12
            intWord = 4872
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 13
            intWord = 4873
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 14
            intWord = 4874
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 15
            intWord = 4875
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

            intPos = 16
            intWord = 4876
            strWord = Ec.AppConfig.GetWrd(intWord, , , intType, , True)
            If intType = 77 Then
                intK += 1
                ReDim Preserve objLanguages(intK)
                objLanguages(intK).LanguageID = intPos
                objLanguages(intK).LanguageDescX = strWord.Trim
            End If

        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("GetLanguageList: ", ex)
        End Try


#If TRACE Then
        Log.SYSTEM_CONFIG_LOW(String.Format("Exit Count:({0})", UBound(objLanguages)), LOG_APPNAME, BASE_ERRORNUMBER + 0,
                              startTicks)
#End If
        Return objLanguages
    End Function

    Private Shared bCheckedCatSupport As Boolean = False
    Private Shared bCatSupport As Boolean = False

    Function IsCategoriesSupported(Optional blnRecheck As Boolean = False) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG(String.Format("Enter blnRecheck:({0})", blnRecheck), LOG_APPNAME,
                                                   BASE_ERRORNUMBER + 0)
#End If

        If blnRecheck Or Not bCheckedCatSupport Then
            Using connection = New EaseCore.DAL.Connection(gStrConnectionString)
                bCheckedCatSupport = True
                bCatSupport = EaseCore.DAL.Schema.TableExist(connection, "PDMDocCategories")
                If bCatSupport Then
                    bCatSupport = EaseCore.DAL.Schema.FieldExist(connection, "shmm", "doccategory")
                End If
            End Using
        End If


#If TRACE Then
        Log.SYSTEM_CONFIG(String.Format("Exit ({0})", bCatSupport), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return bCatSupport
    End Function

    Public Function EnableKVI() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.SYSTEM_CONFIG("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnvalue As Boolean = False
        'R.J === Temporarily KVI is disabled for all Retail EASE Customers
        Try
            If Ec.DBConfig.IsKviForLogistics Or Ec.DBConfig.EnableUserDefinedKVI Then 'And Ec.AppConfig.RetailEASE 
                blnRtnvalue = True
            End If
        Catch ex As Exception
            GenerateException("EnableKVI", ex)
        End Try


#If TRACE Then
        Log.SYSTEM_CONFIG(String.Format("Exit ({0})", blnRtnvalue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnvalue
    End Function

    Public Function GetLanguageID(ByVal strLanguage As String) As Int16

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW(String.Format("Enter strLanguage:({0})", strLanguage), LOG_APPNAME,
                                                 BASE_ERRORNUMBER + 0)
#End If

        'NOTE: Get the language ID from the Language name 
        Dim intLangID As Int16 = 0
        Try
            Select Case strLanguage.Trim.ToUpper
                Case "CHINESE"
                    intLangID = 1
                Case "TURKISH"
                    intLangID = 2
                Case "KOREAN"
                    intLangID = 3
                Case "PORTUGUESE"
                    intLangID = 4
                Case "GERMAN"
                    intLangID = 5
                Case "HINDI"
                    intLangID = 6
                Case "JAPANESE"
                    intLangID = 7
                Case "SPANISH"
                    intLangID = 8
                Case "RUSSIAN"
                    intLangID = 9
                Case "VIETNAMESE"
                    intLangID = 10
                Case "PORTUGESE"
                    intLangID = 11
                Case "ITALIAN"
                    intLangID = 12
                Case "THAI", "FRENCH" 'TODO: DH THAI in JUNO, FRENCH in MASTER
                    intLangID = 13
                Case "MALASIAN"
                    intLangID = 14
                Case "FINNISH"
                    intLangID = 15
                Case "ROMANIAN"
                    intLangID = 16
                Case Else 'English
                    intLangID = 0
            End Select
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
            GenerateException("AppConfig: GetLanguageID", ex)
        End Try


#If TRACE Then
        Log.UTILITY_LOW(String.Format("Exit ({0})", intLangID), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return intLangID
    End Function

    Public Function GetLanguagefromID(ByVal intLangID As Int16) As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_LOW(String.Format("Enter intLangID:({0})", intLangID), LOG_APPNAME,
                                                 BASE_ERRORNUMBER + 0)
#End If

        'NOTE: Get the language ID from the Language name 
        Dim strLanguage As String = ""
        Try
            Select Case intLangID
                Case 1
                    strLanguage = "Chinese"
                Case 2
                    strLanguage = "Turkish"
                Case 3
                    strLanguage = "Korean"
                Case 4
                    strLanguage = "Portuguese"
                Case 5
                    strLanguage = "German"
                Case 6
                    strLanguage = "Hindi"
                Case 7
                    strLanguage = "Japanese"
                Case 8
                    strLanguage = "Spanish"
                Case 9
                    strLanguage = "Russian"
                Case 10
                    strLanguage = "Vietnamese"
                Case 11
                    strLanguage = "Portugese"
                Case 12
                    strLanguage = "Italian"
                Case 13
                    strLanguage = "French" 'TODO: DH THAI in JUNO, FRENCH in MASTER
                Case 14
                    strLanguage = "Malasian"
                Case 15
                    strLanguage = "Finnish"
                Case 16
                    strLanguage = "Romanian"
                Case Else
                    strLanguage = "English"
            End Select
        Catch ex As Exception
            GenerateException("AppConfig: GetLanguageID", ex)
        End Try


#If TRACE Then
        Log.UTILITY_LOW(String.Format("Exit ({0})", strLanguage), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strLanguage
    End Function

    Function GetVersionFromText(ByVal strFilename As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.FILE_DIR_IO_LOW(String.Format("Enter strFilename:{0}", strFilename), LOG_APPNAME,
                                                     BASE_ERRORNUMBER + 0)
#End If
        Dim versionText As String = ""
        Dim lineCount As Integer = 0

        If System.IO.File.Exists(strFilename) = True Then
            Dim objReader As New System.IO.StreamReader(strFilename)

            versionText = versionText & objReader.ReadLine()
        Else
            versionText = ""
        End If


#If TRACE Then
        Log.FILE_DIR_IO_LOW(String.Format("Exit: ({0})", versionText), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return versionText
    End Function

    Public Function GetTimeStampFromText(ByVal strFilename As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.FILE_DIR_IO_LOW(String.Format("Enter strFilename:{0}", strFilename), LOG_APPNAME,
                                                     BASE_ERRORNUMBER + 0)
#End If

        Dim TimeStampText As String = ""
        Dim lineCount As Integer = 0

        If System.IO.File.Exists(strFilename) = True Then
            Dim objReader As New System.IO.StreamReader(strFilename)
            TimeStampText = objReader.ReadLine()
            TimeStampText = objReader.ReadLine()
        Else
            TimeStampText = ""
        End If


#If TRACE Then
        Log.FILE_DIR_IO_LOW(String.Format("Exit: ({0})", TimeStampText), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return TimeStampText
    End Function

    Public Function GetAdminGroupName(ByVal intAdminID As Int16) As String
#If TRACE Then
        Dim startTicks As Long = Log.EASESYS_IO_MED(String.Format("Enter: intAdminID:({0})", intAdminID), LOG_APPNAME,
                                            BASE_ERRORNUMBER + 0)
#End If
        'Admin Group ID: 0-None, 1-Administrator (Local), 2-Admin (Corporate), 3-Admin (IT), 4- Supervisor, 5-> LPA - Admin, 6 -> LPA-Auditor, 7-> LPA Manager
        Dim strTemp As String = ""
        Select Case intAdminID
            Case 1
                strTemp = GetWrd(260)
            Case 2
                strTemp = GetWrd(261)
            Case 3
                strTemp = GetWrd(262)
            Case 4
                strTemp = GetWrd(5555) 'Supervisor
            Case 5
                strTemp = GetWrd(5588) & " " & GetWrd(6785)    'LPA Administrator
            Case 6
                strTemp = GetWrd(5588) & " " & GetWrd(5493)    'LPA Auditor
            Case 7
                strTemp = GetWrd(5588) & " " & GetWrd(7159)    'LPA Manager
            Case 8
                strTemp = GetWrd(3793)    'QC
            Case 9
                strTemp = GetWrd(9707)    'QA
            Case 10
                strTemp = GetWrd(2602)    'Customer

            Case Else
                strTemp = GetWrd(263)
        End Select
#If TRACE Then
        Log.EASESYS_IO_MED(String.Format("Exit: ({0})", strTemp), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strTemp
    End Function
End Class
