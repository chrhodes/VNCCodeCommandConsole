Option Infer On
Option Strict On

Imports System.Drawing
Imports EaseCore
Imports System.Linq
Imports EASEClass7.DataContracts
Imports EASEClass7.DataContracts.AML_VehiclePart
Imports EaseCore.DAL
Imports EaseCore.Extensions
Imports AML_VehicleGlossary = EASEClass7.DataBrokers.AML_VehicleGlossary

Public Class AML
    Private Shared ReadOnly BASE_ERRORNUMBER As Integer = ErrorNumbers.EASECLASS_AML
    Private Const LOG_APPNAME As String = "EASECLASS"

#Region "Enums, Fields, Properties, Structures"
    Public gApplicationDirectory As String = ""

    Public Structure stBannerMessagesStations
        Public MessageID As Int32
        <VBFixedString(15)> Public Stage As String
        Public StationSeq As Int32
    End Structure

    Public Structure stBannerMsgList 'holds the list of banner messages
        '*** any changes in the structure should be updated in ClearBannerMessageObject SUB ***
        Public MessageID As Int32
        <VBFixedString(250)> Public MessageText As String
        <VBFixedString(30)> Public RaisedBy As String
        <VBFixedString(30)> Public ApprovedBy As String
        <VBFixedString(250)> Public MessageReason As String
        <VBFixedString(60)> Public MessageType As String
        Public LineID As Int16
        <VBFixedString(15)> Public StationFrom As String
        <VBFixedString(15)> Public StationTo As String
        Public StartDate As Date
        Public EndDate As Date
        Public LastUpdated As Date
        Public MessageStatus As Int16           'MessageStatus: 0-Draft/Pending Approval, 1-Approved
        '*** any changes in the structure should be updated in ClearBannerMessageObject SUB ***
    End Structure

    Public Structure stBuildLogData
        <VBFixedString(60)> Public DescX As String
    End Structure

    Public Structure stBuildPartVerifyList
        'Any change should be updated in ClearBuildPartVerifyObject too  -----------------------
        Public BuildNo As String
        Public Mseq As Int32
        Public StationNO As String
        Public PartNumber As String
        Public VerifyStatus As Int16 'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
        Public UserID As String
        Public OperatorPosition As Int16
        Public Verify_Station As String
        Public Spare2 As String   'holds Subheader reference number KM: 05/29/2014
        Public Spare3 As String
        Public PartDescX As String
        Public Tracking As String

        'Any change should be updated in ClearBuildPartVerifyObject too  -----------------------
    End Structure

    Public Structure stFLMShiftFrequency
        Public IDX As Byte
        Public Shift1 As Byte
        Public Shift2 As Byte
        Public Shift3 As Byte
        <VBFixedString(30)> Public DescX As String
        Public FrequencyType As Byte 'FrequencyType: 1-Daily, 2- Weekly, 3-Monthly
    End Structure

    Public Structure stLastOpenSubHeader
        Public ID As Int32
        <VBFixedString(1)> Public Rectype As String
        Public Seq As Int16
        <VBFixedString(6)> Public OPNO As String
        <VBFixedString(80)> Public SubHeader As String
        <VBFixedString(100)> Public CommentX As String
        Public OperatorPosition As Int16
    End Structure

    Public Structure stLines
        Public LineID As Int16
        Public PlantID As Int16
        <VBFixedString(80)> Public DescX As String
    End Structure

    Public Structure stLineStatus        'holds the list of build numbers by station for the line
        '** Any changes should be changed in ClearLineStatusObject **'
        Public LineID As Integer
        <VBFixedString(12)> Public StationNO As String
        <VBFixedString(10)> Public BuildNO As String
        Public EngineMoveDate As Date
        Public QualityChecksStatus As Int16     'QualityChecksStatus: -1 -> InComplete, 0-FAIL, 1->PASS
        Public AndonCall As Integer 'Unsigned integer '0-default, 1-first call, 2-second call
        Public OperatorCount As Integer ' AfterEAESClass7 Consolidation:7th August, set in RefreshBuildNOInLineStatus (modaml1.vb, amlviewease)
        Public OSMRequired As Boolean
        '** Any changes should be changed in ClearLineStatusObject **'
    End Structure

    Public Structure stORCGroupItems
        Public GroupID As Int16
        Public CategoryID As Int16
        <VBFixedString(25)> Public Category As String
        Public CategorySeq As Int16
    End Structure

    ''' <summary>
    ''' Holds list of Overridecategory GRoup
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure stOverridecategoryGroup
        Public GroupID As Int16
        <VBFixedString(25)> Public GroupName As String
        Public Seq As Int16
    End Structure

    Public Structure stPartsVerifyList  'holds the list of parts that needs to be verified

        '***************** Any change should be updated in ClearPartVerifyListObject *****************
        <VBFixedString(20)> Public PartNumber As String
        <VBFixedString(60)> Public Supplier_PartNO As String
        '***************** Any change should be updated in ClearPartVerifyListObject *****************
    End Structure

    Public Structure stPLCBuildList
        Public LineID As Int16
        <VBFixedString(6)> Public BuildNo As String
    End Structure

    Public Structure stPSSDocuments

        '** Any change should be updated DateRead ClearPSSDocumentsObject
        <VBFixedString(80)> Public DocDescx As String
        Public TTKey As Int32
        Public CompleteTTKey As Int32
        Public MDMDocID As Long
        Public DocRev As Int32
        Public ReleaseDate As Int32
        <VBFixedString(6)> Public OPNO As String
        <VBFixedString(80)> Public SHDescx As String


        'Public ID As Long
        '<VBFixedString(1)> Public Rectype As String
        'Public SEQ As Int16
        Public SHID As Long
        Public SharedSHID As Long

        '        <VBFixedString(6)> Public ReferenceNO As String
        Public CheckedRead As Byte
        Public PassStatus As Int16
        Public DateRead As Date


        '** Any change should be updated into ClearPSSDocumentsObject
    End Structure

    Public Structure stPSSDocuments_READ
        <VBFixedString(30)> Public USERID As String
        <VBFixedString(6)> Public StationNO As String
        <VBFixedString(6)> Public OPNO As String
        Public LineID As Int16
        Public MDMDocID As Long
        Public TTKey As Int32
        Public CompleteTTKey As Int32
        Public CheckedRead As Byte
        Public PassStatus As Int16
        Public DateRead As Date
    End Structure

    <Serializable()>
    Public Structure stQualityChecks
        Public PlanType As Int16
        Public RecordSeq As Int16
        Public RemedyPlanSeq As Int16
        Public UserInput As String
        Public UserInputFailed As String        'holds the first fail data
        Public SuperVisorAction As String
        Public FailureDesc As String
        Public Spare1 As String
        Public PassFail As Int16                                    'PDMRPItemValues->PassFail: -1 -> InComplete, 0-FAIL, 1->PASS, 2-Override
        Public CommentX As String
        Public RemedyPlanDesc As String         'RemedyPlanDesc
        Public SubHeaderDesc As String          'SubHeaderDesc
        Public OPNO As String                    'Operation NO
        Public StationNO As String
        Public UserInput_Time_Stamp As Date                                             'hold the date in string format
        Public LowerValue As Single
        Public UpperValue As Single
        Public CheckMethod As String
        Public Datex As Date
        Public InvalidEntry As Boolean
        Public OperatorID As String                                 'holds user name who completed the QC/FLM record
        Public Shift As Byte
        Public FrequencyType As Int16

    End Structure

    Public Structure stQualityChecksSearch
        '** Any changes should be changed in ClearQualityChecksSearchObject **'
        <VBFixedString(6)> Public BuildNo As String
        <VBFixedString(10)> Public StationNO As String
        Public RecordSeq As Int16
        Public PlanType As Int16
        Public PassFail As Integer      'PDMRPItemValues->PassFail: -1 -> InComplete, 0-FAIL, 1->PASS, 2-Override
        Public RecordType As Int16      'Record type: 0-Quality checks/Remedy Plan ,1- FLM record
        <VBFixedString(40)> Public PartNO As String
        <VBFixedString(6)> Public OPNO As String
        <VBFixedString(80)> Public SubHeader As String
        Public Shift As Byte
        '** Any changes should be changed in ClearQualityChecksSearchObject **'
    End Structure

    Public Structure stRemedyDocs           'holds records from ophmm table (subheader record and the remedyplan document)
        Public PartID As Long
        Public RecType As String
        Public Seq As Integer
        Public OPNO As String
        Public FileName As String
        Public MDMDocID As Int32
        Public TTKey As Int32
        Public CompleteTTKey As Int32
        Public Marked As Integer
        Public PartNO As String         'option no
        Public MDMDocSeq As Int32
        Public MDMDocRectype As Integer
        Public RemedyPlanDesc As String         'RemedyPlnDesc
        Public PlanType As Int16        'PlanType: 0-Audit Record, 1-Operator Check
        Public SHID As Int32
        Public Shift As Integer
        Public FrequencyType As Integer
    End Structure

    Public Structure stRemedyPlanItemTemplateAndValues           'holds the remedy plan details entered in pdm editor
        Public DOCID As Long
        <VBFixedString(1)> Public DocRecType As String
        Public DocSeq As Integer
        Public Seq As Integer
        Public PlanType As Integer                          'PlanType: 0-Audit Record, 1-Operator Check
        <VBFixedString(80)> Public FailureDesc As String
        <VBFixedString(6)> Public OrigOpNo As String
        <VBFixedString(1)> Public Severity As String
        Public LowerValue As Single
        Public UpperValue As Single
        <VBFixedString(5)> Public Unit As String
        <VBFixedString(1)> Public StopBuild As String
        <VBFixedString(80)> Public Comment As String
        <VBFixedString(1)> Public InputReqd As String
        <VBFixedString(15)> Public CheckMethod As String
        <VBFixedString(10)> Public Spare1 As String         'used only for AML
    End Structure

    <Serializable()>
    Public Structure stSafetyCheck
        '** Any changes should be changed in ClearSafetyCheckObject **'
        <VBFixedString(15)> Public OperatorID As String 'UserID
        Public WSID As Long
        <VBFixedString(45)> Public WSDescX As String
        <VBFixedString(15)> Public WSCode As String
        Public DOCID As Long
        <VBFixedString(1)> Public DocRecType As String
        Public DocSeq As Integer
        <VBFixedString(80)> Public DocDesc As String
        Public RevNum As Short
        <VBFixedString(15)> Public RevDate As Date
        Public DateRead As String 'Date
        Public CheckedRead As Boolean
        Public PassStatus As Int16 '0->Fail, 1-> Pass, -1-> Cancel
        Public UpdateRecord As Boolean


        Public PlantID As Int16                                 'PlantID 
        <VBFixedString(15)> Public WorkgroupID As String        'WorkGroupID
        <VBFixedString(15)> Public WC As String                 'WorkCenterID
        <VBFixedString(30)> Public Engineer As String
        Public DocStatus As Int16
        Public EnggFunction As Int16                            'EnggFunction-> 1-Layout, 2-Manufacturing, 3-Maintenance, 4-QC, 5-Safety,6-Training
        Public MSeq As Int32
        Public InfoType As Int32        '1-> JobAid, 3-> Ref Docs, 2 -> Graphics, 4-> Video
        <VBFixedString(80)> Public Filename As String
        <VBFixedString(255)> Public FilepathX As String
        Public TTKey As Int32
        Public DateX As Int32 'same as revdate
        Public CompleteTTKey As Int32
        <VBFixedString(15)> Public KeyField As String
        <VBFixedString(15)> Public KeyFieldValue As String
        Public PrintOrder As Int16          'if printorder is 77, then it's a shared doc (not stored in location text tables)
        '** Any changes should be changed in ClearSafetyCheckObject **'
    End Structure

    Public Structure stShiftTimes
        Public ShiftID As Byte
        <VBFixedString(20)> Public ShiftName As String
        Public StartTime As Single
        Public EndTime As Single
        Public Enabled As Byte
    End Structure

    Public Structure stSHOptionCodeList       'hold the list generic options for the subheader
        <VBFixedString(10)> Public OptionCode As String
        <VBFixedString(40)> Public OptionName As String
    End Structure

    Public Structure stSpecificOptions  'holds the specific options list (configured in ClientEditor)
        Public OFID As Int16
        Public ID As Int32
        <VBFixedString(10)> Public Code As String
        <VBFixedString(80)> Public Name As String
    End Structure

    Public Structure stSQL
        <VBFixedString(300)> Public SQL As String
        <VBFixedString(60)> Public KeyField As String
    End Structure

    Public Structure stStations
        Public LineID As Integer
        <VBFixedString(10)> Public StationNO As String
        <VBFixedString(10)> Public StationStart As String
        <VBFixedString(10)> Public StationEnd As String
        Public Operators As Integer
        Public OSMRequired As Int16
        Public Desc As String
    End Structure

    Public Structure stSubHeaderOptions
        '*** Any change in the structure should be updated in ClearSubHeaderOptionsObject sub too ***
        Public ID As Long
        <VBFixedString(1)> Public RecType As String
        Public Seq As Integer
        <VBFixedString(6)> Public OPNO As String
        Public SHID As Int32
        Public Mseq As Integer
        Public SHOptionsFlag As Int32
        <VBFixedString(60)> Public SubModel As String
        <VBFixedString(1)> Public Body As String
        <VBFixedString(1)> Public GearBox As String
        <VBFixedString(4)> Public YearX As String
        <VBFixedString(4)> Public Territory As String
        <VBFixedString(1)> Public Performance As String
        <VBFixedString(1)> Public Drive As String
        <VBFixedString(6)> Public SpecificOptions1 As String
        <VBFixedString(6)> Public SpecificOptions2 As String
        <VBFixedString(6)> Public SpecificOptions3 As String
        '*** Any change in the structure should be updated in ClearSubHeaderOptionsObject sub too ***
    End Structure

    Public Structure stUserLastVisitedScreen
        '** Any changes should be changed in ClearUserLastVisitedScreenObject **'
        <VBFixedString(20)> Public IPAddress As String
        <VBFixedString(30)> Public UserID As String
        <VBFixedString(10)> Public BuildNo As String
        <VBFixedString(10)> Public StationNO As String
        Public ID As Int32
        <VBFixedString(1)> Public RecType As String
        Public SEQ As Int32
        <VBFixedString(6)> Public OPNO As String
        <VBFixedString(80)> Public SubHeader As String
        Public PageID As Int16
        Public Spare1 As Int16
        Public OperatorPosition As Integer            ' AfterEAESClass7 Consolidation:7th August
        '** Any changes should be changed in ClearUserLastVisitedScreenObject **'
    End Structure

    Public Structure stVehBOM
        '** Any changes should be changed in ClearVehicleBOMObject **'
        <VBFixedString(4)> Public OptionCode As String
        <VBFixedString(2)> Public NeworMod As String
        <VBFixedString(30)> Public PartNumber As String
        <VBFixedString(40)> Public PartDescription As String
        <VBFixedString(12)> Public CPSC As String
        Public LineNumber As Integer   'holds the reference number
        <VBFixedString(12)> Public StationNumber As String
        <VBFixedString(1)> Public Body As String
        <VBFixedString(1)> Public Drive As String
        <VBFixedString(1)> Public Gear As String
        <VBFixedString(1)> Public Performance As String    'Base or Performance
        <VBFixedString(40)> Public ModelYear As String      'another option code
        Public ModelCode As String
        Public Model1 As Integer
        <VBFixedString(8)> Public TorqueMin As String
        <VBFixedString(8)> Public TorqueMax As String
        <VBFixedString(8)> Public Quantity As String
        <VBFixedString(8)> Public KitNO As String
        Public JointClass As Int16
        Public Angle As Double
        Public NominalTorqueValue As Double
        Public Link As String

        <VBFixedString(40)> Public Validation As String     'holds serial format
        <VBFixedString(2)> Public Tracking As String     'holds Tracking#: Tracking: ST - Serial Tracking, VT-Verification Tracking, BT- Batch Tracking, NT - No Tracking
        Public SupplierCode As String

        '** Any changes should be changed in ClearVehicleBOMObject **'
    End Structure

    Public Structure stVehicleParts
        '** Any changes should be changed in ClearVehiclePartsObject **'
        'IND As String * 2
        <VBFixedString(30)> Public PartNumber As String
        <VBFixedString(40)> Public ShortDesc As String

        <VBFixedString(40)> Public Validation As String     'holds serial format
        <VBFixedString(2)> Public Tracking As String     'holds Tracking#: Tracking: ST - Serial Tracking, VT-Verification Tracking, BT- Batch Tracking, NT - No Tracking
        'LongDesc As String * 60
        'SupplierCode As String * 20
        'SupplierName As String * 40
        'Branch As String * 6
        'SupplierPartNumber As String * 30
        'StatusCode As String * 4
        'SourceCode As String * 4
        'ControlItem As String * 4

        '** Any changes should be changed in ClearVehiclePartsObject **'
    End Structure

    Public Structure stWIPData
        '*** Any changes should be updated in ClearWIPDataObject *** 
        Public EngineID As Int32
        Public LineID As Integer
        <VBFixedString(6)> Public BuildNo As String
        Public StartDate As Date
        <VBFixedString(30)> Public Engineer As String
        Public ModelNumber As Integer
        '*** Any changes should be updated in ClearWIPDataObject *** 
    End Structure

    Structure stImportSharedOps
        Public UniqueNO As Int32
        <VBFixedString(60)> Public OperationDesc As String
    End Structure

    Structure stOptionsFamily
        Public ID As Int16
        <VBFixedString(10)> Public Code As String
        <VBFixedString(40)> Public Name As String
    End Structure

    Public ReadOnly Property DisplayPSSDocumentsDays As Int16
        Get
            Return 60
        End Get
    End Property

    Public ReadOnly Property BodyShopLineId As Integer = 50

    Public ReadOnly Property MainAssemblyLineId As Integer = 20

    Public ReadOnly Property PaintLineId As Integer = 10

#End Region

#Region "Constructors, Initialization, and Load"


#End Region

#Region "Event Handlers"


#End Region

#Region "Main Methods"

    Public Function GetPartShortageReasons() As List(Of PartShortageReason)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = New List(Of PartShortageReason)
        Dim shortageReasons = From t In GetPartShortageReasonsDataTable().AsEnumerable()
                              Select New PartShortageReason With
                              {
                                  .ShortageReasonID = Data.GetDataRowValue(t("shortagereasonid"), 0),
                                  .ShortageReason = Data.GetDataRowValue(t("shortagereason"), "")
                              }
        If (shortageReasons.Any()) Then
            result = shortageReasons.ToList()
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", result.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result
    End Function

    Public Function GetPartShortageReasonsDataTable() As DataTable
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim table As DataTable = Nothing
        Using connection = New Connection(gStrConnectionString)
            Dim sql = "SELECT * FROM partshortagereasons"
            table = connection.GetDataIntoDataTable(sql, Nothing)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return table
    End Function

    Public Function GetPartShortageReason(ByVal shortageReasonID As Integer) As PartShortageReason

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter shortageReasonID:({0})", shortageReasonID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = GetPartShortageReasons().FirstOrDefault(Function(x) x.ShortageReasonID = shortageReasonID)

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result

    End Function


    Public Function UpdatePartShortageReason(ByVal shortageReason As PartShortageReason) As String
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter shortageReason:({0})", shortageReason), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim validationResult = (ValidatePartShortageReason(shortageReason))
        If (String.IsNullOrEmpty(validationResult)) Then
            Using connection = New Connection(gStrConnectionString)
                Dim sql = String.Format("INSERT INTO partshortagereasons VALUES ('{0}')", shortageReason.ShortageReason.ReplaceSingleQuote())
                If (shortageReason.ShortageReasonID > 0) Then
                    sql = String.Format("UPDATE partshortagereasons SET shortagereason = '{0}' WHERE shortagereasonid = {1}", shortageReason.ShortageReason.ReplaceSingleQuote(), shortageReason.ShortageReasonID)
                End If

                connection.ExecuteNonQuery(sql)
            End Using
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit validationResult:({0})", validationResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return validationResult
    End Function

    Public Function DeletePartShortageReason(ByVal shortageReasonID As Integer) As Boolean
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter shortageReasonID:({0})", shortageReasonID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result As Boolean
        Using connection = New Connection(gStrConnectionString)
            Dim sql = String.Format("DELETE FROM partshortagereasons WHERE shortagereasonid = {0})", shortageReasonID)

            connection.ExecuteNonQuery(sql)
            result = True
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit result:({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return result
    End Function


    Public Function GetBuildPartsShortagesDisplayDataTable(Optional ByVal blnAddDummyRow As Boolean = False) As DataTable

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnAddDummyRow:({0})", blnAddDummyRow), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objDataTable As New DataTable

        Dim colDB = New DataColumn("Part #", GetType(String))   '0 -PartnoNO
        objDataTable.Columns.Add(colDB)
        colDB = New DataColumn("Part Description", GetType(String))   '1 -Part Desc
        objDataTable.Columns.Add(colDB)
        colDB = New DataColumn("Quantity", GetType(String))   '2 -QTY
        objDataTable.Columns.Add(colDB)
        colDB = New DataColumn("User", GetType(String))   '3 -User
        objDataTable.Columns.Add(colDB)
        colDB = New DataColumn("Last Updated", GetType(String))   '4 - Last updated
        objDataTable.Columns.Add(colDB)
        colDB = New DataColumn("Mseq", GetType(String))   '5 - Mseq
        objDataTable.Columns.Add(colDB)

        If blnAddDummyRow Then
            Dim objGridTableRow = objDataTable.NewRow()
            objGridTableRow(0) = ""
            objDataTable.Rows.Add(objGridTableRow)
        End If


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objDataTable
    End Function

    Public Function GetVehicleOrdersDropDown() As List(Of VehicleOrders)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As List(Of VehicleOrders)
        Using connection = New Connection(gStrConnectionString)
            Using table = connection.GetDataIntoDataTable("SELECT DISTINCT yearx, model, descriptionx FROM aml_vehicleorders ORDER BY yearx, descriptionx", Nothing)
                Dim result = From t In table.AsEnumerable()
                             Select New VehicleOrders() With
                            {
                                .Year = Data.GetDataRowValue(t("yearx"), 0),
                                .Description = Data.GetDataRowValue(t("Descriptionx"), ""),
                                .Model = Data.GetDataRowValue(t("model"), "")
                            }
                returnValue = result.ToList()
            End Using
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function GetVehicleOrderBuildNumbersByModelAndYear(ByVal model As String, ByVal year As Integer) As List(Of String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter model:({0}) year:({1})", model, year), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As List(Of String)
        Using connection = New Connection(gStrConnectionString)
            Using table = connection.GetDataIntoDataTable(String.Format("SELECT DISTINCT buildnumber FROM aml_vehicleorders WHERE descriptionx = '{0}' AND yearx = {1} ORDER BY buildnumber", model, year), Nothing)
                Dim result = table.AsEnumerable().Select(Function(x) x(0).ToString())
                returnValue = result.ToList()
            End Using
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Sub AddBuildLogErrors(ByRef msgID As Int32, ByRef msgSeq As Int32, ByVal messages As List(Of String))
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter msgID:({0}) messages:({1}) msgSeq:({2})", msgID, messages.Count, msgSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Try
            Using connection = New Connection(gStrConnectionString)
                If msgID = 0 Then
                    msgID = connection.ExecuteScalar(Of Integer)("SELECT MAX(msgid) FROM aml_buildlog", 0)
                    msgID += 1
                    msgSeq = 1
                Else
                    msgSeq = connection.ExecuteScalar(Of Integer)(String.Format("SELECT MAX(msgseq) FROM aml_buildlog WHERE msgid = {0}", msgID), 0)
                    msgSeq += 1
                End If
            End Using

            Dim sqlList As New List(Of String)()
            For Each message In messages
                Dim qb = New DAL.QueryBuilder("aml_buildlog", Connection.DatabaseTypes.SQL)
                qb.QueryType = DAL.QueryBuilder.QueryTypes.Insert
                qb.DataFields.Add("msgid", msgID)
                qb.DataFields.Add("msgseq", msgSeq)
                qb.DataFields.Add("msgtype", 2)
                qb.DataFields.Add("descx", Left(message, 60))
                qb.DataFields.Add("datex", DAL.QueryBuilder.FieldBase.SpecialValues.Now)
                msgSeq += 1
                sqlList.Add(qb.ToString())
            Next

            Ec.IO.RunSQLInArray(sqlList)

        Catch ex As Exception
            Call GenerateException("AddBuildLogErrors", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit msgID:({0}) msgSeq:({1})", msgID, msgSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Sub AddBuildLog(ByRef intMsgID As Int32, ByRef intMsgSeq As Int32,
            ByVal intMsgType As Int16, ByVal strMsg As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intMsgType:({0}) strMsg:({1})", intMsgType, strMsg), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'intMsgType: 1-information, 2-error

        Dim strSQL As String = ""
        Try
            Using connection = New Connection(gStrConnectionString)
                If intMsgID = 0 Then
                    strSQL = "select max(msgid) from AML_BuildLog"
                    intMsgID = connection.ExecuteScalar(Of Integer)(strSQL, 0) + 1
                    intMsgSeq = 1
                End If
                If strMsg.Trim.Length > 60 Then
                    strMsg = Left(strMsg, 60)
                End If

                strSQL = "insert into AML_BuildLog(msgid,msgseq,msgtype,Descx,datex) values (" &
                            intMsgID & "," & intMsgSeq & "," & intMsgType & ",'" &
                            strMsg.ReplaceSingleQuote() & "',getdate())"
                connection.ExecuteNonQuery(strSQL)
                intMsgSeq += 1
            End Using
        Catch ex As Exception
            Call GenerateException("AddBuildLog" & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit intMsgID:({0}) intMsgSeq:({1})", intMsgID, intMsgSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function AnyPartVerifyForBuild(ByVal strBuildNO As String, ByVal strStationNO As String, Optional ByRef blnInComplete As Boolean = False,
                                          Optional ByRef intIncompleteCount As Integer = 0,
                                          Optional intOperatorPosition As Int16 = -1) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) intOperatorPosition:({2})",
                                                          strBuildNO, strStationNO, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'check if any records exist for checking 
        Dim blnRtnValue As Boolean = False
        blnInComplete = False
        intIncompleteCount = 0

        If (Not String.IsNullOrEmpty(strBuildNO)) Then

            Try

                Dim strSQL = String.Format("SELECT verifystatus FROM aml_buildpartsverify WHERE buildnumber = {0} AND {1}", strBuildNO.Trim(), Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO))
                If intOperatorPosition <> -1 Then
                    strSQL &= " AND operatorposition=" & intOperatorPosition
                End If

                Using connection = New EaseCore.DAL.Connection(gStrConnectionString)
                    Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                        If (Not IsNothing(table)) Then
                            For Each row As DataRow In table.Rows
                                blnRtnValue = True
                                Dim intResult = Extensions.Data.GetDataRowValue(row("verifystatus"), 0) 'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride

                                If intResult = -1 Then  'incomplete parts found
                                    intIncompleteCount += 1
                                    blnInComplete = True
                                End If
                            Next
                        End If
                    End Using
                End Using
            Catch ex As Exception
                GenerateException("AnyPartVerifyForBuild", ex)
            End Try

        End If
#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) blnInComplete:({1}) intIncompleteCount:({2})", blnRtnValue, blnInComplete, intIncompleteCount), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Public Function CheckAllSerialsComplete(ByVal strBuildNo As String, ByVal strStationNO As String, ByVal strPartNO As String,
                                            ByVal Optional intOperatorPosition As Integer = -1,
                                            ByVal Optional intPartVerifyMSeq As Integer = 0) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) strPartNO:({2}) intOperatorPosition:({3}) intPartVerifyMSeq:({4})",
                                                          strBuildNo, strStationNO, strPartNO, intOperatorPosition, intPartVerifyMSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = True

        Try
            Dim strSQL = "select serialstatus from AML_BuildPartsSerials where BuildNumber=" & strBuildNo & " and " &
                Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)

            If strPartNO.Trim <> "" Then
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("PartNumber", strPartNO)
            End If
            If intOperatorPosition <> -1 Then
                strSQL &= " and operatorposition=" & intOperatorPosition
            End If
            If intPartVerifyMSeq > 0 Then
                strSQL &= " and mseq=" & intPartVerifyMSeq
            End If

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each row As DataRow In table.Rows
                            Dim intTemp = Data.GetDataRowValue(row("serialstatus"), 0) 'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override

                            If intTemp = 1 Then
                                blnRtnValue = True
                            Else
                                blnRtnValue = False
                                Exit For
                            End If
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("CheckAllSerialsComplete", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue

    End Function

    Public Function BootBagStation(ByVal strStationNO As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0})", strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Select Case strStationNO.Trim
            Case "5510"
                blnRtnValue = True
        End Select


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Sub BuildFLMRecords(Optional ByVal intLineID As Integer = 0,
                    Optional ByVal strStationno As String = "",
                    Optional ByVal strArea As String = "",
                    Optional ByVal intShift As Integer = 1)        'create FLM record for the day (for each station)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) strStationno:({1}) strArea:({2}) intShift:({3})", intLineID, strStationno, strArea, intShift), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strErrLoc As String = "Start"

        Try
            'create FLM record for the day (for each station)
            'processing only one station.
            If strStationno.Trim = "" Then GoTo ExitThisSub
            If strArea.Trim = "" Then GoTo ExitThisSub

            strErrLoc = "GetFLMDate"
            Dim strBuildNO = GetFLMDate() 'FLM records are by day and station(not by ESN and Station)

            'get the FLM records for the machine/station
            strErrLoc = "GetMDMStationFLMRecords"
            Dim objRPDocs = GetMDMStationFLMRecords(strArea, strStationno, intShift)

            'now insert the FLM checks for every ESN
            'Record type: 0-Quality checks/Remedy Plan ,1- FLM record
            strErrLoc = "MassUpdateRemedyPlanRecords"
            MassUpdateRemedyPlanRecords(strBuildNO, objRPDocs, 1, strStationno, , intShift)
ExitThisSub:
        Catch ex As Exception
            GenerateException("BuildFLMRecords: " & strErrLoc, ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function BuildPartVerificationAndSerialList(ByVal strBuildNo As String, ByVal objBUILDBOM() As stVehBOM,
                                                       ByVal intLineID As Integer,
                                                       ByVal Optional strProcessOneStationNO As String = "",
                                                       ByVal Optional objSH() As Parts.stSubHeader = Nothing,
                                                       ByVal Optional blnRefreshBootBaggingOnly As Boolean = False,
                                                       ByVal Optional checkStationSignoff As Boolean = False) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) intLineID:({1}) strProcessOneStationNO:({2}) blnRefreshBootBaggingOnly:({3})",
                                                          strBuildNo, intLineID, strProcessOneStationNO, blnRefreshBootBaggingOnly), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Dim objSQLList As New List(Of String)
        Dim objLineStations(0) As stStations
        Dim intBootBagSeq = 0

        Dim strErrLoc As String = "Start"
        Try

            If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations
                'to avoid database call

                ReDim objLineStations(1)
                objLineStations(1).LineID = intLineID
                objLineStations(1).StationNO = strProcessOneStationNO

            Else
                'default option
                'get the list of stations
                objLineStations = GetLineStations(intLineID)

            End If


            strErrLoc = "GetPartVerifyList"
            Dim partVerifyTable = GetPartVerifyList()

            Dim strSQL = ""
            Dim intMSeq = 0
            If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations
                Using connection = New Connection(gStrConnectionString)
                    If BootBagStation(strProcessOneStationNO) Then
                        strSQL = "select max(mseq) from AML_BuildBootBag where " & Ec.GeneralFunc.GetQueryFieldCondition("buildnumber", strBuildNo)
                        intBootBagSeq = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                        intBootBagSeq += 1
                    End If

                    'processing for just one station, get the Max SEQ
                    strSQL = "select max(mseq) from AML_BuildPartsSerials where " & Ec.GeneralFunc.GetQueryFieldCondition("buildnumber", strBuildNo)
                    intMSeq = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                    intMSeq += 1
                End Using
            End If

            If blnRefreshBootBaggingOnly Then
                GoTo ProcessBootBagParts
            End If


            'delete existing data
            strSQL = "delete from AML_BuildPartsVerify where " & Ec.GeneralFunc.GetQueryFieldCondition("buildnumber", strBuildNo)
            If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strProcessOneStationNO)
            End If
            objSQLList.Add(strSQL)

            strSQL = "delete from AML_BuildPartsSerials where " & Ec.GeneralFunc.GetQueryFieldCondition("buildnumber", strBuildNo)
            If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strProcessOneStationNO)
            End If
            objSQLList.Add(strSQL)

            strErrLoc = "Loop Build BOM: " & UBound(objBUILDBOM)

            Dim processedPartStations As New List(Of String)
            For intK = 1 To UBound(objBUILDBOM)
                Dim strStationNO = objBUILDBOM(intK).StationNumber.Trim
                Dim strPartNO = Ec.GeneralFunc.QPTrim(objBUILDBOM(intK).PartNumber.Trim)
                Dim strPartDescX = objBUILDBOM(intK).PartDescription.Trim
                Dim strTracking = objBUILDBOM(intK).Tracking.Trim
                Dim strValidation = objBUILDBOM(intK).Validation.Trim
                Dim intSerialCount = Strings.ToInt32(objBUILDBOM(intK).Quantity)

                If objBUILDBOM(intK).Quantity = "0" Then GoTo SkipThisRecord 'duplicate record, ignore it. KM: 05/29/2014

                strErrLoc = "Loop " & intK & " of " & UBound(objBUILDBOM) & ", PartNO: " & strPartNO.Trim & ", Station: " & strStationNO.Trim

                If (objLineStations.Where(Function(x) Not String.IsNullOrEmpty(x.StationNO) AndAlso x.StationNO.ToLower().Trim() = strStationNO.ToLower().Trim()).Count() = 0) Then
                    GoTo SkipThisRecord
                End If

                If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations. This condition not required, already covered while getting the BUILDBOM
                    'no need for this, just in case.
                    If strStationNO.Trim.ToLower <> strProcessOneStationNO.Trim.ToLower Then
                        GoTo SkipThisRecord
                    End If
                End If


                strErrLoc = "CheckPartExistInVerifyList: Loop " & intK & " of " & UBound(objBUILDBOM) & ", PartNO: " & strPartNO.Trim & ", Station: " & strStationNO.Trim
                'verify the build part number found in the verification part list
                Dim partVerify = partVerifyTable.Select(String.Format("partnumber='{0}'", strPartNO))
                If (partVerify.Count() = 0) Then
                    GoTo SkipThisRecord
                End If

                strErrLoc = "CheckSubHeaderReferenceNumberMatch: Loop " & intK & " of " & UBound(objBUILDBOM) & ", PartNO: " & strPartNO.Trim & ", Station: " & strStationNO.Trim

                Dim strLineNumber = objBUILDBOM(intK).LineNumber.ToString.Trim
                strLineNumber = FormatSubHeaderReferenceNumber(strLineNumber)

                Dim intOperatorPosition = 1
                If Not CheckSubHeaderReferenceNumberMatch(objSH, strStationNO, strLineNumber, intOperatorPosition) Then  'get the operator position by reference
                    'the part# is not applicable for the specific build number , subheader reference number doesn't match
                    'but check to see if station found, something about Dan's email on Jun 4, 2014
                    If (objLineStations.Where(Function(x) Not String.IsNullOrEmpty(x.StationNO) AndAlso x.StationNO.ToLower().Trim() = strStationNO.ToLower().Trim()).Count = 0) Then
                        GoTo SkipThisRecord
                    End If
                End If

                Dim partStation = strPartNO.Trim & "-" & strStationNO.Trim & "-" & intOperatorPosition
                If (processedPartStations.Contains(partStation)) Then
                    'duplicate part already added, skip this one
                    GoTo SkipThisRecord
                End If

                'Tracking: ST - Serial Tracking, VT-Verification Tracking, BT- Batch Tracking, NT - No Tracking
                Select Case strTracking.Trim.ToUpper
                    Case "ST", "VT", "BT", "PC", "BC", "VC", "UC"
                    Case Else
                        GoTo SkipThisRecord
                End Select

                processedPartStations.Add(partStation)

                Dim strSupplierPartNO = Ec.GeneralFunc.QPTrim(Trim(Data.GetDataRowValue(partVerify(0)("supplier_partno"), "")))

                'intPartPos += 1
                intMSeq += 1
                strSQL = "insert into AML_BuildPartsVerify(BuildNumber,StationNO,PartNumber,MSEQ,LastUpdated,DateCreated,VerifyStatus," &
                    "Verify_Station,spare1,spare2,operatorposition,partdescx) values(  " &
                    strBuildNo.Trim & ",'" & strStationNO.Trim & "','" & Strings.ReplaceSingleQuote(strPartNO) & "'," &
                    intMSeq & ", getdate(),getdate(),-1," &
                    "'Y','" & strSupplierPartNO.Trim & "','" & strLineNumber & "'," & intOperatorPosition & ",'" &
                    Strings.ReplaceSingleQuote(strPartDescX) & "')"

                objSQLList.Add(strSQL)

                strErrLoc = "Add Serial Count DATA : Loop " & intK & " of " & UBound(objBUILDBOM) & ", PartNO: " & strPartNO.Trim & ", Station: " & strStationNO.Trim

                'Tracking: ST - Serial Tracking, VT-Verification Tracking, BT- Batch Tracking, NT - No Tracking
                Select Case strTracking.Trim.ToUpper
                    Case "ST", "BT", "UC", "PC", "VC", "BC"
                        'intSerialCount = 3
                        For intYY = 1 To intSerialCount

                            ''AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override
                            strSQL = "insert into AML_BuildPartsSerials(BuildNumber,StationNO,PartNumber,MSEQ,SERIALSEQ,LastUpdated,DateCreated,Serial_Req,Serial_Format,Supplier_PartNO,AML_PartNO," &
                                    " SCAN_SERIALNO,SerialStatus,spare2,operatorposition,tracking) values(  " &
                                    strBuildNo.Trim & ",'" & strStationNO.Trim & "','" & Strings.ReplaceSingleQuote(strPartNO) & "'," &
                                    intMSeq & "," & intYY & ", getdate(),getdate()," &
                                    "1,'" & strValidation.Trim & "','" & strSupplierPartNO.Trim & "',' '" &
                                    ",' ',-1,'" & strLineNumber.Trim & "'," & intOperatorPosition & ",'" & strTracking.Trim.ToUpper & "')"

                            objSQLList.Add(strSQL)
                        Next
                End Select

SkipThisRecord:

            Next


ProcessBootBagParts:
            strSQL = "delete from AML_BuildBootBag where " & Ec.GeneralFunc.GetQueryFieldCondition("buildnumber", strBuildNo)
            If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strProcessOneStationNO)
            End If
            objSQLList.Add(strSQL)

            For intK = 1 To UBound(objBUILDBOM)
                Dim strStationNO = objBUILDBOM(intK).StationNumber.Trim
                If Not BootBagStation(strStationNO) Then
                    GoTo SkipThisRecord2
                End If

                Dim intZZ = 0
                Select Case objBUILDBOM(intK).KitNO.Trim.ToUpper
                    Case "BB01"
                    Case "BB02"
                        intZZ = 1
                    Case Else
                        GoTo SkipThisRecord2

                End Select

                'Select objBUILDBOM
                Dim strPartNO = Ec.GeneralFunc.QPTrim(objBUILDBOM(intK).PartNumber.Trim)
                Dim strPartDescX = objBUILDBOM(intK).PartDescription.Trim

                Dim sngTemp = Strings.ToSingle(objBUILDBOM(intK).Quantity)
                Dim quantity = ""
                If sngTemp > 0 Then
                    quantity = sngTemp.ToString
                Else
                    quantity = objBUILDBOM(intK).Quantity
                End If

                intBootBagSeq += 1
                strSQL = "insert into AML_BuildBootBag (buildnumber,StationNO, PartNumber, mseq, PartDescX, InsideBox, datecreated, LastUpdated,verifystatus,kitno,quantity,operatorposition)" &
                    " values (" & strBuildNo & ",'" & strStationNO.Trim & "','" & strPartNO.ReplaceSingleQuote() & "'," & intBootBagSeq &
                    ", '" & strPartDescX.ReplaceSingleQuote() & "'," & intZZ & ",getdate(),getdate(),-1,'" &
                    objBUILDBOM(intK).KitNO.ReplaceSingleQuote() & "','" & quantity & "', 1)"
                objSQLList.Add(strSQL)
SkipThisRecord2:
            Next




            If objSQLList.Count > 0 Then
                Ec.IO.RunSQLInArray(objSQLList.ToArray())
            End If

            blnRtnValue = True
        Catch ex As Exception
            GenerateException("BuildPartVerificationAndSerialList: " & strErrLoc, ex)
        Finally
            objLineStations = Nothing
            objSQLList = Nothing
        End Try



#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Sub BuildQualityCheckRecords(ByVal strBuildNO As String, ByVal strStation As String,
                        ByVal intLineID As Integer, ByVal strBuildNOModel As String, ByVal strArea As String,
                        Optional ByRef strLogData As String = "",
                        Optional ByVal blnCreateQCFromWIPScreen As Boolean = False,
                        Optional ByRef blnResult As Boolean = False,
                        Optional ByRef objSH() As Parts.stSubHeader = Nothing,
                        ByVal Optional blnGetSubHeaderOBJECTOnly As Boolean = False)



#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStation:({1}) intLineID:({2}) strBuildNOModel:({3}) strArea:({4}) blnCreateQCFromWIPScreen:({5}) blnGetSubHeaderOBJECTOnly:({6})",
                                                          strBuildNO, strStation, intLineID, strBuildNOModel, strArea, blnCreateQCFromWIPScreen, blnGetSubHeaderOBJECTOnly), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        '** Version 7 Changes **  - PENDING

        'blnGetSubHeaderOBJECTOnly used to create the Part Serial records for the manual stations. Default is FALSE
        '
        '---------------------------------------------------------------------------------------------------------------
        'The QC records can be created in two ways.
        '1. WIP section creates the QC for all stations in certain lines where PLC is integrated (DB9 & DBS, Vantage)
        '2. All other lines create the QC on the fly, when the operator scans the Build#
        '---------------------------------------------------------------------------------------------------------------

        'strBuildNOModel - param not used anywhere  ** obsolete **
        'blnProcessAllStations is true, when calling from amlwipedit screen, adding new build# to WIP

        Dim strSQL As String = ""
        Dim intK As Int32 = 0, strSQLCond3 As String = "", strSQLCond4 As String = ""
        Dim strSubHeader As String = "", strFileName As String = ""
        Dim strRectype As String = "", strOPNO As String = ""
        Dim intPrevID As Int32 = 0, strPrevRectype As String = "", intPrevSeq As Int16 = 0, strPrevOPNO As String = ""
        Dim blnIDChanged As Boolean = False, intMDMDocID As Int32 = 0, intSHID As Int32 = 0
        Dim objRPDocs(0) As stRemedyDocs, intYY As Int32 = 0
        Dim objRPDocsTemp(0) As stRemedyDocs
        Dim strTemp As String = "", intI As Int32 = 0
        Dim intTTKey As Int32 = 0, intMDMDocRectype As Int16 = 0, intMDMDocSeq As Int16 = 0
        Dim strErrLoc As String = "Start", strEngineNO As String = ""
        Dim strTemp2 As String = "", intZZ As Int32 = 0
        Dim intFLMFrequencyID As Byte = 0
        Dim objSearchResult(0) As Parts.stPartSearchResult
        Dim strTempLog As String = "" ', strOperationSQLCond As String = ""
        Dim strOpListSQLCond As String = ""
        Dim strMMTable As String = "ophmm"
        Dim strPlatenID As String = ""
        ''** Version 7 Changes **  - TEST

        'Dim strCond3 As String = ""
        Try

            ReDim objSH(0)


            If DBConfig.Version7 Then
                strMMTable = "subhdr"
            End If
            strLogData = "" : blnResult = False
            strStation = Trim(strStation)
            If strStation = "" Then
                If Not blnCreateQCFromWIPScreen Then
                    'not covered/not applicable (06/25/2012) at this time, just exit out.
                    'we process only one station at a time.
                    GoTo ExitThisSub
                End If
                strOpListSQLCond = GetAutoRefreshingStationsListForSQL(intLineID)  'operations conditions
            End If

            Dim intLinkedModel = 0
            Dim intBuildNoModelNumber = GetBuildNoModelNumber(strBuildNO, True, intLinkedModel)
            '-----------------------------------------------------------------------------------------------
            'get the list of parts for the Area, Model (may contain more than one operation ( OP#: 5050,5050-1,5050-2, 5050-3)

            'READ ME: strStation is not used and blnAvoidOPNO must be true - Get the Plan
            objSearchResult = GetEASEPartNo(strArea, intBuildNoModelNumber, strStation, 0, True, strSQL, intLinkedModel)
            '-----------------------------------------------------------------------------------------------
            strLogData = "OPTIONS SQL: " & strSQL & "<br>"

            If UBound(objSearchResult) = 0 Then
                strLogData &= "**** NO OPTIONS FOUND **** "
                GoTo DeleteExistingRecordsAndExitOut
            End If
            Dim strPartNO = objSearchResult(1).PartNO.Trim

            strLogData &= "<br>Total number of Options Found: " & UBound(objSearchResult) & "<br>"

            'filter the id,rectype,seq,opno to get the docs
            strSQLCond4 = "" : strTemp2 = ""

            '*********** READ ME ****************
            'AML uses only plan for every model, so skip the remaining plans
            intZZ = 1
            strTemp2 = " (" & strMMTable & ".id=" & objSearchResult(intZZ).ID & " and " & strMMTable & ".rectype='" & objSearchResult(intZZ).Rectype.Trim & "' and " & strMMTable & ".seq=" & objSearchResult(intZZ).Seq & ")" '& _
            '" and " & Trim(DBConfig.QueryFunctions.Upper) & "(ophmm.opno) = '" & objSearchResult(intZZ).OPNO.ToUpper.Trim & "')"
            If strSQLCond4.Trim <> "" Then
                strSQLCond4 &= " or "
            End If
            strSQLCond4 &= strTemp2
            strSQLCond4 = " and (" & strSQLCond4 & ") "

            'EW-855: VH 5 Body Shop and MA1 station refresh
            If intLineID = BodyShopLineId Then
                strPlatenID = GetPlatenIDForBuild(strBuildNO, strStation)
            End If

            'Get the list of subheaders based on specific/Generic option for the Build#/Model
            strTempLog = ""
            objSH = GetSubheaderList(objSearchResult(intZZ).ID, objSearchResult(intZZ).Rectype, objSearchResult(intZZ).Seq, strStation, strBuildNO, strStation, True, , True, , strTempLog, , blnCreateQCFromWIPScreen, strOpListSQLCond, , , , strPlatenID)

            If blnGetSubHeaderOBJECTOnly Then
                'Need subheders only.
                GoTo ExitThisSub
            End If

            Dim intID = objSearchResult(intZZ).ID : strRectype = objSearchResult(intZZ).Rectype
            Dim intSeq = objSearchResult(intZZ).Seq

            strLogData &= "<br><b>SUB-HEADERS LOG </b><BR>"
            strLogData &= "<b>SH Count: " & UBound(objSH) & "</b><br>"
            strLogData &= "<br>" & strTempLog.Trim & "<br>"
            If UBound(objSH) = 0 Then
                ' no matching subheaders found, just exit out
                GoTo DeleteExistingRecordsAndExitOut
            End If
            Dim strSQLCond2 = ""
            If blnCreateQCFromWIPScreen Then
                strSQLCond2 = " and " & strOpListSQLCond 'GetAutoRefreshingStationsListForSQL()
                'strSQLCond2 = " and (" & Trim(DBConfig.QueryFunctions.Upper) & "(ophmm.opno) in (" & strSQLCond2.Trim & ")) "
            Else            'processing single station

                Dim intMaxOperators = GetOperatorCountForStation(intLineID, strStation)
                strTemp2 = "'" & UCase(Trim(strStation)) & "'"
                For intZZ = 1 To intMaxOperators
                    strTemp2 &= ",'" & UCase(Trim(strStation)) & "-" & Trim(intZZ.ToString) & "'"
                Next
                strSQLCond2 &= strTemp2
                strSQLCond2 = " and (" & Trim(DBConfig.QueryFunctions.Upper) & "(" & strMMTable & ".opno) in (" & strSQLCond2.Trim & ")) "
            End If

            strErrLoc = "Get Docs List"

            'If Easeclass.DBConfig.Version7 Then

            strSQL = "select shmm.shid, shmm.mdmdocid,shmm.ttkey,shmm.completettkey,shmm.mdmdocseq,shmm.mseq " &
                " from shmm where shmm.mdmdocid > 0 "

            strSQLCond3 = ""  'get the list of subheaders
            For intZZ = 1 To UBound(objSH)
                If strSQLCond3.Trim <> "" Then strSQLCond3 &= ", "
                If DBConfig.SharedSubHeader AndAlso objSH(intZZ).SharedSHID > 0 Then
                    strSQLCond3 &= objSH(intZZ).SharedSHID
                Else
                    strSQLCond3 &= objSH(intZZ).SubHdrID
                End If
            Next
            strSQL &= " and shid in (" & strSQLCond3 & ") order by shid,mseq"

            strLogData &= "<br>Operation Ref Docs SQL: " & strSQL.Trim & "<br>"

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    strErrLoc = "Loop Ref Docs"
                    If (Not IsNothing(table)) Then
                        For Each row As DataRow In table.Rows
                            '** Version 7 Changes **  - TEST
                            'If Not CheckSubHeaderExistFromMemory(objSH, strSubHeader, strOPNO) Then GoTo skipthisrecord

                            intMDMDocID = Data.GetDataRowValue(row("mdmdocid"), 0)
                            If intMDMDocID = 0 Then GoTo SkipThisRecord 'not a shared document, skip it, may a subheader record again.

                            intSHID = Data.GetDataRowValue(row("shid"), 0)

                            intYY += 1
                            ReDim Preserve objRPDocs(intYY)

                            objRPDocs(intYY).PartID = intID
                            objRPDocs(intYY).RecType = strRectype.Trim
                            objRPDocs(intYY).Seq = intSeq
                            objRPDocs(intYY).OPNO = ""
                            objRPDocs(intYY).Shift = 0

                            objRPDocs(intYY).FileName = "" 'strSubHeader.Trim 'strFileName.Trim
                            objRPDocs(intYY).Marked = -2
                            objRPDocs(intYY).MDMDocRectype = 0
                            objRPDocs(intYY).PlanType = 0
                            objRPDocs(intYY).RemedyPlanDesc = ""
                            objRPDocs(intYY).SHID = intSHID

                            objRPDocs(intYY).CompleteTTKey = Data.GetDataRowValue(row("completettkey"), 0)
                            objRPDocs(intYY).MDMDocID = Data.GetDataRowValue(row("mdmdocid"), 0)
                            objRPDocs(intYY).MDMDocSeq = Data.GetDataRowValue(row("mdmdocseq"), 0)
                            objRPDocs(intYY).TTKey = Data.GetDataRowValue(row("ttkey"), 0)

                            objRPDocs(intYY).FrequencyType = 0
                            strLogData &= "<br>ID: " & intID.ToString & ", OPNO: " & strOPNO.Trim & ", SH: " & strSubHeader.Trim & ", DocID:" & objRPDocs(intYY).MDMDocID & ", TTKey: " & objRPDocs(intYY).TTKey
SkipThisRecord:
                        Next
                    End If
                End Using

                If intYY = 0 Then GoTo DeleteExistingRecordsAndExitOut 'no records found, clear existing records

                'got the list of operation documents (for the subheader list), now sort it
                intYY = 0 : ReDim objRPDocsTemp(0)
                For intZZ = 1 To UBound(objSH)
                    intSHID = objSH(intZZ).SubHdrID
                    If DBConfig.SharedSubHeader AndAlso objSH(intZZ).SharedSHID > 0 Then
                        intSHID = objSH(intZZ).SharedSHID
                    End If
                    For intI = 1 To UBound(objRPDocs)
                        If objRPDocs(intI).SHID = intSHID Then
                            'may have more than one document
                            intYY += 1
                            ReDim Preserve objRPDocsTemp(intYY)
                            objRPDocsTemp(intYY) = objRPDocs(intI)
                            objRPDocsTemp(intYY).PartID = intID
                            objRPDocsTemp(intYY).RecType = strRectype
                            objRPDocsTemp(intYY).Seq = intSeq
                            objRPDocsTemp(intYY).OPNO = objSH(intZZ).OPNO

                            objRPDocsTemp(intYY).SHID = objSH(intZZ).SubHdrID
                            objRPDocsTemp(intYY).FileName = objSH(intZZ).SubHdrDesc.Trim
                            objRPDocsTemp(intYY).OPNO = objSH(intZZ).OPNO
                            objRPDocsTemp(intYY).PartNO = strPartNO.Trim

                        End If

                    Next
                Next
                ReDim objRPDocs(0)
                objRPDocs = objRPDocsTemp
                ReDim objRPDocsTemp(0)

                'now we got all the shared documents, now get the remedy plan documents

                strErrLoc = "Loop RPDOCS"
                strTemp = ""
                For intI = 1 To UBound(objRPDocs)

                    If objRPDocs(intI).Marked = -2 Then
                        If Trim(strTemp) <> "" Then strTemp &= " or "
                        strTemp &= "( pdmpartmm.docid = " & objRPDocs(intI).MDMDocID &
                            " and pdmpartmm.docrectype='" & objRPDocs(intI).MDMDocRectype & "'" &
                            " and pdmpartmm.docseq=" & objRPDocs(intI).MDMDocSeq & ")"
                        '" and pdmpartmm.ttkey=" & objRPDocs(intI).TTKey & ")"

                    End If

                Next

                'docgroupid=1, remedy plan docs
                strSQL = "select docdesc,docid,ttkey,docseq,docrectype,docdesc from pdmpartmm " &
                    " where docgroupid=1 and doctype=3 and (" & strTemp & ") order by docid"

                strLogData &= "<br>MDM SQL: " & strSQL.Trim

                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    strErrLoc = "Loop Quality Check Data"
                    If (Not IsNothing(table)) Then
                        For Each row As DataRow In table.Rows
                            intMDMDocID = Data.GetDataRowValue(row("docid"), 0)
                            intTTKey = Data.GetDataRowValue(row("ttkey"), 0)
                            intMDMDocRectype = Data.GetDataRowValue(Of Int16)(row("docrectype"), 0)
                            intMDMDocSeq = Data.GetDataRowValue(Of Int16)(row("docseq"), 0)
                            strTemp = Data.GetDataRowValue(row("docdesc"), "")

                            For intI = 1 To UBound(objRPDocs)
                                strErrLoc = "Loop Quality Check Data : Validation"
                                If objRPDocs(intI).Marked = -2 And objRPDocs(intI).MDMDocID = intMDMDocID _
                            And objRPDocs(intI).TTKey = intTTKey Then

                                    'matching document found.
                                    objRPDocs(intI).Marked = 1
                                    objRPDocs(intI).MDMDocRectype = intMDMDocRectype
                                    objRPDocs(intI).MDMDocSeq = intMDMDocSeq
                                    objRPDocs(intI).RemedyPlanDesc = Trim(strTemp)
                                    'Exit For       'may have more than one document, don't exit for
                                End If
                            Next intI
                            strErrLoc = "Loop Quality Check Data 2"
                        Next
                    End If
                End Using
            End Using

            'discard the non remedy plan docs
            '-----------------------------------------------------------------------
            'Don't make any changes to this block, doing so may cause isses in
            'storing the header data
            '-----------------------------------------------------------------------
            strErrLoc = "Clean Object"
            ReDim objRPDocsTemp(0)
            intK = 0
            For intI = 1 To UBound(objRPDocs)
                strTemp = Trim(objRPDocs(intI).FileName)
                'If Left(strTemp, 3) = "***" Or objRPDocs(intI).Marked = 1 Then
                If objRPDocs(intI).Marked = 1 Then
                    intK += 1
                    ReDim Preserve objRPDocsTemp(intK)
                    objRPDocsTemp(intK) = objRPDocs(intI)
                End If

SkipThisRecord2:
            Next

DeleteExistingRecordsAndExitOut:
            strErrLoc = "Refresh Object"
            'refresh the object
            objRPDocs = objRPDocsTemp

            strErrLoc = "Log Data"
            'Record type: 0-Quality checks/Remedy Plan ,1- FLM record
            strLogData &= "<br><b>Total QC Records: " & UBound(objRPDocs) & "</b>"
            For intK = 1 To UBound(objRPDocs)
                strLogData &= "<br>" & objRPDocs(intK).OPNO.Trim & "-" & objRPDocs(intK).RemedyPlanDesc.Trim
            Next
            strErrLoc = "MassUpdateRemedyPlanRecords"
            MassUpdateRemedyPlanRecords(strBuildNO, objRPDocs, 0, strStation, blnCreateQCFromWIPScreen)

            strErrLoc = "MassUpdateRemedyPlanRecords-complet"
            If UBound(objRPDocs) > 0 Then
                blnResult = True
            End If


ExitThisSub:

        Catch ex As Exception
            GenerateException("BuildQualityCheckRecords: " & strErrLoc & "<br>" & strSQL, ex)
        Finally
            objRPDocs = Nothing
            objRPDocsTemp = Nothing
            objSearchResult = Nothing
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit strLogData:({0}) blnResult:({1}) objSH.Count:({2})", strLogData, blnResult, UBound(objSH)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function CheckBootBagPartExist2(ByVal strBuildNO As String, ByVal strStationNO As String, ByVal strPartNO As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2})", strBuildNO, strStationNO, strPartNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'AML_BuildBootBag: VerifyStatus: -1 Incomplete, 1-Complete
        Dim strSQL As String = "select count(1) from AML_BuildBootBag where " &
            " buildnumber=" & strBuildNO & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) &
            " and verifystatus <> 1 and " & Ec.GeneralFunc.GetQueryFieldCondition("partnumber", strPartNO)
        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0) > 0
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue

    End Function

    Public Function CheckBootBagParts(strBuildNO As String, strStationNO As String, Optional ByRef blnBootBagPartsStatusIncomplete As Boolean = True) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1})", strBuildNO, strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'returns, if bootbag exists or not.
        blnBootBagPartsStatusIncomplete = False
        Dim blnRtnValue As Boolean = False
        Try
            Dim strSQL = String.Format("select verifystatus from AML_BuildBootBag where buildnumber={0} and {1}", strBuildNO, Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO))

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each row As DataRow In table.Rows
                            blnRtnValue = True
                            Dim intStatus = Data.GetDataRowValue(row("verifystatus"), -1)

                            Select Case intStatus
                                Case -1
                                    'AML_BuildBootBag: VerifyStatus: -1 Incomplete, 1-Complete
                                    blnBootBagPartsStatusIncomplete = True
                                    Exit For  'no need to loop again
                            End Select
                        Next
                    End If
                End Using
            End Using

        Catch ex As Exception
            GenerateException("CheckBootBagParts", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) blnBootBagPartsStatusIncomplete:({1})", blnRtnValue, blnBootBagPartsStatusIncomplete), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue

    End Function

    Public Function CheckBuildNumberExists(ByVal strBuildNO As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0})", strBuildNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        If (Not String.IsNullOrEmpty(strBuildNO.Trim)) Then
            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.ExecuteScalar(Of Integer)(String.Format("select COUNT(*) from aml_vehicleorders WITH (NOLOCK) where buildnumber={0}", strBuildNO.Trim), 0)
                blnRtnValue = result > 0
            End Using
        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function


    Public Function CheckFLMStatusAndAndonTimer(ByVal strCurrentDate As String, ByVal strStationNO As String, ByVal intShift As Int16,
                                                ByVal intLineID As Integer, ByRef strLastUpdated As String,
                                                ByRef intTimerSeconds As Int16, ByVal blnIndependentStations As Boolean) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strCurrentDate:({0}) strStationNO:({1}) intShift:({2}) intLineID:({3}) blnIndependentStations:({4})",
                                                          strCurrentDate, strStationNO, intShift, intLineID, blnIndependentStations), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim blnRtnValue As Boolean = True

        Try
            intTimerSeconds = 5 : strLastUpdated = ""
            'PDMRPDATA->Record type: 0-Quality checks/Remedy Plan ,1- FLM record
            strSQL = "select count(1) from pdmrpdata WITH (NOLOCK) where " &
                    Ec.GeneralFunc.GetQueryFieldCondition("engineno", strCurrentDate) &
                    " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) &
                    " and resultflag=-1 and recordtype=1"
            If intShift > 0 Then
                strSQL &= " and shift=" & intShift
            End If

            Using connection = New Connection(gStrConnectionString)
                Dim intTemp = connection.ExecuteScalar(Of Integer)(strSQL, 0)

                If intTemp > 0 Then
                    blnRtnValue = False
                End If

                strSQL = "select timeseconds from AML_Andon where lineid=" & intLineID.ToString
                If blnIndependentStations Then
                    strSQL = "select timeseconds from AML_ANDON_Stations where lineid=" & intLineID.ToString & " and " & Ec.GeneralFunc.GetQueryFieldCondition("STATIONNO", strStationNO)
                End If

                ' TODO(crhodes)
                ' This is probably where we need to get the time from the AML_LineStatus_BodyShop table for Line 50

                intTimerSeconds = connection.ExecuteScalar(Of Int16)(strSQL, 0)
                If intTimerSeconds <= 0 Then intTimerSeconds = 10 'default to 10 seconds
            End Using

        Catch ex As Exception
            GenerateException("CheckFLMStatusAndAndonTimer", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) strLastUpdated:({1}) intTimerSeconds:({2})", blnRtnValue, strLastUpdated, intTimerSeconds), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Function CheckForAutoRefreshingLine(ByVal intLineID As Integer) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0})", intLineID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Any changes in this should be considered for  
        '   CheckForAutoRefreshingLine()
        '   CheckForAutoRefreshingStations. 

        Dim blnRtnValue As Boolean = False

        If intLineID = 3 Or intLineID = 8 Or intLineID = 50 Then
            blnRtnValue = True
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Public Function CheckForAutoRefreshingStations(ByVal strStation As String,
                Optional ByVal blnExcludeFirstStation As Boolean = False,
                Optional ByRef blnFirstStationInLine As Boolean = False) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStation:({0}) blnExcludeFirstStation:({1})", strStation, blnExcludeFirstStation), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Any changes in this should be considered for  
        '   CheckForAutoRefreshingLine()
        '   CheckForAutoRefreshingStations. 

        Dim intStationNO As Int32 = 0
        Dim blnRtnValue As Boolean = False  ' Default is no auto refresh

        Try
            blnFirstStationInLine = False
            intStationNO = Strings.ToInt32(strStation)

            ' 3000 - 3350 is T&F 1
            ' 4000 - 4400 is T&F 2
            '
            ' NB.  Not all stations in T&F lines auto-refresh, e.g. 3500 - Door, 3600 - PowerTrain, 3700 - Bumper
            ' Revisit with AML to see if can expose this through configuration to avoid hard coding.

            'automation is available oinly at these stations

            Select Case intStationNO
                Case 3000 To 3350, 4000 To 4400
                    blnRtnValue = True
                Case Else
            End Select

            ' First station in line (Cake Stands) work differently in frmPoll.

            If intStationNO = 3000 Or intStationNO = 4000 Then
                blnFirstStationInLine = True
            End If

            If blnExcludeFirstStation And blnRtnValue Then
                If intStationNO = 3000 Or intStationNO = 4000 Then
                    blnRtnValue = False
                End If
            End If
        Catch ex As Exception
            GenerateException("CheckForAutoRefreshingStations", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) FirstStationInLine:({1})", blnRtnValue, blnFirstStationInLine), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue

    End Function

    Public Function CheckForFLMRecords(ByVal strCurrentDate As String, ByVal strStationNO As String,
                                       Optional intShift As Integer = 0,
                                       Optional blnCheckFLMStatus As Boolean = False, Optional ByRef blnFLMStatus As Boolean = False) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strCurrentDate:({0}) strStationNO:({1}) intShift:({2}) blnCheckFLMStatus:({3})", strCurrentDate, strStationNO, intShift, blnCheckFLMStatus), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Dim currentDateCondition = Ec.GeneralFunc.GetQueryFieldCondition("engineno", strCurrentDate)
        Dim stationNoCondition = Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)
        Dim strSQL As String = String.Format("SELECT COUNT(1) FROM PDMRPDATA WITH (NOLOCK) WHERE {0} AND {1} AND recordtype = 1", currentDateCondition, stationNoCondition)

        If intShift <> 0 Then
            Select Case intShift
                Case 1
                    strSQL &= " AND (shift IS NULL OR shift=0 OR shift=1) "
                Case Else
                    strSQL &= " AND shift=" & intShift
            End Select
        End If

        Try
            Using connection = New Connection(gStrConnectionString)
                Dim intCounter = connection.ExecuteScalar(strSQL, 0)
                If intCounter > 0 Then
                    blnRtnValue = True
                End If

                If blnCheckFLMStatus Then
                    intCounter = 0
                    blnFLMStatus = True
                    strSQL = String.Format("SELECT COUNT(1) FROM PDMRPDATA WITH (NOLOCK) WHERE {0} AND {1} AND resultflag = -1 AND recordtype = 1", currentDateCondition, stationNoCondition)
                    If intShift > 0 Then
                        strSQL &= " AND shift=" & intShift
                    End If

                    intCounter = connection.ExecuteScalar(strSQL, 0)
                    If intCounter > 0 Then
                        blnFLMStatus = False
                    End If
                End If
            End Using
        Catch ex As Exception
            GenerateException("CheckForFLMRecords: " & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) FLMStatus:({1})", blnRtnValue, blnFLMStatus), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Public Function CheckForQualityCheckRecords(ByVal strBuildNo As String, ByVal strStationNO As String,
            ByVal strOptionNO As String, ByVal strOPNO As String, ByVal strSubheader As String,
            Optional ByRef intRecCount As Int16 = 0) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) strOptionNO:({2}) strOPNO:({3}) strSubheader:({4})", strBuildNo, strStationNO, strOptionNO, strOPNO, strSubheader), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        intRecCount = 0
        If DBConfig.Version7 Then
        Else
            If strSubheader.Trim <> "" AndAlso Left(Trim(strSubheader), 3) <> "***" Then
                strSubheader = "***" & strSubheader.Trim
            End If
        End If


        'Record type: 0-Quality checks/Remedy Plan ,1- FLM record
        Dim strSQL = "select count(1) from pdmrpdata WITH (NOLOCK) where " & Ec.GeneralFunc.GetQueryFieldCondition("engineno", strBuildNo) &
                        " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)
        If strOptionNO.Trim <> "" Then
            strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("optionno", strOptionNO)
        End If
        If strOPNO.Trim <> "" Then
            strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("opno", strOPNO)
        End If
        strSQL &= " and recordtype=0 "        'and resultflag = -1

        If strSubheader.Trim <> "" Then
            strSQL &= " and " & Trim(DBConfig.QueryFunctions.Upper) & "(subheader)='" & UCase(Strings.ReplaceSingleQuote(Trim(strSubheader))) & "'"
        End If

        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            intRecCount = connection.ExecuteScalar(Of Int16)(strSQL, 0)
            If intRecCount > 0 Then
                returnValue = True
            End If
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) intRecCount:({1})", returnValue, intRecCount), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function CheckLastStationNumber(ByVal strStationNO As String, ByVal Optional intLineID As Integer = 0) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0}) intLineID:({1})", strStationNO, intLineID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False

        ' Apply only to T&F_LINE_1 (3) and T&F_LINE_2 (8)
        If intLineID = 3 And Trim(strStationNO) = "3350" Then
            blnRtnValue = True
        ElseIf intLineID = 8 And Trim(strStationNO) = "4400" Then
            blnRtnValue = True
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue

    End Function

    Public Function CheckOverrideAccessRightsForUser(strUserID As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserID:({0})", strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = False
        Dim strSQL As String = "select COUNT(1) from euserrights where " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID) & " and misc1='1'"
        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.ExecuteScalar(Of Integer)(strSQL, 0)
            returnValue = (result > 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit retVal:({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue

    End Function

    Public Function CheckPartSerialExist(strBuildNO As String, strStationNO As String, intPartVerifyMSeq As Integer) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) intPartVerifyMSeq:({2})", strBuildNO, strStationNO, intPartVerifyMSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select COUNT(*) from AML_BuildPartsSerials where BuildNumber=" & strBuildNO &
                " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) &
                " and mseq=" & intPartVerifyMSeq

        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            returnValue = (connection.ExecuteScalar(Of Integer)(strSQL, 0) > 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function CheckPSSDocumentReadStatus(objSH() As Parts.stSubHeader, strStationNO As String, strUSERID As String, Optional ByRef blnDocsFound As Boolean = False) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objSH.Length:({0}) strStationNO:({1}) strUSERID:({2})",
                                                          objSH.Length, strStationNO, strUSERID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objPSS(0) As stPSSDocuments, blnStatus As Boolean = True
        Dim strErrorLoc As String = "start"

        Try
            blnDocsFound = False
            strErrorLoc = "GetMDMProcessDocumentsList"
            'get the list of documents for the subheaders (for the station)
            objPSS = GetMDMProcessDocumentsList(objSH, strStationNO, strUSERID) 'last 90 days only
            If UBound(objPSS) = 0 Then GoTo ExitThisSub

            strErrorLoc = "Process"
            blnDocsFound = True
            blnStatus = True
            For intK = 1 To UBound(objPSS)
                strErrorLoc = "Process: " & intK
                If objPSS(intK).PassStatus <> 1 Then
                    blnStatus = False
                    Exit For
                End If
            Next

ExitThisSub:
        Catch ex As Exception
            GenerateException("CheckPSSDocumentReadStatus: " & strErrorLoc, ex)
        Finally
            objPSS = Nothing
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) blnDocsFound:({1})", blnStatus, blnDocsFound), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnStatus
    End Function

    Public Function CheckSafetyDocCompleted(ByVal objSafetyCheck() As stSafetyCheck, ByVal lngDocID As Long,
                                            ByVal intDocSeq As Int16, ByRef intPointer As Integer,
                                            intDocRev As Int16) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objSafetyCheck.Length:({0}) lngDocID:({1}) intDocSeq:({2}) intDocRev:({3})", objSafetyCheck.Length, lngDocID, intDocSeq, intDocRev), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Try
            intPointer = 0
            For intI = 1 To UBound(objSafetyCheck)
                If objSafetyCheck(intI).DOCID = lngDocID And objSafetyCheck(intI).DocSeq = intDocSeq And Trim(objSafetyCheck(intI).DateRead) <> "" And objSafetyCheck(intI).RevNum = intDocRev Then
                    intPointer = intI
                    If objSafetyCheck(intI).PassStatus = 0 Or objSafetyCheck(intI).PassStatus = 1 Then
                        blnRtnValue = True
                        Exit For
                    End If
                End If
            Next intI
        Catch ex As Exception
            GenerateException("CheckSafetyDocCompleted", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) intPointer:({1})", blnRtnValue, intPointer), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    <Obsolete("Execute sp_checkpartverify, except for aml3 seems like it does not have a build number on it.")>
    Public Function CheckScannedPart(ByRef strPartNO As String, Optional ByRef strSerialNO As String = "") As String


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strPartNo:({0})", strPartNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        strPartNO = Trim(strPartNO)

        Dim strErrorMsg As String = "", intPos = 0
        Dim strTemp2 As String = ""

        strSerialNO = ""
        Try

            If strPartNO.Trim = "" Then
                strErrorMsg = "PART " & strPartNO.Trim & " INCORRECT FOR BUILD."
                GoTo ExitThisSub
            End If

            If strPartNO.Trim.Contains("*") Then
                strErrorMsg = "PART " & strPartNO.Trim & " INCORRECT FOR BUILD."
                GoTo ExitThisSub
            End If

            'Barcode Errors
            'single barcode contains both the part number and the serial number, separated by a �/�.
            'EX: ED23-6007-AA/A3A3A333333
            'EX: AML SAMPLE: 6G33-7002-AL/01655  (Part#: 6G33-7002-AL, Serial: 01655)
            If strPartNO.Trim.Contains("/") Then
                intPos = InStr(strPartNO, "/")
                strTemp2 = strPartNO.Trim
                If intPos > 0 Then
                    strPartNO = Left(strTemp2, intPos - 1) 'partno  /6G33-7002-AL

                    strSerialNO = Right(strTemp2, Len(strTemp2) - intPos)  'Serial / 01655
                End If
            End If


            If strPartNO.Trim.Contains(" ") Then '8D45-29474-AA 12737493
                intPos = InStr(strPartNO, " ")

                strTemp2 = strPartNO.Trim
                If intPos > 0 Then
                    strPartNO = Left(strTemp2, intPos - 1) 'partno  /8D45-29474-A
                End If
            End If

        Catch ex As Exception
            GenerateException("CheckScannedPart", ex)
        End Try


ExitThisSub:


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) strPartNO:({1}) strSerialNO:({2})", strErrorMsg, strPartNO, strSerialNO), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strErrorMsg
    End Function

    Public Function CheckSerialRequiredForPartNumber(strBuildNo As String, strStationNO As String, intPartVerifyMSeq As Integer) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) intPartVerifyMSeq:({2})", strBuildNo, strStationNO, intPartVerifyMSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)

#End If

        'AML_BuildPartsSerials: SerialS7tatus: -1 Incomplete, 1-Complete, 2-GL Override
        Dim intTemp As Int32 = GetSerialRecordPosition(strBuildNo, strStationNO, "", intPartVerifyMSeq)
        Dim blnRtnValue As Boolean = False
        If intTemp > 0 Then
            blnRtnValue = True
        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Function CheckStationSignoffStatus_ForAllOperatorPositions(strBuildNO As String, strStationNO As String, intStationOperatorsCount As Int16) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) intStationOperatorsCount:({2})", strBuildNO, strStationNO, intStationOperatorsCount), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'the intStationOperators holds the number of operators for the Stations
        ' AfterEAESClass7 Consolidation:7th August

        Dim returnValue = False
        Try
            Dim strSQL = String.Format("select count(signoffstatus) from AML_StationSignoffStatus where BuildNumber={0} and {1} and signoffstatus=1", strBuildNO.Trim, Ec.GeneralFunc.GetQueryFieldCondition("StationNO", strStationNO))
            Using connection = New Connection(gStrConnectionString)
                Dim intCompleteCount = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                returnValue = (intCompleteCount = intStationOperatorsCount)
            End Using
        Catch ex As Exception
            GenerateException(ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function CheckVersion7UserExist(ByVal strUserID As String, Optional ByRef strUserName As String = "") As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserID:({0})", strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.ExecuteScalar(Of String)("select username from euser where " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID), "")
            strUserName = result
            blnRtnValue = result <> ""
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) strUserName:({1})", blnRtnValue, strUserName), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue

    End Function

    Public Function CheckWIPBuildNumberExist(ByVal strBuildNO As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0})", strBuildNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult = False
        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.ExecuteScalar(Of Integer)(String.Format("select COUNT(*) from MES_ENGINESLIST where {0}", Ec.GeneralFunc.GetQueryFieldCondition("engineno", strBuildNO.Trim)), 0)
            blnResult = result > 0
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function ClearPartSerials(strBuildNO As String, strStationNO As String, fintPartVerifyMSeq As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) fintPartVerifyMSeq:({2})", strBuildNO, strStationNO, fintPartVerifyMSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False

        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override
                    strSQL = "update AML_BuildPartsSerials set SerialStatus=-1, SCAN_SERIALNO='', " &
                            "USERID =' ',lastupdated=getdate() " &
                            " where BuildNumber=" & strBuildNO & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) &
                            " and mseq=" & fintPartVerifyMSeq
                    '" and " & Ec.GeneralFunc.GetQueryFieldCondition("PartNumber", strParTNO)
                    connection.ExecuteNonQuery(strSQL)

                    'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
                    strSQL = "update aml_buildpartsverify set VerifyStatus=-1,lastupdated=getdate(),userid= '' " &
                               " where " & Ec.GeneralFunc.GetQueryFieldCondition("BuildNumber", strBuildNO) &
                               " and " & Ec.GeneralFunc.GetQueryFieldCondition("StationNO", strStationNO) &
                               " and mseq=" & fintPartVerifyMSeq

                    'strSQL &= " and (" & Ec.GeneralFunc.GetQueryFieldCondition("PartNumber", strParTNO) & " or " & _
                    'Ec.GeneralFunc.GetQueryFieldCondition("spare1", strParTNO) & ")"
                    connection.ExecuteNonQuery(strSQL)

                    connection.CommitTransaction()
                    blnResult = True
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Call GenerateException("ClearPartSerials: " & strSQL, ex)
                End Try
            End Using
        Catch ex As Exception
            Call GenerateException("ClearPartSerials: " & strSQL, ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult

    End Function

    Public Sub CompleteWIPData(ByVal arrBuild() As String, Optional ByRef blnResult As Boolean = False)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter arrBuild.Length:({0})", arrBuild.Length), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim intK As Int16 = 0

        Try
            blnResult = False
            Dim sqlList = arrBuild.Skip(1).Select(Function(x) GetWIPCompleteSQL(x))
            If sqlList.Count > 0 Then
                Ec.IO.RunSQLInArray(sqlList.ToArray(), blnResult)
            End If
        Catch ex As Exception
            Call GenerateException("CompleteWIPData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit blnResult:({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function DeleteBannerMessages(ByVal ArrMsgIDs() As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter ArrMsgIDs.Length:({0})", ArrMsgIDs.Length), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Try
            Dim strSQL = "DELETE FROM bannermessages WHERE msgid = {0}"
            Dim sqlList = ArrMsgIDs.Skip(1).Select(Function(x) String.Format(strSQL, x))
            If sqlList.Count > 0 Then
                Ec.IO.RunSQLInArray(sqlList.ToArray(), blnRtnValue)
            End If
        Catch ex As Exception
            Call GenerateException("DeleteBannerMessages", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue

    End Function

    Public Function DeleteBuildData(ByVal strBuildNO As String, ByVal strStationNO As String,
                Optional ByVal blnDeleteWIPData As Boolean = False) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) blnDeleteWIPData:({2})", strBuildNO, strStationNO, blnDeleteWIPData), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'like QC, WIp. Signoff status
        'blnDeleteWIPData : true, if deleting the build# from WIP screen

        Dim blnResult As Boolean = False

        Try
            Dim sqlList = New List(Of String)()
            Dim stationNoSQL = If(strStationNO.Trim <> "", String.Format(" and stationno = '{0}}'", Strings.ReplaceSingleQuote(strStationNO.Trim)), "")

            Dim strSQL = "delete from pdmrpdata where engineno='" & Strings.ReplaceSingleQuote(strBuildNO) & "'" & stationNoSQL
            sqlList.Add(strSQL)

            strSQL = "delete from pdmrpitemvalues where engineno='" & Strings.ReplaceSingleQuote(strBuildNO) & "'" & stationNoSQL
            sqlList.Add(strSQL)

            If blnDeleteWIPData Then
                strSQL = "delete from MES_ENGINESLIST where engineno='" & Strings.ReplaceSingleQuote(strBuildNO) & "'"
                sqlList.Add(strSQL)
            End If

            strSQL = "delete from AML_StationSignoffStatus where BuildNumber=" & strBuildNO.Trim & stationNoSQL
            sqlList.Add(strSQL)

            strSQL = "delete from aml_buildpartsverify where BuildNumber=" & strBuildNO.Trim & stationNoSQL
            sqlList.Add(strSQL)

            strSQL = "delete from aml_buildpartsserials where BuildNumber=" & strBuildNO.Trim & stationNoSQL
            sqlList.Add(strSQL)

            Ec.IO.RunSQLInArray(sqlList.ToArray(), blnResult)
        Catch ex As Exception
            Call GenerateException("DeleteQualityChecksData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Sub DeleteIPCAndFLMRecords(ByVal strBuildNO As String, ByVal ArrStations() As String,
    Optional ByVal blnDeleteFLMRecords As Boolean = False)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) blnDeleteFLMRecords:({1}) ArrStations.Length:({2})",
                                                          strBuildNO, blnDeleteFLMRecords, ArrStations.Length), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim blnResult As Boolean = False

        Try
            Dim strFLMDate As String = GetFLMDate()

            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    For intK = 1 To UBound(ArrStations)
                        Dim strStationNO = ArrStations(intK)
                        If strStationNO.Trim = "" Then
                            'never happens, just in case 
                            GoTo SkipThisStation
                        End If

                        strSQL = "delete from pdmrpdata where engineno='" & Strings.ReplaceSingleQuote(strBuildNO) & "'"
                        strSQL &= " and stationno='" & Strings.ReplaceSingleQuote(strStationNO.Trim) & "'"
                        connection.ExecuteNonQuery(strSQL)

                        strSQL = "delete from pdmrpitemvalues where engineno='" & Strings.ReplaceSingleQuote(strBuildNO) & "'"
                        strSQL &= " and stationno='" & Strings.ReplaceSingleQuote(strStationNO.Trim) & "'"
                        connection.ExecuteNonQuery(strSQL)

                        If blnDeleteFLMRecords Then
                            strBuildNO = strFLMDate
                            strSQL = "delete from pdmrpdata where engineno='" & Strings.ReplaceSingleQuote(strBuildNO) & "'"
                            strSQL &= " and stationno='" & Strings.ReplaceSingleQuote(strStationNO.Trim) & "'"
                            connection.ExecuteNonQuery(strSQL)

                            strSQL = "delete from pdmrpitemvalues where engineno='" & Strings.ReplaceSingleQuote(strBuildNO) & "'"
                            strSQL &= " and stationno='" & Strings.ReplaceSingleQuote(strStationNO.Trim) & "'"
                            connection.ExecuteNonQuery(strSQL)
                        End If
SkipThisStation:
                    Next
                    connection.CommitTransaction()
                    blnResult = True
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Throw ex
                End Try
            End Using
        Catch ex As Exception
            Call GenerateException("DeleteIPCAndFLMRecords: " & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub DeleteUserLastVisitedScreen(ByVal strIPAddress As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strIPAddress:({0})", strIPAddress), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        strSQL = " delete from AMLSwitchLogin where " &
                    Ec.GeneralFunc.GetQueryFieldCondition("ipaddress", strIPAddress)

        Ec.IO.RunSQL(strSQL)

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function FormatSubHeaderReferenceNumber(ByVal strRefNumber As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strRefNumber:({0})", strRefNumber), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strRtnValue As String = "", intCharCount As Int16 = 5  ''ES-10: AML Support: Week 39

        If strRefNumber.Trim.Length < intCharCount Then
            strRtnValue = StrDup(intCharCount - Len(strRefNumber.Trim), "0") & strRefNumber
        Else
            strRtnValue = strRefNumber.Trim
        End If
        If DBConfig.Version7 Then
            'ES-10: AML Support: Week 39
            If strRtnValue.Trim = "0000" Or strRtnValue.Trim = "00000" Then
                strRtnValue = ""
            End If
        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strRtnValue
    End Function

    Public Function GetAMLModelsList(Optional ByVal intIDx As Int16 = 0) As List(Of AML_ModelList)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = New List(Of AML_ModelList)

        Try
            Dim strSQL = "SELECT * FROM AML_ModelList {0} ORDER BY modelnumber"
            strSQL = String.Format(strSQL, If(intIDx > 0, String.Format("WHERE id = {0}", intIDx), ""))

            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.GetDataIntoClassOf(Of AML_ModelList)(strSQL)
                If (result IsNot Nothing) Then returnValue = result.ToList()
            End Using
        Catch ex As Exception
            GenerateException("GetAMLModelsList", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", returnValue.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue

    End Function

    Public Function GetAndonCallStatus(ByVal intLineID As Integer, ByVal strStationNO As String) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) strStationNO:({1})", intLineID, strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strLineStatusTableName = GetLineStatusTableName(intLineID)
        Dim strSQL = "select andoncall from " & strLineStatusTableName & " where lineid=" & intLineID.ToString & " and stationno='" & strStationNO.Trim & "'"

        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(strSQL, 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue

    End Function

    Public Function GetAndonLineStatusFlag(ByVal intLineID As Integer, Optional ByRef dtEngineMoved As Date = Nothing) As Int16

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0})", intLineID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Andon Table.LineStatus : 1-Line Indexsing, 2- indexcomplete, 0-default value
        'AML Andon table holds the data by LINE
        'AML AML_ANDON_Stations table holds the data by Line and Stations

        Dim strSQL As String = "select linestatus,lastupdated from aml_andon where lineid=" & intLineID.ToString
        Dim intLineStatus As Int16 = 0
        Try
            Using connection = New Connection(gStrConnectionString)
                Using dataTable = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(dataTable) AndAlso dataTable.Rows.Count > 0) Then
                        Dim dbReader As DataRow = dataTable.Rows(0)
                        intLineStatus = Data.GetDataRowValue(Of Int16)(dbReader("linestatus"), 0)
                        dtEngineMoved = Data.GetDataRowValue(dbReader("lastupdated"), DateTime.MinValue)
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetAndonLineStatusFlag", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) ({1})", intLineStatus, dtEngineMoved), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return intLineStatus
    End Function

    Public Function GetAndonLineStatusFlag_IndependentStation(ByVal intLineID As Int16, strStationNO As String, Optional ByRef dtEngineMoved As Date = Nothing) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) strStationNO:({1})", intLineID, strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'AML Andon table holds the data by LINE
        'AML AML_ANDON_Stations table holds the data by Line and Stations

        'Andon Table.LineStatus : 1-Line Indexsing, 2- indexcomplete, 0-default value
        Dim strSQL As String = "select linestatus,lastupdated from AML_ANDON_Stations where lineid=" & intLineID.ToString &
             " And " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)

        Dim intLineStatus = 0
        Try
            dtEngineMoved = Date.MinValue

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim row As DataRow = table.Rows(0)

                        intLineStatus = Data.GetDataRowValue(row("linestatus"), 0)
                        dtEngineMoved = Data.GetDataRowValue(row("lastupdated"), DateTime.MinValue)
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetAndonLineStatusFlag_independentStation", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit intLineStatus:({0}) dtEngineMoved:({1})", intLineStatus, dtEngineMoved), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return intLineStatus
    End Function

    Public Function GetBannerMessageForStation(ByVal strStationNO As String) As stBannerMsgList()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION_LOW(String.Format("Enter strStationNO:({0})", strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'get the messages for the station
        Dim objMBL(0) As stBannerMsgList
        Dim strSQL As String = ""
        Try
            strSQL = "select BannerMessages.* from BannerMessages,BannerMessages_Stations " &
                        " where BannerMessages.messagestatus=1 " &
                        " and BannerMessages.msgid=BannerMessages_Stations.msgid " &
                        " and " & Ec.GeneralFunc.GetQueryFieldCondition("BannerMessages_Stations.stationno", strStationNO) &
                        " and (getdate() between startdate and enddate+1)"

            objMBL = GetBannerMessageList(, , , , strSQL)
        Catch ex As Exception
            GenerateException("GetBannerMessageForStation", ex)
        End Try


#If TRACE Then
        Log.OPERATION_LOW(String.Format("Exit Count:({0})", UBound(objMBL)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objMBL
    End Function

    Public Function GetBannerMessageList(Optional ByVal intMsgID As Int32 = 0,
                Optional ByVal blnForBanner As Boolean = False,
                Optional ByVal strSQLCondition As String = "",
                Optional ByVal intLineID As Int16 = 0,
                Optional ByVal strAltSQL As String = "") As stBannerMsgList()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION_LOW(String.Format("Enter intMsgID:({0}) blnForBanner:({1}) strSQLCondition:({2}) intLineID:({3}) strAltSQL:({4})",
                                                          intMsgID, blnForBanner, strSQLCondition, intLineID, strAltSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objMBL() As stBannerMsgList
        ReDim objMBL(0)
        Dim strSQL As String = ""
        Dim strDateFormat As String = GetDateFormat()
        Dim strCond As String = " where "

        Try

            If strAltSQL.Trim <> "" Then
                strSQL = strAltSQL
                GoTo GetTheData
            End If

            strSQL = "select * from BannerMessages "
            If intMsgID > 0 Then
                strSQL &= strCond & " msgid=" & intMsgID
                strCond = " and "
            End If
            If intLineID > 0 Then
                strSQL &= strCond & " lineid=" & intLineID
                strCond = " and "
            End If
            If blnForBanner Then
                'MessageStatus: 0-Draft/Pending Approval, 1-Approved

                'strSQLCond = " (lastupdated between to_date('" & Trim(Format(dtFrom.Value, "MM/dd/yyyy")) & " 0:0:0' ,'MM/DD/YYYY HH24:MI:SS') and  " & _
                '    " to_date('" & Trim(Format(dtTo.Value, "MM/dd/yyyy")) & " 23:59:59','MM/DD/YYYY HH24:MI:SS'))"

                'three SQL conditions for the dates
                '1. Current date between  start and end date
                '2. Start date is today
                '3. End date is today

                strSQL &= strCond & " ((" & GetDateBuiltInFunction() & " BETWEEN startdate AND enddate) or " &
                            "(to_char(sysdate,'MM/DD/YYYY')=to_char(startdate,'MM/DD/YYYY') or " &
                            "to_char(sysdate,'MM/DD/YYYY')=to_char(enddate,'MM/DD/YYYY')))" &
                            " and messagestatus=1"
                strCond = " and "
            End If

            If strSQLCondition.Trim <> "" Then
                strSQL &= strCond & strSQLCondition
                strCond = " and "
            End If
GetTheData:
            strSQL &= " order by lastupdated desc"

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objMBL(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1

                            objMBL(counter).MessageID = Data.GetDataRowValue(row("msgid"), 0)
                            objMBL(counter).MessageText = Data.GetDataRowValue(row("messagetext"), "")
                            objMBL(counter).RaisedBy = Data.GetDataRowValue(row("raisedby"), "")
                            objMBL(counter).ApprovedBy = Data.GetDataRowValue(row("approvedby"), "")
                            objMBL(counter).MessageReason = Data.GetDataRowValue(row("messagereason"), "")
                            objMBL(counter).MessageType = Data.GetDataRowValue(row("messagetype"), "")
                            objMBL(counter).LineID = Data.GetDataRowValue(Of Int16)(row("lineid"), 1)
                            objMBL(counter).StationFrom = Data.GetDataRowValue(row("stationfrom"), "")
                            objMBL(counter).StationTo = Data.GetDataRowValue(row("stationto"), "")
                            objMBL(counter).StartDate = Data.GetDataRowValue(row("startdate"), DateTime.MinValue)
                            objMBL(counter).EndDate = Data.GetDataRowValue(row("enddate"), DateTime.MinValue)
                            objMBL(counter).LastUpdated = Data.GetDataRowValue(row("lastupdated"), DateTime.MinValue)
                            'MessageStatus: 0-Draft/Pending Approval, 1-Approved
                            objMBL(counter).MessageStatus = Data.GetDataRowValue(Of Int16)(row("messagestatus"), 0)
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetBannerMessageList", ex)
        End Try


#If TRACE Then
        Log.OPERATION_LOW(String.Format("Exit Count:({0})", UBound(objMBL)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objMBL
    End Function

    Public Function GetBannerMessageStationsList(
                ByVal intMessageID As Int32,
                Optional ByVal strAddlSQL As String = "") As stBannerMessagesStations()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION_LOW(String.Format("Enter intMessageID:({0}) strAddlSQL:({1})", intMessageID, strAddlSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objBMS() As stBannerMessagesStations
        ReDim objBMS(0)

        Try
            Dim strSQL = "select * from BannerMessages_Stations "
            If intMessageID > 0 Then
                strSQL &= " where msgid=" & intMessageID
            End If
            If strAddlSQL.Trim <> "" Then
                strSQL &= " where " & strAddlSQL
            End If
            strSQL &= " order by stationseq"

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objBMS(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1

                            objBMS(counter).MessageID = Data.GetDataRowValue(row("MsgID"), 0)
                            objBMS(counter).Stage = Data.GetDataRowValue(row("stationno"), "")
                            objBMS(counter).StationSeq = Data.GetDataRowValue(row("stationseq"), 0)
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetBannerMessageStationsList", ex)
        End Try


#If TRACE Then
        Log.OPERATION_LOW(String.Format("Exit Count:({0})", UBound(objBMS)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objBMS
    End Function

    Public Function GetBootBagPartQuantity(strBuildNO As String, strStationNO As String, strPartNO As String, Optional ByRef strBoot As String = "") As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2})", strBuildNO, strStationNO, strPartNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strQuantity = ""
        Try
            strBoot = ""
            Dim strSQL = "select quantity,insidebox from AML_BuildBootBag where " &
                   " buildnumber=" & strBuildNO & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) &
                   " and " & Ec.GeneralFunc.GetQueryFieldCondition("partnumber", strPartNO)

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each row As DataRow In table.Rows
                            strQuantity = Data.GetDataRowValue(row(0), "")

                            'BootBag.InsideBag: TRUE-Inside box, False - BOOT
                            Select Case Data.GetDataRowValue(row(1), -1)
                                Case 0
                                    strBoot = "Boot"
                                Case 1
                                    strBoot = "Bag"
                            End Select
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetBootBagPartQuantity", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0} strBoot:({1})", strQuantity, strBoot), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strQuantity

    End Function

    Public Function GetBootBagPartsList(ByVal strBuilNO As String, strStationNO As String, Optional intOperatorPosition As Int16 = 0) As List(Of AML_BuildBootBag)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuilNO:({0}) strStationNO:({1}) intOperatorPosition:({2})", strBuilNO, strStationNO, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL = "select * from AML_BuildBootBag where BuildNumber=" & strBuilNO
        If strStationNO.Trim <> "" Then
            strSQL &= $" and stationno = '{strStationNO}'"
        End If
        If intOperatorPosition <> 0 Then
            strSQL &= $" and operatorposition={intOperatorPosition}"
        End If
        strSQL &= " order by partnumber"

        Dim returnValue = New List(Of AML_BuildBootBag)

        Try
            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.GetDataIntoClassOf(Of AML_BuildBootBag)(strSQL)
                If (result IsNot Nothing) Then returnValue = result.ToList()
            End Using
        Catch ex As Exception
            GenerateException("GetBootBagPartsList", ex)
        End Try

ExitThisSub:


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", returnValue.Count()), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetBuildLogData(ByVal intMsgID As Int32, ByVal intMsgType As Int16) As stBuildLogData()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intMsgID:({0}) intMsgType:({1})", intMsgID, intMsgType), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objBLD() As stBuildLogData
        ReDim objBLD(0)
        Try
            Dim strSQL = "select descx from aml_buildlog where msgid=" & intMsgID & " and msgtype=" & intMsgType.ToString & " order by msgseq"
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objBLD(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1
                            objBLD(counter).DescX = Data.GetDataRowValue(row("descx"), "").Trim()
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetBuildLogData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objBLD)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objBLD

    End Function

    Public Function GetBuildNoModelNumber(ByVal buildNo As String, Optional getLinkedModel As Boolean = False, Optional ByRef linkedModel As Integer = 0) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) blnGetLinkedModel:({1})", buildNo, getLinkedModel), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim intTemp = 0
        Try
            If buildNo.Trim = "" Or buildNo.Trim = "0" Then
                GoTo ExitThisSub
            End If

            Using connection = New Connection(gStrConnectionString)
                Dim strSQL = String.Format("SELECT TOP 1 model1 FROM aml_vehicleorders WITH (NOLOCK) WHERE buildnumber={0}", buildNo.Trim())
                Dim result = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                If (Not IsNothing(result) AndAlso Not IsDBNull(result)) Then
                    intTemp = result
                End If

                If getLinkedModel Then
                    strSQL = String.Format("SELECT TOP 1 modelnumberlinked FROM aml_modellist WITH (NOLOCK) WHERE modelnumber={0}", intTemp)
                    result = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                    If (Not IsNothing(result) AndAlso Not IsDBNull(result)) Then
                        linkedModel = result
                    End If
                End If
            End Using

ExitThisSub:

        Catch ex As Exception
            GenerateException(ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) intLinkedModel:({1})", intTemp, linkedModel), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return intTemp
    End Function

    Public Function GetBuildNumberModel(ByVal intBuildNO As Int32) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intBuildNO:({0})", intBuildNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As String = ""
        Using connection = New Connection(gStrConnectionString)
            Dim strSQL As String = "select descriptionx from aml_vehicleorders where buildnumber='" & intBuildNO.ToString & "'"
            returnValue = connection.ExecuteScalar(Of String)(strSQL, "")
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function GetBuildNumberPartSerialDetail(Mseq As String, StationNo As String, PartNo As String) As List(Of AML_BuildPartSerials)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter Mseq:({0}) StationNo:({1}) PartNo:({2})", Mseq, StationNo, PartNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = New List(Of AML_BuildPartSerials)
        Try
            Using connection = New Connection(gStrConnectionString)
                Dim strSQL = "SELECT * FROM AML_BuildPartsSerials where Mseq =" & Mseq & " and StationNO ='" & StationNo & "' and PartNumber ='" & PartNo & "'"
                Dim result = connection.GetDataIntoClassOf(Of AML_BuildPartSerials)(strSQL)
                If (result IsNot Nothing) Then
                    returnValue = result.ToList()
                End If
            End Using
        Catch ex As Exception
            GenerateException("GetBuildNumberPartSerialList", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", returnValue.Count()), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function GetBuildNumberPartSerialList(strBuildNO As String, strStationNO As String,
                                                 Optional strPartNO As String = "",
                                                 Optional intSerialStatus As Integer = 0,
                                                 Optional intSerialSEQ As Integer = -22,
                                                 Optional intOperatorPosition As Integer = -1,
                                                 Optional intPartMVerifySeq As Integer = 0) As List(Of AML_BuildPartSerials)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2}) intSerialStatus:({3}) intSerialSEQ:({4}) intOperatorPosition:({5}) intPartMVerifySeq:({6})",
                                                          strBuildNO, strStationNO, strPartNO, intSerialStatus, intSerialSEQ, intOperatorPosition, intPartMVerifySeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override

        Dim returnValue = New List(Of AML_BuildPartSerials)
        Try
            Using connection = New Connection(gStrConnectionString)
                Dim strSQL = GetBuildPartSerialQuery("", strBuildNO, strStationNO, strPartNO, intSerialStatus, intSerialSEQ, intOperatorPosition, intPartMVerifySeq)
                Dim result = connection.GetDataIntoClassOf(Of AML_BuildPartSerials)(strSQL)
                If (result IsNot Nothing) Then
                    returnValue = result.ToList()
                End If
            End Using
        Catch ex As Exception
            GenerateException("GetBuildNumberPartSerialList", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", returnValue.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetBuildNumberPartVerifyList(strBuildNO As String,
                                                 strStationNO As String,
                                                 Optional strPartNO As String = "",
                                                 Optional blnSerialsRequired As Boolean = False,
                                                 Optional blnVerifyStation As Boolean = False,
                                                 Optional blnGetInCompleteRecordOnly As Boolean = False,
                                                 Optional intPartVerifyStatus As Int16 = -22,
                                                 Optional strOrderBy As String = "",
                                                 Optional ByVal strCallFrom As String = "",
                                                 Optional intOperatorPosition As Int16 = -1) As stBuildPartVerifyList()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2}) blnSerialsRequired:({3}) blnVerifyStation:({4}) " &
                                                          "blnGetInCompleteRecordOnly:({5}) intPartVerifyStatus:({6}) strOrderBy:({7}) strCallFrom:({8}) intOperatorPosition:({9})",
                                                          strBuildNO, strStationNO, strPartNO, blnSerialsRequired, blnVerifyStation,
                                                          blnGetInCompleteRecordOnly, intPartVerifyStatus, strOrderBy, strCallFrom, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objBPVL(0) As stBuildPartVerifyList
        Dim intCount = 0
        'Build SQL statement
        Dim strSQL = GetBuildPartVerifyQuery("all", strBuildNO, strStationNO, strPartNO, blnVerifyStation, blnGetInCompleteRecordOnly, intPartVerifyStatus, strOrderBy, strCallFrom, intOperatorPosition)
        Try
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim Preserve objBPVL(table.Rows.Count)
                        For Each row As DataRow In table.Rows
                            intCount += 1
                            objBPVL(intCount).BuildNo = ""
                            objBPVL(intCount).StationNO = Data.GetDataRowValue(row("StationNO"), "")
                            objBPVL(intCount).PartNumber = Data.GetDataRowValue(row("PartNumber"), "")
                            objBPVL(intCount).Mseq = Data.GetDataRowValue(row("MSEQ"), 0)
                            objBPVL(intCount).UserID = Data.GetDataRowValue(row("USERID"), "")
                            objBPVL(intCount).OperatorPosition = Data.GetDataRowValue(Of Int16)(row("OPERATORPOSITION"), 0)
                            objBPVL(intCount).Mseq = Data.GetDataRowValue(row("mseq"), 0)
                            objBPVL(intCount).VerifyStatus = Data.GetDataRowValue(Of Int16)(row("VerifyStatus"), -1) 'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
                            objBPVL(intCount).Spare2 = Data.GetDataRowValue(row("Spare2"), "")  ''holds Subheader reference number KM: 05/29/2014
                            objBPVL(intCount).Spare3 = Data.GetDataRowValue(row("Spare3"), "")
                            objBPVL(intCount).PartDescX = Data.GetDataRowValue(row("PartDescX"), "")
                            objBPVL(intCount).Verify_Station = Data.GetDataRowValue(row("Verify_Station"), "")
                            objBPVL(intCount).Tracking = Data.GetDataRowValue(row("Tracking"), "")
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetBuildNumberPartVerifyList", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objBPVL)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objBPVL
    End Function

    Public Sub GetDescriptionFromGlossary(ByVal strBodyCode As String,
                                          ByVal strModelCode As String,
                                          ByVal strTerritoryCode As String,
                                          ByVal strDriveCode As String,
                                          ByVal strGearboxCode As String,
                                          ByRef strBodeDescX As String,
                                          ByRef strModelDescX As String,
                                          ByRef strTerritoryDescX As String,
                                          ByRef strDriveDescX As String,
                                          ByRef strGearBoxDescX As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBodyCode:({0}) strModelCode:({1}) strTerritoryCode:({2}) strDriveCode:({3}) strGearboxCode:({4})",
                                                          strBodyCode, strModelCode, strTerritoryCode, strDriveCode, strGearboxCode), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        strBodeDescX = "" : strModelDescX = "" : strTerritoryDescX = "" : strDriveDescX = "" : strGearBoxDescX = ""

        Dim strSQL As String = ""
        Dim strSQLCond As String = ""

        Dim strBodyGroup As String = "bodystyle"
        Dim strModelGroup As String = "model"
        Dim strTerritoryGroup As String = "territory"
        Dim strDriveGroup As String = "drive"
        Dim strGearBoxGroup As String = "gearbox"
        Dim strTempGroup As String = "", strDescX As String = ""

        Try
            strSQL = "select groupx,descriptionx from aml_vehicleglossary where groupX <> 'option' "

            If (Not String.IsNullOrEmpty(strBodyCode)) Then
                strSQLCond = "(groupx='" & strBodyGroup.Trim & "' and valuex='" & Strings.ReplaceSingleQuote(strBodyCode) & "') "
            End If

            If (Not String.IsNullOrEmpty(strModelCode)) Then
                If strSQLCond.Trim <> "" Then strSQLCond &= " or "
                strSQLCond &= "(groupx='" & strModelGroup.Trim & "' and valuex='" & Strings.ReplaceSingleQuote(strModelCode) & "') "
            End If

            If (Not String.IsNullOrEmpty(strTerritoryCode)) Then
                If strSQLCond.Trim <> "" Then strSQLCond &= " or "
                strSQLCond &= "(groupx='" & strTerritoryGroup.Trim & "' and valuex='" & Strings.ReplaceSingleQuote(strTerritoryCode) & "') "
            End If

            If (Not String.IsNullOrEmpty(strDriveCode.Trim)) Then
                If strSQLCond.Trim <> "" Then strSQLCond &= " or "
                strSQLCond &= "(groupx='" & strDriveGroup.Trim & "' and valuex='" & Strings.ReplaceSingleQuote(strDriveCode) & "') "
            End If

            If (Not String.IsNullOrEmpty(strGearboxCode.Trim)) Then
                If strSQLCond.Trim <> "" Then strSQLCond &= " or "
                strSQLCond &= "(groupx='" & strGearBoxGroup.Trim & "' and valuex='" & Strings.ReplaceSingleQuote(strGearboxCode) & "') "
            End If

            If strSQLCond.Trim = "" Then GoTo ExitThisSub

            strSQL &= " and " & strSQLCond

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        For Each row As DataRow In table.Rows
                            strTempGroup = Data.GetDataRowValue(row("groupx"), "")
                            If strTempGroup.Trim = "" Then GoTo SkipThisRecord

                            strDescX = Data.GetDataRowValue(row("descriptionx"), "")
                            Select Case strTempGroup.Trim.ToLower
                                Case strBodyGroup
                                    strBodeDescX = strDescX.Trim
                                Case strModelGroup
                                    strModelDescX = strDescX.Trim
                                Case strTerritoryGroup
                                    strTerritoryDescX = strDescX.Trim
                                Case strDriveGroup
                                    strDriveDescX = strDescX.Trim
                                Case strGearBoxGroup
                                    strGearBoxDescX = strDescX.Trim
                            End Select
SkipThisRecord:
                        Next
                    End If
                End Using
            End Using
ExitThisSub:
        Catch ex As Exception
            GenerateException("GetDescriptionFromGlossary", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit strBodeDescX:({0}) strModelDescX:({1}) strTerritoryDescX:({2}) strDriveDescX:({3}) strGearBoxDescX:({4})",
                                 strBodeDescX, strModelDescX, strTerritoryDescX, strDriveDescX, strGearBoxDescX), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub


    Public Function GetEASELinkedModel(ByVal intModelNumber As Integer) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intModelNumber:({0})", intModelNumber), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = 0
        Dim strSQL As String = "select ModelNumberLinked from aml_modellist where ModelNumber=" & intModelNumber
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(strSQL, 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetEASEPartNoByArea(ByVal area As String) As Parts.stPartSearchResult()
        Dim partSearchCriteria = New PartSearchCriteria
        partSearchCriteria.RecType = "0"
        partSearchCriteria.GroupTech = area.Trim() & "*"

        Dim partSearchResults() As Parts.stPartSearchResult
        ReDim partSearchResults(0)

        Try
            partSearchResults = Ec.Parts.GetPartSearchRecords(partSearchCriteria)
        Catch ex As Exception
            GenerateException(ex)
        End Try

        Return partSearchResults
    End Function

    Public Function GetEASEPartNo(ByVal strArea As String,
                                  ByVal intBuildNoModelNumber As Integer,
                                  ByVal strStationNO As String,
                                  ByVal intOperatorLoginPosition As Integer,
                                  Optional ByVal blnAvoidOPNO As Boolean = False,
                                  Optional ByRef strSQL As String = "",
                                  Optional intLinkedModelNO As Integer = 0) As Parts.stPartSearchResult()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strArea:({0}) intBuildNoModelNumber:({1}) strStationNO:({2}) intOperatorLoginPosition:({3}) blnAvoidOPNO:({4}) intLinkedModelNO:({5})",
                                                          strArea, intBuildNoModelNumber, strStationNO, intOperatorLoginPosition, blnAvoidOPNO, intLinkedModelNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'blnAvoidOPNO is used only for getting the QC status count for the station

        Dim partSearchCriteria As New PartSearchCriteria
        Dim objSearchResult(0) As Parts.stPartSearchResult
        Dim blnSecondSearch As Boolean = False

        Try
            partSearchCriteria.RecType = "0"
            partSearchCriteria.GroupTech = strArea.Trim & "*"
            'objSearchCriteria.UserDef1 = HttpContext.Current.Session("Model") & "*"

            'the models are linked, so get the parent model for the current model and use it
            'NOTE: UserDef1 is used to hold the model field for the plan, any changes in this has to
            'be updated in GetEASEPartInfo and ShowOPBOM. GetModelForThePlan subs

            'objSearchCriteria.UserDef1 = GetLinkedModel(strModel.Trim) & "*"

            If intBuildNoModelNumber = 0 Then
                'for PSS Report only
            Else 'default
                If intLinkedModelNO > 0 Then
                    partSearchCriteria.UserDef1 = intLinkedModelNO.ToString()
                Else
                    partSearchCriteria.UserDef1 = GetEASELinkedModel(intBuildNoModelNumber).ToString()
                End If

            End If


            If blnAvoidOPNO Then
                'SPECIAL CASE   
            Else            'DEFAULT OPTION
                If intOperatorLoginPosition > 0 Then
                    partSearchCriteria.OPNO = strStationNO.Trim & "-" & intOperatorLoginPosition.ToString
                Else
                    partSearchCriteria.OPNO = strStationNO.Trim
                End If
            End If


SearchAgain:
            objSearchResult = Ec.Parts.GetPartSearchRecords(partSearchCriteria, strSQL)

            If UBound(objSearchResult) = 0 AndAlso blnSecondSearch = False And intOperatorLoginPosition = 1 Then
                'ok, the operation number may not '-1' suffixed to it
                'try search again
                partSearchCriteria.OPNO = strStationNO.Trim
                blnSecondSearch = True        'to avoid looping
                GoTo SearchAgain
            End If

        Catch ex As Exception
            GenerateException("GetEASEPartNo", ex)
            'gErrorMsg = GeneralError(gUser, "aml1: GetEASEPartNo", ex, True)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0}) strSQL:({1})", UBound(objSearchResult), strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objSearchResult

    End Function


    Public Function GetImportLogConfigurationData(Optional ByRef strImportDirectory As String = "") As String()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrEmailList(0) As String
        Try
            strImportDirectory = ""

            Dim strSQL = "select data,keyx from client where keyx between 153 and 163 and wl=99"
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim counter = 0
                        For Each reader As DataRow In table.Rows
                            Dim intKeyX = Data.GetDataRowValue(reader("keyx"), 0)
                            Select Case intKeyX
                                Case 153, 154, 155, 156, 157, 158
                                    If Not IsDBNull(reader("data")) Then
                                        counter += 1
                                        ReDim Preserve ArrEmailList(counter)
                                        ArrEmailList(counter) = Trim(Data.GetDataRowValue(reader("data"), ""))
                                    End If
                                Case 163
                                    strImportDirectory = Trim(Data.GetDataRowValue(reader("data"), ""))
                            End Select
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetImportLogConfigurationData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0}) strImportDirectory:({1})", UBound(ArrEmailList), strImportDirectory), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrEmailList

    End Function

    Public Function GetLastBuildRunDate() As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(datex) from aml_buildlog where msgtype=3"
        Dim returnValue = ""
        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.ExecuteScalar(Of Integer)(strSQL, 0)
            If (Not IsNothing(result) AndAlso Not IsDBNull(result)) Then
                returnValue = result.ToString()
            End If
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetLastOpenedSubheader(ByVal strUserID As String, ByVal strBuildNO As String,
            ByVal strStationNO As String, ByVal intOperatorPosition As Int16) As stLastOpenSubHeader


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserID:({0}) strBuildNO:({1}) strStationNO:({2}) intOperatorPosition:({3})", strUserID, strBuildNO, strStationNO, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim objLOS As New stLastOpenSubHeader

        Try

            objLOS.ID = 0
            objLOS.OPNO = ""
            objLOS.Rectype = ""
            objLOS.Seq = 0
            objLOS.SubHeader = ""
            objLOS.CommentX = ""
            objLOS.OperatorPosition = 1

            strSQL = "select id,rectype,seq,opno,subheader,commentx,operatorposition" &
                        " from AMLUserLog " &
                        " where " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID) &
                        " And " & Ec.GeneralFunc.GetQueryFieldCondition("buildno", strBuildNO) &
                        " And " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) &
                        " And operatorposition=" & intOperatorPosition &
                        "   order by mseq desc"     'to get the last mseq

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim row As DataRow = table.Rows(0)

                        objLOS.ID = Data.GetDataRowValue(row("id"), 0)
                        objLOS.Rectype = Data.GetDataRowValue(row("rectype"), "")
                        objLOS.Seq = Data.GetDataRowValue(Of Int16)(row("seq"), 0)
                        objLOS.OPNO = Data.GetDataRowValue(row("opno"), "")
                        objLOS.SubHeader = Data.GetDataRowValue(row("subheader"), "")
                        objLOS.CommentX = Data.GetDataRowValue(row("commentx"), "")
                        objLOS.OperatorPosition = Data.GetDataRowValue(Of Int16)(row("OperatorPosition"), 1)

                        If Trim(UCase(objLOS.CommentX)) = "EASEAUTO" Then
                            objLOS.CommentX = ""
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetLastOpenedSubheader", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ID:({0}) OPNO:({1}) SubHeader:({2})", objLOS.ID, objLOS.OPNO, objLOS.SubHeader), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objLOS

    End Function

    Public Function GetLineDescription(ByVal intLineID As Integer) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0})", intLineID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select description from lines where lineid=" & intLineID.ToString
        Dim returnValue = ""
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of String)(strSQL, "")
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue

    End Function

    Public Function GetLineID(ByVal strLineDescription As String) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strLineDescription:({0})", strLineDescription), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL = "select lineid from lines where " & Ec.GeneralFunc.GetQueryFieldCondition("description", strLineDescription)
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue

    End Function

    Public Function GetLines(Optional ByVal intLineID As Int16 = 0) As stLines()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0})", intLineID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objLines(0) As stLines

        Try
            Dim strSQL = "select * from lines "
            If intLineID > 0 Then       'only used while building FLM/QC 
                strSQL &= " where lineid=" & intLineID.ToString
            End If
            strSQL &= " order by description"

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim Preserve objLines(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1

                            objLines(counter).LineID = Data.GetDataRowValue(Of Int16)(row("lineid"), 0)
                            objLines(counter).DescX = Data.GetDataRowValue(row("description"), "")
                            objLines(counter).PlantID = Data.GetDataRowValue(Of Int16)(row("plantid"), 0)

                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetLines", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", objLines), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objLines

    End Function

    Public Function GetLineStations(ByVal intLineID As Integer, Optional ByVal blnOSM As Boolean = False, Optional ByVal intWSID As Int16 = 0) As stStations()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) blnOSM:({1}) intWSID:({2})", intLineID, blnOSM, intWSID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objLineStations(0) As stStations
        Try
            Dim strSQL = "select lineid,absnno,absndesc,absnstart,absnend,operators,osmrequired from stations "
            If intLineID > 0 Then
                strSQL &= " where lineid=" & intLineID
            End If

            If blnOSM Then
                strSQL = "select stations.lineid lineid,stations.absnno absnno,stations.absnstart absnstart,stations.absnend absnend,stations.operators operators,stations.osmrequired osmrequired,stations.absndesc absndesc from stations,lines ,pdmworkstation"
                strSQL &= " where lines.lineid=stations.lineid"
                strSQL &= " And pdmworkstation.wc= lines.description"
                strSQL &= " And pdmworkstation.wscode=stations.absnno"

                If intLineID > 0 Then
                    strSQL &= " And lines.lineid=" & intLineID
                End If
                If intWSID > 0 Then
                    strSQL &= " And pdmworkstation.wsid=" & intWSID
                End If
            End If
            strSQL &= " order by lineid, absnno"  'don't change the sort order, used in frmpoll page.

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim objLineStations(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1
                            objLineStations(counter).LineID = Data.GetDataRowValue(Of Int16)(row("lineid"), 0)
                            objLineStations(counter).StationNO = Data.GetDataRowValue(row("absnno"), "")
                            objLineStations(counter).StationStart = Data.GetDataRowValue(row("absnstart"), "")
                            objLineStations(counter).StationEnd = Data.GetDataRowValue(row("absnend"), "")
                            objLineStations(counter).Operators = Data.GetDataRowValue(Of Byte)(row("operators"), 1)
                            objLineStations(counter).OSMRequired = Data.GetDataRowValue(Of Int16)(row("osmrequired"), 0)
                            objLineStations(counter).Desc = Data.GetDataRowValue(row("absndesc"), "")
                        Next
                    End If
                End Using
            End Using

        Catch ex As Exception
            GenerateException("GetLineStations", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objLineStations)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objLineStations
    End Function

    Public Function GetLineStatus(Optional ByVal lineID As Integer? = Nothing, Optional ByVal stationNo As String = Nothing) As stLineStatus()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objLS(0) As stLineStatus
        Try
            Dim whereClause = New List(Of String)()
            If (lineID.HasValue) Then whereClause.Add(String.Format("lineid = {0}", lineID))
            If (Not String.IsNullOrEmpty(stationNo)) Then whereClause.Add(String.Format("stationno = '{0}'", stationNo))
            Dim whereStatement = ""
            If whereClause.Any() Then
                whereStatement = String.Format(" WHERE {0}", String.Join(" AND ", whereClause))
            End If

            Dim tableQueryList = New List(Of String)()
            Dim amlLineStatusSql = "select lineid, stationno, buildno, enginemovedate, ipcstatus, andoncall, readstatus from AML_LineStatus {0}"
            Dim amlOtherLinesSql = "select lineid, stationno, CAST(buildno as varchar(10)), enginemovedate, ipcstatus, andoncall, readstatus FROM {0} {1}"

            tableQueryList.Add(String.Format(amlLineStatusSql, whereStatement))
            tableQueryList.Add(String.Format(amlOtherLinesSql, "aml_linestatus_bodyshop", whereStatement))
            If (Ec.AppConfig.AstonMartinStAthan) Then
                tableQueryList.Add(String.Format(amlOtherLinesSql, "aml_linestatus_paint", whereStatement))
                tableQueryList.Add(String.Format(amlOtherLinesSql, "aml_linestatus_ma", whereStatement))
            End If

            'READ ME. DON'T CHANGE THE ORDER BY CLAUSE
            'AFFECTS ANDON LINE MOVE (FRMPOLL.ASPX)
            Dim sql = String.Format("{0} ORDER BY lineid,stationno", String.Join(Environment.NewLine & "UNION" & Environment.NewLine, tableQueryList))

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(sql, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim objLS(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1

                            objLS(counter).AndonCall = Data.GetDataRowValue(Of Byte)(row("andoncall"), 0)
                            objLS(counter).BuildNO = Data.GetDataRowValue(row("BuildNo"), "")
                            objLS(counter).EngineMoveDate = Data.GetDataRowValue(row("EngineMoveDate"), Date.MinValue)
                            objLS(counter).LineID = Data.GetDataRowValue(Of Int16)(row("lineid"), 0)
                            objLS(counter).OperatorCount = 1
                            objLS(counter).QualityChecksStatus = Data.GetDataRowValue(Of Int16)(row("IPCStatus"), 0)
                            objLS(counter).StationNO = Data.GetDataRowValue(row("StationNO"), "")
                            objLS(counter).OSMRequired = False
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetLineStatus: ", ex)
        End Try

ExitThisSub:
#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objLS)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objLS
    End Function



    Public Function GetMDMProcessDocumentsList(objSH() As Parts.stSubHeader,
                                               strStationNO As String, strUSERID As String,
                                               Optional blnSkipReadDATA As Boolean = False) As stPSSDocuments()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objSH.Length:({0}) strStationNO:({1}) strUSERID:({2}) blnSkipReadDATA:({3})", objSH.Length, strStationNO, strUSERID, blnSkipReadDATA), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intK As Int32 = 0
        Dim objPSS(0) As stPSSDocuments
        Dim objPSSFinal(0) As stPSSDocuments, intFF As Int32 = 0
        Dim intTemp As Int32 = 0
        Dim strErrLoc As String = "start"
        Dim objPSS_Read(0) As stPSSDocuments_READ
        Dim intYY As Int32 = 0, intSHID As Int32 = 0
        Dim strSQL As String = ""

        Try
            Dim intCurrentDateInNumFormat = Dates.GetCurrentDateInNumFormat - DisplayPSSDocumentsDays()
            Dim strSHIDList As String = ""
            For intK = 1 To UBound(objSH)
                If strSHIDList.Trim <> "" Then strSHIDList &= ", "
                If DBConfig.SharedSubHeader Then
                    If objSH(intK).SharedSHID > 0 Then
                        strSHIDList &= objSH(intK).SharedSHID
                    Else
                        strSHIDList &= objSH(intK).SubHdrID
                    End If
                Else
                    strSHIDList &= objSH(intK).SubHdrID
                End If
            Next
            If strSHIDList.Trim = "" Then GoTo ExitThisSub

            strSQL = "select pdmpartmm.docdesc,pdmpartmm.TTKey,pdmpartmm.CompleteTTKey " &
                        " ,pdmpartmm.revnumber,pdmpartmm.datex,pdmpartmm.docid,shmm.shid  " &
                        " from shmm,pdmpartmm where " &
                        " upper(substring(pdmpartmm.docdesc,1,3))='PSS' AND"
            If intCurrentDateInNumFormat > 0 Then
                strSQL &= " pdmpartmm.datex > " & intCurrentDateInNumFormat & " and "
            End If
            strSQL &= "  pdmpartmm.docgroupid=0 and" &
                        " pdmpartmm.DocID=shmm.MDMDocID and" &
                        " pdmpartmm.ttkey=shmm.ttkey and" &
                        " pdmpartmm.CompleteTTKey=shmm.CompleteTTKey and" &
                        " pdmpartmm.docrectype='0'  and " &
                        " shmm.mdmdocid>0 and shmm.shid in (" & strSHIDList.Trim & ") order by shmm.shid, shmm.mseq"

            strErrLoc = "Get Conn"
            Dim strOPNO As String = ""

            intK = 0
            Using connection = New Connection(gStrConnectionString)
                strErrLoc = "Get CoMM"
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim Preserve objPSS(table.Rows.Count)
                        For Each row As DataRow In table.Rows
                            intK += 1

                            objPSS(intK).OPNO = ""
                            objPSS(intK).SHDescx = ""
                            objPSS(intK).DateRead = Date.MinValue

                            strErrLoc = "Loop DocDesc"
                            objPSS(intK).DocDescx = Data.GetDataRowValue(row(0), "")
                            objPSS(intK).TTKey = Data.GetDataRowValue(row(1), 0)
                            objPSS(intK).CompleteTTKey = Data.GetDataRowValue(row(2), 0)
                            objPSS(intK).MDMDocID = Data.GetDataRowValue(Of Long)(row(5), 0)
                            objPSS(intK).DocRev = Data.GetDataRowValue(row(3), 0)
                            objPSS(intK).ReleaseDate = Data.GetDataRowValue(row(4), 0)
                            objPSS(intK).SHID = Data.GetDataRowValue(Of Long)(row(6), 0)
                        Next
                    End If
                End Using
            End Using

            strErrLoc = "Loop SH"
            'now we got all the documents, now get the list of documents by station/ operations
            intFF = 0
            For intK = 1 To UBound(objSH)  'loop every sh
                strErrLoc = "Loop SH " & intK
                intSHID = objSH(intK).SubHdrID
                If DBConfig.SharedSubHeader AndAlso objSH(intK).SharedSHID > 0 Then
                    intSHID = objSH(intK).SharedSHID
                End If
                strOPNO = objSH(intK).OPNO.Trim

                'get the docs
                For intYY = 1 To UBound(objPSS)   'loop every mdm document
                    strErrLoc = "Loop PSS " & intYY
                    If objPSS(intYY).SHID = intSHID Or objPSS(intYY).SharedSHID = intSHID Then
                        strErrLoc = "CheckMDMDocumentExistInMemory " & intYY
                        If Not CheckMDMDocumentExistInMemory(objPSSFinal, objPSS(intYY).MDMDocID, objPSS(intYY).TTKey, objPSS(intYY).CompleteTTKey, objPSS(intYY).OPNO) Then
                            'check same document is not added in multiple sh in same operation.

                            intFF += 1
                            ReDim Preserve objPSSFinal(intFF)

                            objPSSFinal(intFF) = objPSS(intYY)
                            objPSSFinal(intFF).OPNO = strOPNO  'set the operation numbers
                            objPSSFinal(intFF).SHID = objSH(intK).SubHdrID
                            objPSSFinal(intFF).SharedSHID = objSH(intK).SharedSHID
                        End If
                        Exit For
                    End If
                Next
            Next

            strErrLoc = "Refresh PSS"
            ReDim objPSS(0)

            objPSS = objPSSFinal

            If blnSkipReadDATA Then
                GoTo ExitThisSub
            End If

            If intK > 0 Then

                strErrLoc = "GetPSSDocumentsReadData: Station: " & strStationNO & ", OP: " & strUSERID

                'get the list of documents read by the operaotr
                objPSS_Read = GetPSSDocumentsReadData(strStationNO, strUSERID)

                If UBound(objPSS_Read) = 0 Then GoTo ExitThisSub

                For intK = 1 To UBound(objPSS)
                    strErrLoc = "objPSS: Station: " & intK
                    Dim intPos = GetPSSRecordPositionFromMemory(objPSS_Read, strStationNO, objPSS(intK).MDMDocID, objPSS(intK).TTKey, objPSS(intK).CompleteTTKey, strUSERID)

                    If intPos > 0 Then
                        strErrLoc = "Reset Value"
                        objPSS(intK).CheckedRead = objPSS_Read(intPos).CheckedRead
                        objPSS(intK).DateRead = objPSS_Read(intPos).DateRead
                        objPSS(intK).PassStatus = objPSS_Read(intPos).PassStatus
                    End If
                Next
            End If
ExitThisSub:
        Catch ex As Exception
            GenerateException("GetMDMProcessDocumentsList: " & strErrLoc & ", SQL: " & strSQL, ex)
        Finally
            objPSS_Read = Nothing
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objPSS)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objPSS

    End Function

    Public Function GetModelsList() As String()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrModels() As String
        ReDim ArrModels(0)

        Try
            Dim strSQL = "select distinct(model) from aml_vehicleorders"
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim ArrModels(table.Rows.Count)
                        Dim counter = 0
                        For Each reader As DataRow In table.Rows
                            counter += 1
                            Dim strTemp = Data.GetDataRowValue(reader(0), "")
                            If Trim(strTemp) = "" Then GoTo SkipThisRecord

                            ArrModels(counter) = strTemp
SkipThisRecord:
                        Next
                        ReDim ArrModels(counter)
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetModelsList", ex)
        End Try

        If UBound(ArrModels) = 0 Then
            ReDim ArrModels(5)
            ArrModels(1) = "AMV8"
            ArrModels(2) = "Cygnet"
            ArrModels(3) = "DB9 Coup"
            ArrModels(4) = "DB9 DBS"
            ArrModels(5) = "DB9 Vol."
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(ArrModels)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return ArrModels
    End Function

    Public Function GetNewORCGroupName(intCategoryID As Int16) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intCategoryID:({0})", intCategoryID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Get New and unique Question id
        Dim strSQL As String = "select CATEGORYDESC from AMLOverrideCategory where categoryid=" & intCategoryID
        Dim returnValue = ""
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of String)(strSQL, "")
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue

    End Function

    Private Function GetPartVerifyList() As DataTable

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As DataTable = Nothing
        Try
            Dim strSQL = "SELECT partnumber, ISNULL(supplier_partno, '') AS supplier_partno FROM aml_partsverify ORDER BY partnumber"
            Using connection = New Connection(gStrConnectionString)
                returnValue = connection.GetDataIntoDataTable(strSQL, Nothing)
            End Using
        Catch ex As Exception
            GenerateException(ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function GetPlantID(ByVal strWC As String, ByVal strStationNO As String) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strWC:({0}) strStationNO:({1})", strWC, strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            Dim strSQL As String = "select plantid from pdmworkstation where wc='" & strWC.Trim & "' and wscode='" & strStationNO.Trim & "'"
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue

    End Function

    Public Function GetPlatenIDForBuild(buildNo As String, stationNo As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNo:({1})", buildNo, stationNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = ""
        If (Not String.IsNullOrEmpty(buildNo) AndAlso Not String.IsNullOrEmpty(stationNo)) Then
            Using connection = New Connection(gStrConnectionString)
                Dim strSQL = "select PLATTENID from AML_LINESTATUS_BODYSHOP WITH (NOLOCK) where BUILDNO = " & Strings.ToInt32(buildNo.Trim) & " and " & Ec.GeneralFunc.GetQueryFieldCondition("STATIONNO", stationNo)
                returnValue = connection.ExecuteScalar(Of String)(strSQL, "")
            End Using
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue

    End Function

    Public Function GetPLCBuildList() As stPLCBuildList()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objPLCList(0) As stPLCBuildList

        Try
            Dim strSQL = "SELECT lineid, buildno FROM aml_plc_buildlist WHERE easeread = 0 AND lineid > 0 AND buildno <> ''"
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        Dim counter = 0
                        ReDim Preserve objPLCList(table.Rows.Count)
                        For Each reader As DataRow In table.Rows
                            counter += 1

                            objPLCList(counter).LineID = Data.GetDataRowValue(Of Int16)(reader("lineid"), 0)
                            objPLCList(counter).BuildNo = Data.GetDataRowValue(reader("buildno"), "")
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetPLCBuildList", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objPLCList)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objPLCList

    End Function

    Public Function GetPSSDocumentsReadData(strStationNO As String,
                                            strUserID As String,
                                            Optional intLineID As Int16 = 0,
                                            Optional intCheckRead As Int16 = 0) As stPSSDocuments_READ()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0}) strUserID:({1}) intLineID:({2}) intCheckRead:({3})", strStationNO, strUserID, intLineID, intCheckRead), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim objPSS_Read(0) As stPSSDocuments_READ, intK As Int32 = 0
        Dim strJoin As String = " where ", strOrderby As String = " order by userid,mdmdocid"

        Try
            strSQL = "select userid,mdmdocid,ttkey,completettkey,checkedread,pass_status,dateread,stationno,LINEID,opno from AML_PSSDocsCheck  "

            If (Connection.GetDatabaseTypeFromConnectionString(gStrConnectionString) = Connection.DatabaseTypes.SQL) Then
                strSQL += " with (NOLOCK) "
            End If

            If strStationNO.Trim <> "" Then
                strSQL &= strJoin & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)
                strJoin = " and "
            End If
            If strUserID.Trim <> "" Then
                strSQL &= strJoin.Trim & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID)
                strJoin = " and "
            End If

            Select Case intCheckRead
                Case -1   'incomplete
                    strSQL &= strJoin.Trim & " Pass_Status<>1 "
                    strJoin = " and "
                Case 1  'pass
                    strSQL &= strJoin.Trim & " Pass_Status=1 "
                    strJoin = " and "
                Case Else 'all

            End Select
            If intLineID > 0 Then
                strSQL &= strJoin.Trim & " lineid=" & intLineID
                strJoin = " and "
                strOrderby = " order by lineid,mdmdocid,dateread  "
            End If

            strSQL &= strOrderby

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim objPSS_Read(table.Rows.Count)
                        For Each row As DataRow In table.Rows
                            intK += 1

                            objPSS_Read(intK).USERID = Data.GetDataRowValue(row(0), "")
                            objPSS_Read(intK).MDMDocID = Data.GetDataRowValue(Of Long)(row(1), 0)
                            objPSS_Read(intK).TTKey = Data.GetDataRowValue(row(2), 0)
                            objPSS_Read(intK).CompleteTTKey = Data.GetDataRowValue(row(3), 0)
                            objPSS_Read(intK).CheckedRead = Data.GetDataRowValue(Of Byte)(row(4), 0)
                            objPSS_Read(intK).PassStatus = Data.GetDataRowValue(Of Int16)(row(5), -1) ' AML_PSSDocsCheck:PassStatus -> 1-Pass, -1-Incomplete
                            objPSS_Read(intK).DateRead = Data.GetDataRowValue(row(6), DateTime.MinValue)
                            objPSS_Read(intK).StationNO = Data.GetDataRowValue(row(7), "")
                            objPSS_Read(intK).LineID = Data.GetDataRowValue(Of Int16)(row(8), 0)
                            objPSS_Read(intK).OPNO = Data.GetDataRowValue(row(9), "")

                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetPSSDocumentsReadData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objPSS_Read)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objPSS_Read

    End Function

    Public Function GetPSSRecordPositionFromMemory(ByVal objPSS_Read() As stPSSDocuments_READ,
                                     ByVal strStationNO As String,
                                    ByVal lngMDMDocID As Long, ByRef intTTKey As Int32,
                                    intCompleteTTKey As Int32, strUserID As String) As Integer


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objPSS_Read.Length:({0}) strStationNO:({1}) lngMDMDocID:({2}) intCompleteTTKey:({3}) strUserID:({4})",
                                                          objPSS_Read.Length, strStationNO, lngMDMDocID, intCompleteTTKey, strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intRtnValue = 0
        Try

            For intI = 1 To UBound(objPSS_Read)
                'If objSafetyCheck(intI).DOCID = lngDocID And objSafetyCheck(intI).DocSeq = intDocSeq And objSafetyCheck(intI).CheckedRead = True Then
                If strStationNO.Trim.ToLower = objPSS_Read(intI).StationNO.Trim.ToLower AndAlso
                    strUserID.Trim.ToLower = objPSS_Read(intI).USERID.Trim.ToLower AndAlso objPSS_Read(intI).MDMDocID = lngMDMDocID _
                    And objPSS_Read(intI).TTKey = intTTKey And objPSS_Read(intI).CompleteTTKey = intCompleteTTKey Then

                    intRtnValue = intI
                    Exit For
                End If
            Next intI
        Catch ex As Exception
            GenerateException("GetPSSRecordPositionFromMemory", ex)
        End Try



#If TRACE Then
        Log.OPERATION(String.Format("Exit intTTKey:({0})", intTTKey), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return intRtnValue

    End Function

    Public Function GetQCCompleBuildBumberList(ByVal strStationNO As String, ByVal strAddlCond As String) As String()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0}) strAddlCond:({1})",
                                                          strStationNO, strAddlCond), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim ArrBuildNO(0) As String
        Dim intK = 0

        Try
            'PDMRPDATA: Result Flag: -1 -> incomplete, 1->pass, 2->fail
            Dim strSQL = "SELECT DISTINCT(engineno) FROM pdmrpdata WHERE engineno IN (" & strAddlCond & ") " &
                    " AND stationno='" & strStationNO.Trim & "' AND resultflag IN (1,2) AND engineno IS NOT NULL"
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim ArrBuildNO(table.Rows.Count)
                        For Each row As DataRow In table.Rows
                            intK += 1
                            ArrBuildNO(intK) = Data.GetDataRowValue(row(0), "")
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetQCIncompleBuildBumberList", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(ArrBuildNO)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrBuildNO
    End Function

    Public Function GetQualityChecksData(ByVal objQCS As stQualityChecksSearch,
            Optional ByVal strAddlSQL As String = "", Optional ByVal strOrderBy As String = "") As stQualityChecks()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objQCS.PartNO:({0}) objQCS.BuildNo:({1}) strAddlSQL:({2}) strOrderBy:({3})",
                                                          objQCS.PartNO, objQCS.BuildNo, strAddlSQL, strOrderBy), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'GetQualityChecksData, GetFLMData, First Line Maintenance
        'PassFail: -1 -> InComplete, 0-FAIL, 1->PASS
        'PlanType: 0-Audit Record, 1-Operator Check
        'Record type: 0-Quality checks/Remedy Plan ,1- FLM record

        Dim objQCData(0) As stQualityChecks
        Dim strTemp As String = ""

        Try
            'PDMRPDATA->Record type: 0-Quality checks/Remedy Plan ,1- FLM record
            'get all the records for the ESN and Station
            'pdmrpdata.auditornotes,pdmrpdata.reworkinsnotes,pdmrpdata.reworknotes,pdmrpdata.supervisornotes,pdmrpdata.OperatorNotes
            Dim strSQL = "select pdmrpdata.datex,pdmrpdata.plantype,pdmrpdata.stationno," &
                        " pdmrpdata.subheader,pdmrpdata.remedyplandesc," &
                        " pdmrpdata.optionno,pdmrpdata.opno,pdmrpdata.shift,pdmrpdata.FrequencyType,pdmrpitemvalues.* " &
                        " from pdmrpitemvalues,pdmrpdata WITH (NOLOCK) where " &
                        " pdmrpitemvalues.engineno=pdmrpdata.engineno and pdmrpitemvalues.stationno=pdmrpdata.stationno " &
                        " and pdmrpitemvalues.recordseq=pdmrpdata.recordseq and " &
                        Ec.GeneralFunc.GetQueryFieldCondition("pdmrpdata.engineno", objQCS.BuildNo.Trim) &
                        " and pdmrpdata.recordtype= " & objQCS.RecordType

            If Trim(objQCS.StationNO) <> "" Then
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("pdmrpdata.stationno", objQCS.StationNO)
            End If

            If objQCS.RecordSeq > 0 Then
                strSQL &= " and pdmrpdata.recordseq=" & objQCS.RecordSeq
            End If
            If objQCS.PlanType <> -1 Then
                strSQL &= " and pdmrpdata.plantype=" & objQCS.PlanType
            End If

            If Trim(objQCS.PartNO) <> "" Then
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("pdmrpdata.optionno", objQCS.PartNO)
            End If

            If Trim(objQCS.OPNO) <> "" Then
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("pdmrpdata.opno", objQCS.OPNO)
            End If

            If Trim(objQCS.SubHeader) <> "" Then
                strTemp = objQCS.SubHeader.Trim
                If DBConfig.Version7 Then
                Else
                    If Left(strTemp, 3) <> "***" Then strTemp = "***" & strTemp.Trim
                End If
                strSQL &= " and " & Trim(DBConfig.QueryFunctions.Upper) & "(subheader)='" & UCase(Strings.ReplaceSingleQuote(Trim(strTemp))) & "'"
            End If

            If objQCS.RecordType = 1 Then
                'flm check
                If objQCS.Shift > 0 Then
                    strSQL &= " and pdmrpdata.shift=" & objQCS.Shift
                End If
            End If

            If strAddlSQL.Trim <> "" Then
                strSQL &= " and " & strAddlSQL.Trim
            End If

            Select Case objQCS.PassFail
                ''PassFail: -1 -> InComplete, 0-FAIL, 1->PASS
                Case -1, 0, 1
                    strSQL &= " and pdmrpitemvalues.passfail=" & objQCS.PassFail
                Case 99
                    strSQL &= " and pdmrpitemvalues.passfail in (-1,0)"
            End Select

            '** don't change the order by seq, affects UpdateQualityChecksData SUB
            If strOrderBy.Trim <> "" Then
                'THIS ONE SHOULD BE USED ONLY FOR REPORTING PURPOSES, NOT FOR ANYTHING ELSE.
                strSQL &= strOrderBy
            Else        'default for all calls
                strSQL &= " order by pdmrpitemvalues.recordseq,pdmrpitemvalues.RemedyPlanSeq "
            End If


            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim objQCData(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1
                            objQCData(counter).PlanType = Data.GetDataRowValue(Of Int16)(row("plantype"), 0)
                            objQCData(counter).StationNO = Data.GetDataRowValue(row("stationno"), "")
                            objQCData(counter).RecordSeq = Data.GetDataRowValue(Of Int16)(row("recordseq"), 0)
                            objQCData(counter).RemedyPlanSeq = Data.GetDataRowValue(Of Int16)(row("remedyplanseq"), 0)
                            objQCData(counter).SuperVisorAction = Data.GetDataRowValue(row("supervisoraction"), "")
                            objQCData(counter).UserInput = Data.GetDataRowValue(row("userinput"), "")
                            objQCData(counter).FailureDesc = Data.GetDataRowValue(row("failuredesc"), "")
                            objQCData(counter).Spare1 = Data.GetDataRowValue(row("spare1"), "")
                            objQCData(counter).PassFail = Data.GetDataRowValue(Of Int16)(row("passfail"), -1)
                            objQCData(counter).RemedyPlanDesc = Data.GetDataRowValue(row("RemedyPlanDesc"), "")
                            objQCData(counter).OPNO = Trim(Data.GetDataRowValue(row("opno"), ""))
                            objQCData(counter).SubHeaderDesc = Trim(Data.GetDataRowValue(row("subheader"), ""))
                            objQCData(counter).CommentX = Trim(Data.GetDataRowValue(row("commentx"), ""))
                            objQCData(counter).CheckMethod = Trim(Data.GetDataRowValue(row("checkmethod"), ""))
                            objQCData(counter).LowerValue = Data.GetDataRowValue(Of Single)(row("LowerValue"), 0)
                            objQCData(counter).UpperValue = Data.GetDataRowValue(Of Single)(row("UpperValue"), 0)
                            objQCData(counter).UserInputFailed = Trim(Data.GetDataRowValue(row("UserInputFailed"), ""))
                            objQCData(counter).UserInput_Time_Stamp = Data.GetDataRowValue(row("UserInput_Time_Stamp"), DateTime.MinValue)
                            objQCData(counter).InvalidEntry = False
                            objQCData(counter).Datex = Data.GetDataRowValue(row("datex"), DateTime.MinValue)
                            objQCData(counter).OperatorID = Trim(Data.GetDataRowValue(row("operatorid"), ""))
                            objQCData(counter).Shift = Data.GetDataRowValue(Of Byte)(row("Shift"), 0)
                            objQCData(counter).FrequencyType = Data.GetDataRowValue(Of Int16)(row("FrequencyType"), 0)
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetQualityChecksData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objQCData)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objQCData
    End Function

    Public Sub GetQualityChecksResultsSummary(ByVal strBuildNO As String,
                    ByVal strStationNO As String, ByVal strRecordSeqList As String,
                    ByRef intPass As Integer, ByRef intFail As Integer, ByRef intInComplete As Integer)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strRecordSeqList:({2})", strBuildNO, strStationNO, strRecordSeqList), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'get the QC results summary for a list of qc records (based on sh selection criteria)

        Dim strSQL As String = "", intResultFlag As Int16 = 0
        intPass = 0 : intFail = 0 : intInComplete = 0

        Try
            strSQL = "select resultflag from pdmrpdata WITH (NOLOCK) where " & Ec.GeneralFunc.GetQueryFieldCondition("engineno", strBuildNO) & " and " &
                                    Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) & " and " &
                                    " recordtype=0 AND resultflag IS NOT NULL "
            If strRecordSeqList.Trim <> "" Then
                strSQL &= " and recordseq in (" & strRecordSeqList & ")"
            End If

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        intInComplete = table.AsEnumerable().Count(Function(x) Data.GetDataRowValue(x("resultflag"), 0) = -1)
                        intPass = table.AsEnumerable().Count(Function(x) Data.GetDataRowValue(x("resultflag"), 0) = 1)
                        intFail = table.AsEnumerable().Count(Function(x) Data.GetDataRowValue(x("resultflag"), 0) = 2)
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetQualityChecksResultsSummary", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit InComplete:({0}) Pass:({1}) Fail:({2})", intInComplete, intPass, intFail), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function GetRemedyPlanItemTemplate(ByVal lngDocID As Long, ByVal intTTKey As Int32,
            Optional ByVal intRPType As Int16 = -1,
            Optional ByVal strAddlSQL As String = "",
            Optional ByVal intDocSeq As Int32 = 0,
            Optional ByVal strDocRectype As String = "") As stRemedyPlanItemTemplateAndValues()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intTTKey:({0}) intRPType:({1}) strAddlSQL:({2}) intDocSeq:({3}) strDocRectype:({4})",
                                                          intTTKey, intRPType, strAddlSQL, intDocSeq, strDocRectype), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objRPITAV(0) As stRemedyPlanItemTemplateAndValues
        Dim strSQL As String = ""
        Dim strOrderBy As String = " order by pdmremedyplan.docid,pdmremedyplan.seq "

        Try
            strSQL = "select pdmremedyplan.* from pdmremedyplan,pdmpartmm with (NOLOCK) where  " &
                    " pdmremedyplan.docid=pdmpartmm.docid and " &
                    " pdmremedyplan.docseq=pdmpartmm.docseq "

            If lngDocID > 0 Then
                strSQL &= " and pdmpartmm.docid=" & lngDocID &
                          " and pdmpartmm.ttkey=" & intTTKey
            End If
            If intDocSeq > 0 Then
                strSQL &= " and pdmremedyplan.docseq=" & intDocSeq
            End If
            If strDocRectype.Trim <> "" Then
                strSQL &= " and pdmremedyplan.docrectype='" & strDocRectype.Trim & "'"
            End If

            If intRPType <> -1 Then
                strSQL &= " and pdmremedyplan.plantype=" & intRPType
            End If

            If Trim(strAddlSQL) <> "" Then
                strSQL &= " and (" & Trim(strAddlSQL) & ")"
                'DONT CHANGE THE ORDER BY CLAUSE
                strOrderBy = " order by pdmremedyplan.docid,pdmremedyplan.docrectype,pdmremedyplan.docseq,pdmremedyplan.seq "
            End If
            strSQL &= strOrderBy

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objRPITAV(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1

                            objRPITAV(counter).Spare1 = ""     'used in import process
                            objRPITAV(counter).DOCID = Data.GetDataRowValue(Of Long)(row("docid"), 0)
                            objRPITAV(counter).DocRecType = Data.GetDataRowValue(row("docrectype"), "0").Trim()
                            objRPITAV(counter).DocSeq = Data.GetDataRowValue(row("docseq"), 0)
                            objRPITAV(counter).CheckMethod = Trim(Data.GetDataRowValue(row("checkmethod"), ""))
                            objRPITAV(counter).Comment = Trim(Data.GetDataRowValue(row("comments"), ""))
                            objRPITAV(counter).FailureDesc = Trim(Data.GetDataRowValue(row("failuredesc"), ""))
                            objRPITAV(counter).InputReqd = Trim(Data.GetDataRowValue(row("inputrequired"), ""))
                            objRPITAV(counter).LowerValue = Data.GetDataRowValue(Of Single)(row("lowervalue"), 0)
                            objRPITAV(counter).OrigOpNo = Trim(Data.GetDataRowValue(row("origopno"), ""))
                            'PlanType: 0-Audit Record, 1-Operator Check
                            objRPITAV(counter).PlanType = Data.GetDataRowValue(row("PlanType"), 0)
                            objRPITAV(counter).Seq = Data.GetDataRowValue(row("seq"), 0)
                            objRPITAV(counter).Severity = Trim(Data.GetDataRowValue(row("severity"), ""))
                            objRPITAV(counter).StopBuild = Trim(Data.GetDataRowValue(row("stopbuild"), ""))
                            objRPITAV(counter).Unit = Trim(Data.GetDataRowValue(row("unitno"), ""))
                            objRPITAV(counter).UpperValue = Data.GetDataRowValue(Of Single)(row("uppervalue"), 0)
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetRemedyPlanItemTemplate", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objRPITAV)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objRPITAV

    End Function

    Public Function GetSafetyMainImageSQL(ByVal strStationNO As String, ByVal strWorkCenter As String,
            ByVal strOption As String, Optional ByVal strRecType As String = "0",
            Optional ByVal strOrderBY As String = "") As String

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter strStationNO:({0}) strWorkCenter:({1}) strOption:({2}) strRecType:({3}) strOrderBY:({4})",
                                                           strStationNO, strWorkCenter, strOption, strRecType, strOrderBY), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        Dim strFields As String = "pdmworkstationmm1.wsid,pdmworkstationmm1.completettkey"
        If strOption.Trim = "recordcount" Then
            strFields = "count(pdmworkstationmm1.wsid)"
        End If

        strSQL = "select " & strFields.Trim & " from pdmworkstationmm1, pdmworkstation" &
                " where pdmworkstationmm1.wsid = pdmworkstation.wsid " &
                " and " & Ec.GeneralFunc.GetQueryFieldCondition("pdmworkstation.wscode", strStationNO) &
                " and " & Ec.GeneralFunc.GetQueryFieldCondition("pdmworkstation.wc", strWorkCenter) &
                " and pdmworkstationmm1.docstatus=1 and pdmworkstationmm1.docdesc like '%SAFETY_MAN%' " &
                " and pdmworkstationmm1.completettkey > 0 and pdmworkstationmm1.docrectype='" & strRecType.Trim & "'"
        If strOrderBY.Trim <> "" Then
            strSQL &= " " & strOrderBY
        End If

#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strSQL
    End Function

    Public Sub GetSafetyManDocumentInfo(ByVal strStationNO As String, ByVal strWorkCenter As String, ByRef intWSID As Int32, ByRef intCompleteTTKey As Int32)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0}) strWorkCenter:({1})", strStationNO, strWorkCenter), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strRecType As String = "0"
        Dim strOrderBY As String = ""
        Dim blnGetHistoryDocument As Boolean = False

        intWSID = 0 : intCompleteTTKey = 0

        Try
            Using connection = New Connection(gStrConnectionString)

GetDocumentAgain:

                Dim strSQL = GetSafetyMainImageSQL(strStationNO, strWorkCenter, "", strRecType, strOrderBY)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim reader As DataRow = table.Rows(0)
                        intWSID = Data.GetDataRowValue(reader("wsid"), 0)
                        intCompleteTTKey = Data.GetDataRowValue(reader("completettkey"), 0)
                    End If
                End Using

                If intWSID = 0 AndAlso blnGetHistoryDocument = False Then
                    'the safety doc is in unreleased state. Check for history and 
                    'if history exists, show the history doc instead
                    blnGetHistoryDocument = True
                    strRecType = "2"
                    strOrderBY = " order by docid, docseq desc"
                    GoTo GetDocumentAgain
                End If
            End Using

ExitThisSub:

        Catch ex As Exception
            GenerateException("GetSafetyManDocumentInfo", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit intWSID:({0}) intCompleteTTKey:({1})", intWSID, intCompleteTTKey), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function GetSafetyManImageDocument(ByVal intWSID As Int32, ByVal intCompleteTTKey As Int32, ByVal strFileName As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intWSID:({0}) intCompleteTTKey:({1}) strFileName:({2})",
                                                          intWSID, intCompleteTTKey, strFileName), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objParams As New Text.stTextParams
        Try
            Ec.Text.ClearTextParamsObject(objParams)
            objParams.Param1 = intWSID.ToString()
            objParams.Param2 = intCompleteTTKey.ToString()

            If Not Ec.Text.ReadTextFromDB(14, objParams, strFileName) Then
                'failed to create the document, exit out
                strFileName = ""
            End If

ExitThisSub:
        Catch ex As Exception
            GenerateException("GetSafetyManImageDocument", ex)
        Finally
            objParams = Nothing
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", strFileName), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strFileName
    End Function

    Public Function GetSerialRecordPosition(strBuildNo As String, strStationNO As String, strPartNO As String, intPartVerifyMSeq As Integer,
                                            Optional blnCheckForIncompleteStatus As Boolean = True) As Integer


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) strPartNO:({2}) intPartVerifyMSeq:({3}) blnCheckForIncompleteStatus:({4})",
                                                          strBuildNo, strStationNO, strPartNO, intPartVerifyMSeq, blnCheckForIncompleteStatus), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select serialseq from AML_BuildPartsSerials where buildnumber=" & strBuildNo.Trim &
                " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)
        Dim strJoin As String = "  "
        Dim strTemp2 As String = ""

        If intPartVerifyMSeq > 0 Then
            strTemp2 = " mseq=" & intPartVerifyMSeq
            strJoin = " or "
        End If

        If strPartNO.Trim <> "" Then
            strTemp2 &= strJoin & Ec.GeneralFunc.GetQueryFieldCondition("partnumber", strPartNO)
        End If

        If strTemp2.Trim <> "" Then
            strSQL &= " and (" & strTemp2 & ") "
        End If
        If blnCheckForIncompleteStatus Then
            strSQL &= " and serialstatus=-1 "
        End If

        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetStationLineID(ByVal stationNo As String) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0})", stationNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(GetStationLineIDSQL(stationNo), 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function GetStationLineIDSQL(ByVal stationNo As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0})", stationNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = "select lineid from stations where " & Ec.GeneralFunc.GetQueryFieldCondition("absnno", stationNo)

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result
    End Function

    Public Function GetStationLineDescription(ByVal connection As Connection, ParamArray ByVal stationNoList As String()) As Dictionary(Of String, String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0})", UBound(stationNoList)), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim stations = stationNoList.Select(Function(x) String.Format("'{0}'", x)).ToArray()
        Dim strSQL As String = String.Format("SELECT DISTINCT s.absnno, l.description FROM stations s INNER JOIN lines l ON s.lineid = l.lineid WHERE s.absnno IN ({0})", String.Join(",", stations))
        Dim result = New Dictionary(Of String, String)()
        Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
            If (Not IsNothing(table)) Then
                For Each row As DataRow In table.Rows
                    result.Add(row("absnno").ToString().Trim(), row("description").ToString().Trim())
                Next
            End If
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result

    End Function

    Public Function GetStationSignoffStatus(strBuildNO As String, strStationNO As String, intOperatorPosition As Integer) As Int16

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) intOperatorPosition:({2})",
                                                          strBuildNO, strStationNO, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As Int16 = 0
        Try
            Dim buildNo = 0
            If (Integer.TryParse(strBuildNO, buildNo)) Then
                'AML_StationSignoffStatus: SignoffStatus: -1 incomplete, 1- complete, 0-Default value
                Dim strSQL = "select signoffstatus from AML_StationSignoffStatus where BuildNumber=" & buildNo &
                        " and " & Ec.GeneralFunc.GetQueryFieldCondition("StationNO", strStationNO) &
                        " and OPERATORPOSITION=" & intOperatorPosition.ToString
                Using connection = New Connection(gStrConnectionString)
                    returnValue = connection.ExecuteScalar(Of Int16)(strSQL, 0)
                End Using
            End If
        Catch ex As Exception
            GenerateException("GetStationSignoffStatus", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetSubHeaderID(ByVal intID As Int32, ByVal strRectype As String, ByVal intSeq As Int16, ByVal strOPNO As String, ByVal strSubHeader As String) As Int32

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRectype:({1}) intSeq:({2}) strOPNO:({3}) strSubHeader:({4})",
                                                          intID, strRectype, intSeq, strOPNO, strSubHeader), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strAddlFields As String = ""
        If DBConfig.SharedSubHeader Then strAddlFields = ",sharedshid"
        Dim strSQL As String = "select shid" & strAddlFields & " from subhdr where id=" & intID & " and rectype='" & strRectype.Trim & "' and seq=" & intSeq & " and " & Ec.GeneralFunc.GetQueryFieldCondition("opno", strOPNO) & " and " & Ec.GeneralFunc.GetQueryFieldCondition("descx", strSubHeader)
        Dim intSHID As Int32 = 0, intSharedSHID As Int32 = 0
        Try
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim reader As DataRow = table.Rows(0)
                        intSHID = Data.GetDataRowValue(reader(0), 0)
                        If DBConfig.SharedSubHeader Then
                            intSharedSHID = Data.GetDataRowValue(reader(1), 0)

                            If intSharedSHID > 0 Then intSHID = intSharedSHID
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetSubHeaderID", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", intSHID), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return intSHID
    End Function

    Public Function GetSubeaderListByPartNo(ByVal buildNo As String, ByVal area As String) As Parts.stSubHeader()
        Dim mmTable As String = "ophmm"
        Dim intLinkedModel = 0
        Dim intBuildNoModelNumber = GetBuildNoModelNumber(buildNo, True, intLinkedModel)
        'READ ME: strStation is not used and blnAvoidOPNO must be true - Get the Plan
        Dim partSearchResults = GetEASEPartNo(area, intBuildNoModelNumber, "", 0, True, , intLinkedModel)
        If (UBound(partSearchResults) = 0) Then Return Nothing

        Dim partSearchResult = partSearchResults.Skip(1).First()
        Return _
            GetSubheaderList(partSearchResult.ID, partSearchResult.Rectype, partSearchResult.Seq, "", buildNo, "", True, , True,
                             , blnCreateQCFromWIPScreen:=True)
    End Function

    Public Function GetSubheaderList(ByVal lngID As Long,
                ByVal strRectype As String,
                ByVal intSeq As Integer,
                ByVal strOPNO As String,
                ByVal strBuildNO As String,
                ByVal strStationNO As String,
                Optional ByVal blnGetSHForAllStations As Boolean = False,
                Optional ByRef strQCRecordSeqList As String = "",
                Optional ByVal blnSkipQCData As Boolean = False,
                Optional ByVal blnOVERRIDEFilters As Boolean = False,
                Optional ByRef strTempLog As String = "",
                Optional ByVal blnGetLogData As Boolean = False,
                Optional ByVal blnCreateQCFromWIPScreen As Boolean = False,
                Optional ByVal strOpListSQLCond As String = "",
                Optional ByVal blnSkipDefaultSubHeader As Boolean = False,
                Optional ByVal strCustomOrderBy As String = "",
                Optional intCustomBuildNoModelNumber As Int16 = 0,
                Optional strPlatenID As String = "") As Parts.stSubHeader()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter lngID:({0}) strRectype:({1}) intSeq:({2}) strOPNO:({3}) strBuildNO:({4}) strStationNO:({5}) " &
                                                          "blnGetSHForAllStations:({6}) blnSkipQCData:({7}) blnOVERRIDEFilters:({8}) " &
                                                          "blnGetLogData:({9}) blnCreateQCFromWIPScreen:({10}) strOpListSQLCond:({11}) " &
                                                          "blnSKipDefaultSubHeader:({12}) strCustomOrderBy:({13}) intCustomBuildNoModelNumber:({14}) strPlatenID:({15})",
                                                          lngID, strRectype, intSeq, strOPNO, strBuildNO, strStationNO,
                                                          blnGetSHForAllStations, blnSkipQCData, blnOVERRIDEFilters,
                                                          blnGetLogData, blnCreateQCFromWIPScreen, strOpListSQLCond,
                                                          blnSkipDefaultSubHeader, strCustomOrderBy, intCustomBuildNoModelNumber, strPlatenID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        strQCRecordSeqList = ""

        Dim objSH(0) As Parts.stSubHeader
        Dim objReturnSH(0) As Parts.stSubHeader
        Dim intCount = 0
        Dim strErrorLocation As String = "Start" 'strGenericOptions As String = "", 
        Dim objSHOALL(0) As stSubHeaderOptions
        Dim objSHOptions(0) As stSubHeaderOptions
        Dim blnMatch As Boolean = False
        Dim objQCsData(0) As stQualityChecks
        Dim strSQLCond As String = ""
        Dim blnProcessMultipleStations As Boolean = False
        Dim strSharedSHIDList As String = ""

        Try
            Dim intLineID = GetStationLineID(strStationNO)

            If blnCreateQCFromWIPScreen Then
                strOPNO = "" 'need to get all operations.
                strStationNO = ""

                'strSQLCond = GetAutoRefreshingStationsListForSQL()  'operations conditions
                strSQLCond = strOpListSQLCond
                blnProcessMultipleStations = True

                GoTo SkipBuildOpQuery
            End If

            If blnGetSHForAllStations Then

                If strOPNO.Trim <> "" Then
                    Dim intOperatorCount = GetOperatorCountForStation(intLineID, strStationNO)
                    'blnGetSHForAllStations is used only on aml1 page to get the count of
                    'failed,incomplete,pass qc records.DEFAULT IS FALSE
                    Dim strTemp2 = "'" & UCase(Trim(strOPNO)) & "'"
                    For intZZ = 1 To intOperatorCount 'GetMaxNumberOfOperatorPositions()
                        strTemp2 &= ",'" & UCase(Trim(strOPNO)) & "-" & Trim(intZZ.ToString) & "'"
                    Next

                    If strTemp2.Trim <> "" Then
                        strSQLCond = " opno in (" & strTemp2 & ")"
                    End If

                End If

            End If
SkipBuildOpQuery:

            ' HACK(crhodes)
            ' Maybe evaluate PlatenID here and remove if 0 or 99

            objSH = Ec.Parts.GetOpSubheaders(lngID, strRectype, intSeq, strOPNO, , True, strSQLCond, blnSkipDefaultSubHeader, strCustomOrderBy, strPlatenID)

            If UBound(objSH) = 0 Then
                GoTo ExitThisFunction
            End If



            If blnOVERRIDEFilters Then
                objReturnSH = objSH
                GoTo ExitThisFunction 'no filters, testing only, 
            End If


            'EW-708: AML: ViewEASE changes list
            If DBConfig.SharedSubHeader Then
                strSharedSHIDList = ""

                For intK = 1 To UBound(objSH)
                    If objSH(intK).SharedSHID > 0 Then
                        If strSharedSHIDList.Trim <> "" Then strSharedSHIDList &= ", "
                        strSharedSHIDList &= objSH(intK).SharedSHID
                    End If
                Next
            End If

            If intCustomBuildNoModelNumber > 0 AndAlso strBuildNO.Trim = "" Then
                'used only in PSS Documents Reports only
                objReturnSH = objSH
                GoTo StartProcessingQualityChecks
            End If

            ' HACK(crhodes)
            ' I think the platen filter should go here.

            'get the list of subheader options configured in Process Plan
            objSHOALL = GetSubHeaderOptionsList(lngID, strRectype, intSeq, strOPNO, strSQLCond, strSharedSHIDList)

            If UBound(objSHOALL) = 0 Then
                'no options configured, assume all subheaders are ok
                objReturnSH = objSH
                GoTo StartProcessingQualityChecks
            End If

            If strBuildNO.Trim = "" Then
                ReDim objReturnSH(0)  'why return subheader, if the build# blank.
                GoTo ExitThisFunction
            End If

            Dim objVehicleOrder = GetVehicleOrderDetails(strBuildNO)     'get the vehicle details
            Dim intBuildNOModel1 = objVehicleOrder.Model1.Value
            Dim objAMLModels = GetAMLModelsList()   'get the list of aml models configured in ClientEditor
            Dim strBuildNoModel = objAMLModels.First(Function(x) x.ModelNumber = intBuildNOModel1).ModelCode.Trim()
            Dim vehicleOptions = GetVehicleOptions(strBuildNO)  'holds both specific, other options for the engine

StartProcessSHList:

            For intK = 1 To UBound(objSH)

                Dim intSHID = objSH(intK).SubHdrID

                If DBConfig.SharedSubHeader AndAlso objSH(intK).SharedSHID > 0 Then
                    intSHID = objSH(intK).SharedSHID
                End If

                ''EW-855: VH 5 Body Shop and MA1 station refresh
                ' TODO(crhodes) EW-1166
                ' Leaving this code in but commented out until we get final acceptance from Leon.  
                ' Fix is above in passing PlatenID into GetOpSubHeaders(...)

                'If strPlatenID.Trim <> "" Then
                '    If objSH(intK).PlatenID.Trim.ToLower <> strPlatenID.Trim.ToLower Then
                '        GoTo SkipThisSubheader
                '    End If
                'End If

                Dim blnOptionFound = False
                Dim blnSubModelMatchFound = False

                objSHOptions = GetSubHeaderOptionsFromMemory(objSHOALL, intSHID)     'valadd fields holds the subheader id for AML

                If UBound(objSHOptions) = 0 Then
                    'no options for the subheader, so add it
                    blnOptionFound = True
                    GoTo SkipProcessingSpecificOptions
                End If

                'loop thru each subheader options for the subheader and look for the matching submodel record
                Dim blnAllOptionRecordFound = False 'the sub-model has 'all' selected in process plan (i.e applicable to all sub-models)
                Dim intPosX = 0

                For intYY = 1 To UBound(objSHOptions)

                    If objSHOptions(intYY).SubModel.Trim <> "" AndAlso strBuildNoModel.Trim.ToLower = objSHOptions(intYY).SubModel.Trim.ToLower Then
                        intPosX = intYY
                        blnSubModelMatchFound = True
                        Exit For  'if matching sub-model found, don't use 'all' option record
                    End If

                    If objSHOptions(intYY).SubModel.Trim = "" Then
                        blnAllOptionRecordFound = True

                        'no EXIT FOR HERE, FINDING MATCHING SUB-MODEL IS CRUICIAL
                    End If

                    If blnAllOptionRecordFound = False AndAlso (strBuildNoModel.Trim = objSHOptions(intYY).SubModel) = False Then
                        Dim err As String = String.Format("<br>[vh5]An invalid option code was selected from Process Plan. Expected ModelNo '{0}', actual ModelNo '{1}'.[vh5]<br>", strBuildNoModel, objSHOptions(intYY).SubModel)
                        strTempLog &= err
                    End If

                Next

                If blnSubModelMatchFound = False And blnAllOptionRecordFound = False Then
                    GoTo SkipThisSubheader
                End If

                'OK ** NOW PROCESS FOR THE MATCHING SUB-MODEL OPTION RECORDS OR 'ALL' OPTION RECORDS IN THE LOOP

                For intYY = 1 To UBound(objSHOptions)

                    Dim blnSHMatch = True

                    If blnSubModelMatchFound Then
                        'matching sub-model found, check the other criterias
                        intYY = intPosX
                    Else
                        'process only 'ALL' sub-models

                        If objSHOptions(intYY).SubModel.Trim <> "" Then     'different sub-model, skip them
                            GoTo SkipThisSHOptionRecord
                        End If

                    End If

                    'sub-header model record found. the position is in intyy

                    Dim strSHBody = objSHOptions(intYY).Body.Trim
                    Dim strSHGear = objSHOptions(intYY).GearBox.Trim
                    Dim strSHDrive = objSHOptions(intYY).Drive.Trim
                    Dim strSHTerritory = objSHOptions(intYY).Territory.Trim
                    Dim strSHPerformance = objSHOptions(intYY).Performance.Trim
                    Dim strSHYear = objSHOptions(intYY).YearX.Trim

                    Dim strSpecific1 = objSHOptions(intYY).SpecificOptions1.Trim
                    Dim strSpecific2 = objSHOptions(intYY).SpecificOptions2.Trim
                    Dim strSpecific3 = objSHOptions(intYY).SpecificOptions3.Trim

                    If strSHBody.Trim <> "" Then
                        If strSHBody.Trim.ToLower <> objVehicleOrder.Body.Trim.ToLower Then
                            blnSHMatch = False
                        End If
                    End If

                    If strSHGear.Trim <> "" Then
                        If strSHGear.Trim.ToLower <> objVehicleOrder.GearBox1.Trim.ToLower Then
                            blnSHMatch = False
                        End If
                    End If

                    If strSHDrive.Trim <> "" Then
                        If strSHDrive.Trim.ToLower <> objVehicleOrder.Drive1.Trim.ToLower Then
                            blnSHMatch = False
                        End If
                    End If

                    If strSHTerritory.Trim <> "" Then
                        If strSHTerritory.Trim.ToLower <> objVehicleOrder.Territory1.Trim.ToLower Then
                            blnSHMatch = False
                        End If
                    End If

                    If strSHPerformance.Trim <> "" Then
                        If strSHPerformance.Trim.ToLower <> objVehicleOrder.Performance1.Trim.ToLower Then
                            blnSHMatch = False
                        End If
                    End If

                    If strSHYear.Trim <> "" Then
                        If strSHYear.ToInt16() <> objVehicleOrder.YearX Then
                            blnSHMatch = False
                        End If
                    End If

                    If blnSHMatch = False Then
                        If blnSubModelMatchFound Then
                            'ok no match found, skip this sh
                            GoTo SkipThisSubheader
                        Else
                            GoTo SkipThisSHOptionRecord
                        End If
                    End If

                    'no specific options configured for the submodel
                    If strSpecific1.Trim = "" And strSpecific2.Trim = "" And strSpecific3.Trim = "" Then
                        blnOptionFound = True
                        Exit For
                    End If

                    'check the specific option match
                    If CheckForSHSpecificOptions(vehicleOptions, strSpecific1, strSpecific2, strSpecific3) Then
                        blnOptionFound = True
                        Exit For
                    End If

                    If blnSubModelMatchFound Then
                        'looking for exact sub-model filtering criteria
                        Exit For
                    End If
SkipThisSHOptionRecord:

                Next intYY

SkipProcessingSpecificOptions:

                If blnOptionFound Then
                    'add the subheader in the array
                    intCount += 1
                    ReDim Preserve objReturnSH(intCount)
                    objReturnSH(intCount) = objSH(intK)

                End If

SkipThisSubheader:

            Next intK

StartProcessingQualityChecks:

            If blnSkipQCData Then GoTo ExitThisFunction

            Dim objQCS As New stQualityChecksSearch
            ClearQualityChecksSearchObject(objQCS)
            objQCS.BuildNo = strBuildNO
            objQCS.StationNO = strStationNO
            objQCS.PartNO = Ec.Parts.GetPartNO(lngID)
            If Not blnGetSHForAllStations Then
                objQCS.OPNO = strOPNO
            End If
            objQCS.PassFail = -22       '-22 to get all records
            objQCS.RecordType = 0       'Record type: 0-Quality checks/Remedy Plan ,1- FLM record
            objQCsData = GetQualityChecksData(objQCS, strSQLCond)

            If UBound(objQCsData) = 0 Then GoTo ExitThisFunction 'no quality checks, exit out.

            For intK = 1 To UBound(objReturnSH)
                Dim intQCStatus = 1

                For intYY = 1 To UBound(objQCsData)

                    If blnProcessMultipleStations Then
                        'creating QC for multiple stations from WIP screen

                        blnMatch = (objReturnSH(intK).SubHdrDesc.Trim.ToLower = objQCsData(intYY).SubHeaderDesc.Trim.ToLower And objReturnSH(intK).OPNO.Trim.ToLower = objQCsData(intYY).OPNO.Trim.ToLower)
                    Else
                        If blnGetSHForAllStations Then
                            'used only when getting the qc summary data for aml1 page. 
                            blnMatch = (objReturnSH(intK).SubHdrDesc.Trim.ToLower = objQCsData(intYY).SubHeaderDesc.Trim.ToLower And objReturnSH(intK).OPNO.Trim.ToLower = objQCsData(intYY).OPNO.Trim.ToLower)
                        Else        'default option
                            blnMatch = objReturnSH(intK).SubHdrDesc.Trim.ToLower = objQCsData(intYY).SubHeaderDesc.Trim.ToLower
                        End If
                    End If

                    If blnMatch Then

                        objQCsData(intYY).CommentX = "matchingshfound"      'used only to get the qc summary data (pass, fail, incomplete in aml1.aspx)

                        'a subheader record can have more than one qc
                        Select Case objQCsData(intYY).PassFail     '.PassFail :-1 -> InComplete, 0-FAIL, 1->PASS
                            Case -1
                                intQCStatus = -1  'qc incomplete
                            Case 0
                                intQCStatus = 0     'one record fail, so total fail
                                Exit For
                            Case 1
                        End Select

                    End If

                Next intYY

                objReturnSH(intK).QualityChecksStatus = intQCStatus
            Next intK

            If blnGetSHForAllStations Then
                'used only to get the QC results summary in aml1.aspx (not the default option)

                strQCRecordSeqList = ""

                For intYY = 1 To UBound(objQCsData)
                    If objQCsData(intYY).CommentX = "matchingshfound" Then
                        If strQCRecordSeqList.Trim <> "" Then strQCRecordSeqList &= ","
                        strQCRecordSeqList &= objQCsData(intYY).RecordSeq
                    End If
                Next

            End If

ExitThisFunction:
        Catch ex As Exception
            Log.Error(ex, LOG_APPNAME)
        Finally
            objSHOALL = Nothing
            objSHOptions = Nothing
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0}) strQCRecordSeqList:({1})", UBound(objReturnSH), strQCRecordSeqList), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objReturnSH
    End Function

    Public Function GetSubHeaderOptionsList(ByVal lngID As Long, ByVal strRectype As String, ByVal intSeq As Integer,
                                            ByVal strOPNO As String, Optional ByVal strSQLCond As String = "",
                                            Optional strSharedSHIDList As String = "") As stSubHeaderOptions()
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strRectype:({0}) intSeq:({1}) strOPNO:({2}) strSQLCond:({3}) strSharedSHIDList:({4})", strRectype, intSeq, strOPNO, strSQLCond, strSharedSHIDList), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'get the list of the subheader options configured in the process plan

        Dim objSHO(0) As stSubHeaderOptions
        Dim strOrderby As String = " order by shid,mseq"
        Try

            Dim strSQL = String.Format("SELECT * FROM aml_shoptions WHERE id = {0} AND rectype = '{1}' AND seq = {2}", lngID, strRectype.Trim, intSeq)
            If strSQLCond.Trim <> "" Then
                strSQL &= " AND " & strSQLCond
                'processing multiple stations.
                strOrderby = " ORDER BY id,rectype,seq,opno,shid,mseq"
            Else
                strSQL &= " AND opno='" & Trim(strOPNO) & "'"
            End If

            If strSharedSHIDList.Trim <> "" Then
                'EW-708: AML: ViewEASE changes list/ Shared Subheaders
                strSQL &= " UNION ALL "
                strSQL &= "select * from AML_SHOptions where shid in (" & strSharedSHIDList & ") "
            End If

            strSQL &= strOrderby ' " order by shid,mseq"

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim Preserve objSHO(table.Rows.Count)
                        Dim counter = 0
                        For Each reader As DataRow In table.Rows
                            counter += 1

                            objSHO(counter).ID = Data.GetDataRowValue(reader("id"), 0)
                            objSHO(counter).RecType = Data.GetDataRowValue(reader("rectype"), "")
                            objSHO(counter).Seq = Data.GetDataRowValue(Of Int16)(reader("seq"), 0)
                            objSHO(counter).OPNO = Data.GetDataRowValue(reader("opno"), "")
                            objSHO(counter).SHID = Data.GetDataRowValue(reader("shid"), 0)
                            objSHO(counter).Mseq = Data.GetDataRowValue(Of Int16)(reader("mseq"), 0)
                            objSHO(counter).SubModel = Data.GetDataRowValue(reader("submodel"), "")
                            objSHO(counter).Body = Data.GetDataRowValue(reader("body"), "")
                            objSHO(counter).GearBox = Data.GetDataRowValue(reader("gearbox"), "")
                            objSHO(counter).Performance = Data.GetDataRowValue(reader("performance"), "")
                            objSHO(counter).YearX = Data.GetDataRowValue(reader("yearx"), "")
                            objSHO(counter).Territory = Data.GetDataRowValue(reader("territory"), "")
                            objSHO(counter).Drive = Data.GetDataRowValue(reader("drive"), "")
                            objSHO(counter).SpecificOptions1 = Data.GetDataRowValue(reader("specoptions1"), "")
                            objSHO(counter).SpecificOptions2 = Data.GetDataRowValue(reader("specoptions2"), "")
                            objSHO(counter).SpecificOptions3 = Data.GetDataRowValue(reader("specoptions3"), "")
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetSubHeaderOptionsList", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objSHO)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objSHO
    End Function

    Public Function GetSubHeaderReferenceNumber(ByVal intPartID As Int32, ByVal strRectype As String,
            ByVal intSeq As Int16, ByVal strOPNO As String, ByVal strSubHeader As String, Optional ByVal lngSHID As Long = 0) As String


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intPartID:({0}) strRectype:({1}) intSeq:({2}) strOPNO:({3}) strSubHeader:({4})", intPartID, strRectype, intSeq, strOPNO, strSubHeader), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim returnValue = ""
        Dim intPos = 0
        Using connection = New Connection(gStrConnectionString)
            If DBConfig.Version7 Then  '** Version 7 Changes **  - DONE
                strSQL = "select model2 from subhdr where shid=" & lngSHID.ToString

                Dim result = connection.ExecuteScalar(Of String)(strSQL, "")
                If (Not IsNothing(result) AndAlso Not IsDBNull(result)) Then
                    returnValue = result
                    intPos = InStr(returnValue, ".")
                    If intPos > 0 Then
                        returnValue = Left(returnValue, intPos - 1)
                    End If
                End If
            Else
                If Left(strSubHeader.Trim, 3) <> "***" Then
                    strSubHeader = "***" & strSubHeader.Trim
                End If
                strSQL = "select rhlhflag from element where left(eldesc,3)='***'" &
                                " and id=" & intPartID.ToString &
                                " and rectype='" & strRectype.Trim & "' " &
                                " and seq=" & intSeq.ToString &
                                " and opno='" & strOPNO.Trim & "' " &
                                " and eldesc='" & Strings.ReplaceSingleQuote(strSubHeader.Trim) & "' "

                Dim result = connection.ExecuteScalar(Of String)(strSQL, "")
                If (Not IsNothing(result) AndAlso Not IsDBNull(result)) Then
                    returnValue = result
                End If
            End If
        End Using
#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Private Function GetTreeViewData_Oracle() As DataSet

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objDS As New DataSet("Treeview")
        Dim objDataColumn() As DataColumn

        Try
            Using connection = New Connection(gStrConnectionString)
                Dim strTableName = "lines"
                Dim strSQL = "select pdmworkcenter.wc wc,pdmworkcenter.descx descx from lines,pdmworkcenter where pdmworkcenter.wc=lines.description order by lines.lineid"

                Dim adapter = connection.GetAdapter(connection.GetCommand(strSQL))
                Helpers.Fill(adapter, strTableName, objDS)
                ReDim objDataColumn(0)
                objDataColumn(0) = objDS.Tables(strTableName).Columns("wc")
                objDS.Tables(strTableName).PrimaryKey = objDataColumn

                strTableName = "stations"
                strSQL = "select pdmworkstation.wsid wsid,pdmworkstation.wscode wscode,pdmworkstation.wc wc from stations,lines ,pdmworkstation "
                strSQL &= " where(lines.lineid = stations.lineid)"
                strSQL &= " and pdmworkstation.wc= lines.description"
                strSQL &= " and pdmworkstation.wscode=stations.absnno"
                strSQL &= " order by stations.lineid,stations.absnno"

                adapter = connection.GetAdapter(connection.GetCommand(strSQL))
                Helpers.Fill(adapter, strTableName, objDS)

                ReDim objDataColumn(1)
                objDataColumn(0) = objDS.Tables(strTableName).Columns("wc")
                objDataColumn(1) = objDS.Tables(strTableName).Columns("wsid")
                objDS.Tables(strTableName).PrimaryKey = objDataColumn
            End Using
        Catch ex As Exception
            Call GenerateException(ex)
        Finally
            objDataColumn = Nothing
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objDS
    End Function

    Public Function GetTreeViewData_OSM() As DataSet 'IDataAdapter


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objDataSet As New DataSet

        Try
            Select Case Connection.GetDatabaseTypeFromConnectionString(gStrConnectionString)
                Case Connection.DatabaseTypes.Oracle, Connection.DatabaseTypes.OracleManagedProvider, Connection.DatabaseTypes.OracleManagedProviderBridge
                    objDataSet = GetTreeViewData_Oracle()
                Case Connection.DatabaseTypes.SQL
                    objDataSet = GetTreeViewData_SQLServer()
            End Select
        Catch ex As Exception
            Call GenerateException("GetTreeViewData_Oracle", ex)
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objDataSet

    End Function

    Private Function GetTreeViewData_SQLServer() As DataSet

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim objDS As New DataSet("Treeview")
        Dim objDataColumn(0) As DataColumn

        Try
            Using connection = New Connection(gStrConnectionString)
                Dim strTableName = "lines"
                strSQL = "select pdmworkcenter.wc wc,pdmworkcenter.descx descx from lines,pdmworkcenter where pdmworkcenter.wc=lines.description order by lines.lineid"
                Dim adapter = connection.GetAdapter(strSQL)
                Helpers.Fill(adapter, strTableName, objDS)
                ReDim objDataColumn(0)
                objDataColumn(0) = objDS.Tables(strTableName).Columns("wc")
                objDS.Tables(strTableName).PrimaryKey = objDataColumn

                strTableName = "stations"
                strSQL = "select pdmworkstation.wsid wsid,pdmworkstation.wscode wscode,pdmworkstation.wc wc from stations,lines ,pdmworkstation"
                strSQL &= " where(lines.lineid = stations.lineid)"
                strSQL &= " and pdmworkstation.wc= lines.description"
                strSQL &= " and pdmworkstation.wscode=stations.absnno"
                strSQL &= " order by stations.lineid,stations.absnno"

                adapter = connection.GetAdapter(strSQL)
                Helpers.Fill(adapter, strTableName, objDS)

                ReDim objDataColumn(2)
                objDataColumn(0) = objDS.Tables(strTableName).Columns("wc")
                objDataColumn(1) = objDS.Tables(strTableName).Columns("wscode")
                objDataColumn(1) = objDS.Tables(strTableName).Columns("wsid")
                objDS.Tables(strTableName).PrimaryKey = objDataColumn
            End Using
        Catch ex As Exception
            Call GenerateException("GetTreeViewData_SQLServer", ex)
        Finally
            objDataColumn = Nothing
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objDS
    End Function

    Public Function GetUserLastVisitedScreen(ByVal strIPAddress As String,
            ByVal struserID As String, ByVal intOperatorPosition As Integer) As stUserLastVisitedScreen


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strIPAddress:({0}) struserID:({1}) intOperatorPosition:({2})", strIPAddress, struserID, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        Dim objULVS As New stUserLastVisitedScreen
        Try
            ClearUserLastVisitedScreenObject(objULVS)


            strSQL = "select * from AMLSwitchLogin where " &
                    Ec.GeneralFunc.GetQueryFieldCondition("ipaddress", strIPAddress) &
                    " and " & Ec.GeneralFunc.GetQueryFieldCondition("userid", struserID) &
                    " and operatorposition=" & intOperatorPosition.ToString &
                    " and DATEDIFF(hh, datex, getdate())  <4"       'get the records updated only in the last four hours

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim reader As DataRow = table.Rows(0)
                        objULVS.IPAddress = Trim(Data.GetDataRowValue(reader("ipaddress"), ""))
                        objULVS.UserID = Trim(Data.GetDataRowValue(reader("userid"), ""))
                        objULVS.BuildNo = Trim(Data.GetDataRowValue(reader("buildno"), ""))
                        objULVS.StationNO = Trim(Data.GetDataRowValue(reader("stationno"), ""))
                        objULVS.ID = Data.GetDataRowValue(reader("id"), 0)
                        objULVS.RecType = Trim(Data.GetDataRowValue(reader("rectype"), ""))
                        objULVS.SEQ = Data.GetDataRowValue(reader("seq"), 0)
                        objULVS.OPNO = Trim(Data.GetDataRowValue(reader("opno"), ""))
                        objULVS.SubHeader = Trim(Data.GetDataRowValue(reader("subheader"), ""))
                        objULVS.PageID = Data.GetDataRowValue(Of Int16)(reader("pageid"), 0)
                        objULVS.Spare1 = Data.GetDataRowValue(Of Int16)(reader("spare1"), 0)
                        objULVS.OperatorPosition = Data.GetDataRowValue(Of Int16)(reader("operatorposition"), 0)
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetUserLastVisitedScreen", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objULVS
    End Function

    Public Function GetVehicleBOMForTorqueVISReport(ByVal model As Integer, ByVal modelYear As String, Optional ByVal jointClass As Integer? = Nothing, Optional ByVal excludeParts As List(Of String) = Nothing) As stVehBOM()

        Dim objVBOM() As stVehBOM
        ReDim objVBOM(0)
        Try
            Dim strSQL = "SELECT DISTINCT " +
                            "aml_vehiclebom.partnumber, body, drive, gearbox, CAST(null AS varchar(10)) AS linex, optioncode, trackstation, cpsc, performance, model, modelyear1, modelyear2, modelyear3, modelyear4, modelyear5, torquemin, torquemax,	quantity, shortdesc, tracking, [validation], kitno, angle, jointclass, suppliercode " +
                            "FROM aml_vehiclebom LEFT JOIN aml_vehicleparts ON aml_vehiclebom.partnumber = aml_vehicleparts.partnumber WHERE {0}"
            Dim whereClauses As New List(Of String)()

            whereClauses.Add(String.Format("(aml_vehiclebom.model='{0}')", model))

            Dim modelYearClause = String.Join(" OR ", Enumerable.Range(1, 5).Select(Function(x) String.Format("aml_vehiclebom.modelyear{0}='{1}'", x, modelYear)))
            whereClauses.Add(String.Format("({0})", modelYearClause))

            If (Not IsNothing(jointClass)) Then
                If (jointClass = 0) Then
                    whereClauses.Add("(jointClass IN (1, 2))")
                Else
                    whereClauses.Add(String.Format("(jointClass = {0})", jointClass))
                End If
            End If

            If (Not IsNothing(excludeParts) AndAlso excludeParts.Count > 0) Then
                whereClauses.Add(String.Format("(aml_vehiclebom.partnumber NOT IN ({0}))", String.Join(", ", excludeParts.Select(Function(x) "'" + x + "'"))))
            End If

            strSQL = String.Format(strSQL, String.Join(" AND ", whereClauses))
            objVBOM = LoadVehicleBomIntoArray(strSQL)
        Catch ex As Exception
            GenerateException(ex)
        End Try

        Return objVBOM

    End Function

    Public Function GetVehicleBOM(ByVal vehicleOptions As List(Of AML_VehicleOption),
                                  ByVal strModelYear As String,
                                  ByVal strDrive As String,
                                  ByVal strBody As String,
                                  ByVal strGear As String,
                                  ByVal strPerformance As String,
                                  ByVal strBuildNO As String,
                                  ByVal intRangeFrom As Integer,
                                  ByVal intRangeTo As Integer,
                                  ByVal intStationNO As Integer,
                                  ByVal intModel1 As Integer,
                                  ByVal strRefNumber As String,
                                  Optional ByVal strAddlCond As String = "",
                                  Optional ByRef strSQL As String = "",
                                  Optional strQueryPartNo As String = "",
                                  Optional includeEmptyOptionCodes As Boolean = True,
                                  Optional ByVal blnPartShortageExclusion As Boolean = False,
                                  Optional ByVal jointClass As Integer? = Nothing) As stVehBOM()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objVOO.Count:({13}) strModelYear:({0}) strDrive:({1}) strBody:({2}) strGear:({3}) strPerformance:({4}) strBuildNO:({5}) intRangeFrom:({6}) intRangeTo:({7}) intStationNO:({8}) intModel1:({9}) strRefNumber:({10}) strAddlCond:({11}) strQueryPartNo:({12}))", strModelYear, strDrive, strBody, strGear, strPerformance, strBuildNO, intRangeFrom, intRangeTo, intStationNO, intModel1, strRefNumber, strAddlCond, strQueryPartNo, vehicleOptions.Count()), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objVBOM() As stVehBOM
        ReDim objVBOM(0)
        Try
            Dim whereClauses As New List(Of String)()

            strSQL = "SELECT * FROM aml_vehiclebom " &
                        "LEFT JOIN aml_vehicleparts ON aml_vehiclebom.partnumber = aml_vehicleparts.partnumber "
            If (vehicleOptions.Any()) Then
                Dim optionCodes = vehicleOptions.Where(Function(x) Not String.IsNullOrEmpty(x.OptionCode1)).Select(Function(x) String.Format("'{0}'", Strings.ReplaceSingleQuote(x.OptionCode1))).ToList()
                If (includeEmptyOptionCodes) Then optionCodes.Add("''")
                whereClauses.Add(String.Format("aml_vehiclebom.optioncode IN ({0})", String.Join(",", optionCodes)))
            End If

            If Trim(strModelYear) <> "" Then
                Dim modelYearClause = String.Join(" OR ", Enumerable.Range(1, 5).Select(Function(x) String.Format("aml_vehiclebom.modelyear{0}='{1}'", x, strModelYear)))
                whereClauses.Add(String.Format("({0})", modelYearClause))
            End If
            If intStationNO <> -1 Then
                whereClauses.Add(String.Format("(trackstation = '{0}')", intStationNO.ToString().PadLeft(4, "0"c)))
            End If
            If strRefNumber.Trim <> "" Then
                whereClauses.Add(String.Format("(case when SUBSTRING(linex, PATINDEX('%[^0]%', linex+'.'), LEN(linex)) = '' then '0' else SUBSTRING(linex, PATINDEX('%[^0]%', linex+'.'), LEN(linex)) end = '{0}')", Strings.ReplaceSingleQuote(strRefNumber).TrimStart("0"c)))
            End If
            If strDrive.Trim <> "" Then
                whereClauses.Add(String.Format("(drive = '{0}' OR drive = '')", Strings.ReplaceSingleQuote(strDrive)))
            End If
            If strBody.Trim <> "" Then
                whereClauses.Add(String.Format("(body = '{0}' OR body = '')", Strings.ReplaceSingleQuote(strBody)))
            End If
            If strGear.Trim <> "" Then
                whereClauses.Add(String.Format("(gearbox = '{0}' OR gearbox = '')", Strings.ReplaceSingleQuote(strGear)))
            End If
            If strPerformance.Trim <> "" Then
                whereClauses.Add(String.Format("(performance = '{0}' OR performance = '')", Strings.ReplaceSingleQuote(strPerformance)))
            End If
            If intModel1 <> -1 Then
                whereClauses.Add(String.Format("(model = {0})", intModel1))
            End If
            If strQueryPartNo.Trim <> "" Then
                whereClauses.Add(String.Format("(aml_vehicleparts.partnumber = '{0}')", Strings.ReplaceSingleQuote(strQueryPartNo)))
            End If
            If strAddlCond.Trim <> "" Then
                whereClauses.Add(strAddlCond)
            End If
            If (Not IsNothing(jointClass)) Then
                If (jointClass = 0) Then
                    whereClauses.Add("(jointClass IN (1, 2))")
                Else
                    whereClauses.Add(String.Format("(jointClass = {0})", jointClass))
                End If
            End If
            ' part shortage supplier exclusion
            If blnPartShortageExclusion Then
                whereClauses.Add("suppliercode not in ( select suppliercode from PartShortageExclusions)")
            End If


            strSQL &= " WHERE " & String.Join(" AND ", whereClauses)
            strSQL &= " ORDER BY trackstation, aml_vehiclebom.partnumber"

            objVBOM = LoadVehicleBomIntoArray(strSQL)

ExitThisSub:

        Catch ex As Exception
            GenerateException("GetVehicleBOM", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0}) strSQL:({1})", objVBOM.Count(), strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objVBOM
    End Function

    Private Function LoadVehicleBomIntoArray(ByVal sql As String) As stVehBOM()
        Dim objVBOM(0) As stVehBOM
        Using connection = New Connection(gStrConnectionString)
            Using table = connection.GetDataIntoDataTable(sql, Nothing)
                If (Not IsNothing(table)) Then
                    Dim intCount As Int32 = 0
                    ReDim Preserve objVBOM(table.Rows.Count)
                    For Each row As DataRow In table.Rows
                        intCount = intCount + 1

                        objVBOM(intCount).PartNumber = Trim(Data.GetDataRowValue(row("partnumber"), ""))
                        objVBOM(intCount).Body = Trim(Data.GetDataRowValue(row("body"), ""))
                        objVBOM(intCount).Drive = Trim(Data.GetDataRowValue(row("drive"), ""))
                        objVBOM(intCount).Gear = Trim(Data.GetDataRowValue(row("gearbox"), ""))
                        objVBOM(intCount).LineNumber = Data.GetDataRowValue(row("linex"), 0)
                        objVBOM(intCount).OptionCode = Trim(Data.GetDataRowValue(row("optioncode"), ""))
                        objVBOM(intCount).StationNumber = Trim(Data.GetDataRowValue(row("trackstation"), ""))
                        objVBOM(intCount).CPSC = Trim(Data.GetDataRowValue(row("cpsc"), ""))
                        objVBOM(intCount).Performance = Trim(Data.GetDataRowValue(row("performance"), ""))
                        objVBOM(intCount).Model1 = Data.GetDataRowValue(row("model"), 0)

                        Dim modelYearList = New List(Of String)()
                        modelYearList.Add(Data.GetDataRowValue(row("modelyear1"), ""))
                        modelYearList.Add(Data.GetDataRowValue(row("modelyear2"), ""))
                        modelYearList.Add(Data.GetDataRowValue(row("modelyear3"), ""))
                        modelYearList.Add(Data.GetDataRowValue(row("modelyear4"), ""))
                        modelYearList.Add(Data.GetDataRowValue(row("modelyear5"), ""))

                        objVBOM(intCount).ModelYear = String.Join(", ", modelYearList.Where(Function(x) Not String.IsNullOrEmpty(x)))
                        objVBOM(intCount).PartDescription = ""

                        objVBOM(intCount).TorqueMin = Trim(Data.GetDataRowValue(row("TorqueMin"), ""))
                        objVBOM(intCount).TorqueMax = Trim(Data.GetDataRowValue(row("TorqueMax"), ""))
                        objVBOM(intCount).Quantity = Trim(Data.GetDataRowValue(row("quantity"), ""))
                        objVBOM(intCount).PartDescription = Trim(Data.GetDataRowValue(row("shortdesc"), ""))
                        objVBOM(intCount).Tracking = Trim(Data.GetDataRowValue(row("tracking"), ""))
                        objVBOM(intCount).Validation = Trim(Data.GetDataRowValue(row("validation"), ""))
                        objVBOM(intCount).KitNO = Trim(Data.GetDataRowValue(row("kitno"), ""))
                        objVBOM(intCount).Angle = Data.GetDataRowValue(Of Double)(row("angle"), 0)
                        objVBOM(intCount).JointClass = Data.GetDataRowValue(Of Int16)(row("JointClass"), 0)
                        objVBOM(intCount).SupplierCode = Trim(Data.GetDataRowValue(row("SupplierCode"), ""))

                        If (Not String.IsNullOrEmpty(objVBOM(intCount).TorqueMin) AndAlso Not String.IsNullOrEmpty(objVBOM(intCount).TorqueMax)) Then
                            objVBOM(intCount).NominalTorqueValue = (Convert.ToDouble(objVBOM(intCount).TorqueMin) + Convert.ToDouble(objVBOM(intCount).TorqueMax)) / 2
                        End If
                    Next
                End If
            End Using
        End Using
        Return objVBOM
    End Function

    Public Function GetVehicleOptions(ByVal strBuildNO As String) As List(Of AML_VehicleOption)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0})", strBuildNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = New List(Of AML_VehicleOption)
        Try

            'get the specific options
            Dim strSQL = String.Format("SELECT * FROM aml_vehicleoptions WHERE buildnumber='{0}' AND optioncode1 IS NOT NULL", Strings.ReplaceSingleQuote(strBuildNO))
            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.GetDataIntoClassOf(Of AML_VehicleOption)(strSQL)
                If (result IsNot Nothing) Then returnValue = result.ToList()
            End Using
        Catch ex As Exception
            GenerateException("GetVehicleOptions", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue

    End Function

    Public Function GetVehicleOrderDetails(ByVal buildNos As List(Of String)) As List(Of AML_VehicleOrders)
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter buildNos:({0})", buildNos.Count()), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = New List(Of AML_VehicleOrders)
        Try

            Dim strSQL = String.Format("SELECT * FROM aml_vehicleorders " &
                                        "LEFT JOIN aml_vehicletrim ON aml_vehicleorders.buildnumber = aml_vehicletrim.buildnumber " &
                                        "WHERE aml_vehicleorders.buildnumber IN ({0})", String.Join(", ", buildNos.Where(Function(x) Not String.IsNullOrEmpty(x) AndAlso x.Trim() <> "")))

            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.GetDataIntoClassOf(Of AML_VehicleOrders)(strSQL)
                If (result IsNot Nothing) Then returnValue = result.ToList()
            End Using
        Catch ex As Exception
            GenerateException("GetVehicleOrderDetails", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function GetVehicleOrderDetails(ByVal strBuildNO As String) As AML_VehicleOrders

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0})", strBuildNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As AML_VehicleOrders = Nothing
        Try
            Dim strSQL = "SELECT TOP 1 vo.*, vt.optiondescription FROM aml_vehicleorders vo " &
                     "LEFT JOIN aml_vehicletrim vt ON vo.buildnumber = vt.buildnumber And environment = 'EC' WHERE vo.buildnumber=" & strBuildNO.ReplaceSingleQuote()

            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.GetDataIntoClassOf(Of AML_VehicleOrders)(strSQL)
                If (result IsNot Nothing) Then returnValue = result.FirstOrDefault()
            End Using
        Catch ex As Exception
            GenerateException("GetVehicleOrderDetails", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit IsNothing({0})", IsNothing(returnValue)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue

    End Function

    Public Function GetVehicleOrderNumber(strBuildNO As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0})", strBuildNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select ordernumber from AML_VehicleOrders where buildnumber=" & strBuildNO.ReplaceSingleQuote()
        Dim returnValue = ""
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of String)(strSQL, "")
        End Using
#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Sub GetViewEASEAccessInfo_Version7(ByVal strUserID As String,
                                              ByRef strGroupName As String,
                                              ByRef strRHFields As String,
                                              ByRef strOPFields As String,
                                              ByRef intGroupID As Int32,
                                              Optional ByRef strViewEASEUserTypeID As String = "")

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter: strUserID:({0})", strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        Try
            strGroupName = "" : strRHFields = "" : strOPFields = "" : strViewEASEUserTypeID = ""
            intGroupID = 0
            strSQL = "select easegroups.groupid, rh_fields, op_fields, groupname from easegroups where groupid in " &
                " (select groupid from euserrights where " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID) & ")"

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim reader As DataRow = table.Rows(0)
                        strRHFields = Data.GetDataRowValue(reader("rh_fields"), "")
                        intGroupID = Data.GetDataRowValue(reader("groupid"), 0)
                        strOPFields = Data.GetDataRowValue(reader("op_fields"), "")
                        strGroupName = Data.GetDataRowValue(reader("groupname"), "")
                    End If
                End Using
                strSQL = "select easeaccess from euserrights where " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID) & " and groupid=" & intGroupID.ToString
                strViewEASEUserTypeID = connection.ExecuteScalar(Of Integer)(strSQL, 0).ToString()
            End Using
        Catch ex As Exception
            GenerateException("GetViewEASEAccessInfo_Version7", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit: strGroupName:({0}) strRHFields:({1}) strOPFields:({2}) intGroupID:({3}) strViewEASEUserTypeID:({4})",
                                                          strGroupName, strRHFields, strOPFields, intGroupID, strViewEASEUserTypeID), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function GetWIPData(Optional ByVal intEngineComplete As Int16 = 0,
            Optional ByVal intLineID As Int16 = 0,
            Optional ByVal blnNewEngines As Boolean = False) As stWIPData()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intEngineComplete:({0}) intLineID:({1}) blnNewEngines:({2})",
                                                          intEngineComplete, intLineID, blnNewEngines), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strJoin As String = " where "
        Dim objWIPData() As stWIPData
        ReDim objWIPData(0)
        Try
            Dim strSQL = "select engineid,startdate,startengineer,engineno,spare1,spare2 from MES_ENGINESLIST "
            If intEngineComplete <> -1 Then
                'EngineComplete: 0-Active, 1-Complete, -1 - All
                strSQL &= " where EngineComplete =" & intEngineComplete
                strJoin = " and "
            End If
            If intLineID > 0 Then
                strSQL &= strJoin & " spare2=" & intLineID
                strJoin = " and "
            End If
            If blnNewEngines Then  'get the engines in the last two days only/ three days -KM/Will's request
                strSQL &= strJoin & " (startdate >= (getdate()-3))"
                'strSQL &= strJoin & " (startdate between getdate()-3 and getdate()+1)"
                'strSQL &= strJoin & " (startdate <= getdate()-3) " ' and getdate()+1)"
                strJoin = " and "
            End If
            strSQL &= " order by engineno"

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objWIPData(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1
                            objWIPData(counter).EngineID = Data.GetDataRowValue(row("EngineID"), 0)
                            objWIPData(counter).BuildNo = Data.GetDataRowValue(row("engineno"), "")
                            objWIPData(counter).Engineer = Data.GetDataRowValue(row("startengineer"), "")
                            objWIPData(counter).StartDate = Data.GetDataRowValue(row("startdate"), DateTime.MinValue)
                            objWIPData(counter).ModelNumber = Data.GetDataRowValue(Of Int16)(row("spare1"), 0)
                            objWIPData(counter).LineID = Data.GetDataRowValue(Of Int16)(row("spare2"), 0)
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetWIPData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objWIPData)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objWIPData
    End Function

    Private Function GetWIPEngineID() As Int32

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(EngineID) from MES_ENGINESLIST"
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        returnValue += 1

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetWSID(ByVal strWC As String, ByVal strStationNO As String) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strWC:({0}) strStationNO:({1})", strWC, strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select wsid from pdmworkstation where wc='" & strWC.Trim & "' and wscode='" & strStationNO.Trim & "'"
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function IndependentRefreshingStation(ByVal strStation As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStation:({0})", strStation), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        'EW-855: VH 5 Body Shop and MA1 station refresh
        Select Case strStation.Trim
            Case "3500", "3510", "3520", "3530", "3540", "3600", "3610", "3620", "3630", "3640", "3650", "3660", "3670", "3680"
                blnRtnValue = True
        End Select


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Function IsAdminOrSupervisorRole(ByVal strUserID As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserID:({0})", strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select admingroupid from euser where " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID) & " and userdisabled<>1"
        Dim returnValue = False
        Try
            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                Select Case result
                    Case 1, 2, 3, 4
                        returnValue = True
                End Select
            End Using

        Catch ex As Exception
            Call GenerateException("IsAdminOrSupervisorRole", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function LoadSafetyDocsCheck(ByVal strUserID As String, ByVal lngWSID As Long,
                                        Optional ByVal blnIncludeStationInfo As Boolean = False,
                                        Optional ByVal strTemp As String = "", Optional ByVal blnOnlyActive As Boolean = True,
                                        Optional ByVal intPassStatus As Int16 = -22, Optional ByVal intRelDate As Integer = 0,
                                        Optional ByVal intCheckRead As Int16 = -1, Optional ByVal strSortBy As String = "",
                                        Optional ByVal intDateTO As Int32 = 0) As stSafetyCheck()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserID:({0}) lngWSID:({1}) blnIncludeStationInfo:({2}) strTemp:({3}) blnOnlyActive:({4}) " &
                                                          "intPassStatus:({5}) intRelDate:({6}) intCheckRead:({7}) strSortBy:({8}) intDateTO:({9})",
                                                          strUserID, lngWSID, blnIncludeStationInfo, strTemp, blnOnlyActive,
                                                          intPassStatus, intRelDate, intCheckRead, strSortBy, intDateTO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objSafetyCheck() As stSafetyCheck
        ReDim objSafetyCheck(0)
        Dim strSQL As String = ""
        Dim intYY As Int16 = 0

        Try
            If blnIncludeStationInfo = True Then
                strSQL = "select pdmworkstationmm1.docrectype as 'MDMDocRecType',pdmworkstation.wscode, pdmworkstation.wsdesc,AML_SafetyDocsCheck.*"
                strSQL &= " from AML_SafetyDocsCheck, PDMWorkstation,PDMWorkstationmm1 "
                strSQL &= " where PDMWorkstation.wsid = PDMWorkstationmm1.wsid"
                strSQL &= " and PDMWorkstationmm1.wsid = AML_SafetyDocsCheck.wsid"
            Else
                strSQL = "select AML_SafetyDocsCheck.* from AML_SafetyDocsCheck, PDMWorkstationmm1 "
                strSQL &= " where PDMWorkstationmm1.wsid = AML_SafetyDocsCheck.wsid"
            End If
            'strSQL = "select AML_SafetyDocsCheck.* from AML_SafetyDocsCheck, PDMWorkstationmm1 "
            'strSQL &= " where PDMWorkstationmm1.wsid = AML_SafetyDocsCheck.wsid"
            strSQL &= " and PDMWorkstationmm1.docid = AML_SafetyDocsCheck.docid"
            strSQL &= " and PDMWorkstationmm1.docseq = AML_SafetyDocsCheck.docseq"
            If blnOnlyActive = True Then
                strSQL &= " and PDMWorkstationmm1.docrectype ='0'"
            End If
            If Trim(strTemp) = "" Then
                strSQL &= " and PDMWorkstationmm1.wsid = " & lngWSID
            Else
                strSQL &= " and PDMWorkstationmm1.wsid in (" & strTemp.Trim & ")"
            End If

            If Trim(strUserID) <> "" Then
                strSQL &= " and " & Trim(DBConfig.QueryFunctions.Upper) & "(AML_SafetyDocsCheck.userid)='" & UCase(Strings.ReplaceSingleQuote(Trim(strUserID))) & "'"
            End If
            If intRelDate > 0 AndAlso intDateTO > 0 Then
                strSQL &= " and " & " (dateread between " & CStr(intRelDate) & " and " & CStr(intDateTO) & ") "
            End If
            If intPassStatus <> -22 Then
                strSQL &= " and AML_SafetyDocsCheck.pass_status =" & intPassStatus
            End If

            If intCheckRead <> -1 Then
                strSQL &= " and AML_SafetyDocsCheck.checkedread = " & intCheckRead
            End If
            strSQL &= strSortBy


            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objSafetyCheck(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            If IsDBNull(row("docid")) Then GoTo NextRecord
                            If IsDBNull(row("docseq")) Then GoTo NextRecord

                            counter += 1

                            If Trim(strUserID) = "" Then
                                objSafetyCheck(counter).OperatorID = Data.GetDataRowValue(row("userid"), "")
                            Else
                                objSafetyCheck(counter).OperatorID = strUserID.Trim
                            End If

                            objSafetyCheck(counter).WSID = lngWSID
                            objSafetyCheck(counter).DOCID = Data.GetDataRowValue(Of Long)(row("docid"), 0)
                            If blnIncludeStationInfo = True Then
                                ' in the aml qc report
                                objSafetyCheck(counter).DocRecType = Data.GetDataRowValue(row("MDMDocRecType"), "")
                            Else
                                'in home page
                                objSafetyCheck(counter).DocRecType = Data.GetDataRowValue(row("docrectype"), "")
                            End If

                            objSafetyCheck(counter).DocSeq = Data.GetDataRowValue(row("docseq"), 0)
                            objSafetyCheck(counter).DocDesc = Trim(Data.GetDataRowValue(row("docdesc"), ""))
                            objSafetyCheck(counter).RevNum = Data.GetDataRowValue(Of Int16)(row("revnumber"), 0)
                            objSafetyCheck(counter).PassStatus = Data.GetDataRowValue(Of Int16)(row("pass_status"), 0)

                            Select Case objSafetyCheck(counter).PassStatus
                                Case 0, 1
                                    objSafetyCheck(counter).CheckedRead = Data.GetDataRowValue(row("checkedread"), False)
                            End Select


                            If Not IsDBNull(row("revdate")) Then objSafetyCheck(counter).RevDate = Convert.ToDateTime(Dates.NumToDate(Data.GetDataRowValue(row("revdate"), 0)))
                            If Not IsDBNull(row("dateread")) Then
                                If IsNumeric(row("dateread")) = True And Data.GetDataRowValue(row("dateread"), 0) > 0 Then
                                    objSafetyCheck(counter).DateRead = Dates.NumToDate(Data.GetDataRowValue(row("dateread"), 0))
                                End If
                            End If

                            If objSafetyCheck(counter).CheckedRead = True And IsDate(objSafetyCheck(counter).DateRead) Then
                                objSafetyCheck(counter).UpdateRecord = False
                            End If

                            If blnIncludeStationInfo = True Then
                                objSafetyCheck(counter).WSCode = Trim(Data.GetDataRowValue(row("wscode"), ""))
                                objSafetyCheck(counter).WSDescX = Trim(Data.GetDataRowValue(row("wsdesc"), ""))
                            End If
NextRecord:
                        Next
                        ReDim Preserve objSafetyCheck(counter)
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("LoadSafetyDocsCheck", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objSafetyCheck)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objSafetyCheck
    End Function

    Public Sub ReadSerialScanStationList(ByRef intLineID1 As Integer, ByRef strStation1 As String,
                                                ByRef intLineID2 As Integer, ByRef strStation2 As String,
                                                ByRef intLineID3 As Integer, ByRef strStation3 As String)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Dim intKeyx As Integer = 0
        Try

            strSQL = "select wl,data,keyx from client where keyx in (237,238,239) and dp=99 order by keyx"
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each reader As DataRow In table.Rows
                            intKeyx = 0
                            If Not IsDBNull(reader(2)) Then intKeyx = Data.GetDataRowValue(reader(2), 0)

                            Select Case intKeyx
                                Case 237
                                    intLineID1 = Data.GetDataRowValue(reader(0), 0)
                                    strStation1 = Data.GetDataRowValue(reader(1), "")
                                Case 238
                                    intLineID2 = Data.GetDataRowValue(reader(0), 0)
                                    strStation2 = Data.GetDataRowValue(reader(1), "")
                                Case 239
                                    intLineID3 = Data.GetDataRowValue(reader(0), 0)
                                    strStation3 = Data.GetDataRowValue(reader(1), "")
                            End Select
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("ReadSerialScanStationList", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Enter intLineID1:({0}) strStation1:({1}) intLineID2:({2}) strStation2:({3}) intLineID3:({4}) strStation3:({5})",
                                                          intLineID1, strStation1, intLineID2, strStation2, intLineID3, strStation3), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function SaveBannerMessage(ByRef objMBL As stBannerMsgList,
            ByVal intAddEdit As Int16,
            ByVal objBMS() As stBannerMessagesStations,
            ByVal blnUpdateMessageStations As Boolean) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objMBL.LineID:({0}) intAddEdit:({1}) objBMS.Length:({2}) blnUpdateMessageStations:({3}) objMBL.MessageID:({4})",
                                                          objMBL.LineID, intAddEdit, objBMS.Length, blnUpdateMessageStations, objMBL.MessageID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Dim strSQL As String = ""

        Dim intMsgID As Int32 = 0
        Try
            intMsgID = objMBL.MessageID

            If intAddEdit = 0 Then          'add message
                intMsgID = GetMaximumMessageID()
            End If

            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    If intAddEdit = 0 Then          'add message

                        objMBL.MessageID = intMsgID

                        strSQL = "insert into BannerMessages(msgid,messagetext,raisedby," &
                        " messagereason,messagetype, " &
                        " lineid,stationfrom,stationto, startdate,enddate," &
                        " createddate,lastupdated,messagestatus) values(" &
                        objMBL.MessageID & ",'" & Strings.ReplaceSingleQuote(objMBL.MessageText, True) & "','" &
                        Strings.ReplaceSingleQuote(objMBL.RaisedBy.Trim) & "','" &
                        Strings.ReplaceSingleQuote(objMBL.MessageReason, True) & "','" &
                        Strings.ReplaceSingleQuote(objMBL.MessageType, True) & "'," &
                        objMBL.LineID & ",'" & Trim(objMBL.StationFrom) & "','" &
                        Trim(objMBL.StationTo) & "'," &
                        FormatDateForQuery(objMBL.StartDate) & ", " &
                        FormatDateForQuery(objMBL.EndDate) & ", " &
                        " " & GetDateBuiltInFunction() & "," & GetDateBuiltInFunction() & "," & objMBL.MessageStatus & ")"
                    ElseIf intAddEdit = 1 Then      'edit message

                        strSQL = "update BannerMessages set messagetext='" &
                        Strings.ReplaceSingleQuote(objMBL.MessageText, True) & "'," &
                        "raisedby='" & Strings.ReplaceSingleQuote(objMBL.RaisedBy, True) & "'," &
                        "messagereason='" & Strings.ReplaceSingleQuote(objMBL.MessageReason, True) & "'," &
                        "messagetype='" & Strings.ReplaceSingleQuote(objMBL.MessageType, True) & "'," &
                        "lineid=" & objMBL.LineID & "," &
                        "stationfrom='" & Strings.ReplaceSingleQuote(objMBL.StationFrom, True) & "'," &
                        "stationto='" & Strings.ReplaceSingleQuote(objMBL.StationTo, True) & "'," &
                        "startdate=" & FormatDateForQuery(objMBL.StartDate) & "," &
                        "enddate=" & FormatDateForQuery(objMBL.EndDate) & "," &
                        "lastupdated=" & GetDateBuiltInFunction() & ",messagestatus=" & objMBL.MessageStatus &
                        " where msgid=" & objMBL.MessageID
                    ElseIf intAddEdit = 2 Then      'Approve message
                        'MessageStatus: 0-Draft/Pending Approval, 1-Approved
                        strSQL = "update BannerMessages set messagestatus=1," &
                        "approvedby='" & Strings.ReplaceSingleQuote(objMBL.ApprovedBy, True) & "'" &
                        " where msgid=" & objMBL.MessageID

                    End If

                    If Trim(strSQL) <> "" Then
                        connection.ExecuteNonQuery(strSQL)

                        If blnUpdateMessageStations Then
                            strSQL = "delete from BannerMessages_Stations where msgid=" & intMsgID
                            connection.ExecuteNonQuery(strSQL)

                            For intK = 1 To UBound(objBMS)
                                strSQL = "insert into BannerMessages_Stations(msgid, stationno,stationseq) values (" &
                                            intMsgID & ",'" & Trim(objBMS(intK).Stage) & "'," & objBMS(intK).StationSeq & ")"
                                connection.ExecuteNonQuery(strSQL)
                            Next

                        End If
                        blnRtnValue = True
                    End If
                    connection.CommitTransaction()
                    blnRtnValue = True
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Throw ex
                End Try
            End Using
        Catch ex As Exception
            Call GenerateException("SaveBannerMessage: " & strSQL, ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Public Function SaveSafetyDocCheck(ByVal objSafetyCheck() As stSafetyCheck) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objSafetyCheck.Length:({0})", objSafetyCheck.Length), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        Dim strSQL As String = ""
        Dim strSQLCond As String = ""

        Try
            Dim lngWSID As Long = 0, strUserID As String = ""
            Dim lngDocID As Long = 0, intDocSeq As Integer = 0

            If UBound(objSafetyCheck) > 0 Then
                lngWSID = objSafetyCheck(1).WSID 'All the records should have the same wsid & userid
                strUserID = Trim(objSafetyCheck(1).OperatorID)
            End If

            Dim sqlList = New List(Of String)()
            For intK = 1 To UBound(objSafetyCheck)
                'Update only today's changes
                'If objSafetyCheck(intK).CheckedRead = True And objSafetyCheck(intK).DateRead = Extensions.Dates.CurrentDate Then
                '09/18/12
                If objSafetyCheck(intK).UpdateRecord = True Then 'And objSafetyCheck(intK).CheckedRead = True Then
                    lngDocID = objSafetyCheck(intK).DOCID
                    intDocSeq = objSafetyCheck(intK).DocSeq

                    strSQLCond = " where wsid = " & lngWSID
                    strSQLCond &= " and " & Trim(DBConfig.QueryFunctions.Upper) & "(userid)='" & UCase(Strings.ReplaceSingleQuote(Trim(strUserID))) & "'"
                    strSQLCond &= " and docid = " & lngDocID & " and docseq = " & intDocSeq

                    strSQL = "delete from AML_SafetyDocsCheck " & strSQLCond
                    sqlList.Add(strSQL)

                    strSQL = "insert into AML_SafetyDocsCheck(userid,wsid,wscode,docid,docrectype,docseq,docdesc,revnumber"
                    strSQL &= ",revdate,dateread,checkedread,pass_status) values("
                    strSQL &= "'" & Strings.ReplaceSingleQuote(objSafetyCheck(intK).OperatorID, True) & "'"
                    strSQL &= "," & objSafetyCheck(intK).WSID
                    strSQL &= ",'" & Strings.ReplaceSingleQuote(objSafetyCheck(intK).WSCode, True) & "'"
                    strSQL &= "," & objSafetyCheck(intK).DOCID
                    strSQL &= ",'" & Trim(objSafetyCheck(intK).DocRecType) & "'"
                    strSQL &= "," & objSafetyCheck(intK).DocSeq
                    strSQL &= ",'" & Strings.ReplaceSingleQuote(objSafetyCheck(intK).DocDesc, True) & "'"
                    strSQL &= "," & objSafetyCheck(intK).RevNum
                    strSQL &= "," & Dates.DateToNum(objSafetyCheck(intK).RevDate)
                    strSQL &= "," & Dates.DateToNum(objSafetyCheck(intK).DateRead)
                    If objSafetyCheck(intK).CheckedRead = True Then
                        strSQL &= ",1"
                    Else
                        strSQL &= ",0"
                    End If
                    strSQL &= "," & objSafetyCheck(intK).PassStatus
                    strSQL &= ")"

                    sqlList.Add(strSQL)
                End If
            Next intK

            Ec.IO.RunSQLInArray(sqlList, blnRtnValue)

        Catch ex As Exception
            GenerateException("AstonMartin: SaveSafetyDocCheck", ex)
        Finally
            objSafetyCheck = Nothing
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Public Sub UpdateBannerMessageState(ByVal intMsgID As Int32, ByVal intMessageStatus As Int16, ByVal strApprover As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intMsgID:({0}) intMessageStatus:({1}) strApprover:({2})", intMsgID, intMessageStatus, strApprover), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'MessageStatus: 0-Draft/Pending Approval, 1-Approved
        Dim strSQL As String = "update BannerMessages set messagestatus=" &
                intMessageStatus & ", approvedby='" & Strings.ReplaceSingleQuote(strApprover) & "'" &
                " where msgid=" & intMsgID

        Ec.IO.RunSQL(strSQL)

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub UpdateBootBagPart(strBuildNO As String, strStationNO As String, strPartNO As String, strUserID As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2}) strUserID:({3})", strBuildNO, strStationNO, strPartNO, strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        strSQL = "update AML_BuildBootBag set verifystatus=1,lastupdated=getdate(),userid='" & Strings.ReplaceSingleQuote(strUserID) & "' where " &
            " buildnumber=" & strBuildNO & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) &
            " and verifystatus <> 1 and " & Ec.GeneralFunc.GetQueryFieldCondition("partnumber", strPartNO)
        Ec.IO.RunSQL(strSQL)


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function UpdateBuildList_PLC_LOG(ByVal strBuildNo As String, ByVal strLog As String, Optional ByVal blnErrorLog As Boolean = False) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strLog:({1}) blnErrorLog:({2})",
                                                          strBuildNo, strLog, blnErrorLog), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False

        Try
            Dim sqlList = New List(Of String)()
            sqlList.Add("update aml_plc_buildlist set easeread=1 where buildno='" & strBuildNo.Trim & "'")
            sqlList.Add("delete from aml_plc_buildlist_log where buildno='" & strBuildNo.Trim & "'")
            If strLog.Trim.Length > 255 Then
                strLog = Left(strLog.Trim, 250)     'to cover '
            End If
            sqlList.Add("insert into aml_plc_buildlist_log(buildno,completetime,logdescription) values ('" & strBuildNo.Trim & "',getdate(),'" & Strings.ReplaceSingleQuote(strLog) & "')")
            Ec.IO.RunSQLInArray(sqlList.ToArray(), blnResult)
        Catch ex As Exception
            Call GenerateException("UpdateBuildList_PLC_LOG: " & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult

    End Function

    Public Sub UpdatePSSDocumentReadStatus(objPSS_Read() As stPSSDocuments_READ)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objPSS_Read.Length:({0})", objPSS_Read.Length), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False
        Dim strTemp As String = ""

        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    For intK = 1 To UBound(objPSS_Read)
                        strTemp = ""
                        If objPSS_Read(intK).DateRead <> Date.MinValue Then
                            strTemp = "getdate()" 'Format(objPSS_Read(intK).DateRead, "MMM dd yyyy hh:mm:ss:ffftt")        'get the date in MMM dd yyyy hh:mmtt format
                            'strTemp = " CONVERT(DATETIME,'" & strTemp.Trim & "',109) "          '100 - sql server format: mon dd yyyy hh:mi:ss:mmmAM (or PM)

                        End If

                        strSQL = "delete from AML_PSSDocsCheck where  " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", objPSS_Read(intK).StationNO) & " and " &
                        " MDMDocID=" & objPSS_Read(intK).MDMDocID & " and CompleteTTKey=" & objPSS_Read(intK).CompleteTTKey & " and " &
                        Ec.GeneralFunc.GetQueryFieldCondition("UserID", objPSS_Read(intK).USERID)
                        connection.ExecuteNonQuery(strSQL)

                        strSQL = "insert into AML_PSSDocsCheck (StationNO, UserID,MDMDocID,TTKey,CompleteTTKey,checkedread,pass_status,lineid,opno"
                        If strTemp <> "" Then
                            strSQL &= ", dateread "
                        End If
                        strSQL &= ") values ('" & objPSS_Read(intK).StationNO & "','" & Strings.ReplaceSingleQuote(objPSS_Read(intK).USERID) & "'," &
                        objPSS_Read(intK).MDMDocID & "," & objPSS_Read(intK).TTKey & "," & objPSS_Read(intK).CompleteTTKey & "," & objPSS_Read(intK).CheckedRead & "," & objPSS_Read(intK).PassStatus &
                        "," & objPSS_Read(intK).LineID & ",'" & objPSS_Read(intK).OPNO.Trim & "'"
                        If strTemp <> "" Then
                            strSQL &= ",  " & strTemp
                        End If

                        strSQL &= ") "

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
            Call GenerateException("UpdatePSSDocumentReadStatus", ex)
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function UpdateQSMRequiredByStation(ByVal objLineStations() As stStations) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objLineStations.Length:({0})", objLineStations.Length), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean = False
        Dim strSQL As String = ""
        Try
            Dim sqlList = New List(Of String)()

            For intk = 1 To UBound(objLineStations)
                strSQL = "update stations set osmrequired=" & objLineStations(intk).OSMRequired & " where absnno='" & objLineStations(intk).StationNO & "' and lineid=" & objLineStations(intk).LineID
                sqlList.Add(strSQL)
            Next

            Ec.IO.RunSQLInArray(sqlList, blnResult)

        Catch ex As Exception
            GenerateException("AstonMartin : UpdateQSMRequiredByStation", ex)
        End Try



#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function UpdateQualityChecksData(ByVal objQCData() As stQualityChecks,
                                            ByVal strBuildNO As String,
                                            ByVal strStationNO As String,
                                            ByVal intRecordType As Int16,
                                            Optional blnQCOverride As Boolean = False,
                                            Optional intShift As Byte = 0) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objQCData.Length:({0}) strBuildNO:({1}) strStationNO:({2}) intRecordType:({3}) blnQCOverride:({4}) intShift:({5})",
                                                          objQCData.Length, strBuildNO, strStationNO, intRecordType, blnQCOverride, intShift), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Record type: 0-Quality checks/Remedy Plan ,1- FLM record

        Dim strSQL As String = "", blnResult As Boolean = False
        Dim intPrevRecordSeq As Int16 = -1, intRecordSeq As Int16 = 0
        Dim intResultFlag As Int16 = -1
        Dim intPassFail As Int16 = 0
        Dim blnFail As Boolean = False, blnIncomplete As Boolean = False

        Try
            Dim sqlList = New List(Of String)()
            For intK = 1 To UBound(objQCData)


                intRecordSeq = objQCData(intK).RecordSeq
                If intPrevRecordSeq = -1 Then
                    intPrevRecordSeq = intRecordSeq 'only for the first record
                End If

                If intRecordSeq <> intPrevRecordSeq Then    'update the result in the header record
                    If blnFail Then intResultFlag = 2
                    If blnIncomplete Then intResultFlag = -1

                    'PDMRPDATA: Result Flag: -1 -> incomplete, 1->pass, 2->fail
                    'PDMRPDATA->Record type: 0-Quality checks/Remedy Plan ,1- FLM record

                    strSQL = "update pdmrpdata set resultflag=" & intResultFlag
                    If intShift > 0 Then
                        strSQL &= ", shift=" & intShift
                    End If
                    strSQL &= " where " & Ec.GeneralFunc.GetQueryFieldCondition("engineno", strBuildNO) & " and " &
                                    Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) & " and " &
                                    " recordtype= " & intRecordType.ToString & " and " &
                                    " recordseq=" & intPrevRecordSeq     '** change in recordseq variable
                    sqlList.Add(strSQL)

                    intPrevRecordSeq = intRecordSeq
                    intResultFlag = -1      'default value (** incomplete)
                    blnFail = False : blnIncomplete = False

                End If

                intPassFail = -1
                If Not objQCData(intK).InvalidEntry Then
                    intPassFail = objQCData(intK).PassFail
                End If

                'PDMRPITemValues->Record type: 0-Quality checks/Remedy Plan ,1- FLM record
                'PDMRPItemValues->PassFail: -1 -> InComplete, 0-FAIL, 1->PASS, 2-Override
                strSQL = "update pdmrpitemvalues set UserInput_Time_Stamp=getdate(), passfail=" & objQCData(intK).PassFail & ",userinput='" &
                    Strings.ReplaceSingleQuote(objQCData(intK).UserInput) & "'," &
                    " operatorid='" & Strings.ReplaceSingleQuote(objQCData(intK).OperatorID) & "'"
                If blnQCOverride Then
                    strSQL &= ", supervisoraction='" & Strings.ReplaceSingleQuote(objQCData(intK).SuperVisorAction) & "'"
                End If
                strSQL &= " where " &
                    Ec.GeneralFunc.GetQueryFieldCondition("engineno", strBuildNO) & " and " &
                    Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) & " and " &
                    " recordseq=" & intRecordSeq & " and remedyplanseq=" & objQCData(intK).RemedyPlanSeq &
                     " and recordtype=" & intRecordType.ToString

                '-------------------------------------------------------------------------------------------------
                'Any change in this block should be updated in UpdateQualityChecksData and GetQCStatusFromMemory
                Select Case intPassFail 'objQCData(intK).PassFail
                    Case 0
                        'PDMRPDATA: Result Flag: -1 -> incomplete, 1->pass, 2->fail
                        intResultFlag = 2
                        blnFail = True
                    Case 1
                        'PDMRPDATA: Result Flag: -1 -> incomplete, 1->pass, 2->fail
                        If intResultFlag = 0 Then
                            'any fail is a fail (ex: 4 items, 3 items pass, 1 fail, the result flag is fail)
                        Else
                            intResultFlag = 1
                        End If
                    Case -1
                        blnIncomplete = True
                End Select
                '-------------------------------------------------------------------------------------------------
                sqlList.Add(strSQL)
            Next intK

            If UBound(objQCData) > 0 Then

                If blnFail Then intResultFlag = 2
                If blnIncomplete Then intResultFlag = -1


                'to update the last header record
                strSQL = "update pdmrpdata set resultflag=" & intResultFlag
                If intShift > 0 Then
                    strSQL &= ", shift=" & intShift
                End If
                strSQL &= " where " & Ec.GeneralFunc.GetQueryFieldCondition("engineno", strBuildNO) & " and " &
                            Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO) & " and " &
                            " recordtype=" & intRecordType.ToString & " and " &
                            " recordseq=" & intRecordSeq  '** change in recordseq variable

                sqlList.Add(strSQL)
                Ec.IO.RunSQLInArray(sqlList, blnResult)
            End If
        Catch ex As Exception
            Call GenerateException("UpdateQualityChecksData", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult

    End Function

    Public Sub UpdateStationSignoffStatus(buildNo As String,
                                               stationNo As String,
                                               operatorPosition As Int16,
                                               stationSignoffStatus As Int16,
                                               partVerifySerialStatus As Int16,
                                               userID As String,
                                               Optional wipComplete As Boolean = False)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) intOperatorPosition:({2}) intStationSignoffStatus:({3}) intPartVerifySerialStatus:({4}) strUserID:({5}) blnWIPComplete:({6})",
                                                          buildNo, stationNo, operatorPosition, stationSignoffStatus, partVerifySerialStatus, userID, wipComplete), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        If (String.IsNullOrEmpty(buildNo) OrElse String.IsNullOrEmpty(stationNo)) Then
#If TRACE Then
            Log.OPERATION("Exit (BuildNo or StationNo empty)", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
            Return
#End If
        End If

        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    'AML_StationSignoffStatus: SignoffStatus: -1 incomplete, 1- complete, 0-Default value
                    Dim strSQL = "update AML_StationSignoffStatus set signoffstatus=" & stationSignoffStatus & ",LastUpdated=getdate(), " &
                                " USERID='" & Strings.ReplaceSingleQuote(userID.Trim) & "' where BuildNumber=" & buildNo.Trim &
                                " and " & Ec.GeneralFunc.GetQueryFieldCondition("StationNO", stationNo) &
                                " and OPERATORPOSITION=" & operatorPosition.ToString
                    connection.ExecuteNonQuery(strSQL)

                    strSQL = "update AML_BuildPartsSerials set SerialStatus=" & partVerifySerialStatus & ", " &
                        "USERID ='" & Strings.ReplaceSingleQuote(userID) & "',lastupdated=getdate() " &
                        " where BuildNumber=" & buildNo & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", stationNo) &
                        " and SerialStatus=-1 " &
                        " and OPERATORPOSITION=" & operatorPosition.ToString
                    connection.ExecuteNonQuery(strSQL)

                    'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
                    strSQL = "update AML_BuildPartsVerify set verifystatus=" & partVerifySerialStatus & ", " &
                        "USERID ='" & Strings.ReplaceSingleQuote(userID) & "',lastupdated=getdate() " &
                        " where BuildNumber=" & buildNo & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", stationNo) &
                        " and verifystatus=-1 " &
                        " and OPERATORPOSITION=" & operatorPosition.ToString
                    connection.ExecuteNonQuery(strSQL)

                    If BootBagStation(stationNo) Then
                        strSQL = "update AML_BuildBootBag set verifystatus=" & partVerifySerialStatus & ", " &
                                    "USERID ='" & Strings.ReplaceSingleQuote(userID) & "',lastupdated=getdate() " &
                                    " where BuildNumber=" & buildNo & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", stationNo) &
                                    " and verifystatus=-1 " &
                                    " and OPERATORPOSITION=" & operatorPosition.ToString
                    End If


                    If wipComplete Then
                        'complete the build# from the WIP.
                        strSQL = GetWIPCompleteSQL(buildNo)
                        connection.ExecuteNonQuery(strSQL)
                    End If
                    connection.CommitTransaction()
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Throw
                End Try
            End Using
        Catch ex As Exception
            Call GenerateException("UpdateStationSignoffStatus", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function ValidateLeoniRCMPartCheck(strBuildNo As String, strStationNO As String, strScannedPartNO As String,
                                                ByRef strScannedPartNOResult As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) strScannedPartNO:({2})",
                                                          strBuildNo, strStationNO, strScannedPartNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnMatchingSerialFound As Boolean = False
        'Dim objBPS(0) As stBuildPartSerials
        ' Dim strSupplierPartNO As String = "",
        Dim strRtnValue As String = ""
        Dim blnMatch As Boolean = False ', strAMLPartNO As String = ""
        Dim strSerialFormat As String = "", strTemp2 As String = ""
        Dim strCheckChar As String = ""
        Dim strTempSS As String = ""
        Dim strPP As String = "", strAlpha As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Try

            strScannedPartNOResult = ""

            ''get the part serial record
            'objBPS = GetBuildNumberPartSerialList(strBuildNo, strStationNO, strScannedPartNO, , intSerialSEQ)
            'If UBound(objBPS) = 0 Then  'never happens, just in case
            '    strRtnValue = "Unable to get the Serial record from AML_BuildPartsSerials table, Build# " & strBuildNo.Trim & ", Station: " & strStationNO.Trim & ", Part#: " & strScannedPartNO.Trim & ", Serial Seq: " & intSerialSEQ.ToString
            '    GoTo ExitThisSub
            'End If


            'check the supplier partno


            'strSupplierPartNO = Ec.GeneralFunc.QPTrim(objBPS(1).Supplier_PartNO)
            'strAMLPartNO = Ec.GeneralFunc.QPTrim(objBPS(1).AML_PartNO)
            'strSerialFormat = Ec.GeneralFunc.QPTrim(objBPS(1).Serial_Format)

            'Leoni Harness
            'Barcode number - 20120503114707002832209500
            'GoldCrest Part number - 2832209
            '2 in position 17, 8 in position 18

            If strScannedPartNO.Trim.Length > 17 Then
                strRtnValue = ""

                If strScannedPartNO.Substring(16, 1) = "2" And strScannedPartNO.Substring(17, 1) = "8" Then
                    strTemp2 = Right(strScannedPartNO.Trim, strScannedPartNO.Trim.Length - 16)
                    If strTemp2.Length > 7 Then
                        strTemp2 = Left(strTemp2, 7)

                        strScannedPartNOResult = strTemp2
                        blnMatchingSerialFound = True
                        GoTo ExitThisSub
                    Else
                        strRtnValue = "Leoni Part#: Can't find the matching part#. Serial: " & strScannedPartNO.Trim & ", Leoni: " & strTemp2
                    End If

                    'If strTemp2.Trim.ToLower = strAMLPartNO.Trim.ToLower Or strTemp2.Trim.ToLower = strSupplierPartNO.Trim.ToLower Then
                    '    strScannedPartNOResult = strTemp2
                    '    blnMatchingSerialFound = True
                    '    GoTo ExitThisSub
                    'Else
                    '    strRtnValue = "Leoni Part#: Can't find the matching part#. Serial: " & strScannedPartNO.Trim & ", Leoni: " & strTemp2
                    'End If

                Else
                    strRtnValue = "Leoni Part#: 17th position doesn't contain '2' and/or 18th position doesn't have '8', Part#: " & strScannedPartNO.Trim
                End If

            End If

            'RCM Modules: F8V or F9V followed by 1 numeric, 1 alpha and 2 numeric characters.
            'Test Part: F8V1A00TESTRCM, F9V1A00TESTRCMTEST,F8VABCDEFGHIJKL

            If strScannedPartNO.Trim.ToLower.Contains("f8v") Or strScannedPartNO.Trim.ToLower.Contains("f9v") Then
                strRtnValue = ""
                blnMatch = False
                strCheckChar = "f8v" : strTempSS = ""
                Dim intYY = InStr(strScannedPartNO.Trim.ToLower, strCheckChar)
                If intYY > 0 Then
                    blnMatch = True
                Else
                    strCheckChar = "f9v"
                    intYY = InStr(strScannedPartNO.Trim.ToLower, strCheckChar)
                    blnMatch = True
                End If

                If blnMatch Then
                    strTempSS = "" : blnMatch = False : strTemp2 = ""

                    If strScannedPartNO.Length >= 6 Then 'intYY + 8 Then '4: followed by 1 numeric, 1 alpha and 2 numeric characters: 3: f8v
                        'intYY += 3
                        intYY += 2

                        strTemp2 = strScannedPartNO.Trim.Substring(intYY, 1)  '1st check:  1 NUMERIC

                        blnMatch = False
                        Select Case strTemp2.Trim.ToUpper
                            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                blnMatch = True
                        End Select

                        If blnMatch Then
                            '2nd check:  1 ALPHA
                            blnMatch = False
                            intYY += 1
                            strTemp2 = strScannedPartNO.Trim.Substring(intYY, 1)  '1:  1 NUMERIC

                            If InStr(strAlpha, strTemp2.ToUpper) > 0 Then
                                blnMatch = True
                            Else
                                Select Case strTemp2.Trim.ToUpper
                                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                        blnMatch = True
                                End Select
                            End If

                            If blnMatch Then '3rd check: 1 numeric characters
                                blnMatch = False

                                intYY += 1
                                strTemp2 = strScannedPartNO.Trim.Substring(intYY, 1)  '1:  1 NUMERIC

                                Select Case strTemp2.Trim.ToUpper
                                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                        blnMatch = True
                                End Select


                                If blnMatch Then  '3rd check: Second numeric characters
                                    blnMatch = False

                                    intYY += 1
                                    strTemp2 = strScannedPartNO.Trim.Substring(intYY, 1)  '1:  1 NUMERIC
                                    Select Case strTemp2.Trim.ToUpper
                                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                            blnMatch = True
                                    End Select
                                End If
                            End If
                        Else
                            GoTo DisplayErrorMessage
                        End If
                    Else
DisplayErrorMessage:
                        strRtnValue = "RCM Modules: incorrect data,  1 numeric, 1 alpha and 2 numbers followed by f8v or f9v, Serial: " & strScannedPartNO.Trim
                    End If
                End If

                If blnMatch Then
                    If strCheckChar.Trim = "f8v" Then
                        blnMatchingSerialFound = True
                        strScannedPartNOResult = "31334738"
                    ElseIf strCheckChar.Trim = "f9v" Then
                        blnMatchingSerialFound = True
                        strScannedPartNOResult = "31334739"

                    End If
                End If
            End If

            '20120503114707002832209500

ExitThisSub:
        Catch ex As Exception
            GenerateException("ValidateLeoniRCMPartCheck", ex)
        Finally
            'objBPS = Nothing
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) strScannedPartNOResult:({1})", strRtnValue, strScannedPartNOResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function

    Public Sub WriteAMLLineIndexLog(ByVal intLineID As Integer, ByVal strStationNO As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) strStationNO:({1})", intLineID, strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(keyx) from AML_Andon_IndexLog"

        Try
            'EW-855: VH 5 Body Shop and MA1 station refresh 
            Using connection = New Connection(gStrConnectionString)
                Dim keyX = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                keyX += 1
                strSQL = "insert into AML_Andon_IndexLog(keyx,datex,lineid,stationno) " &
                        " values (" & keyX & ",getdate()," & intLineID & ",'" & strStationNO.Trim & "')"
                connection.ExecuteNonQuery(strSQL)
            End Using
        Catch ex As Exception
            GenerateException("WriteAMLLineIndexLog: " & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub WriteBuildPartVerify(strBuildNO As String, strStationNO As String, strPartNO As String, strUserID As String, intPartVerifyMSeq As Integer)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2}) strUserID:({3}) intPartVerifyMSeq:({4})",
                                                          strBuildNO, strStationNO, strPartNO, strUserID, intPartVerifyMSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Try
            'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
            strSQL = "UPDATE aml_buildpartsverify SET verifystatus = 1, lastupdated = gETDATE(), userid= '" & Strings.ReplaceSingleQuote(strUserID, True) & "' " &
                        " WHERE buildnumber = " & strBuildNO &
                        " AND stationno = '" & strStationNO & "'" &
                        " AND mseq = " & intPartVerifyMSeq

            Ec.IO.RunSQL(strSQL)

        Catch ex As Exception
            GenerateException("WriteBuildPartVerify", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub WriteBuildPartVerifyGLOveride(strBuildNO As String, strStationNO As String, strPartNO As String, strUserID As String, intPartVerifyMSeq As Integer)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2}) strUserID:({3}) intPartVerifyMSeq:({4})",
                                                          strBuildNO, strStationNO, strPartNO, strUserID, intPartVerifyMSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Try
            'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
            strSQL = "UPDATE aml_buildpartsverify SET verifystatus = 2, lastupdated = gETDATE(), userid= '" & Strings.ReplaceSingleQuote(strUserID, True) & "' " &
                        " WHERE buildnumber = " & strBuildNO &
                        " AND stationno = '" & strStationNO & "'" &
                        " AND mseq = " & intPartVerifyMSeq

            Ec.IO.RunSQL(strSQL)

        Catch ex As Exception
            GenerateException("WriteBuildPartVerify", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function WriteImportLogConfigurationData(ByVal ArrEmailList() As String, ByVal strImportDirectory As String) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter ArrEmailList.Length:({0}) strImportDirectory:({1})", ArrEmailList.Length, strImportDirectory), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False
        Dim intStartPos As Int16 = 152

        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    'clear existing data (153-158) , can extend upto 162
                    For intK = 1 To 6
                        strSQL = "update client set data=' ' where keyx=" & intK + intStartPos
                        connection.ExecuteNonQuery(strSQL)
                    Next

                    'strImportDirectory
                    strSQL = "update client set wl=99,data='" & strImportDirectory.Trim & " ' where keyx=163"  'import directory
                    connection.ExecuteNonQuery(strSQL)

                    'now update it
                    For intK = 1 To UBound(ArrEmailList)
                        strSQL = "update client set wl=99,data='" & Strings.ReplaceSingleQuote(ArrEmailList(intK)) & " ' where keyx=" & intK + intStartPos
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
            Call GenerateException("WriteLogEmailSendersList", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit blnResult:({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult

    End Function

    Public Function WriteLineStatus(ByVal objLS() As stLineStatus,
                    Optional ByVal updateAndOnReadStatus As Boolean = False,
                    Optional ByVal lineId As Integer = 0,
                    Optional ByVal blnUpdateWIPStatus As Boolean = False,
                    Optional ByVal blnIndependentStations As Boolean = False) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objLS.Length:({0}) blnUpdateAndonReadStatus:({1}) intLineID:({2}) blnUpdateWIPStatus:({3}) blnIndependentStations:({4})",
                                                          objLS.Length, updateAndOnReadStatus, lineId, blnUpdateWIPStatus, blnIndependentStations), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False
        Dim dtDateX As Date = Date.Now
        Dim strLineStatusTableName As String = GetLineStatusTableName(lineId)
        Dim strStationNo = ""

        Try
            Dim sqlList = New List(Of String)()
            For intK = 1 To UBound(objLS)

                If objLS(intK).LineID = 0 Then GoTo SkipThisRecord 'the line id should never be 0

                If lineId > 0 AndAlso objLS(intK).LineID <> lineId Then
                    GoTo SkipThisRecord
                End If

                strStationNo = Trim(objLS(intK).StationNO)

                ' TODO(crhodes)
                ' Until we understand how IPCStatus, AndonCALL, and other fields in AML_LINESTATUS_BODYSHOP are used
                ' let's just skip updating anything in the table and just let Premier manage the table.

                If (Ec.AppConfig.AstonMartinStAthan AndAlso (lineId = BodyShopLineId OrElse lineId = MainAssemblyLineId OrElse lineId = PaintLineId)) Then
                    strSQL = $"UPDATE {strLineStatusTableName} SET ipcstatus = {objLS(intK).QualityChecksStatus}, andoncall = {objLS(intK).AndonCall} WHERE lineid = {lineId} AND stationno = '{strStationNo}'"
                    sqlList.Add(strSQL)
                ElseIf lineId = BodyShopLineId Then
                    'Gaydon should not update this line
                Else
                    strSQL = $"delete from {strLineStatusTableName} where lineid = {objLS(intK).LineID} and StationNO='{strStationNo}'"
                    sqlList.Add(strSQL)

                    If IsDate(objLS(intK).EngineMoveDate) AndAlso objLS(intK).EngineMoveDate <> Date.MinValue Then
                        dtDateX = objLS(intK).EngineMoveDate
                    End If

                    Dim strTemp = Format(dtDateX, "MMM dd yyyy hh:mm:ss:ffftt")        'get the date in MMM dd yyyy hh:mmtt format
                    strTemp = " CONVERT(DATETIME,'" & strTemp.Trim & "',109) "          '100 - sql server format: mon dd yyyy hh:mi:ss:mmmAM (or PM)
                    Dim intBuildNo = If(Trim(objLS(intK).BuildNO) = "", 0, Trim(objLS(intK).BuildNO).ToInt32())
                    strSQL = "insert into " & strLineStatusTableName & "(LineID,StationNO,Buildno,EngineMoveDate,IPCStatus,AndonCALL,readstatus) values(" &
                            objLS(intK).LineID & ",'" & strStationNo & "'," & intBuildNo & "," & strTemp & "," &
                            objLS(intK).QualityChecksStatus & "," & objLS(intK).AndonCall & ",0)"

                    sqlList.Add(strSQL)
                End If

SkipThisRecord:
            Next

            If updateAndOnReadStatus And lineId > 0 Then
                'AML Andon table holds the data by LINE
                strSQL = "update aml_andon set readstatus=1 where lineid=" & lineId.ToString
                If blnIndependentStations Then
                    'EW-855: VH 5 Body Shop and MA1 station refresh
                    'AML AML_ANDON_Stations table holds the data by Line and Stations
                    strSQL = "update AML_ANDON_Stations set readstatus=1 where lineid=" & lineId.ToString & " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNo)
                End If

                sqlList.Add(strSQL)
            End If
            Ec.IO.RunSQLInArray(sqlList, blnResult)

        Catch ex As Exception
            Call GenerateException("WriteLineStatus: " & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function

    Public Function WriteSerialRecordForPartNumber(buildNo As String,
                                                   stationNo As String,
                                                   partNo As String,
                                                   serialNo As String,
                                                   userId As String,
                                                   operatorPosition As Integer,
                                                   Optional validate As Boolean = False,
                                                   Optional ByRef errorMessage As String = "",
                                                   Optional partMVerifySeq As Integer = 0,
                                                   Optional pageCallFrom As String = "",
                                                   Optional previousSerialNo As String = "") As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) strPartNO:({2}) strSerialNO:({3}) " &
                                                          "strUserID:({4}) intOperatorPosition:({5}) blnValidate:({6}) " &
                                                          "intPartMVerifySeq:({7}) strPageCallFrom:({8}) strPrevSerialNo:({9})",
                                                          buildNo, stationNo, partNo, serialNo,
                                                          userId, operatorPosition, validate,
                                                          partMVerifySeq, pageCallFrom, previousSerialNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As Boolean = False
        Dim checkForIncompleteStatus As Boolean = False
        Dim strPartNO2 As String = ""
        If pageCallFrom.Equals("partverify", StringComparison.CurrentCultureIgnoreCase) Or pageCallFrom.Equals("qcreport", StringComparison.CurrentCultureIgnoreCase) Then
            checkForIncompleteStatus = True
            strPartNO2 = partNo.Trim
        End If
        errorMessage = ""

        'need the part number to get the serial mseq , Need the Part number to get the serial seq
        Dim serialSeq = GetSerialRecordPosition(buildNo, stationNo, strPartNO2, partMVerifySeq, checkForIncompleteStatus)
        If serialSeq > 0 Then

            If validate Then
                errorMessage = ValidateSerialForBuild(buildNo, stationNo, partNo, serialSeq, serialNo, partMVerifySeq)
                If errorMessage.Trim <> "" Then
                    GoTo ExitWithNoSave
                End If
            End If


            'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override
            Dim sql = $"update AML_BuildPartsSerials set SerialStatus=1, SCAN_SERIALNO='{serialNo.ReplaceSingleQuote()}', USERID ='{userId.ReplaceSingleQuote()}',lastupdated=getdate() "

            If previousSerialNo.Trim <> "" Then
                sql &= ", serial_prev1='" & previousSerialNo.Trim & "' "
            End If
            sql &= $" where BuildNumber={buildNo} and {Ec.GeneralFunc.GetQueryFieldCondition("stationno", stationNo)} and mseq= {partMVerifySeq} and serialseq= {serialSeq}"

            Ec.IO.RunSQL(sql)

            returnValue = True
ExitWithNoSave:
        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) strErrorMsg:({1})", returnValue, errorMessage), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function WriteVINRecordForPartNumber(buildNo As String,
                                                   stationNo As String,
                                                   partNo As String,
                                                   serialNo As String,
                                                   userId As String,
                                                   operatorPosition As Integer,
                                                   Optional validate As Boolean = False,
                                                   Optional ByRef errorMessage As String = "",
                                                   Optional partMVerifySeq As Integer = 0,
                                                   Optional pageCallFrom As String = "",
                                                   Optional previousSerialNo As String = "") As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) strPartNO:({2}) strSerialNO:({3}) " &
                                                          "strUserID:({4}) intOperatorPosition:({5}) blnValidate:({6}) " &
                                                          "intPartMVerifySeq:({7}) strPageCallFrom:({8}) strPrevSerialNo:({9})",
                                                          buildNo, stationNo, partNo, serialNo,
                                                          userId, operatorPosition, validate,
                                                          partMVerifySeq, pageCallFrom, previousSerialNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As Boolean = False
        Dim checkForIncompleteStatus As Boolean = False
        Dim strPartNO2 As String = ""
        If pageCallFrom.Equals("partverify", StringComparison.CurrentCultureIgnoreCase) Or pageCallFrom.Equals("qcreport", StringComparison.CurrentCultureIgnoreCase) Then
            checkForIncompleteStatus = True
            strPartNO2 = partNo.Trim
        End If
        errorMessage = ""

        'need the part number to get the serial mseq , Need the Part number to get the serial seq
        Dim serialSeq = GetSerialRecordPosition(buildNo, stationNo, strPartNO2, partMVerifySeq, checkForIncompleteStatus)
        If serialSeq > 0 Then

            If validate Then
                errorMessage = ValidateSerialForBuild(buildNo, stationNo, partNo, serialSeq, serialNo, partMVerifySeq)
                If errorMessage.Trim <> "" Then
                    GoTo ExitWithNoSave
                End If
            End If


            If Ec.IO.AnyRecords($"select * from AML_BuildPartsSerials where BuildNumber={buildNo} and {Ec.GeneralFunc.GetQueryFieldCondition("stationno", stationNo)} and mseq= {partMVerifySeq} and serialseq= {serialSeq}") Then
                'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override
                Dim sql = $"update AML_BuildPartsSerials set SerialStatus=1, SCAN_SERIALNO='{serialNo.ReplaceSingleQuote()}', USERID ='{userId.ReplaceSingleQuote()}',lastupdated=getdate() "

                If previousSerialNo.Trim <> "" Then
                    sql &= ", serial_prev1='" & previousSerialNo.Trim & "' "
                End If
                sql &= $" where BuildNumber={buildNo} and {Ec.GeneralFunc.GetQueryFieldCondition("stationno", stationNo)} and mseq= {partMVerifySeq} and serialseq= {serialSeq}"

                Ec.IO.RunSQL(sql)
            Else
                'Add record if not existing.  Used for VC tracking
                If pageCallFrom.Equals("partverify", StringComparison.CurrentCultureIgnoreCase) Then
                    Dim sql = $"insert into AML_BuildPartsSerials (BuildNumber, StationNO, PartNumber, MSeq, SerialSeq, userId, lastupdated, datecreated, operatorposition, serialstatus,SCAN_SERIALNO) "
                    sql &= $"values({buildNo},'{stationNo}','{partNo}',{partMVerifySeq},{serialSeq},'{userId.ReplaceSingleQuote()}',getdate(),getdate(),{operatorPosition},1,'{serialNo.ReplaceSingleQuote()}')"
                    Ec.IO.RunSQL(sql)
                End If
            End If

            returnValue = True
ExitWithNoSave:
        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) strErrorMsg:({1})", returnValue, errorMessage), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Sub WriteBuildPartSerialsRecordForPartNumber(buildNo As String,
                                                   stationNo As String,
                                                   partNo As String,
                                                   serialNo As String,
                                                   userId As String,
                                                   operatorPosition As Integer,
                                                   Optional validate As Boolean = False,
                                                   Optional ByRef errorMessage As String = "",
                                                   Optional partMVerifySeq As Integer = 0,
                                                   Optional pageCallFrom As String = "",
                                                   Optional previousSerialNo As String = "")

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo:({0}) strStationNO:({1}) strPartNO:({2}) strSerialNO:({3}) " &
                                                          "strUserID:({4}) intOperatorPosition:({5}) blnValidate:({6}) " &
                                                          "intPartMVerifySeq:({7}) strPageCallFrom:({8}) strPrevSerialNo:({9})",
                                                          buildNo, stationNo, partNo, serialNo,
                                                          userId, operatorPosition, validate,
                                                          partMVerifySeq, pageCallFrom, previousSerialNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim checkForIncompleteStatus As Boolean = False
        Dim strPartNO2 As String = ""

        errorMessage = ""

        'need the part number to get the serial mseq , Need the Part number to get the serial seq
        Dim serialSeq = GetSerialRecordPosition(buildNo, stationNo, strPartNO2, partMVerifySeq, checkForIncompleteStatus)
        If serialSeq > 0 Then

            'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override
            Dim sql = $"update AML_BuildPartsSerials set SerialStatus=1, SCAN_SERIALNO='{serialNo.ReplaceSingleQuote()}', USERID ='{userId.ReplaceSingleQuote()}',lastupdated=getdate() "

            If previousSerialNo.Trim <> "" Then
                sql &= ", serial_prev1='" & previousSerialNo.Trim & "' "
            End If
            sql &= $" where BuildNumber={buildNo} and {Ec.GeneralFunc.GetQueryFieldCondition("stationno", stationNo)} and serialseq= {serialSeq}"

            Ec.IO.RunSQL(sql)

        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit WriteBuildPartSerialsRecordForPartNumber"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub


    Public Sub WriteSerialScanStationsList(lineId1 As Int16, station1 As String, lineId2 As Int16, station2 As String,
                                           lineId3 As Int16, station3 As String)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID1:({0}) strStation1:({1}) intLineID2:({2}) strStation2:({3}) intLineID3:({4}) strStation3:({5})",
                                                          lineId1, station1, lineId2, station2, lineId3, station3), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        Try
            Dim sqlList = New List(Of String)()

            'Keyx: 237, 238, 239, 240, 241 : AML: ESN Serial checks  : Holds station list
            'EW-454: OIW - Sub-header Detailed Description
            strSQL = "UPDATE CLIENT set DATA='" & station1.Trim & "'," &
                     " WL= " & CStr(lineId1) & "," &
                     " L=0 ," &
                     " TYPE=0," &
                     " DP= 99" &
                     " where KEYX=237"
            sqlList.Add(strSQL)

            strSQL = "UPDATE CLIENT set DATA='" & station2.Trim & "'," &
                     " WL= " & CStr(lineId2) & "," &
                     " L=0 ," &
                     " TYPE=0," &
                     " DP= 99" &
                     " where KEYX=238"
            sqlList.Add(strSQL)

            strSQL = "UPDATE CLIENT set DATA='" & station3.Trim & "'," &
                     " WL= " & CStr(lineId3) & "," &
                     " L=0 ," &
                     " TYPE=0," &
                     " DP= 99" &
                     " where KEYX=239"
            sqlList.Add(strSQL)
            Ec.IO.RunSQLInArray(sqlList)

        Catch ex As Exception
            Call GenerateException("WriteSerialScanStationsList", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub WriteStationSignoffRecords_Build(intLineID As Integer, strBuildNO As String,
                                                     Optional strProcessOneStationNO As String = "")

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) strBuildNO:({1}) strProcessOneStationNO:({2})",
                                                          intLineID, strBuildNO, strProcessOneStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'write the station signoff records for Build#

        Dim objLineStations(0) As stStations
        Dim strSQL As String = ""
        Dim intTotalOperators = 1 'GetMaxNumberOfOperatorPositions()
        Try
            'get the list of stations
            If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations
                'to avoid database call

                ReDim objLineStations(1)
                objLineStations(1).LineID = intLineID
                objLineStations(1).StationNO = strProcessOneStationNO
                objLineStations(1).Operators = GetOperatorCountForStation(intLineID, strProcessOneStationNO)

            Else
                'default option
                'get the list of stations
                objLineStations = GetLineStations(intLineID)
            End If

            If UBound(objLineStations) = 0 Then
                GoTo SkipWritingRecords
            End If

            Dim sqlList = New List(Of String)()

            For intK = 1 To UBound(objLineStations)

                If strProcessOneStationNO.Trim <> "" Then  'processing records for one station/ manual stations
                    If objLineStations(intK).StationNO.Trim.ToLower <> strProcessOneStationNO.Trim.ToLower Then
                        GoTo SkipThisStationRecord
                    End If
                End If

                intTotalOperators = objLineStations(intK).Operators

                'delete existing record.
                strSQL = "delete from AML_StationSignoffStatus where BuildNumber=" & strBuildNO.Trim &
                            " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", objLineStations(intK).StationNO)
                sqlList.Add(strSQL)

                For intOPPos = 1 To intTotalOperators
                    'AML_StationSignoffStatus: SignoffStatus: -1 incomplete, 1- complete, 0-Default value
                    strSQL = "insert into AML_StationSignoffStatus(BuildNumber,stationno,lastupdated,signoffstatus,OPERATORPOSITION) values (" &
                        strBuildNO.Trim & ",'" & objLineStations(intK).StationNO.Trim & "',getdate(),0," & intOPPos & ")"
                    sqlList.Add(strSQL)
                Next intOPPos
SkipThisStationRecord:
            Next intK
            Ec.IO.RunSQLInArray(sqlList)

SkipWritingRecords:
        Catch ex As Exception
            GenerateException("WriteStationSignoffRecords", ex)
        Finally
            objLineStations = Nothing
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub WriteUserLastVisitedScreen(ByVal objULVS As stUserLastVisitedScreen)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objULVS.BuildNo:({0}) objULVS.StationNO:({1}) objULVS.ID:({2}) objULVS.RecType:({3}) objULVS.SEQ:({4}) objULVS.OPNO:({5}) objULVS.IPAddress:({6})",
                                                          objULVS.BuildNo, objULVS.StationNO, objULVS.ID, objULVS.RecType, objULVS.SEQ, objULVS.OPNO, objULVS.IPAddress), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    strSQL = "delete from AMLSwitchLogin where " &
                             Ec.GeneralFunc.GetQueryFieldCondition("ipaddress", objULVS.IPAddress) &
                             " and " & Ec.GeneralFunc.GetQueryFieldCondition("userid", objULVS.UserID) &
                             " and " & Ec.GeneralFunc.GetQueryFieldCondition("buildno", objULVS.BuildNo) &
                             " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", objULVS.StationNO) &
                             " and operatorposition=" & objULVS.OperatorPosition

                    connection.ExecuteNonQuery(strSQL)

                    If objULVS.IPAddress.Trim <> "" Then
                        Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(EASEClass7.QueryBuilder.QueryType.Insert,
                                                                                  "AMLSwitchLogin")
                        queryBuilder.AddField("ipaddress", objULVS.IPAddress, False)
                        queryBuilder.AddField("userid", objULVS.UserID, False)
                        queryBuilder.AddField("buildno", objULVS.BuildNo, False)
                        queryBuilder.AddField("stationno", objULVS.StationNO, False)
                        queryBuilder.AddField("datex", "getdate()", True)
                        queryBuilder.AddField("id", objULVS.ID, True)
                        queryBuilder.AddField("rectype", objULVS.RecType, False)
                        queryBuilder.AddField("seq", objULVS.SEQ, True)
                        queryBuilder.AddField("OPNO", objULVS.OPNO, False)
                        queryBuilder.AddField("subheader", objULVS.SubHeader, False)
                        queryBuilder.AddField("pageid", objULVS.PageID, True)
                        queryBuilder.AddField("spare1", objULVS.Spare1, True)
                        queryBuilder.AddField("operatorposition", objULVS.OperatorPosition, True)

                        strSQL = queryBuilder.GenerateQuery
                        connection.ExecuteNonQuery(strSQL)

                    End If

                    connection.CommitTransaction()
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Throw ex
                End Try
            End Using
        Catch ex As Exception
            Call GenerateException("WriteUserLastVisitedScreen", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function WriteWIPData(ByVal objWIPData As stWIPData) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter objWIPData.EngineID:({0}) objWIPData.LineID:({1}) objWIPData.BuildNo:({2}) objWIPData.StartDate:({3}) objWIPData.Engineer:({4}) objWIPData.ModelNumber:({5})",
                                                          objWIPData.EngineID, objWIPData.LineID, objWIPData.BuildNo, objWIPData.StartDate, objWIPData.Engineer, objWIPData.ModelNumber), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnRtnValue As Boolean = False
        Dim intEngineID As Int32 = 0
        Try
            intEngineID = GetWIPEngineID()
            strSQL = "insert into MES_ENGINESLIST(engineid,lineid,engineno,enginecomplete,startdate,startengineer,Spare1,spare2 ) values (" &
            intEngineID & "," & objWIPData.LineID & ",'" & Strings.ReplaceSingleQuote(objWIPData.BuildNo) & "',0,getdate()" & ",'" & Strings.ReplaceSingleQuote(objWIPData.Engineer, True) & "'," & objWIPData.ModelNumber & "," & objWIPData.LineID & ")"


            Ec.IO.RunSQL(strSQL)

            blnRtnValue = True
        Catch ex As Exception
            GenerateException("GetWIPData", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue

    End Function

    Public Function GetStationDefaultLocation(ByVal id As Integer) As AML_StationDefaultLocation

        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.GetDataIntoClassOf(Of AML_StationDefaultLocation)(String.Format("SELECT * FROM aml_stationdefaultlocations WHERE stationdefaultid = {0}", id))
            If (IsNothing(result)) Then Return Nothing
            Return result.FirstOrDefault()
        End Using

    End Function

    Public Function GetStationDefaultLocation(ByVal ipAddress As String) As AML_StationDefaultLocation

        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.GetDataIntoClassOf(Of AML_StationDefaultLocation)(String.Format("SELECT * FROM aml_stationdefaultlocations WHERE ipaddress = '{0}'", ipAddress))
            If (IsNothing(result)) Then Return Nothing
            Return result.FirstOrDefault()
        End Using

    End Function

    Public Function GetSharedStation(ByVal lineId As Integer, ByVal station As String) As String
        If String.IsNullOrEmpty(station) OrElse String.IsNullOrEmpty(gStrConnectionString) Then Return ""

        Using connection = New Connection(gStrConnectionString)
            Dim sharedLocation = connection.ExecuteScalar(Of String)(String.Format("SELECT absnno FROM stations WHERE lineid='{0}' and sharedstation = '{1}'", lineId, station), "").Trim

            If sharedLocation Is Nothing Then Return "" Else Return sharedLocation
        End Using

    End Function

    Public Function GetStationDefaultLocations() As List(Of AML_StationDefaultLocation)

        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.GetDataIntoClassOf(Of AML_StationDefaultLocation)("SELECT * FROM aml_stationdefaultlocations")
            If (IsNothing(result)) Then Return Nothing
            Return result.ToList()
        End Using

    End Function

    Public Function SaveStationDefaultLocation(ByVal item As AML_StationDefaultLocation) As String
        Using connection = New Connection(gStrConnectionString)
            Return SaveStationDefaultLocation(connection, item)
        End Using
    End Function

    Public Function SaveStationDefaultLocation(ByVal connection As Connection, ByVal item As AML_StationDefaultLocation) As String

        If (IsNothing(item)) Then Return ""
        If (String.IsNullOrEmpty(item.IPAddress)) Then Return ""

        Dim existInStationNo = connection.ExecuteScalar(Of String)(String.Format("SELECT stationno FROM aml_stationdefaultlocations WHERE ipaddress = '{0}' AND stationno <> '{1}'", item.IPAddress, item.StationNo), "")
        If existInStationNo <> "" Then
            Return String.Format("IP Address: {0} already exists for Station No: {1}", item.IPAddress, existInStationNo)
        End If

        Dim sql = ""
        If (item.StationDefaultID = 0) Then
            sql = String.Format("INSERT INTO aml_stationdefaultlocations VALUES ('{0}', '{1}', {2}, '{3}')", item.Area, item.StationNo, item.LineID, item.IPAddress)
        Else
            sql = String.Format("UPDATE aml_stationdefaultlocations SET ipaddress = '{0}', stationno = '{1}', area = '{3}', lineid = {4} WHERE stationdefaultid = {2}", item.IPAddress, item.StationNo, item.StationDefaultID, item.Area, item.LineID)
        End If
        connection.ExecuteNonQuery(sql)

        Return ""

    End Function

#End Region

#Region "Utility Methods"


    Public Sub ClearBannerMessageObject(ByRef objBML As stBannerMsgList)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objBML.MessageID = 0
        objBML.MessageText = ""
        objBML.RaisedBy = ""
        objBML.ApprovedBy = ""
        objBML.MessageReason = ""
        objBML.MessageType = ""
        objBML.LineID = 1
        objBML.StationFrom = Date.MinValue.ToString()
        objBML.StationTo = Date.MinValue.ToString()
        objBML.StartDate = Date.MinValue
        objBML.EndDate = Date.MinValue
        objBML.LastUpdated = Date.MinValue
        objBML.MessageStatus = 0        'MessageStatus: 0-Draft/Pending Approval, 1-Approved


#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub ClearLineStatusObject(ByRef objLS As stLineStatus)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objLS.LineID = 0
        objLS.StationNO = ""
        objLS.BuildNO = ""
        objLS.EngineMoveDate = Date.MinValue
        objLS.QualityChecksStatus = 0     '-1 -> InComplete, 0-FAIL, 1->PASS
        objLS.AndonCall = 0                 '0-default, 1-first call, 2-second call 
        objLS.OperatorCount = 1 'default number of operator positions for the stations ' AfterEAESClass7 Consolidation:7th August

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub ClearPartVerifyListObject(ByRef objPVL As stPartsVerifyList)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objPVL.PartNumber = ""
        'objPVL.GRN_Req = ""
        'objPVL.Verify_Station = ""
        'objPVL.Serial_Req = ""
        'objPVL.Serial_Format = ""
        'objPVL.Serial_AutoGen = ""
        objPVL.Supplier_PartNO = ""

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub ClearPSSDocument_Read_Object(ByRef objPSSRead As stPSSDocuments_READ)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objPSSRead.USERID = ""
        objPSSRead.StationNO = ""
        objPSSRead.LineID = 0
        objPSSRead.OPNO = ""
        objPSSRead.MDMDocID = 0
        objPSSRead.TTKey = 0
        objPSSRead.CompleteTTKey = 0
        objPSSRead.CheckedRead = 0
        objPSSRead.PassStatus = -1   ' AML_PSSDocsCheck:PassStatus -> 1-Pass, -1-Incomplete
        objPSSRead.DateRead = Date.MinValue

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub ClearQualityChecksSearchObject(ByRef objQCS As stQualityChecksSearch)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objQCS.BuildNo = ""
        objQCS.OPNO = ""
        objQCS.PartNO = ""
        objQCS.PassFail = -22       '-1 -> InComplete, 0-FAIL, 1->PASS
        objQCS.PlanType = -1
        objQCS.RecordSeq = 0
        objQCS.RecordType = -1      'Record type: 0-Quality checks/Remedy Plan ,1- FLM record
        objQCS.StationNO = ""
        objQCS.SubHeader = ""
        objQCS.Shift = 0

        'objQCS.FromDateX = Date.MinValue
        'objQCS.ToDateX = Date.MinValue

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub ClearSafetyCheckObject(ByRef objSafetyCheck As stSafetyCheck)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objSafetyCheck.OperatorID = ""
        objSafetyCheck.WSID = 0
        objSafetyCheck.WSCode = ""
        objSafetyCheck.WSDescX = ""
        objSafetyCheck.DOCID = 0
        objSafetyCheck.DocRecType = "0"
        objSafetyCheck.DocSeq = 0
        objSafetyCheck.DocDesc = ""
        objSafetyCheck.RevNum = 1
        objSafetyCheck.RevDate = Date.MinValue
        objSafetyCheck.DateRead = "" 'Date.MinValue
        objSafetyCheck.CheckedRead = False
        objSafetyCheck.PassStatus = -1 'default to cancel
        objSafetyCheck.UpdateRecord = True


        objSafetyCheck.PlantID = -1
        objSafetyCheck.WorkgroupID = ""
        objSafetyCheck.WC = ""
        objSafetyCheck.Engineer = ""
        objSafetyCheck.DocStatus = -1
        objSafetyCheck.EnggFunction = -1
        objSafetyCheck.MSeq = 0
        objSafetyCheck.InfoType = -1
        objSafetyCheck.Filename = ""
        objSafetyCheck.FilepathX = ""
        objSafetyCheck.TTKey = 0
        objSafetyCheck.CompleteTTKey = 0
        objSafetyCheck.KeyField = ""
        objSafetyCheck.KeyFieldValue = ""
        objSafetyCheck.PrintOrder = 0

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Sub ClearUserLastVisitedScreenObject(ByRef objULVS As stUserLastVisitedScreen)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objULVS.IPAddress = ""
        objULVS.UserID = ""
        objULVS.BuildNo = ""
        objULVS.StationNO = ""
        objULVS.ID = 0
        objULVS.RecType = ""
        objULVS.SEQ = 0
        objULVS.OPNO = ""
        objULVS.SubHeader = ""
        objULVS.PageID = 1      'home page
        objULVS.Spare1 = 0
        objULVS.OperatorPosition = 0

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub
    Public Sub ClearVehicleBOMObject(ByRef objVBOM As stVehBOM)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objVBOM.Body = ""
        objVBOM.CPSC = ""
        objVBOM.Drive = ""
        objVBOM.Gear = ""
        objVBOM.LineNumber = 0
        objVBOM.NeworMod = ""
        objVBOM.OptionCode = ""
        objVBOM.PartNumber = ""
        objVBOM.StationNumber = ""
        objVBOM.Performance = ""
        objVBOM.ModelYear = ""
        objVBOM.Model1 = 0
        objVBOM.PartDescription = ""
        objVBOM.TorqueMax = ""
        objVBOM.TorqueMin = ""
        objVBOM.Quantity = ""
        objVBOM.KitNO = ""

        objVBOM.Validation = ""
        objVBOM.Tracking = ""
        objVBOM.SupplierCode = ""

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub
    Public Sub ClearWIPDataObject(ByRef objWIPData As stWIPData)

#If TRACE Then
        Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        objWIPData.ModelNumber = 0
        objWIPData.BuildNo = ""
        objWIPData.EngineID = 0
        objWIPData.StartDate = Date.MinValue
        objWIPData.Engineer = ""
        objWIPData.LineID = 0

#If TRACE Then
        Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function FormatDateForQuery(ByVal dtDate As Date, Optional ByVal blnIncludeTime As Boolean = False) As String
#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_MED(String.Format("Enter dtDate:({0}) blnIncludeTime:({1})", dtDate, blnIncludeTime), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strRtnValue As String = ""
        Dim strDateFormat As String = "MM/dd/yyyy" '** DON'T CHANGE THIS, USED IN SQL, MAY BE DIFFERENT FORMAT IN ORACLE/SQL SERVER DBS **

        Select Case Connection.GetDatabaseTypeFromConnectionString(gStrConnectionString)
            Case Connection.DatabaseTypes.Oracle, Connection.DatabaseTypes.OracleManagedProvider, Connection.DatabaseTypes.OracleManagedProviderBridge
                strRtnValue = " TO_Date('" & dtDate.ToString(strDateFormat) & "','MM/DD/YYYY') "

                If blnIncludeTime Then
                    strRtnValue = " TO_Date('" & dtDate.ToString(strDateFormat) & " 0:0:0','MM/DD/YYYY HH24:MI:SS') "
                End If
            Case Else
                strRtnValue = " convert(datetime,'" & dtDate.ToString(strDateFormat) & "',101) "  ''MM/DD/YYYY'
        End Select


#If TRACE Then
        Log.UTILITY_MED(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function

    Private Function GetDateBuiltInFunction() As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Database type (0-access, 1-Microsoft Oracle .NET Provider, 2 - Sql Server, 3- Oracle .NET Provider)
        Dim strRtnValue As String = "sysdate"
        Select Case Connection.GetDatabaseTypeFromConnectionString(gStrConnectionString)
            Case Connection.DatabaseTypes.Oracle, Connection.DatabaseTypes.OracleManagedProvider, Connection.DatabaseTypes.OracleManagedProviderBridge
            Case Else
                strRtnValue = "getdate()"
        End Select


#If TRACE Then
        Log.UTILITY_MED(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function

    Public Function GetDateFormat(Optional ByVal blnIncludeTime As Boolean = False) As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_MED(String.Format("Enter blnIncludeTime:({0})", blnIncludeTime), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strDateFormat As String = "MM/dd/yyyy"
        strDateFormat = "MMM-dd-yyyy"
        If blnIncludeTime Then
            strDateFormat &= " hh:mm:ss tt"
        End If


#If TRACE Then
        Log.UTILITY_MED(String.Format("Exit ({0})", strDateFormat), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strDateFormat
    End Function

    Public Function GetMaximumMessageID() As Int32

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(msgid) from BannerMessages"
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        returnValue += 1

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetRandomNumber() As String

#If TRACE Then
        Dim startTicks As Long = Log.UTILITY_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim dtDate As Date = Now, strTemp As String = ""
        strTemp = dtDate.Year & dtDate.Month & dtDate.Day & dtDate.Hour & dtDate.Minute & dtDate.Second & dtDate.Millisecond


#If TRACE Then
        Log.UTILITY_MED(String.Format("Exit ({0})", strTemp.Trim()), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return Trim(strTemp)
    End Function

    Public Function GetRecordCountQuery(ByVal strSQL As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter strSQL:({0})", strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'get the complete sql and replace all fields with count(1) and return the sQL
        'EX: Input is 'select partno from partxref', Output is select count(1) from partxref
        Dim strSQL2 As String = "", strTemp As String = ""
        Dim intPos2 As String = ""

        Try
            strSQL = Trim(strSQL)
            If Left(LCase(strSQL), 6) <> "select" Then          'never happens, defensive coding (to avoid typos)
                strSQL2 = strSQL
                GoTo ExitThisSub
            End If

            Dim intPos = InStr(LCase(strSQL), "from")
            strTemp = Right(strSQL, (strSQL.Length - intPos) + 1)


            intPos = InStr(LCase(strTemp), " order by")
            If intPos > 0 Then
                strTemp = Left(strTemp, intPos)
            End If


            intPos = InStr(LCase(strTemp), " group by")
            If intPos > 0 Then
                strTemp = Left(strTemp, intPos)
            End If



            strSQL2 = "select count(1) " & strTemp.Trim ' Right(strSQL, (strSQL.Length - intPos) + 1)


ExitThisSub:


        Catch ex As Exception
            GenerateException("GetRecordCountQuery", ex)
        End Try


#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strSQL2), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strSQL2

    End Function

    Public Function TubsLine(intLineID As Integer) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0})", intLineID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnResult As Boolean
        blnResult = (intLineID = 7)


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult
    End Function

#End Region

#Region "Protected Methods"


#End Region

#Region "Private Methods"
    Private Function ValidatePartShortageReason(ByVal [shortageReason] As PartShortageReason) As String
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = ""
        Dim list = GetPartShortageReasons()
        If (shortageReason.ShortageReasonID > 0) Then
            If (list.Where(Function(x) x.ShortageReason.Equals(shortageReason.ShortageReason, StringComparison.CurrentCultureIgnoreCase) AndAlso x.ShortageReasonID <> shortageReason.ShortageReasonID).Count() > 0) Then
                result = String.Format("Part Shortage Reason: {0} already exists.", shortageReason.ShortageReason)
            End If
        Else
            If (list.Where(Function(x) x.ShortageReason.Equals(shortageReason.ShortageReason, StringComparison.CurrentCultureIgnoreCase)).Count() > 0) Then
                result = String.Format("Part Shortage Reason: {0} already exists.", shortageReason.ShortageReason)
            End If
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit result:({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result
    End Function



    Private Function CheckForSHSpecificOptions(ByVal vehicleOptions As List(Of AML_VehicleOption), ByVal strSpecific1 As String, ByVal strSpecific2 As String, ByVal strSpecific3 As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strSpecific1:({0}) strSpecific2:({1}) strSpecific3:({2})",
                                                          strSpecific1, strSpecific2, strSpecific3), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False
        'no specific options configured for the submodel
        If strSpecific1.Trim = "" And strSpecific2.Trim = "" And strSpecific3.Trim = "" Then
            blnRtnValue = True
            GoTo ExitThisFunction
        End If
        Dim blnSpec1 As Boolean = False, blnSpec2 As Boolean = False, blnSpec3 As Boolean = False


        For Each vehicleOption In vehicleOptions

            If blnSpec1 Then GoTo CheckForSpec2
            If strSpecific1.Trim = "" Then
                blnSpec1 = True
            Else
                If strSpecific1.Trim.ToLower = vehicleOption.OptionCode1.Trim.ToLower Then
                    blnSpec1 = True
                End If
            End If
CheckForSpec2:

            If blnSpec2 Then GoTo CheckForSpec3
            If strSpecific2.Trim = "" Then
                blnSpec2 = True
            Else
                If strSpecific2.Trim.ToLower = vehicleOption.OptionCode1.Trim.ToLower Then
                    blnSpec2 = True
                End If
            End If

CheckForSpec3:
            If strSpecific3.Trim = "" Then
                blnSpec3 = True
            Else
                If strSpecific3.Trim.ToLower = vehicleOption.OptionCode1.Trim.ToLower Then
                    blnSpec3 = True
                End If
            End If

            If blnSpec1 And blnSpec2 And blnSpec3 Then
                blnRtnValue = True
                Exit For
            End If
        Next

ExitThisFunction:


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Private Function CheckMDMDocumentExistInMemory(objPSS() As stPSSDocuments, intMDMDocID As Long, intTTKey As Int32,
                                                   intCompleteTTKey As Int32, strOPNO As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intMDMDocID:({0}) intTTKey:({1}) intCompleteTTKey:({2}) strOPNO:({3})",
                                                          intMDMDocID, intTTKey, intCompleteTTKey, strOPNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = objPSS.Where(Function(x) x.MDMDocID = intMDMDocID AndAlso x.TTKey = intTTKey AndAlso x.CompleteTTKey = intCompleteTTKey AndAlso x.OPNO.ToLower().Trim() = strOPNO.ToLower().Trim()).Count > 0

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnRtnValue
    End Function

    Private Function CheckSerialExist(strSerialNO As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override
        Dim strSQL As String = "select COUNT(serialseq) from AML_BuildPartsSerials where " & Ec.GeneralFunc.GetQueryFieldCondition("scan_serialno", strSerialNO) & " and serialstatus in (1,2)"
        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0) > 0
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Private Function CheckSubHeaderReferenceNumberMatch(objSH() As Parts.stSubHeader, strStationNO As String, strReferenceNO As String,
                                                        ByRef intOperatorPosition As Integer) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0}) strReferenceNO:({1})",
                                                          strStationNO, strReferenceNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False

        intOperatorPosition = 1
        strReferenceNO = FormatSubHeaderReferenceNumber(strReferenceNO)
        For intK = 1 To UBound(objSH)

            If InStr(objSH(intK).OPNO.Trim.ToUpper, strStationNO.Trim.ToUpper) > 0 Then 'If objSH(intK).OPNO.Trim.ToUpper = strStationNO.Trim.ToUpper Then
                'need to ignore the operation position, thats why used instr instead of straight match

                'match the operation number
                'If instr(objSH(intK).OPNO.Trim.ToUpper,strStationNO.Trim.ToUpper) > 0 Then
                Dim strTemp2 = FormatSubHeaderReferenceNumber(objSH(intK).Model2.ToString)

                If strTemp2.Trim = strReferenceNO.Trim Then
                    blnRtnValue = True

                    If objSH(intK).OPNO.Trim.Length > 2 Then  'defensive coding, just in case
                        Select Case Right(objSH(intK).OPNO, 2)
                            Case "-1"
                                intOperatorPosition = 1
                            Case "-2"
                                intOperatorPosition = 2
                            Case "-3"
                                intOperatorPosition = 3
                        End Select
                    End If
                    Exit For
                End If
            End If
        Next

FinalExit:


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) intOperatorPosition:({1})", blnRtnValue, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue

    End Function

    Private Function FirstDayOfMonth() As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim holidayList = New List(Of Date)
        holidayList.Add(New Date(2015, 7, 27))
        holidayList.Add(New Date(2015, 7, 28))
        holidayList.Add(New Date(2015, 7, 29))
        holidayList.Add(New Date(2015, 7, 30))
        holidayList.Add(New Date(2015, 7, 31))
        holidayList.Add(New Date(2015, 8, 1))
        holidayList.Add(New Date(2015, 8, 2))
        holidayList.Add(New Date(2015, 8, 3))
        holidayList.Add(New Date(2015, 8, 4))
        holidayList.Add(New Date(2015, 8, 5))
        holidayList.Add(New Date(2015, 8, 6))
        holidayList.Add(New Date(2015, 8, 7))
        holidayList.Add(New Date(2015, 8, 31))
        holidayList.Add(New Date(2015, 12, 25))
        holidayList.Add(New Date(2015, 12, 26))
        holidayList.Add(New Date(2015, 12, 27))
        holidayList.Add(New Date(2015, 12, 28))
        holidayList.Add(New Date(2015, 12, 29))
        holidayList.Add(New Date(2015, 12, 30))
        holidayList.Add(New Date(2015, 12, 31))
        holidayList.Add(New Date(2016, 1, 1))

        Dim workDays = 1
        Dim objDate As Date = Date.Now
        Dim datevalue As Date = New Date(objDate.Year, objDate.Month, 1)
        Dim dayIsFound As Boolean = False
        Dim blnRtnValue As Boolean = False

        While Not dayIsFound
            If datevalue.DayOfWeek = DayOfWeek.Saturday Then
            ElseIf datevalue.DayOfWeek = DayOfWeek.Sunday Then
            Else
                For Each holiday In holidayList
                    If DateDiff(DateInterval.Day, datevalue, holiday) = 0 Then
                        GoTo SkipThisDate
                    End If
                Next
                dayIsFound = True
                GoTo ExitLooping
            End If
SkipThisDate:
            workDays += 1
            datevalue = datevalue.AddDays(1)

            If workDays > 30 Then  'never happens, just in case
                Exit While
            End If
ExitLooping:
        End While

        If datevalue.Day = objDate.Day Then
            blnRtnValue = True
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Private Function GetAutoRefreshingStationsListForSQL(ByVal intLineID As Integer) As String


#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter intLineID:({0})", intLineID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objLS(0) As stStations
        Dim strOperationSQLCond As String = ""

        Try
            objLS = GetLineStations(intLineID)      'get all lines and its stations

            For intI = 1 To UBound(objLS)
                Dim strTemp2 = ""
                If CheckForAutoRefreshingLine(objLS(intI).LineID) Then
                    'get all the stations for line DB9 and Vantage.

                    Dim strOPNO = objLS(intI).StationNO.Trim
                    Dim intOperatorsCount = objLS(intI).Operators
                    strTemp2 = "'" & UCase(Trim(strOPNO)) & "'"
                    For intZZ = 1 To intOperatorsCount 'GetMaxNumberOfOperatorPositions()
                        strTemp2 &= ",'" & UCase(Trim(strOPNO)) & "-" & Trim(intZZ.ToString) & "'"
                    Next
                End If

                If strTemp2.Trim <> "" Then
                    If strOperationSQLCond.Trim <> "" Then strOperationSQLCond &= ","
                    strOperationSQLCond &= strTemp2 ' " opno in (" & strTemp2 & ")"
                End If
            Next

            If strOperationSQLCond.Trim <> "" Then
                'holds all operations list with operator positions (like 8100,8100-1,8200,...)
                strOperationSQLCond = Trim(DBConfig.QueryFunctions.Upper) & "(opno) in ( " & strOperationSQLCond & ") "
            End If
        Catch ex As Exception
            GenerateException("GetAutoRefreshingStationsListForSQL", ex)
        End Try

#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strOperationSQLCond), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strOperationSQLCond
    End Function

    Private Function GetBuildPartSerialQuery(strOption As String, strBuildNO As String, strStationNO As String,
                                             Optional strPartNO As String = "", Optional intSerialStatus As Integer = 0,
                                             Optional intSerialSEQ As Integer = -22,
                                             Optional intOperatorPosition As Integer = -1,
                                             Optional intPartMVerifySeq As Integer = 0) As String
#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter strBuildNO:({0}) strStationNO:({1}) strPartNO:({2}) intSerialStatus:({3}) intSerialSEQ:({4}) intOperatorPosition:({5}) intPartMVerifySeq:({6})",
                                                           strBuildNO, strStationNO, strPartNO, intSerialStatus, intSerialSEQ, intOperatorPosition, intPartMVerifySeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'AML_BuildPartsSerials: SerialStatus: -1 Incomplete, 1-Complete, 2-GL Override
        Dim strSQL As String = ""
        Dim strFieldsList As String = " * "
        If strOption.Trim.ToLower = "anyrecords" Then
            strFieldsList = " count(StationNO) "
        End If
        'Build SQL statement
        strSQL = "SELECT " & strFieldsList & " FROM AML_BuildPartsSerials where " & Ec.GeneralFunc.GetQueryFieldCondition("BuildNumber", strBuildNO)
        If strStationNO.Trim <> "" Then
            strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("StationNO", strStationNO)
        End If


        If strPartNO.Trim <> "" Then
            strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("PartNumber", strPartNO)
        End If
        If intSerialStatus <> 0 Then
            strSQL &= " and SerialStatus=" & intSerialStatus
        End If

        If intSerialSEQ <> -22 Then
            strSQL &= " and serialseq=" & intSerialSEQ
        End If
        If intPartMVerifySeq > 0 Then
            strSQL &= " and mseq=" & intPartMVerifySeq
        End If
        If intOperatorPosition <> -1 Then
            strSQL &= " and operatorposition=" & intOperatorPosition
        End If
        strSQL &= " order by StationNO,operatorposition,PartNumber,serialseq "  'don't change order by used in export to xml (ExportPartSerialsData)

#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strSQL
    End Function

    Private Function GetBuildPartVerifyQueryFromSupplierPartNo(ByVal supplierPartNo As String, ByVal buildNo As String, ByVal stationNo As String) As String
#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter supplierPartNo:({0}) buildNo:({1}) stationNo:({2})", supplierPartNo, buildNo, stationNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL = String.Format("SELECT * FROM aml_buildpartsverify WHERE partnumber = (SELECT TOP 1 partnumber FROM aml_partsverify WHERE supplier_partno = '{0}')  AND buildnumber = '{1}' AND stationno = '{2}'", supplierPartNo, buildNo, stationNo)

#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strSQL
    End Function

    Private Function GetBuildPartVerifyQuery(strOption As String,
                                             strBuildNO As String,
                                             strStationNO As String,
                                             Optional strPartNO As String = "",
                                            Optional blnVerifyStation As Boolean = False,
                                            Optional blnGetInCompleteRecordOnly As Boolean = False,
                                            Optional intVerifyStatus As Int16 = -22,
                                            Optional strOrderBy As String = "",
                                            Optional strCallFrom As String = "",
                                            Optional ByRef intOperatorPosition As Int16 = -1) As String

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter strOption:({0}) strBuildNO:({1}) strStationNO:({2}) strPartNO:({3}) " &
                                                          "blnVerifyStation:({4}) blnGetInCompleteRecordOnly:({5}) intVerifyStatus:({6}) " &
                                                          "strOrderBy:({7}) strCallFrom:({8}) intOperatorPosition:({9})",
                                                          strOption, strBuildNO, strStationNO, strPartNO,
                                                          blnVerifyStation, blnGetInCompleteRecordOnly, intVerifyStatus,
                                                          strOrderBy, strCallFrom, intOperatorPosition), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "SELECT {0} FROM aml_buildpartsverify a left join aml_vehicleparts b  on a.partnumber = b.partnumber WHERE {1} {2}"
        Dim strFieldsList As String = "a.*, b.tracking"
        If strOption.Trim.ToLower = "anyrecords" Then
            strFieldsList = "COUNT(stationno)"
        End If

        'Build SQL statement
        Dim whereClause = New List(Of String)()
        whereClause.Add(Ec.GeneralFunc.GetQueryFieldCondition("a.BuildNumber", strBuildNO))
        If strStationNO.Trim <> "" Then whereClause.Add(Ec.GeneralFunc.GetQueryFieldCondition("a.StationNO", strStationNO))
        If strOrderBy.Trim = "" Then strOrderBy = " order by a.StationNO,a.operatorposition,a.PartNumber "
        If strCallFrom.Trim = "PART-VERIFICATION" Then
            strOrderBy = ""
            If strPartNO.Trim <> "" Then
                whereClause.Add(String.Format("({0} OR {1})", Ec.GeneralFunc.GetQueryFieldCondition("a.PartNumber", strPartNO), Ec.GeneralFunc.GetQueryFieldCondition("a.spare1", strPartNO)))
            End If
        Else
            If strPartNO.Trim <> "" Then
                whereClause.Add(Ec.GeneralFunc.GetQueryFieldCondition("a.PartNumber", strPartNO))
            End If
        End If
        If blnGetInCompleteRecordOnly Then
            'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
            whereClause.Add("a.VerifyStatus=-1")
        End If
        'AML_BuildPartsVerify: VerifyStatus:  '-1 Incomplete, 1-Complete, 2-GLOverride
        If intVerifyStatus <> -22 Then whereClause.Add(String.Format("a.verifystatus = {0}", intVerifyStatus))
        If intOperatorPosition <> -1 Then whereClause.Add(String.Format("a.operatorposition = {0}", intOperatorPosition))


        strSQL = String.Format(strSQL, strFieldsList, String.Join(" AND ", whereClause), strOrderBy)
        If strCallFrom.Trim = "PART-VERIFICATION" Then
            strSQL &= " UNION " & GetBuildPartVerifyQueryFromSupplierPartNo(strPartNO, strBuildNO, strStationNO)
        End If

#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit strSQL:({0})", strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return strSQL
    End Function

    Private Function GetLineStatusTableName(lineId As Integer) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0})", lineId), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim lineStatusTableName = "AML_LineStatus"
        If lineId = BodyShopLineId Then
            lineStatusTableName = "AML_LINESTATUS_BODYSHOP"
        ElseIf Ec.AppConfig.AstonMartinStAthan AndAlso lineId = MainAssemblyLineId Then
            lineStatusTableName = "AML_LINESTATUS_MA"
        ElseIf Ec.AppConfig.AstonMartinStAthan AndAlso lineId = PaintLineId Then
            lineStatusTableName = "AML_LINESTATUS_Paint"
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", lineStatusTableName), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return lineStatusTableName
    End Function

    Private Function GetMDMStationFLMRecords(ByVal strArea As String, ByVal strStationNO As String, intShift As Integer) As stRemedyDocs()


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter ({0}) ({1}) ({2})", strArea, strStationNO, intShift), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'get the FLM records for the station/machine

        Dim strSQL As String = ""
        Dim objRPDocs(0) As stRemedyDocs
        Dim strSQLCond As String = " pdmworkstationmm1.printorder=77 "

        Dim objFLMSF(0) As stFLMShiftFrequency
        Dim intFreqID As Int16 = 0
        Dim dtDate As DateTime = Date.Now, blnGood As Boolean = False
        Dim intFreqType As Int16 = 0

        Dim strErrorLoc As String = "Start"

        Try
            Dim isFirstDayOfMonth = FirstDayOfMonth()

            If DBConfig.Version7 Then
                strSQLCond = " (" & strSQLCond.Trim & " Or pdmworkstationmm1.mdmdocid >0) "
            End If

            'EW-672: AML : FLM Frequencies
            strSQL = "select pdmpartmm.completettkey,pdmpartmm.docid,pdmpartmm.docseq,pdmpartmm.docrectype," &
                      " pdmpartmm.docdesc,pdmpartmm.ttkey,pdmpartmm.flmfrequencyid from " &
                      " pdmpartmm, pdmworkstationmm1, pdmworkstation" &
                      " with (NOLOCK)  " &
                      " where " &
                      " pdmpartmm.docid = pdmworkstationmm1.docid " &
                      " And pdmpartmm.ttkey=pdmworkstationmm1.ttkey " &
                      " And  " & strSQLCond &
                      " And pdmworkstationmm1.infotype=8    " &
                      " And pdmworkstationmm1.wsid=pdmworkstation.wsid   " &
                      " And " & Ec.GeneralFunc.GetQueryFieldCondition("pdmworkstation.wc", strArea) &
                      " And " & Ec.GeneralFunc.GetQueryFieldCondition("pdmworkstation.wscode", strStationNO) &
                      " order by pdmpartmm.docdesc"

            strErrorLoc = "GetConnectionObj"
            Using connection = New Connection(gStrConnectionString)
                strErrorLoc = "GetFLMShiftFrequency"
                objFLMSF = GetFLMShiftFrequency(connection)
                Using objDataTable = connection.GetDataIntoDataTable(strSQL, Nothing)
                    Dim counter = 0
                    strErrorLoc = "Loop Records"
                    For Each objDataRow As DataRow In objDataTable.Rows
                        intFreqType = 0
                        intFreqID = Data.GetDataRowValue(Of Int16)(objDataRow("flmfrequencyid"), 0)

                        If intFreqID > 0 Then
                            blnGood = False
                            For intYY = 1 To UBound(objFLMSF)
                                If objFLMSF(intYY).IDX = intFreqID Then
                                    'match found.
                                    intFreqType = objFLMSF(intYY).FrequencyType

                                    'FrequencyType: 1-Daily, 2- Weekly, 3-Monthly
                                    Select Case intFreqType
                                        Case 1
                                            'good to have this FLM
                                            blnGood = True

                                        Case 2
                                            If dtDate.DayOfWeek = 1 Then  'monday
                                                blnGood = True
                                            End If

                                        Case 3
                                            blnGood = isFirstDayOfMonth
                                    End Select

                                    If objFLMSF(intYY).Shift1 = 1 And objFLMSF(intYY).Shift2 = 1 And objFLMSF(intYY).Shift3 = 1 Then
                                        'applicable to all shifts
                                    Else
                                        If blnGood Then
                                            Select Case intShift
                                                Case 1
                                                    If objFLMSF(intYY).Shift1 = 1 Then
                                                    Else
                                                        blnGood = False
                                                    End If
                                                Case 2
                                                    If objFLMSF(intYY).Shift2 = 1 Then
                                                    Else
                                                        blnGood = False
                                                    End If
                                                Case 3
                                                    If objFLMSF(intYY).Shift3 = 1 Then
                                                    Else
                                                        blnGood = False
                                                    End If
                                            End Select
                                        End If
                                    End If

                                    Exit For
                                End If
                            Next

                            If Not blnGood Then
                                GoTo SkipThisRecord
                            End If
                        End If


                        counter += 1
                        ReDim Preserve objRPDocs(counter)

                        objRPDocs(counter).PartID = 0
                        objRPDocs(counter).RecType = "0"
                        objRPDocs(counter).OPNO = ""
                        objRPDocs(counter).Seq = 0
                        objRPDocs(counter).FileName = ""
                        objRPDocs(counter).Marked = -2
                        objRPDocs(counter).Shift = intShift

                        objRPDocs(counter).PlanType = 0
                        objRPDocs(counter).RemedyPlanDesc = ""
                        objRPDocs(counter).PartNO = ""
                        objRPDocs(counter).FrequencyType = intFreqType
                        objRPDocs(counter).CompleteTTKey = Data.GetDataRowValue(objDataRow("completettkey"), 0)
                        objRPDocs(counter).MDMDocID = Data.GetDataRowValue(objDataRow("docid"), 0)
                        objRPDocs(counter).MDMDocSeq = Data.GetDataRowValue(objDataRow("docseq"), 0)
                        objRPDocs(counter).MDMDocRectype = Data.GetDataRowValue(Of Int16)(objDataRow("docrectype"), 0)
                        objRPDocs(counter).TTKey = Data.GetDataRowValue(objDataRow("ttkey"), 0)
                        objRPDocs(counter).RemedyPlanDesc = Data.GetDataRowValue(objDataRow("docdesc"), "")
SkipThisRecord:
                    Next
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetMDMStationFLMRecords, Location: " & strErrorLoc & ", SQL: " & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objRPDocs)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objRPDocs
    End Function

    Private Function GetOperatorCountForStation(ByVal lineID As Integer, stationNO As String) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) strStationNO:({1})", lineID, stationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(GetOperatorCountForStationSQL(lineID, stationNO), 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Private Function GetOperatorCountForStationSQL(ByVal lineID As Integer, ByVal stationNo As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intLineID:({0}) strStationNO:({1})", lineID, stationNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = " select operators from stations where lineid=" & lineID & " And " & Ec.GeneralFunc.GetQueryFieldCondition("absnno", stationNo)

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result

    End Function


    Private Function GetRemedyPlanItemsFromMemory(ByVal objRPITAVTEMP() As stRemedyPlanItemTemplateAndValues,
                    ByVal intDocID As Int32, ByVal strDocRectype As String,
                    ByVal intDocSeq As Integer) As stRemedyPlanItemTemplateAndValues()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intDocID:({0}) strDocRectype:({1}) intDocSeq:({2})",
                                                          intDocID, strDocRectype, intDocSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objRPITAV(0) As stRemedyPlanItemTemplateAndValues
        Dim intCount = 0, blnFound As Boolean = False

        Try
            For intK = 1 To UBound(objRPITAVTEMP)
                If objRPITAVTEMP(intK).DOCID = intDocID And Trim(objRPITAVTEMP(intK).DocRecType) = strDocRectype.Trim _
                        And objRPITAVTEMP(intK).DocSeq = intDocSeq Then
                    intCount += 1
                    ReDim Preserve objRPITAV(intCount)
                    objRPITAV(intCount) = objRPITAVTEMP(intK)
                    blnFound = True
                Else
                    If blnFound Then Exit For '
                End If
            Next
        Catch ex As Exception
            GenerateException("GetRemedyPlanItemsFromMemory", ex)
        End Try



#If TRACE Then
        Log.OPERATION(String.Format("ExitCount: ({0})", UBound(objRPITAV)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objRPITAV
    End Function

    Private Function GetRPDataMaxRecordSeq(ByVal strEngineNo As String) As Int32

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strEngineNo:({0})", strEngineNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(recordseq) from pdmrpdata WITH (NOLOCK) where engineno='" & strEngineNo.Trim & "'"
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        If returnValue = 0 Then returnValue = 1

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue

    End Function

    Private Function GetStationNoFromOperationNumber(ByVal strOPNO As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strOPNO:({0})", strOPNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strStationNo = ""
        Try
            strStationNo = strOPNO.Trim
            Dim intPos = InStr(strStationNo.Trim, "-")
            If intPos > 0 Then
                strStationNo = Left(strOPNO.Trim, intPos - 1)
            End If

        Catch ex As Exception
            GenerateException("GetStationNoFromOperationNumber", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", strStationNo), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strStationNo

    End Function

    Private Function GetSubHeaderOptionsFromMemory(ByVal objSHOALL() As stSubHeaderOptions, ByVal intSHID As Int32) As stSubHeaderOptions()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION_LOW(String.Format("Enter intSHID:({0})", intSHID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objSHOptions(0) As stSubHeaderOptions

        If intSHID <= 99 Then GoTo ExitThisSub 'the existing  valadd field contains 1,2,3

        Dim intCounter = 0
        For intK = 1 To UBound(objSHOALL)
            If objSHOALL(intK).SHID > 0 AndAlso objSHOALL(intK).SHID = intSHID Then
                intCounter += 1
                ReDim Preserve objSHOptions(intCounter)
                objSHOptions(intCounter) = objSHOALL(intK)
            End If
        Next
ExitThisSub:


#If TRACE Then
        Log.OPERATION_LOW(String.Format("Exit Count:({0})", UBound(objSHOptions)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objSHOptions
    End Function

    Private Function GetWIPCompleteSQL(strBuildNO As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.DATABASE_IO_LOW(String.Format("Enter strBuildNO:({0})", strBuildNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        strSQL = "update MES_ENGINESLIST set enginecomplete=1, enddate=getdate() where engineno='" & strBuildNO.Trim & "'"



#If TRACE Then
        Log.DATABASE_IO_LOW(String.Format("Exit ({0})", strSQL), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strSQL

    End Function

    'Public Sub ResetPDMRPDataDeleteFlag()
    '    fblnDeletePDMRPData = False
    'End Sub
    Private Sub MassUpdateRemedyPlanRecords(ByVal strEngineNO As String,
                ByVal objRPDocs() As stRemedyDocs, ByVal intRecordType As Int16,
                Optional ByVal strStationNO As String = "",
                Optional ByVal blnCreateQCFromWIPScreen As Boolean = False,
                Optional intShift As Integer = 0)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strEngineNO:({0}) intRecordType:({1}) strStationNO:({2}) blnCreateQCFromWIPScreen:({3}) intShift:({4})",
                                                          strEngineNO, intRecordType, strStationNO, blnCreateQCFromWIPScreen, intShift), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'ByVal ArrBuildList()
        'intRecordType : PDMRPDATA->Record type: 0-Quality checks/Remedy Plan ,1- FLM record

        Dim strErrLoc As String = ""
        Dim objRPITAVALL(0) As stRemedyPlanItemTemplateAndValues
        Dim objRPITAV(0) As stRemedyPlanItemTemplateAndValues
        Dim objRPITAVTEMP(0) As stRemedyPlanItemTemplateAndValues

        Dim intYY As Int32 = 0, intCount As Int32 = 0, intK As Int32 = 0, intI As Int32 = 0
        Dim intKeyLink As Int32 = 0
        Dim intZZ As Int32 = 9, intTemp As Int32 = 0
        Dim intRecordSeq As Int32 = 0, intRemedyPlanType As Int16 = 1  'PlanType: 0-Audit Record, 1-Operator Check
        Dim strSpare1 As String = ""
        Dim strDateFields As String = "", strDateFieldValues As String = ""
        Dim strAddlFields As String = "", strTemp2 As String = ""
        Dim intPos2 As Int16 = 0

        Try
            If strEngineNO.Trim = "" Or strEngineNO.Trim = "0" Then
                GoTo ExitThisSub
            End If
            If UBound(objRPDocs) = 0 Then GoTo DeleteExistingRecordsAndExitOut

            'NOW GET ALL THE REMEDY PLAN HEADER AND ITEM VALUES (ALL IN ONE SHOT, TO IMPROVE PERFORMANCE).
            strErrLoc = "Get Remedy Plan Header and Item Values"
            intYY = 0 : intCount = 0 : intK = 0
            ReDim objRPITAVALL(0)       'redim the array


            For intI = 1 To UBound(objRPDocs)       'get all the remedy plan items in one SQL call
                'get the remedy plan item values

                If strAddlFields.Trim <> "" Then strAddlFields &= " or "
                strAddlFields &= " ( pdmpartmm.docid=" & objRPDocs(intI).MDMDocID &
                                " and pdmpartmm.docseq=" & objRPDocs(intI).MDMDocSeq &
                                " and pdmpartmm.docrectype='" & objRPDocs(intI).MDMDocRectype & "')"
            Next intI
            strErrLoc = "GetRemedyPlanItemTemplate"  'GET remedy plan item values
            objRPITAVTEMP = GetRemedyPlanItemTemplate(0, 0, -1, strAddlFields)

            For intI = 1 To UBound(objRPDocs)

                intKeyLink = intI
                objRPDocs(intI).Marked = intKeyLink '** remedy plan header/item records link **

                'get the remedy plan item values
                strErrLoc = "GetRemedyPlanItemTemplate"
                objRPITAV = GetRemedyPlanItemsFromMemory(objRPITAVTEMP, objRPDocs(intI).MDMDocID, objRPDocs(intI).MDMDocRectype.ToString(), objRPDocs(intI).MDMDocSeq)

                intCount = UBound(objRPITAV)            '(item values for current rp record)
                If intCount = 0 Then GoTo SkipThisRecord4

                intYY = UBound(objRPITAVALL)            'so far in memory (all records)


                intK = intYY + intCount
                ReDim Preserve objRPITAVALL(intK)
                For intK = 1 To intCount        'add the record to ALL array
                    objRPITAVALL(intYY + intK) = objRPITAV(intK)
                    objRPITAVALL(intYY + intK).Comment = intKeyLink.ToString   '** remedy plan header/item records link **
                    objRPITAVALL(intYY + intK).Spare1 = objRPDocs(intI).OPNO.Trim  '** remedy plan header/item records link **
                Next
SkipThisRecord4:
            Next
            ReDim objRPITAV(0)

DeleteExistingRecordsAndExitOut:

            strErrLoc = "Create SQL Statements"

            intRecordSeq = GetRPDataMaxRecordSeq(strEngineNO)

            Dim sqlList = New List(Of String)()
            Dim strSQL = "delete from PDMRPData where " &
                                    Ec.GeneralFunc.GetQueryFieldCondition("engineno", strEngineNO) &
                                    " and recordtype=" & intRecordType
            If strStationNO.Trim <> "" Then
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)
            End If
            If intRecordType = 1 Then
                'flm check
                strSQL &= " and shift=" & intShift
            End If

            sqllist.add(strsql)

            strSQL = "delete from pdmrpitemvalues where " & Ec.GeneralFunc.GetQueryFieldCondition("engineno", strEngineNO)
            strSQL &= "  and recordtype=" & intRecordType      'PDMRPITemValues->Record type: 0-Quality checks/Remedy Plan ,1- FLM record
            If strStationNO.Trim <> "" Then
                strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("stationno", strStationNO)
            End If
            If intRecordType = 1 Then
                'flm check
                strSQL &= " and shift=" & intShift
            End If

            sqllist.add(strsql)

            If UBound(objRPDocs) = 0 Then GoTo SkipThisRecord5

            strErrLoc = "Import PDMPRPData: " & strEngineNO.Trim
            For intI = 1 To UBound(objRPDocs)

                If blnCreateQCFromWIPScreen Then
                    strStationNO = GetStationNoFromOperationNumber(objRPDocs(intI).OPNO.Trim)
                End If

                intTemp = (intRecordSeq + objRPDocs(intI).Marked)

                'PDMRPDATA: Result Flag: -1 -> incomplete, 1->pass, 2->fail
                'PDMRPDATA->Record type: 0-Quality checks/Remedy Plan ,1- FLM record
                strSQL = "insert into pdmrpdata(engineno,stationno,optionno,opno,recordseq,datex," &
                            "subheader,resultflag,docid,docseq,docrectype,Operatorx,remedyplandesc, " &
                            " plantype,ttkey,recordtype,shift,FrequencyType) " &
                            " values('" &
                            Strings.ReplaceSingleQuote(strEngineNO, True) & "','" &
                            Strings.ReplaceSingleQuote(strStationNO, True) & "','" &
                            Strings.ReplaceSingleQuote(objRPDocs(intI).PartNO, True) & "','" &
                            Strings.ReplaceSingleQuote(objRPDocs(intI).OPNO, True) & "'," &
                            intTemp.ToString & ",getdate(),'" & Strings.ReplaceSingleQuote(objRPDocs(intI).FileName, True) & "'," &
                            "-1," & objRPDocs(intI).MDMDocID & "," & objRPDocs(intI).MDMDocSeq & "," &
                            objRPDocs(intI).MDMDocRectype & ",0,'" &
                            Strings.ReplaceSingleQuote(objRPDocs(intI).RemedyPlanDesc, True) & "'," & intRemedyPlanType & "," &
                            objRPDocs(intI).TTKey & "," & intRecordType.ToString & "," & intShift & "," & objRPDocs(intI).FrequencyType & ")"

                sqllist.add(strsql)
            Next

            strErrLoc = "Import RP Item Values: " & strEngineNO.Trim
            For intZZ = 1 To UBound(objRPITAVALL)

                intTemp = (intRecordSeq + Strings.ToInt16(objRPITAVALL(intZZ).Comment))
                If blnCreateQCFromWIPScreen Then
                    strStationNO = GetStationNoFromOperationNumber(objRPITAVALL(intZZ).Spare1)
                End If

                strSpare1 = ""
                If objRPITAVALL(intZZ).LowerValue <> 0 AndAlso objRPITAVALL(intZZ).UpperValue <> 0 Then
                    strSpare1 = String.Format("(Range {0} to {1})", Trim(objRPITAVALL(intZZ).LowerValue.ToString()), Trim(objRPITAVALL(intZZ).UpperValue.ToString()))
                End If

                'save the item values
                'PDMRPItemValues: PassFail: -1 -> incomplete, 1->pass, 2->fail
                'PDMRPITemValues->Record type: 0-Quality checks/Remedy Plan ,1- FLM record

                strSQL = "insert into pdmrpitemvalues(recordseq,remedyplanseq,supervisoraction," &
                            "userinput,engineno,stationno,failuredesc,spare1,passfail,commentx,userinputfailed," &
                            " userinputfailed_operator_clock,  " &
                            " userinput_operator_clock" & strDateFields & ",stopbuild, " &
                            " checkmethod,lowervalue,uppervalue,recordtype,shift) values(" &
                            intTemp & "," & objRPITAVALL(intZZ).Seq & ",' ',' ','" &
                            Strings.ReplaceSingleQuote(strEngineNO, True) & "','" &
                            Strings.ReplaceSingleQuote(strStationNO, True) & "','" &
                            Strings.ReplaceSingleQuote(Trim(objRPITAVALL(intZZ).FailureDesc), True) & "','" &
                            Strings.ReplaceSingleQuote(strSpare1, True) & "',-1" &
                            ",' ',' ',0,0" & strDateFieldValues & ",'" & Trim(objRPITAVALL(intZZ).StopBuild) & "'" &
                            ",'" & Trim(objRPITAVALL(intZZ).CheckMethod) & "'," & objRPITAVALL(intZZ).LowerValue & "," &
                            objRPITAVALL(intZZ).UpperValue & "," & intRecordType.ToString & "," & intShift & ")"

                sqllist.add(strsql)
            Next intZZ
SkipThisRecord5:

            Ec.IO.RunSQLInArray(sqllist)

ExitThisSub:
        Catch ex As Exception
            GenerateException("MassUpdateRemedyPlanRecords: " & strErrLoc, ex)
        Finally
            objRPITAVTEMP = Nothing
            objRPITAVALL = Nothing
            objRPITAV = Nothing
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Private Function ValidateSerialForBuild(strBuildNo As String, strStationNO As String,
                                                strPartNO As String, intSerialSEQ As Integer, ByRef strSerialNO As String,
                                                intPartMVerifySeq As Integer) As String


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBuildNo=({0}) strStationNO=({1}) strPartNO=({2}) intSerialSEQ=({3}) strSerialNO=({4}) intPartMVerifySeq=({5})",
                                                          strBuildNo, strStationNO, strPartNO, intSerialSEQ, strSerialNO, intPartMVerifySeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strRtnValue As String = ""
        Dim strSupplierPartNO As String = ""
        Dim strTemp2 As String = "", strAMLPartNO As String = ""
        Dim strSerialFormat As String = ""
        Dim intMatchCount As Int16 = 0
        Dim strTempSS As String = ""
        Dim strPP As String = "", strAlpha As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim intPos As Int16 = 0, blnMatch As Boolean = False, strCheckChar As String = ""
        Dim blnMatchingSerialFound As Boolean = False

        Try
            'get the part serial record
            Dim buildPartSerials = GetBuildNumberPartSerialList(strBuildNo, strStationNO, "", , intSerialSEQ, , intPartMVerifySeq)
            If (Not buildPartSerials.Any()) Then  'never happens, just in case
                strRtnValue = "Unable to get the Serial record from AML_BuildPartsSerials table, Build# " & strBuildNo.Trim & ", Station: " & strStationNO.Trim & ", Part#: " & strPartNO.Trim & ", Serial Seq: " & intSerialSEQ.ToString & ", MSeq=" & intPartMVerifySeq & ", Serial Seq: " & intSerialSEQ
                GoTo ExitThisSub
            End If


            'check the supplier partno
            strSupplierPartNO = Ec.GeneralFunc.QPTrim(buildPartSerials(0).Supplier_PartNO)
            strAMLPartNO = Ec.GeneralFunc.QPTrim(buildPartSerials(0).AML_PartNO)
            strSerialFormat = Ec.GeneralFunc.QPTrim(buildPartSerials(0).Serial_Format)

            If strSerialFormat IsNot Nothing AndAlso strSerialFormat.Trim <> "" Then  ''ANANANNNNNNNNNNNNN
                'match the format for atleast 7 chars

                If strSerialNO IsNot Nothing Then
                    strTemp2 = strSerialNO.Trim
                End If


                'If strTemp2.Trim.Length < 7 Then
                '    'minimum 7 characters

                '    strRtnValue = "Entered Serial number is less then 7 characters, can't match. Serial#: " & strTemp2.Trim & ", Format: " & strSerialFormat
                '    GoTo ExitThisSub
                'End If

                'If strSerialFormat.Trim.Length <= 7 Then
                '    strRtnValue = "The Serial format contains less than 7 characters, Can't continue. Format: " & strSerialFormat
                '    GoTo ExitThisSub
                'End If

                'take the small length and start the process
                Dim intYY = strTemp2.Trim.Length
                If intYY > strSerialFormat.Trim.Length Then intYY = strSerialFormat.Trim.Length

                intMatchCount = 0
                For intK = 1 To intYY

                    'ANANANNNNNNNNNNNNNXX: A-Alpha, N-Numerix, X Any
                    strTempSS = strSerialFormat.Substring(intK - 1, 1)  'substring position start with 0
                    strPP = strTemp2.Substring(intK - 1, 1) 'substring position start with 0
                    Select Case strTempSS.Trim.ToUpper
                        Case "A"
                            If InStr(strAlpha, strPP.ToUpper) > 0 Then
                                'intMatchCount += 1
                            Else
                                strRtnValue = "NO-MATCH"
                            End If
                        Case "N"
                            Select Case strPP.Trim.ToString
                                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                    'intMatchCount += 1
                                Case Else
                                    strRtnValue = "NO-MATCH"
                            End Select
                        Case "X"
                            'intMatchCount += 1
                        Case Else
                            strRtnValue = "NO-MATCH"
                    End Select

                    If intK = strSerialFormat.Trim.Length Then
                        'match is good, just exit out
                        blnMatchingSerialFound = True
                        Exit For
                    End If

                Next


                If strRtnValue.Trim <> "" Then
                    strRtnValue = "Unable to find the Serial Format match, Can't continue. Format: " & strSerialFormat & ", Serial: " & strTemp2.Trim
                Else
                    If blnMatchingSerialFound Then

                        'Tracking: ST - Serial Tracking, VT-Verification Tracking, BT- Batch Tracking, NT - No Tracking
                        If buildPartSerials(0).Tracking.Trim.ToUpper = "ST" Then
                            'the scanned number must be unique
                            If CheckSerialExist(strSerialNO) Then
                                strRtnValue = "The Serial# '" & strSerialNO.Trim & "' is not unique. Can't continue."
                                GoTo ExitThisSub
                            End If
                        Else
                            'batch tracking, just accept and continue
                        End If
                    End If
                End If
            Else
                    'no serial found, just accept the entered serial#
                    blnMatchingSerialFound = True
            End If



ExitThisSub:
        Catch ex As Exception
            GenerateException("ValidateSerialForBuild", ex)
        End Try

        If blnMatchingSerialFound Then
        Else
            If strRtnValue.Trim = "" Then
                strRtnValue = "NO MATCH: " & strSerialNO.Trim
            End If
        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit strRtnValue:({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function
#End Region

#Region "Models"

    Public Class SHOptionsData
        Public Property ComboBox As Integer
        Public Property Name As String
        Public Property DescX As String
        Public ReadOnly Property DisplayValue As String
            Get
                If (String.IsNullOrEmpty(DescX)) Then Return Name
                Return DescX
            End Get
        End Property

    End Class

    Public Class PartShortageReason
        Public Property ShortageReasonID As Integer
        Public Property ShortageReason As String
    End Class

    <Serializable()>
    Public Class VehicleOrders
        Public Year As Integer
        Public Model As String
        Public Description As String
    End Class


    Public Class PartShortageExclusion
        Public Property ShortageExclusionID As Integer
        Public Property SupplierCode As String
        Public Property Name As String
    End Class

    Public Class OperatorViewModelColors
        Public Property Model As String
        Public Property Color As Color
    End Class

    Public Class TorqueVISReportDetail
        Public Property StationNo As String
        Public Property AreaDescription As String
        Public Property PartID As Long
        Public Property HasDocuments As Boolean
        Public Property SHID As Integer
        Public Property DocumentPosition As Integer
        Public Property SubHeaders As List(Of Parts.stSubHeader)
    End Class

#End Region

    Public Class PSSDetail
        Public Property PartNumber As String
        Public Property PSSLink As String
        Public Property SubheaderID As String
        Public Property OpNO As String
        Public Property ErrorString As String
        Public Property SubHeaders As List(Of Parts.stSubHeader)
    End Class

    Public Class StationOption
        Public Property BuildNumber As String
        Public Property Model As String
        Public Property OptionCode As String
        Public Property OptionDescription As String
        Public Property TrackStation As String
    End Class

    Public Class StationVehicleTrim
        Public Property Selected As Boolean
        Public Property Environment As String
        Public Property Description As List(Of String)
    End Class

    Public Function GetVehiclePart(ByVal partNumber As String) As Base

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter partNumber:({0})", partNumber), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = New Base
        Using connection = New Connection(gStrConnectionString)
            result = connection.GetDataIntoClassOf(Of Base)(String.Format("SELECT * FROM aml_vehicleparts WHERE partnumber='{0}'", partNumber)).FirstOrDefault()
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit GetVehiclePart"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result

    End Function

    Public Function GetTrackStationsByPartNumber(ByVal partNumber As String) As List(Of String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter partNumber:({0})", partNumber), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = New List(Of String)

        Using connection = New Connection(gStrConnectionString)
            Using trackStationTable = connection.GetDataIntoDataTable(String.Format("SELECT DISTINCT trackstation FROM aml_vehiclebom WHERE partnumber='{0}'", partNumber), Nothing)

                For Each row As DataRow In trackStationTable.Rows
                    Dim lineID = connection.ExecuteScalar(Of Integer)(GetStationLineIDSQL(Data.GetDataRowValue(row("trackstation"), "")), 0)
                    Dim operatorCount = connection.ExecuteScalar(Of Integer)(GetOperatorCountForStationSQL(lineID, Data.GetDataRowValue(row("trackstation"), "")), 0)
                    Dim opNoList = Enumerable.Range(1, operatorCount).Select(Function(x) String.Format("{0}-{1}", row("trackstation"), x))

                    result.AddRange(opNoList)
                Next

            End Using
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit GetTrackStationsByPartNumber: ({0})", result.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return result

    End Function

    Public Function GetPSSLinkFromPartNumber(ByVal opNo As String, ByVal partNumber As String, ByVal pageFrom As String) As PSSDetail
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter opNo:({0}) partNumber:({1}) pageFrom:({2})", opNo, partNumber, pageFrom), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = New PSSDetail With {.PartNumber = partNumber}
        Dim stationNo = opNo.Split("-"c)(0)
        Using connection = New Connection(gStrConnectionString)
            Using bomTable = connection.GetDataIntoDataTable(String.Format("SELECT DISTINCT linex, trackstation FROM aml_vehiclebom WHERE trackstation = '{1}' AND partnumber='{0}'", partNumber, stationNo), Nothing)
                Dim referenceNumber = Convert.ToInt32(bomTable.Rows(0)("linex"))
                Using subHeaderTable = connection.GetDataIntoDataTable(String.Format("SELECT TOP 1 * FROM subhdr WHERE model2 = {0} AND rectype='0' AND seq = 0 AND opno = '{1}'", referenceNumber, opNo), Nothing)
                    If (IsNothing(subHeaderTable) OrElse subHeaderTable.Rows.Count = 0) Then
                        result.ErrorString = "No Subheaders found for this part number and selected OpNO."
                    Else
                        Dim shID = If(Data.GetDataRowValue(subHeaderTable(0)("sharedshid"), 0) > 0, Data.GetDataRowValue(subHeaderTable(0)("sharedshid"), 0), Data.GetDataRowValue(subHeaderTable(0)("shid"), 0))
                        Dim documentSQL = Ec.Parts.GetDocumentsSQL(" shmm ", 4, Data.GetDataRowValue(Of Long)(subHeaderTable(0)("ID"), 0), "0", 0, opNo, , , , shID, True, True)
                        result.SubheaderID = shID.ToString()
                        result.OpNO = Math.Truncate(Data.GetDataRowValue(Of Decimal)(subHeaderTable(0)("model2"), 0)).ToString()
                        Dim documentExists = connection.ExecuteScalar(Of Integer)(documentSQL, 0)
                        If (documentExists = 0) Then
                            result.ErrorString = "No PSS Document found for this part number."
                        Else
                            result.PSSLink = String.Format("aml4.aspx?shpos={0}&partid={1}&shid={2}&opno={3}&buildno={4}&pagefrom={5}", subHeaderTable(0)("shseq"), subHeaderTable(0)("ID"), shID, opNo, "", pageFrom)
                        End If
                    End If
                End Using
            End Using
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit GetPSSLinkFromPartNumber: ({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result
    End Function

    Public Sub RefreshOperatorViewModelColors()
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Using connection = New Connection(gStrConnectionString)
            Dim sql = "INSERT INTO easeconfig " +
                        "SELECT DISTINCT 'ViewEASE', 'AstonMartin', 'Operator View Model Colors', descriptionx, 'System.Drawing.Color', '-3318692', null " +
                        "FROM aml_vehicleorders v " +
                        "LEFT JOIN easeconfig e ON v.DescriptionX = e.settingid AND settinggroup = 'Operator View Model Colors'" +
                        "WHERE e.settingID is null " +
                        "ORDER BY 1"
            connection.ExecuteNonQuery(sql)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function GetPartShortageExclusion() As List(Of PartShortageExclusion)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = New List(Of PartShortageExclusion)
        Dim shortageExclusion = From t In GetPartShortageExclusionTable().AsEnumerable()
                                Select New PartShortageExclusion With
                                    {
                                        .ShortageExclusionID = Data.GetDataRowValue(t("shortageExclusionid"), 0),
                                        .SupplierCode = Data.GetDataRowValue(t("SupplierCode"), ""),
                                        .Name = Data.GetDataRowValue(t("Name"), "")
                                    }
        If (shortageExclusion.Any()) Then
            result = shortageExclusion.ToList()
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", result.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result
    End Function

    Public Function GetPartShortageExclusionTable() As DataTable
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim table As DataTable = Nothing
        Using connection = New Connection(gStrConnectionString)
            Dim sql = "SELECT * FROM partshortageExclusions"
            table = connection.GetDataIntoDataTable(sql, Nothing)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return table
    End Function

    Public Function GetPartShortageExclusion(ByVal shortageExclusionID As Integer) As PartShortageExclusion

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter shortageExclusionID:({0})", shortageExclusionID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = GetPartShortageExclusion().FirstOrDefault(Function(x) x.ShortageExclusionID = shortageExclusionID)

#If TRACE Then
        Log.OPERATION(String.Format("Exit"), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result

    End Function

    Private Function ValidatePartShortageExclusion(ByVal [shortageExclusion] As PartShortageExclusion) As String
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = ""
        Dim list = GetPartShortageExclusion()
        If (shortageExclusion.ShortageExclusionID > 0) Then
            If (list.Any(Function(x) x.SupplierCode.Equals(shortageExclusion.SupplierCode, StringComparison.CurrentCultureIgnoreCase) AndAlso x.ShortageExclusionID <> shortageExclusion.ShortageExclusionID)) Then
                result = String.Format("Supplier Code: {0} already exists.", shortageExclusion.SupplierCode)
            End If
        Else
            If (list.Any(Function(x) x.SupplierCode.Equals(shortageExclusion.SupplierCode, StringComparison.CurrentCultureIgnoreCase))) Then
                result = String.Format("Supplier Code: {0} already exists.", shortageExclusion.SupplierCode)
            End If
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit result:({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return result
    End Function

    Public Function UpdatePartShortageExclusion(ByVal shortageExclusion As PartShortageExclusion) As String
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter shortageExclusion:({0})", shortageExclusion), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim validationResult = (ValidatePartShortageExclusion(shortageExclusion))
        If (String.IsNullOrEmpty(validationResult)) Then
            Using connection = New Connection(gStrConnectionString)
                Dim sql = String.Format("INSERT INTO PartShortageExclusions VALUES ('{0}','{1}')", Strings.ReplaceSingleQuote(shortageExclusion.SupplierCode), Strings.ReplaceSingleQuote(shortageExclusion.Name))
                If (shortageExclusion.ShortageExclusionID > 0) Then
                    sql = String.Format("UPDATE PartShortageExclusions SET SupplierCode = '{0}', name ='{1}' WHERE shortageexclusionid = {2}", Strings.ReplaceSingleQuote(shortageExclusion.SupplierCode), Strings.ReplaceSingleQuote(shortageExclusion.Name), shortageExclusion.ShortageExclusionID)
                End If

                connection.ExecuteNonQuery(sql)
            End Using
        End If

#If TRACE Then
        Log.OPERATION(String.Format("Exit validationResult:({0})", validationResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return validationResult
    End Function

    Public Function DeletePartShortageExclusion(ByVal shortageExclusionID As Integer) As Boolean
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter shortageReasonID:({0})", shortageExclusionID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result As Boolean
        Using connection = New Connection(gStrConnectionString)
            Dim sql = String.Format("DELETE FROM partshortageexclusion WHERE shortageExclusionID = {0})", shortageExclusionID)

            connection.ExecuteNonQuery(sql)
            result = True
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit result:({0})", result), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return result
    End Function

    Public Function WriteAML_ModelList(ByVal intModelNo As Int16, ByVal strModelCode As String, ByVal intModelNoLinked As Int16, ByVal intID As Integer) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intModelNo:({0}) strModelCode:({1}) intModelNoLinked:({2}) intID:({3})",
                                                          intModelNo, strModelCode, intModelNoLinked, intID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", strCond As String = " AND "

        Dim blnSuccess As Boolean = True

        Try

            If intID = 0 Then 'new record
                intID = GetNewModelID()
                strSQL = "insert into AML_ModelList(id,modelcode,modelnumber,modelnumberlinked) values (" & intID &
                         " , '" & Strings.ReplaceSingleQuote(strModelCode) & "' , " & intModelNo & " , " &
                        intModelNoLinked & ")"
            Else
                strSQL = "update AML_ModelList set modelcode= '" & Strings.ReplaceSingleQuote(strModelCode) & "' , " &
                        "modelnumber =" & intModelNo & " , modelnumberlinked= " & intModelNoLinked &
                        " where id= " & intID

            End If

            Ec.IO.RunSQL(strSQL)

        Catch ex As Exception
            blnSuccess = False
            GenerateException("WriteAML_ModelList", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnSuccess), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnSuccess
    End Function

    Public Function GetNewModelID() As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(id) from AML_ModelList"
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        returnValue += 1

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Sub DeleteAMLModel(ByVal intIDX As Int16)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intIDX:({0})", intIDX), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim strSQL As String = ""

        Try
            strSQL = "delete from AML_ModelList where id =" & intIDX
            Ec.IO.RunSQL(strSQL)

        Catch ex As Exception
            GenerateException("DeleteAMLModel", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub


    Public Function GetOptionsFamily() As stOptionsFamily()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim objOptionsFamily(0) As stOptionsFamily
        Dim strSQL = "SELECT * FROM optionsfamily ORDER BY ofcode"

        Try
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim Preserve objOptionsFamily(table.Rows.Count)
                        Dim intCount = 0
                        For Each reader As DataRow In table.Rows
                            intCount += 1

                            objOptionsFamily(intCount).ID = Data.GetDataRowValue(Of Int16)(reader("OFID"), 0)
                            objOptionsFamily(intCount).Code = Data.GetDataRowValue(reader("OFCode"), "")
                            objOptionsFamily(intCount).Name = Data.GetDataRowValue(reader("OFName"), "")
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetOptionsFamily", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objOptionsFamily)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objOptionsFamily
    End Function

    Public Function GetSpecificOptions(ByVal intOFID As Int16, Optional ByVal intID As Int16 = 0) As stSpecificOptions()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intOFID:({0}) intID:({1})", intOFID, intID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strCond As String = " where "
        Dim strOrdeBy As String = " order by OptionCode "
        Dim objSpecificOptions(0) As stSpecificOptions
        Dim strSQL = "SELECT * FROM specificoptions"
        If intOFID > 0 Then
            strSQL = strSQL & strCond & " OFID = " & intOFID
            strCond = " and "
        Else
            strOrdeBy = " order by ofid,optioncode"
        End If
        If intID > 0 Then
            strSQL = strSQL & strCond & " OptionID = " & intID
        End If
        strSQL = strSQL & strOrdeBy '" order by OptionCode"

        Try
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim Preserve objSpecificOptions(table.Rows.Count)
                        Dim intCount = 0
                        For Each reader As DataRow In table.Rows
                            intCount += 1

                            objSpecificOptions(intCount).OFID = Data.GetDataRowValue(Of Int16)(reader("OFID"), 0)
                            objSpecificOptions(intCount).ID = Data.GetDataRowValue(reader("OptionID"), 0)
                            objSpecificOptions(intCount).Code = Data.GetDataRowValue(reader("OptionCode"), "")
                            objSpecificOptions(intCount).Name = Data.GetDataRowValue(reader("OptionName"), "")
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetOptionsFamily", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objSpecificOptions)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objSpecificOptions
    End Function

    Public Function OptionFamily_Exist(ByVal strCode As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strCode:({0})", strCode), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL = "select COUNT(*) from optionsfamily where ofcode = '" & Strings.ReplaceSingleQuote(strCode) & " ' "
        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0) > 0
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function SpecificOption_Exist(ByVal intOFID As Int16, ByVal strCode As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intOFID:({0}) strCode:({1})", intOFID, strCode), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL = "select COUNT(*)  from specificoptions where optioncode = '" & Strings.ReplaceSingleQuote(strCode) & " ' "
        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0) > 0
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function UpdateOptionsFamily(ByVal strOFCode As String, ByVal strOFName As String, ByVal intOFID As Int16) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strOFCode:({0}) strOFName:({1}) intOFID:({2})", strOFCode, strOFName, intOFID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean = False, strSQL As String = ""

        Try
            If intOFID = 0 Then 'new optionfamily
                strSQL = "INSERT INTO OPTIONSFAMILY(OFID,OFCODE,OFNAME) values ("
                strSQL = strSQL & GetNewOptionsFamilyID() 'generate new id
                strSQL = strSQL & " , '" & Strings.ReplaceSingleQuote(strOFCode) & " ' "
                strSQL = strSQL & " , '" & Strings.ReplaceSingleQuote(strOFName) & " ' ) "
            Else
                strSQL = "UPDATE OPTIONSFAMILY SET OFNAME= '" & Strings.ReplaceSingleQuote(strOFName) & " ' "
                strSQL = strSQL & " WHERE OFID =" & CStr(intOFID) 
            End If

            Ec.IO.RunSQL(strSQL)
            blnRtnValue = True

        Catch ex As Exception
            blnRtnValue = False
            Call GenerateException("ClientEd: UpdateOptionsFamily", ex)
        End Try



#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Function GetNewOptionsFamilyID() As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(OFID) from OPTIONSFAMILY"
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        returnValue += 1

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetNewSpecificOptionsID(ByVal intOFID As Integer) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intOFID:({0})", intOFID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(OPTIONID) from SPECIFICOPTIONS"
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        returnValue += 1


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function UpdateSpecificOptions(ByVal intOFID As Int16, ByVal intOptID As Int16, ByVal strOptCode As String, ByVal strOptName As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intOFID:({0}) intOptID:({1}) strOptCode:({2}) strOptName:({3})", intOFID, intOptID, strOptCode, strOptName), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnValue As Boolean, strSQL As String = ""

        Try
            If intOptID = 0 Then
                strSQL = "INSERT INTO SPECIFICOPTIONS(OFID,OptionID,OptionCODE,OptionNAME) values ("
                strSQL = strSQL & intOFID
                strSQL = strSQL & "," & GetNewSpecificOptionsID(intOFID) 'generate new id
                strSQL = strSQL & " , '" & Strings.ReplaceSingleQuote(strOptCode) & " ' "
                strSQL = strSQL & " , '" & Strings.ReplaceSingleQuote(strOptName) & " ' ) "
            Else
                strSQL = "UPDATE SPECIFICOPTIONS SET OPTIONNAME= '" & Strings.ReplaceSingleQuote(strOptName) & " ' "
                strSQL = strSQL & "WHERE OFID=" & CStr(intOFID) 
                strSQL = strSQL & " AND OPTIONID = " & intOptID
            End If

            Ec.IO.RunSQL(strSQL)
            blnRtnValue = True

        Catch ex As Exception
            blnRtnValue = False
            Call GenerateException("ClientEd: UpdateSpecificOptions", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Function GetNumberOfSpecificOptions(ByVal intOFID As Int16) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intOFID:({0})", intOFID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL = "select count(OFID) from SPECIFICOPTIONS where  OFID = " & intOFID
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Sub DeleteOptionFamily(ByVal intOFID As Int16)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intOFID:({0})", intOFID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim strSQL As String = ""

        Try

            strSQL = "delete from OPTIONSFAMILY where OFID = " & intOFID
            Ec.IO.RunSQL(strSQL)

            'alos delete its specific options
            strSQL = "delete from SPECIFICOPTIONS where OFID = " & intOFID
            Ec.IO.RunSQL(strSQL)

        Catch ex As Exception
            GenerateException("DeleteOptionFamily", ex)
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function SpecificOptionUsed(ByVal strSpecificOptionCode As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strSpecificOptionCode:({0})", strSpecificOptionCode), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If


        Dim strSQL = "select COUNT(shid) from aml_shoptions where specoptions1 ='" & strSpecificOptionCode.Trim & "' or specoptions2 ='" & strSpecificOptionCode.Trim & "' or specoptions3 ='" & strSpecificOptionCode.Trim & "'"
        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0) > 0
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Sub DeleteSpecificOptions(ByVal intOFID As Int16, ByVal intOptID As Int16)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intOFID:({0}) intOptID:({1})", intOFID, intOptID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim strSQL As String = ""

        Try

            strSQL = "delete from SPECIFICOPTIONS where OFID = " & intOFID
            strSQL = strSQL & " AND OPTIONID = " & intOptID
            Ec.IO.RunSQL(strSQL)

        Catch ex As Exception
            GenerateException("DeleteSpecificOptions", ex)
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If


    End Sub


    Public Function CheckModelUsage(ByVal strModelCode As String) As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strModelCode:({0})", strModelCode), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim intCount As Integer
        Dim strRtnValue As String = "", strSQL As String = "", strTemp As String = ""
        Dim blnResult As Boolean = False
        Try

            strSQL = "select partxref.partno,routehdr.rectype , routehdr.seq from routehdr ,partxref " &
                         "where partxref.id = routehdr.id and routehdr.user1 ='" & Strings.ReplaceSingleQuote(strModelCode) & " ' " &
                         "order by rectype , seq"


            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each reader As DataRow In table.Rows
                            intCount += 1

                            If Not IsDBNull(reader("partno")) AndAlso Not IsDBNull(reader("rectype")) AndAlso Not IsDBNull(reader("seq")) Then

                                strTemp = String.Format("{0}  (Rectype : {1} ,  SEQ :  {2})", reader("partno"), reader("rectype"), reader("seq"))
                                If Trim(strRtnValue) <> "" Then strRtnValue = strRtnValue & vbCrLf
                                strRtnValue = strRtnValue & strTemp

                            End If
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("CheckModelUsage", ex)
        End Try



#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function

    Public Function GetAMLSubHeaderOptionsList(ByVal intID As Long, ByVal strRecType As String, ByVal intSeq As Integer,
                                               ByVal strOPNO As String, ByVal intSHID As Int32, Optional ByVal intMseq As Int16 = 0) As stSubHeaderOptions() ' ByVal strOPNO As String,


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRecType:({1}) intSeq:({2}) strOPNO:({3}) intSHID:({4}) intMseq:({5})",
                                                          intID, strRecType, intSeq, strOPNO, intSHID, intMseq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If


        Dim strSQL As String = "", strCond As String = " AND "
        Dim objSHO() As stSubHeaderOptions
        ReDim objSHO(0)
        Dim strOrderby As String = ""
        Dim strSHIDList As String = ""
        Dim blnCheckSharedSH As Boolean = True
        Dim objSH(0) As SubHeaders.stSubHeaders

        'Build SQL statement
        strSQL = "SELECT * FROM AML_SHOptions" & " where ID = " & intID & " and rectype = '" & Trim(strRecType) &
            "' and seq = " & intSeq

        If strOPNO.Trim <> "" Then
            strSQL &= " and opno='" & Trim(strOPNO) & "' "

            strOrderby = " order by opno, mseq"
        End If



        If intSHID > 0 Then
            strSQL = strSQL & strCond & " shid = " & intSHID
        End If
        If intMseq > 0 Then
            strSQL = strSQL & strCond & " mseq = " & intMseq
        End If

        If strOrderby.Trim = "" Then
            If intSHID = 0 Then
                strSQL = strSQL & " order by shid , mseq"
            Else
                strSQL = strSQL & " order by mseq"
            End If
        Else
            strSQL = strSQL & strOrderby
        End If


        Try
            Using connection = New Connection(gStrConnectionString)
                Dim counter = 0
LoopAgain:
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objSHO(counter + table.Rows.Count)
                        For Each reader As DataRow In table.Rows
                            counter += 1

                            objSHO(counter).ID = Data.GetDataRowValue(reader("id"), 0)
                            objSHO(counter).RecType = Data.GetDataRowValue(reader("rectype"), "")
                            objSHO(counter).Seq = Data.GetDataRowValue(Of Int16)(reader("seq"), 0)
                            objSHO(counter).OPNO = Data.GetDataRowValue(reader("opno"), "")
                            objSHO(counter).SHID = Data.GetDataRowValue(reader("shid"), 0)
                            objSHO(counter).Mseq = Data.GetDataRowValue(Of Int16)(reader("mseq"), 0)
                            objSHO(counter).SubModel = Data.GetDataRowValue(reader("submodel"), "")
                            objSHO(counter).Body = Data.GetDataRowValue(reader("body"), "")
                            objSHO(counter).GearBox = Data.GetDataRowValue(reader("gearbox"), "")
                            objSHO(counter).Performance = Data.GetDataRowValue(reader("performance"), "")
                            objSHO(counter).YearX = Data.GetDataRowValue(reader("yearx"), "")
                            objSHO(counter).Territory = Data.GetDataRowValue(reader("territory"), "")
                            objSHO(counter).Drive = Data.GetDataRowValue(reader("drive"), "")
                            objSHO(counter).SpecificOptions1 = Data.GetDataRowValue(reader("specoptions1"), "")
                            objSHO(counter).SpecificOptions2 = Data.GetDataRowValue(reader("specoptions2"), "")
                            objSHO(counter).SpecificOptions3 = Data.GetDataRowValue(reader("specoptions3"), "")
                        Next
                    End If
                End Using

                If intSHID = 0 And DBConfig.SharedSubHeader AndAlso blnCheckSharedSH Then

                    blnCheckSharedSH = False
                    'EW-825: AML: Operator View Changes
                    objSH = Ec.SubHeaders.GetSubHeadersList(intID, strRecType, intSeq, strOPNO)
                    For intYY = 1 To UBound(objSH)
                        If objSH(intYY).SharedSHID > 0 Then
                            If strSHIDList.Trim <> "" Then strSHIDList &= ", "
                            strSHIDList &= objSH(intYY).SharedSHID
                        End If
                    Next

                    If strSHIDList.Trim <> "" Then
                        strSQL = "SELECT * FROM AML_SHOptions" & " where shid in (" & strSHIDList.Trim & ") "
                        GoTo LoopAgain
                    End If
                End If
            End Using
        Catch ex As Exception
            GenerateException("GetAMLSubHeaderOptionsList", ex)
        Finally
            objSH = Nothing
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objSHO)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objSHO
    End Function

    Public Function WriteAMLSubHeaderOptionsList(ByVal objSHO() As stSubHeaderOptions, Optional ByVal blnDeleteALLSHOptions As Boolean = False, Optional ByVal blnEditRecord As Boolean = False) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnDeleteALLSHOptions:({0}) blnEditRecord:({1})", blnDeleteALLSHOptions, blnEditRecord), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim blnRtnvalue As Boolean = False

        'blnDeleteAllSHOptions is false by default,deletes all options for a subheader
        'blnEditRecord true,only when editing a single options record for a subheader


        Dim strSQL As String = "", blnResult As Boolean = False

        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try

                    For intK = 1 To UBound(objSHO)
                        If intK = 1 Then
                            If blnDeleteALLSHOptions Or blnEditRecord Then        'delete all existing records for the document
                                strSQL = "delete from AML_SHOptions where id = " & objSHO(intK).ID & "  and rectype= '" & Trim(objSHO(intK).RecType) &
                                            "'and seq =" & objSHO(intK).Seq & " and shid =" & objSHO(intK).SHID 'and opno='" & Trim(objSHO(intK).OPNO) & "' no need shid is sufficient

                                If blnEditRecord Then
                                    strSQL = strSQL & " and mseq=" & objSHO(intK).Mseq
                                End If

                                connection.ExecuteNonQuery(strSQL)

                                'write tolog
                            End If
                        End If

                        strSQL = "insert into AML_SHOptions (id,rectype,seq,opno,shid,mseq,submodel,body,gearbox,yearx,territory,performance,drive,specoptions1,specoptions2,specoptions3) values( " & objSHO(intK).ID & ",'" & objSHO(intK).RecType & "'," & objSHO(intK).Seq & ",'" & Trim(objSHO(intK).OPNO) & "'" &
                                 "," & objSHO(intK).SHID & "," & objSHO(intK).Mseq & ",'" & Strings.ReplaceSingleQuote(Trim(objSHO(intK).SubModel)) & "' ," &
                                  "'" & objSHO(intK).Body & "','" & objSHO(intK).GearBox & "','" & objSHO(intK).YearX & "'," &
                                 "'" & Trim(objSHO(intK).Territory) & "','" & objSHO(intK).Performance & "','" & objSHO(intK).Drive & "' ,'" & Trim(objSHO(intK).SpecificOptions1) & "'," & "'" & Trim(objSHO(intK).SpecificOptions2) & "','" & Trim(objSHO(intK).SpecificOptions3) & "'" & ")"

                        connection.ExecuteNonQuery(strSQL)
                    Next

                    connection.CommitTransaction()
                    blnRtnvalue = True
                Catch ex As Exception
                    connection.RollbackTransaction()
                    Throw ex
                End Try
            End Using

        Catch ex As Exception
            GenerateException(" WriteAMLSubHeaderOptionsList", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnRtnvalue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnvalue
    End Function

    Public Function GetOptionDatafromGlossaryTable(ByRef counter As Integer) As String(,)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim arrTemp(0, 2) As String '2 -dimensional array,
        Try
            counter = 0
            Dim strSQL = "select valuex ,descriptionx from aml_vehicleglossary where groupx not in ('country','family','option','station')"

            Using connection = New Connection(gStrConnectionString)
                Using dataTable = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(dataTable) AndAlso dataTable.Rows.Count > 0) Then
                        ReDim arrTemp(dataTable.Rows.Count, 2)
                        For Each row As DataRow In dataTable.Rows
                            counter += 1
                            Dim strCode = Data.GetDataRowValue(row(0), "")
                            Dim strDescX = Data.GetDataRowValue(row(1), "")

                            arrTemp(counter, 1) = Trim(strCode)
                            arrTemp(counter, 2) = Trim(strDescX)
                        Next
                    End If
                End Using
            End Using

        Catch ex As Exception
            GenerateException("GetOptionDatafromGlossary", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0}) counter:({1})", UBound(arrTemp), counter), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return arrTemp
    End Function

    Public Function GetOptionDescriptionFromMemory(ByVal strOptionCodex As String, ByVal arrOptionData(,) As String, ByVal intArrayCount As Integer) As String


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strOptionCodex:({0}) intArrayCount:({1})", strOptionCodex, intArrayCount), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strRtnValue As String = ""

        Try
            If Trim(strOptionCodex) = "" Then GoTo EXITHTIS
            For intYY = 1 To intArrayCount
                If UCase(Trim(strOptionCodex)) = UCase(Trim(arrOptionData(intYY, 1))) Then
                    strRtnValue = arrOptionData(intYY, 2)
                    Exit For
                End If
            Next

EXITHTIS:
        Catch ex As Exception
            GenerateException("frmAMLSHOptions-GetOptionDescriptionFromMemory", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strRtnValue
    End Function

    Public Function RefreshOptionData(ByVal objSHOExistingRecords() As stSubHeaderOptions,
                                     ByVal intComboPos_SM As Int16, ByVal intComboPos_body As Int16, ByVal intComboPos_Gear As Int16,
                                     ByVal intComboPos_Drive As Int16, ByVal intComboPos_Year As Int16, ByVal intComboPos_Terr As Int16,
                                     ByVal intComboPos_Perf As Int16, ByVal strModel As String, ByVal blnAdd As Boolean) As List(Of SHOptionsData)



#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intComboPos_SM:({0}) intComboPos_body:({1}) intComboPos_Gear:({2}) intComboPos_Drive:({3}) intComboPos_Year:({4}) intComboPos_Terr:({5}) intComboPos_Perf:({6}) strModel:({7}) blnAdd:({8})",
                                                          intComboPos_SM, intComboPos_body, intComboPos_Gear, intComboPos_Drive, intComboPos_Year, intComboPos_Terr, intComboPos_Perf, strModel, blnAdd), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strTemp As String = "", strOPtionCodex As String = "", intLocation As Int16 = 1
        Dim arrOptionData(,) As String
        Dim optionsData = New List(Of SHOptionsData)

        Try

            'load the model combo
            Dim strSQL = "select distinct(modelcode) from aml_ModelList order by modelcode"
            Dim intComboBoxKey = intComboPos_SM
            Dim counter = 0

            Using connection = New Connection(gStrConnectionString)
GetDatafromDB:
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each reader As DataRow In table.Rows
                            strTemp = Data.GetDataRowValue(reader(0), "")

                            If intComboBoxKey = intComboPos_SM And blnAdd Then
                                For intYY = 1 To UBound(objSHOExistingRecords)
                                    If Trim(objSHOExistingRecords(intYY).SubModel) = Trim(strTemp) Then
                                        GoTo SKIPTHISRECORD
                                    End If
                                Next
                            End If

                            counter += 1
                            optionsData.Add(New SHOptionsData() With {
                                            .Name = strTemp,
                                            .ComboBox = intComboBoxKey})
SKIPTHISRECORD:
                        Next
                    End If
                End Using


                Ec.GeneralFunc.DelayProcess1(2)  'just delay

                Select Case intLocation
                    Case 1
                        'load the body combo
                        strSQL = "select distinct(body) from aml_vehicleorders where model1 = " & strModel & " order by body "
                        intComboBoxKey = intComboPos_body
                        intLocation = 2
                        GoTo GetDatafromDB
                    Case 2
                        'load the gearbox
                        strSQL = "select distinct(gearbox1)  from aml_vehicleorders where model1 = " & strModel & " order by gearbox1"
                        intComboBoxKey = intComboPos_Gear
                        intLocation = 3
                        GoTo GetDatafromDB
                    Case 3
                        'drive
                        strSQL = "select distinct(drive1) from aml_vehicleorders where model1 = " & strModel & " order by drive1 "
                        intComboBoxKey = intComboPos_Drive
                        intLocation = 4
                        GoTo GetDatafromDB
                    Case 4
                        'year
                        strSQL = "select distinct(yearx) from aml_vehicleorders where model1 = " & strModel & " order by yearx "
                        intComboBoxKey = intComboPos_Year
                        intLocation = 5
                        GoTo GetDatafromDB
                    Case 5
                        'territory
                        strSQL = "select distinct(territory1) from aml_vehicleorders where model1 = " & strModel & " order by territory1 "
                        intComboBoxKey = intComboPos_Terr
                        intLocation = 6
                        GoTo GetDatafromDB
                    Case 6
                        'performance
                        strSQL = "select distinct (performance1) from aml_vehicleorders where model1 = " & strModel & " order by performance1 "
                        intComboBoxKey = intComboPos_Perf
                        intLocation = 7
                        GoTo GetDatafromDB

                End Select
            End Using

            If (optionsData.Count = 0) Then GoTo EXITHISSUB

            arrOptionData = GetOptionDatafromGlossaryTable(counter)

            For counter = 0 To optionsData.Count - 1
                strOPtionCodex = Trim(optionsData(counter).Name)
                If Trim(strOPtionCodex) = "" Then GoTo SKipTHISOPTION
                optionsData(counter).DescX = GetOptionDescriptionFromMemory(strOPtionCodex, arrOptionData, counter)
SKipTHISOPTION:
            Next

EXITHISSUB:
        Catch ex As Exception
            GenerateException(ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return optionsData

    End Function

    Public Sub DeleteAMALSubheaderOptionsList(ByVal intID As Int32, ByVal strRecType As String, ByVal intSeq As Int16, ByVal strOPNO As String, ByVal intSHID As Int32, Optional ByVal strMseqList As String = "")

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRecType:({1}) intSeq:({2}) strOPNO:({3}) intSHID:({4}) strMseqList:({5})",
                                                          intID, strRecType, intSeq, strOPNO, intSHID, strMseqList), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""
        Try
            strSQL = "delete from AML_SHOptions where id = " & intID & " and rectype = '" & Trim(strRecType) & "' and seq = " & intSeq & " and opno = '" & Trim(strOPNO) & "'"

            If intSHID > 0 Then
                strSQL = strSQL & " and shid = " & intSHID
            End If

            If Trim(strMseqList) <> "" Then
                strSQL = strSQL & " and mseq in ( " & strMseqList & " )"
            End If

            Ec.IO.RunSQL(strSQL)

        Catch ex As Exception
            GenerateException("DeleteAMALSubheaderOptionsList", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function AnySafetyDocsFound(strUserID As String) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserID:({0})", strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        If Not Ec.AppConfig.AstonMartin Then Return False
        Dim strSQL = "SELECT COUNT(userid) FROM AML_SafetyDocsCheck WHERE UPPER(userid)='" & UCase(Trim(strUserID)) & "'"
        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0) > 0
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    ''' <summary>
    ''' Get Override Category Group
    ''' </summary>
    ''' <param name="intCategoryID">Category ID</param>
    ''' <returns>ORC Group</returns>
    Public Function GetOverrideCategoryGroups(Optional ByVal intCategoryID As Int16 = 0) As stOverridecategoryGroup()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intCategoryID:({0})", intCategoryID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objOverridecategoryGroup() As stOverridecategoryGroup
        Dim strSQL As String = "select categoryid,seq,categorydesc from AMLOverrideCategory "
        If intCategoryID > 0 Then
            strSQL &= " where categoryid=" & intCategoryID
        End If
        strSQL &= " order by seq"

        ReDim objOverridecategoryGroup(0)
        Try

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objOverridecategoryGroup(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1

                            objOverridecategoryGroup(counter).GroupID = Data.GetDataRowValue(Of Int16)(row("categoryid"), 0)
                            objOverridecategoryGroup(counter).GroupName = Data.GetDataRowValue(row("categorydesc"), "")
                            objOverridecategoryGroup(counter).Seq = Data.GetDataRowValue(Of Int16)(row("seq"), 0)
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetOverrideCategoryGroups", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objOverridecategoryGroup)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objOverridecategoryGroup
    End Function

    ''' <summary>
    ''' Delete Override category Group
    ''' </summary>
    ''' <param name="objOverridecategoryGroup">list of Override category group to be deleted</param>
    ''' <returns>Result: True/False</returns>
    ''' <remarks></remarks>
    Public Function DeleteOverridecategoryGroup(ByVal objOverridecategoryGroup() As stOverridecategoryGroup) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False

        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try

                    For intK = 1 To UBound(objOverridecategoryGroup)
                        strSQL = "delete from AMLOverrideCategory where categoryid=" & objOverridecategoryGroup(intK).GroupID
                        connection.ExecuteNonQuery(strSQL)

                        strSQL = "delete from AMLOverrideCategory where categoryid=" & objOverridecategoryGroup(intK).GroupID
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
            Call GenerateException("DeleteOverridecategoryGroup", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult

    End Function

    ''' <summary>
    ''' Get ORC Group Items
    ''' </summary>
    ''' <param name="intGroupID">Group ID</param>
    ''' <param name="intGroupItemID">Group Item ID</param>
    ''' <returns></returns>
    ''' <remarks>
    Public Function GetORCGroupItems(ByVal intGroupID As Int16, Optional ByVal intGroupItemID As Int16 = 0) As stORCGroupItems()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intGroupID:({0}) intGroupItemID:({1})", intGroupID, intGroupItemID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If


        Dim objORCGItems() As stORCGroupItems
        Dim strSQL As String = "select categoryid,commentid,seq,commentdesc from AMLOverrideCategoryItem"

        If intGroupID > 0 Then
            strSQL &= " where categoryid=" & intGroupID

            If intGroupItemID > 0 Then
                strSQL &= " and commentid=" & intGroupItemID
            End If
        End If

        strSQL &= " order by seq"
        ReDim objORCGItems(0)
        Try

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve objORCGItems(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1

                            objORCGItems(counter).GroupID = Data.GetDataRowValue(Of Int16)(row("categoryid"), 0)
                            objORCGItems(counter).CategoryID = Data.GetDataRowValue(Of Int16)(row("commentid"), 0)
                            objORCGItems(counter).CategorySeq = Data.GetDataRowValue(Of Int16)(row("seq"), 0)
                            objORCGItems(counter).Category = Trim(Data.GetDataRowValue(row("commentdesc"), ""))
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetORCGroupItems", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objORCGItems)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return objORCGItems
    End Function

    Public Function GetNewORCGroupID() As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Get New and unique Question id
        Dim strSQL As String = "select max(categoryid) from AMLOverrideCategory "
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        returnValue += 1


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetNewORCGroupSeq() As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select max(seq) from AMLOverrideCategory "
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using
        returnValue += 1


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function DeleteORCGroupItems(ByVal objORCGroupItems() As stORCGroupItems) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False

        Try

            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try
                    For intK = 1 To UBound(objORCGroupItems)
                        strSQL = "delete from AMLOverrideCategoryItem where categoryid=" & objORCGroupItems(intK).GroupID &
                        " and commentid=" & objORCGroupItems(intK).CategoryID
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
            Call GenerateException("DeleteORCGroupItems", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult

    End Function
    ''' <summary>
    ''' Get New/Max/High/Unique ORC Group Item ID
    ''' </summary>
    ''' <param name="intGroupId">Group ID</param>
    ''' <returns>Unique ID</returns>
    ''' <remarks></remarks>
    Public Function GetNewORCGroupITemID(ByVal intGroupId As Int16) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intGroupId:({0})", intGroupId), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'Get New and unique Question id
        Dim strSQL As String = "select max(commentid) from AMLOverrideCategoryItem where categoryid= " & intGroupId
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using

        returnValue += 1


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue

    End Function
    ''' <summary>
    ''' Write ORC Group Items
    ''' </summary>
    ''' <param name="objORCGroupItems">object  holds the ORCGroupITems</param>
    ''' <returns>Result: True/False</returns>
    ''' <remarks></remarks>
    Public Function WriteORCGroupItem(ByVal objORCGroupItems() As stORCGroupItems) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False

        Try

            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try

                    For intK = 1 To UBound(objORCGroupItems)
                        strSQL = "delete from AMLOverrideCategoryItem where categoryid=" & objORCGroupItems(intK).GroupID &
                            " and commentid=" & objORCGroupItems(intK).CategoryID
                        connection.ExecuteNonQuery(strSQL)

                        strSQL = "insert into AMLOverrideCategoryItem(categoryid,commentid,seq,commentdesc) values(" &
                            objORCGroupItems(intK).GroupID & "," & objORCGroupItems(intK).CategoryID & "," &
                            intK & ",'" & Strings.ReplaceSingleQuote(objORCGroupItems(intK).Category) & "')"
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
            Call GenerateException("WriteORCGroupItem", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult
    End Function
    ''' <summary>
    ''' Write ORC Groups 
    ''' </summary>
    ''' <param name="objORCGroup">object holding ORC groups</param>
    ''' <returns>Result: True/False</returns>
    ''' <remarks></remarks>
    Public Function WriteORCGroups(ByVal objORCGroup() As stOverridecategoryGroup) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "", blnResult As Boolean = False


        Try
            Using connection = New Connection(gStrConnectionString)
                connection.BeginTransaction()
                Try

                    For intK = 1 To UBound(objORCGroup)
                        'just in case, defensive coding.
                        strSQL = "delete from AMLOverrideCategory where categoryid=" & objORCGroup(intK).GroupID
                        connection.ExecuteNonQuery(strSQL)

                        'insert the record, use intk as seq
                        strSQL = "insert into AMLOverrideCategory(categoryid, seq, categorydesc) values(" &
                            objORCGroup(intK).GroupID & "," & intK & ",'" &
                            Strings.ReplaceSingleQuote(objORCGroup(intK).GroupName) & "')"
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
            Call GenerateException("WriteORCGroups", ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return blnResult

    End Function

    Public Function CheckStationNumberExist(strStationNO As String) As Integer

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strStationNO:({0})", strStationNO), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "select count(1) from stations where " & Ec.GeneralFunc.GetQueryFieldCondition("absnno", strStationNO)
        Dim returnValue = 0
        Using connection = New Connection(gStrConnectionString)
            returnValue = connection.ExecuteScalar(Of Integer)(strSQL, 0)
        End Using


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Public Function GetNewSharedSubheaderUniqueID(Optional getCount As Integer = 1) As List(Of Long)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intSHCount:({0}) ", getCount), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If
        Dim result As List(Of Long) = Nothing

        Try
            Dim strSQL = String.Format("SELECT TOP {0} * FROM vw_AML_NextOperationID ORDER BY 1", getCount)
            Using connection = New Connection(gStrConnectionString)
                Using dataTable = connection.GetDataIntoDataTable(strSQL, Nothing)
                    result = dataTable.AsEnumerable().Select(Function(x) Convert.ToInt64(x(0))).ToList()
                End Using
            End Using
        Catch ex As Exception
            Call GenerateException("GetNewSharedSubheaderUniqueID", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", result.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return result
    End Function

    Public Sub GetPartAreaAndModel(intID As Int32, strRectype As String, intSeq As Int16, ByRef strArea As String, ByRef strModel As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRectype:({1}) intSeq:({2})", intID, strRectype, intSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = ""

        Try
            strArea = "" : strModel = ""

            strSQL = "select user0,user1 from routehdr where id=" & intID &
                " and rectype='" & strRectype.Trim & "' and seq=" & intSeq

            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim reader As DataRow = table.Rows(0)
                        strArea = Data.GetDataRowValue(reader(0), "")
                        strModel = Data.GetDataRowValue(reader(1), "")
                    End If
                End Using
            End Using

        Catch ex As Exception
            GenerateException("GetPartAreaAndModel: " & strSQL, ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0}) strArea:({1})  strModel:({2})", intID, strArea, strModel), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function GetShiftTimes() As stShiftTimes()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'EVC-745: AML: Shift and FLM Frequency changes in ClientEditor and MDM
        Dim strSQL As String = ""
        Dim objST(0) As stShiftTimes
        Dim intK = 0
        Try
            strSQL = "SELECT * FROM aml_shifttimes ORDER BY shiftname"
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table)) Then
                        ReDim Preserve objST(table.Rows.Count)
                        For Each row As DataRow In table.Rows
                            intK += 1

                            objST(intK).ShiftID = Data.GetDataRowValue(Of Byte)(row("ShiftID"), 0)
                            objST(intK).ShiftName = Data.GetDataRowValue(row("ShiftName"), "")
                            objST(intK).StartTime = Data.GetDataRowValue(Of Single)(row("StartTime"), 0)
                            objST(intK).EndTime = Data.GetDataRowValue(Of Single)(row("EndTime"), 0)
                            objST(intK).Enabled = Data.GetDataRowValue(Of Byte)(row("enabled"), 0)
                        Next
                    End If
                End Using
            End Using

            If intK = 0 Then
                ReDim Preserve objST(2)
                'first shift
                objST(1).ShiftName = "1"
                objST(1).ShiftID = 1
                objST(1).StartTime = 6.0
                objST(1).EndTime = 18.0
                'second shift
                objST(2).ShiftName = "2"
                objST(2).ShiftID = 2
                objST(2).StartTime = 18.01
                objST(2).EndTime = 1.3
            End If
        Catch ex As Exception
            GenerateException("GetShiftTimes", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objST)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objST
    End Function

    Public Sub WriteShiftTimes(objST() As stShiftTimes)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Try

            Dim sqlList = New List(Of String)()
            Dim strSQL = "delete from AML_ShiftTimes"
            sqlList.Add(strSQL)


            For intK = 1 To UBound(objST)
                Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(EASEClass7.QueryBuilder.QueryType.Insert, "AML_ShiftTimes")

                queryBuilder.AddField("ShiftID", objST(intK).ShiftID, True)
                queryBuilder.AddField("ShiftName", objST(intK).ShiftName, False)
                queryBuilder.AddField("StartTime", objST(intK).StartTime, True)
                queryBuilder.AddField("EndTime", objST(intK).EndTime, True)
                queryBuilder.AddField("Enabled", objST(intK).Enabled, True)
                strSQL = queryBuilder.GenerateQuery()
                sqlList.Add(strSQL)
            Next

            Ec.IO.RunSQLInArray(sqllist)

        Catch ex As Exception
            GenerateException("WriteShiftTimes", ex)
        End Try


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

    End Sub

    Public Function GetFLMDate() As String

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strTemp2 As String = ""
        strTemp2 = Dates.CurrentDate
        If strTemp2.Trim.Length > 0 Then
            strTemp2 = Left(strTemp2, 10)
        End If


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", strTemp2), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return strTemp2
    End Function

    Public Function GetFLMShiftFrequency(Optional intIDX As Int16 = 0) As stFLMShiftFrequency()
        Using connection = New Connection(gStrConnectionString)
            Return GetFLMShiftFrequency(connection)
        End Using
    End Function

    Public Function GetFLMShiftFrequency(ByVal connection As Connection, Optional intIDX As Int16 = 0) As stFLMShiftFrequency()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intIDX:({0}) ", intIDX), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim objFLMSF() As stFLMShiftFrequency
        ReDim objFLMSF(0)

        Try
            Dim strSQL = "SELECT * FROM AML_FLMFrequency {0} ORDER BY descx"
            strSQL = String.Format(strSQL, If(intIDX > 0, String.Format("WHERE id = {0}", intIDX), ""))

            Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                If (Not IsNothing(table)) Then
                    ReDim Preserve objFLMSF(table.Rows.Count)
                    Dim intCounter = 0
                    For Each row As DataRow In table.Rows
                        intCounter += 1

                        objFLMSF(intCounter).IDX = Data.GetDataRowValue(Of Byte)(row("ID"), 0)
                        objFLMSF(intCounter).DescX = Data.GetDataRowValue(row("DescX"), "")
                        objFLMSF(intCounter).Shift1 = Data.GetDataRowValue(Of Byte)(row("ShiftID1"), 0)
                        objFLMSF(intCounter).Shift2 = Data.GetDataRowValue(Of Byte)(row("ShiftID2"), 0)
                        objFLMSF(intCounter).Shift3 = Data.GetDataRowValue(Of Byte)(row("ShiftID3"), 0)
                        'FrequencyType: 1-Daily, 2- Weekly, 3-Monthly
                        objFLMSF(intCounter).FrequencyType = Data.GetDataRowValue(Of Byte)(row("FrequencyType"), 0)
                    Next
                End If
            End Using
        Catch ex As Exception
            GenerateException("GetFLMShiftFrequency", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(objFLMSF)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return objFLMSF

    End Function

    Public Sub DeleteFLMShiftFrequency(strIDList As String)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strIDList:({0}) ", strIDList), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'EVC-745: AML: Shift and FLM Frequency changes in ClientEditor and MDM
        Dim strSQL As String = "delete from AML_FLMFrequency where id in (" & strIDList & ") "
        Ec.IO.RunSQL(strSQL)


#If TRACE Then
        Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
    End Sub

    Public Function WriteFLMShiftFrequency(objFLMSF As stFLMShiftFrequency) As Boolean

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'EVC-745: AML: Shift and FLM Frequency changes in ClientEditor and MDM
        Dim strSQL As String = "", blnResult As Boolean = False
        Dim intIDX = 0, intQueryType As QueryBuilder.QueryType = QueryBuilder.QueryType.Insert

        Try
            Using connection = New Connection(gStrConnectionString)
                If objFLMSF.IDX = 0 Then
                    strSQL = "select max(ID) from AML_FLMFrequency"
                    Dim result = connection.ExecuteScalar(Of Integer)(strSQL, 0)
                    intIDX = result + 1
                Else
                    intIDX = objFLMSF.IDX
                    intQueryType = EASEClass7.QueryBuilder.QueryType.Update
                End If

                Dim queryBuilder = EASEClass7.QueryBuilder.CreateNewQuery(intQueryType, "AML_FLMFrequency")

                queryBuilder.AddField("DescX", objFLMSF.DescX, False)
                queryBuilder.AddField("ShiftID1", objFLMSF.Shift1, True)
                queryBuilder.AddField("ShiftID2", objFLMSF.Shift2, True)
                queryBuilder.AddField("ShiftID3", objFLMSF.Shift3, True)
                'FrequencyType: 1-Daily, 2- Weekly, 3-Monthly
                queryBuilder.AddField("FrequencyType", objFLMSF.FrequencyType, True)
                If intQueryType = QueryBuilder.QueryType.Update Then
                    queryBuilder.AddConditionField("ID", intIDX.ToString(), True)
                Else
                    queryBuilder.AddField("ID", intIDX, True)
                End If
                strSQL = queryBuilder.GenerateQuery

                connection.ExecuteNonQuery(strSQL)
            End Using
        Catch ex As Exception
            Call GenerateException("WriteFLMShiftFrequency", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", blnResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnResult


    End Function


    Public Function AnyFLMFrequencyUsed(ByVal intFreqID As Int32) As Boolean
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter intFreqID:({0}) ", intFreqID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'EVC-745: AML: Shift and FLM Frequency changes in ClientEditor and MDM
        Dim strSQL As String = "select COUNT(FLMFrequencyID) from PDMPARTMM where FLMFrequencyID=" & intFreqID
        Dim returnValue = False
        Using connection = New Connection(gStrConnectionString)
            Dim result = connection.ExecuteScalar(Of Integer)(strSQL, 0)
            returnValue = result > 0
        End Using


#If TRACE Then
        Log.OPERATION(String.Format("Exit ({0})", returnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return returnValue
    End Function

    Private Function GetMiscflag1ListFromSUBHDRTable() As String()

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim strSQL As String = "SELECT DISTINCT(miscflag1) FROM subhdr WHERE miscflag1 IS NOT NULL ORDER BY miscflag1"
        Dim ArrTemp(0) As String
        Try
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        ReDim Preserve ArrTemp(table.Rows.Count)
                        Dim counter = 0
                        For Each row As DataRow In table.Rows
                            counter += 1
                            ArrTemp(counter) = Data.GetDataRowValue(row(0), "")
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            GenerateException("GetMiscflag1ListFromSUBHDRTable", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0})", UBound(ArrTemp)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return ArrTemp
    End Function

    Public Function ImportSharedOperations(objSharesOpsList() As stImportSharedOps, strUserName As String,
                                           ByRef strDuplicateList As String) As Boolean


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserName:({0})", strUserName), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        'EVC-781: AML: Import shared Operations List from Excel file
        Dim intYY As Int32 = 0
        Dim strSQL As String = ""
        Dim blnRtnValue As Boolean = False
        Dim ArrPlatenID(0) As String, blnDuplicate As Boolean = False
        Dim intDupCount As Int32 = 0
        Try
            strDuplicateList = ""
            If UBound(objSharesOpsList) = 0 Then GoTo ExitThisSub

            ArrPlatenID = GetMiscflag1ListFromSUBHDRTable()
            If UBound(ArrPlatenID) > 0 Then
                For intYY = 1 To UBound(objSharesOpsList)
                    blnDuplicate = False

                    For intAA = 1 To UBound(ArrPlatenID)
                        If objSharesOpsList(intYY).UniqueNO = Strings.ToInt32(ArrPlatenID(intAA)) Then
                            blnDuplicate = True
                            Exit For
                        End If


                    Next

                    If blnDuplicate Then
                        If strDuplicateList.Trim <> "" Then strDuplicateList &= ", "
                        strDuplicateList &= objSharesOpsList(intYY).UniqueNO.ToString
                        objSharesOpsList(intYY).OperationDesc = "DUPLICATE-RECORD"
                        intDupCount += 1
                    End If

                Next
            End If

            If intDupCount = UBound(objSharesOpsList) Then
                GoTo ExitThisSub
            End If

            Dim intPartID = Ec.Parts.GetNewPartID()
            Dim intSHID = Ec.SubHeaders.GetNewSubHeaderID()

            Dim sqlList = New List(Of String)()
            For intYY = 1 To UBound(objSharesOpsList)
                If objSharesOpsList(intYY).OperationDesc = "DUPLICATE-RECORD" Then
                    GoTo SkipThisRecord
                End If

                strSQL = "insert into partxref (id,partno,plantid) values (" & intPartID & ",'SHARED-SH-" & intPartID & "',1)"
                sqlList.Add(strSQL)

                strSQL = "insert into routehdr (partdesc,lotsize,eng,user0,user1,user2,user3,matcost,expunit,explot,ccost,scost,change,tcost," &
                    " toolcosts,setuphrs,runhrs,takttime,plantid,masterpart,nop,ContinousFlow,shlevel,authrel,partrev,flag2,id,rectype,seq,flag1,flag3) " &
                    "values ('SHARED SUB-HEADER' ,1,'" & strUserName & "' ,'' ,'' ,'' ,'' ,0,0,0,0,0,'' ,0," &
                    "0,0,0,0,1,'' ,1,0,1,'00' ,1,0," & intPartID & ",'0' ,0,0,99)"
                sqlList.Add(strSQL)

                strSQL = "delete from capphdr7  where id=" & intPartID & " and  UPPER(rectype)='0'  and seq=0"
                sqlList.Add(strSQL)

                strSQL = "insert into ophdr (opseq,opdesc,workcent,maxbatch,datarec,pfdcyc,pfdsetup,realcyc,realsetup,acostrate," &
                    " acyctime,asetuptime,BasicRunTime,BasicSetupTime,stationtime,effectivefrom,userdef1,userdef2,userdef3," &
                    " toolcostsu,ProcessFlag,nomen,altflag,costkey,va,nva,essnva,checkout_engineer,sharedopid," &
                    "engineer,releasedflag,oprev,opreldate,machrecno,id,rectype,seq,opno) values " &
                    "(1,'SHARED' ,'' ,0,0,0,0,1,1,0" &
                    ",0,0,0,0,0,0,'' ,'' ,'' ," &
                    "0,'' ,3,'0' ,0,0,0,0,'' ,0," &
                    "'" & strUserName & "' ,0,1,0,0," & intPartID & ",'0' ,0,'SHARED' )"
                sqlList.Add(strSQL)

                strSQL = "delete from cappophdr7  where id=" & intPartID & " and  UPPER(rectype)='0'  and seq=0 and  UPPER(opno)='SHARED'"
                sqlList.Add(strSQL)
                strSQL = "insert into subhdr (shseq,id,rectype,seq,opno,descx,changeflag,nummen,model1,destination1,destination2," &
                    " option1,option2,excludefromprint,setuptime,cycletime,ganttstartfrom,lbsticky," &
                    " miscflag1,model2,plattenid,SHAREDSHID,shid) values (" &
                    "1," & intPartID & ",'0' ,0,'SHARED' ,'" & Strings.ReplaceSingleQuote(objSharesOpsList(intYY).OperationDesc.Trim) & "' ,'' ,'A' ,0,0,0," &
                    "0,0,0,0,0,0,0," &
                    "'" & objSharesOpsList(intYY).UniqueNO & "' ," & objSharesOpsList(intYY).UniqueNO & ",'' ,0," & intSHID & ")"
                sqlList.Add(strSQL)

                intPartID += 1 : intSHID += 1
SkipThisRecord:
            Next
            strSQL = "update easesys set partno=" & intPartID + 1 & " where " & Ec.GeneralFunc.GetQueryFieldCondition("reckey", "PART")
            sqlList.Add(strSQL)

            strSQL = "update easesys set partno=" & intSHID + 1 & " where reckey='SHDR'"
            sqlList.Add(strSQL)

            Ec.IO.RunSQLInArray(sqllist, blnRtnValue)

ExitThisSub:
        Catch ex As Exception
            Call GenerateException("ImportSharedOperations: " & strSQL, ex)
        End Try


#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0}) blnRtnValue:({1})", UBound(objSharesOpsList), blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return blnRtnValue
    End Function

    Public Function GetUsersList_Version7(Optional ByVal strUserID As String = "",
        Optional ByRef blnFoundUser As Boolean = False) As List(Of AML_User)


#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter strUserID:({0})", strUserID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        blnFoundUser = False
        'Get all the EASE users list
        Dim users = New List(Of AML_User)
        Dim strSQL As String = "select * FROM euser "
        If Trim(strUserID) <> "" Then
            strSQL &= " where " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID)

        End If
        strSQL &= " order by userid"

        Dim strTemp As String = ""

        Try
            Using connection = New Connection(gStrConnectionString)
                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        Dim counter = 0
                        For Each reader As DataRow In table.Rows
                            counter += 1
                            Dim user As New AML_User

                            user.Keyx = counter
                            user.ID = Data.GetDataRowValue(reader("userid"), "")
                            user.Name = Data.GetDataRowValue(reader("username"), "")

                            Dim strPassword = Data.GetDataRowValue(reader("Passwordx"), "")
                            'We are encrypting password in MCR and ViewEASE from March 2007 (only in EASE SP2)
                            strPassword = Ec.User.DecryptUserPassword(strPassword)
                            Ec.GeneralFunc.ReplaceControlCharacters(strPassword, "")
                            user.Password = strPassword
                            user.Department = ""
                            user.Email = Data.GetDataRowValue(reader("email"), "")
                            user.ExpDays = 0
                            user.ExpStartDate = 0
                            user.Loggedon = Data.GetDataRowValue(Of Int16)(reader("Loggedonease"), 0)
                            user.PlanRes = ""
                            user.PlanRes1 = ""
                            user.PlanResNum = 0
                            user.PlanResNum1 = 0

                            'Access Rights - > 1-All,  4-Process Plan rwd, 5-Process Plan rw, 6-Process Plan Read, 7-Readonly-Signoff
                            user.UAccess = 0
                            user.Modules = ""
                            user.PlantId = 0

                            users.Add(user)
                        Next
                        If counter > 0 Then blnFoundUser = True
                    End If
                End Using

                strSQL = "select euserrights.*,easegroups.groupname from euserrights, easegroups where easegroups.plantid=euserrights.plantid and easegroups.groupid=euserrights.groupid "

                If Trim(strUserID) <> "" Then
                    strSQL &= " and " & Ec.GeneralFunc.GetQueryFieldCondition("userid", strUserID)
                End If
                strSQL &= " order by euserrights.plantid,euserrights.groupid,euserrights.userid"

                Using table = connection.GetDataIntoDataTable(strSQL, Nothing)
                    If (Not IsNothing(table) AndAlso table.Rows.Count > 0) Then
                        For Each reader As DataRow In table.Rows

                            strTemp = Data.GetDataRowValue(reader("userid"), "")
                            For Each user In users
                                If user.ID.Trim.ToUpper = strTemp.Trim.ToUpper Then
                                    user.PlantId = Data.GetDataRowValue(Of Int16)(reader("plantid"), 0)
                                    user.Department = Data.GetDataRowValue(reader("groupname"), "")
                                    user.Modules = Data.GetDataRowValue(reader("Modules"), "")
                                    user.UAccess = Data.GetDataRowValue(Of Int16)(reader("easeaccess"), 0)

                                    Exit For
                                End If
                            Next
                        Next
                    End If
                End Using
            End Using

        Catch ex As Exception
            Call GenerateException("GetUsersList_Version7", ex)
        End Try

#If TRACE Then
        Log.OPERATION(String.Format("Exit Count:({0}) blnFoundUser:({1})", users.Count, blnFoundUser), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return users

    End Function

    Public Function GetStationOptionsByBuildAndStation(ByVal buildNumber As String, ByVal stationNumber As String) As List(Of StationOption)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter buildNumber:({0}), stationNumber:({1})", buildNumber, stationNumber), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim result = New List(Of StationOption)

        Dim strSQL As String = "SELECT distinct T1.BuildNumber" +
                                ",t1.Model" +
                                ",t2.OptionCode1" +
                                ",t2.OptionDescription" +
                                ",t4.TrackStation" +
                                " from [dbo].[AML_VehicleOrders] T1" +
                                " join [dbo].[AML_VehicleOptions] t2 on t2.[BuildNumber] = t1.[BuildNumber]" +
                                " join [dbo].[AML_VehicleBOM] t4 on t2.OptionCode1 = t4.OptionCode and convert (Nvarchar,t1.Model1) = convert (Nvarchar,t4.Model)" +
                                "where T1.BuildNumber = '{0}' and TrackStation = '{1}'"

        Using connection = New Connection(gStrConnectionString)
            Using stationOptionTable = connection.GetDataIntoDataTable(String.Format(strSQL, buildNumber, stationNumber), Nothing)

                For Each row As DataRow In stationOptionTable.Rows
                    Dim stationOption As New StationOption With {.BuildNumber = buildNumber,
                        .TrackStation = stationNumber,
                        .Model = Data.GetDataRowValue(row("Model"), ""),
                        .OptionCode = Data.GetDataRowValue(row("OptionCode1"), ""),
                        .OptionDescription = Data.GetDataRowValue(row("OptionDescription"), "")}

                    result.Add(stationOption)
                Next

            End Using
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit GetStationOptionsByBuildAndStation: ({0})", result.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return result

    End Function

    Public Function GetStationVehicleTrims(ByVal station As String) As List(Of StationVehicleTrim)

#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter GetStationVehicleTrims"), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim arrTemp As New List(Of StationVehicleTrim)

        arrTemp = EASEClass7.EASECache.GetVehicleTrims()

        Dim strSQL As String = "Select environment from AML_StationVehicleTrim " &
                                "where StationNumber = '{0}'"

        Using connection = New Connection(gStrConnectionString)
            Using vehicleTrimTable = connection.GetDataIntoDataTable(String.Format(strSQL, station), Nothing)
                If (Not IsNothing(vehicleTrimTable) AndAlso vehicleTrimTable.Rows.Count > 0) Then
                    For Each row As DataRow In vehicleTrimTable.Rows
                        If arrTemp.Any(Function(f) f.Environment.Equals(Trim(Data.GetDataRowValue(row("Environment"), "")))) Then
                            arrTemp.Where(Function(f) f.Environment.Equals(Trim(Data.GetDataRowValue(row("Environment"), "")))).First.Selected = True
                        End If
                    Next
                End If
            End Using
        End Using

#If TRACE Then
        Log.OPERATION(String.Format("Exit GetStationVehicleTrims: ({0})", arrTemp.Count), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        Return arrTemp

    End Function

    Public Function GetVehicleTrimsByBuild(ByVal buildNo As String) As List(Of AML_VehicleTrim)
#If TRACE Then
        Dim startTicks As Long = Log.OPERATION(String.Format("Enter buildNo:({0})", buildNo), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

        Dim returnValue As New List(Of AML_VehicleTrim)
        Try

            Dim strSQL = String.Format("SELECT * from aml_vehicletrim " &
                                        "WHERE buildnumber = '{0}'", buildNo)

            Using connection = New Connection(gStrConnectionString)
                Dim result = connection.GetDataIntoClassOf(Of AML_VehicleTrim)(strSQL)
                If (result IsNot Nothing) Then returnValue = result.ToList
            End Using
        Catch ex As Exception
            GenerateException("GetVehicleOrderDetails", ex)
        End Try

#If TRACE Then
        Log.OPERATION("Exit GetBuildVehicleTrims", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        Return returnValue
    End Function

    Public Function CheckForUniqueSerialCheck(ByVal strBuildNO As String, ByVal strPartNO As String, ByVal strSerialNO As String) As String
        Dim message As String = ""

        Try
            Dim strSQL = String.Format("SELECT Status from AML_UCPARTCHECK " &
                                        "WHERE buildno = '{0}' and partno = '{1}' and serial='{2}'", strBuildNO, strPartNO, strSerialNO)

            Using connection = New Connection(gStrConnectionString)
                message = connection.GetFirstFieldValue(Of String)(strSQL)
            End Using

            If String.IsNullOrEmpty(message) Then
                strSQL = String.Format("SELECT Count(*) from AML_UCPARTCHECK " &
                                        "WHERE buildno = '{0}' and partno = '{1}'", strBuildNO, strPartNO)

                Using connection = New Connection(gStrConnectionString)
                    Dim count As Int16 = connection.ExecuteScalar(Of Int16)(strSQL, 0, "CheckForUniqueSerialCheck")

                    If count > 0 Then
                        message = "Incorrect serial for build"
                    End If
                End Using
            End If
        Catch ex As Exception
            GenerateException("GetVehicleOrderDetails", ex)
        End Try

        Return message
    End Function
End Class
