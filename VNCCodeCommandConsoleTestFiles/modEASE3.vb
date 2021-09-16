Imports Microsoft.VisualBasic
Imports System.Data
Imports EaseCore


Namespace EASEWebApp
    Public Module modEASE3
        Private Const BASE_ERRORNUMBER As Integer = EaseCore.ErrorNumbers.EASEWEBAPP_EASE3
        'Private Const BASE_ERRORNUMBER As Integer = ErrorNumbers.CLIENTEDITOR_ABOUT
        Private Const LOG_APPNAME As String = "EASEWEBAPP"
#Const TRACE = 1

        Public Function GetPlanFieldValue(ByVal intHeaderType As Int16, ByVal intFieldSeq As Int16, Optional ByVal blnUseCurrentTimeUnit As Boolean = False) As String

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intHeaderType:({0}) intFieldSeq:({1}) blnUseCurrentTimeUnit:({2})", intHeaderType, intFieldSeq, blnUseCurrentTimeUnit), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim strRtnValue As String = ""
            Dim intPlanSignoffCount As Int16 = 0, intAuthCount As Int16 = 0
            'Header Field Type: 0-Plan Header, 1-Operation Header, 2-Element Header, 3-Plan Extra fields, 4-Operation Extra fields
            '                 : 5-Tooling records, 6-Machines, 7-Work Center, 8- Custom ODBC Link,9-Material User Defined Fields
            '                 : 10- Plan Display Fields (used only for reporting and display purpose)
            '                 : 11- Operation Display Fields (used only for reporting and display purpose)
            '                 : 12 - Shared Operation header fields
            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")

            If intHeaderType = 0 Then       'plan header
                Select Case intFieldSeq
                    Case 1
                        strRtnValue = objRH.PartNo
                    Case 2
                        strRtnValue = objRH.Desc
                    Case 3
                        strRtnValue = objRH.LotSize
                    Case 4
                        If Ec.AppConfig.CAPP Or Ec.AppConfig.CAPPOnly Then
                            If objRH.Released Then
                                strRtnValue = Ec.AppConfig.GetWrd(3573)
                            Else
                                intPlanSignoffCount = Ec.Parts.GetPartSignoffCount(objRH.AuthRel)
                                If intPlanSignoffCount > 0 Then
                                    'atleast one signoff has been complete
                                    intAuthCount = Ec.DBConfig.GetAuthLevelCount(objRH.PlantID)


                                    '1/2 complete
                                    strRtnValue = intPlanSignoffCount.ToString & "/" & intAuthCount.ToString & " " & Ec.AppConfig.GetWrd(2964)
                                Else
                                    'default
                                    strRtnValue = Ec.AppConfig.GetWrd(3574)
                                End If


                            End If
                        End If
                    Case 5
                        strRtnValue = objRH.Engineer
                    Case 6
                        If Ec.AppConfig.CAPP Or Ec.AppConfig.CAPPOnly Then
                            If Not EASEClass7.DBConfig.OperationSignoff Then
                                strRtnValue = objRH.ReleaseDate
                            End If
                        End If
                    Case 7
                        strRtnValue = objRH.User0
                    Case 8
                        strRtnValue = objRH.User1
                    Case 9
                        strRtnValue = objRH.User2
                    Case 10
                        strRtnValue = objRH.User3
                    Case 11
                        strRtnValue = objRH.Change
                    Case 12
                        strRtnValue = objRH.TaktTime
                        If blnUseCurrentTimeUnit Then
                            strRtnValue = Format(objRH.TaktTime * Ec.AppConfig.TimeUnit, Ec.AppConfig.TimeFormat)
                        End If
                    Case 13
                        strRtnValue = Ec.AppConfig.GetWrd(2907)
                        If objRH.ContinousFlow Then
                            strRtnValue = Ec.AppConfig.GetWrd(2906)
                        End If
                End Select
            ElseIf intHeaderType = 3 Then       'CAPP Fields - plan header
                If intFieldSeq <= UBound(objRH.CAPPFields) Then
                    strRtnValue = objRH.CAPPFields(intFieldSeq)
                End If
            ElseIf intHeaderType = 10 Then       'CAPP Fields - plan header
                If intFieldSeq = 1 Then
                    If Ec.AppConfig.CAPP Or Ec.AppConfig.CAPPOnly Then
                        strRtnValue = objRH.RevNo
                        If objRH.Released Then
                            strRtnValue &= ", " & Trim(objRH.ReleaseDate)
                        End If
                    End If
                ElseIf intFieldSeq = 2 Then
                    strRtnValue = objRH.NOP
                End If
            End If


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", strRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return strRtnValue.Trim

        End Function

        Public Function GetHeaderFieldsFromMemory(ByVal intHeaderType As Int16) As EASEClass7.DBConfig.stHeaderFields()

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intHeaderType:({0})", intHeaderType), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'This function gets the fields for plan header and operation header only
            Dim objSF(0) As EASEClass7.DBConfig.stHeaderFields
            Dim intK As Int16 = 0, intCount As Int16 = 0
            Dim objInMemory(0) As EASEClass7.DBConfig.stHeaderFields
            Try
                'Header Field Type: 0-Plan Header, 1-Operation Header, 2-Element Header, 3-Plan Extra fields, 4-Operation Extra fields
                '                 : 5-Tooling records, 6-Machines, 7-Work Center, 8- Custom ODBC Link,9-Material User Defined Fields
                '                 : 10- Plan Display Fields (used only for reporting and display purpose)
                '                 : 11- Operation Display Fields (used only for reporting and display purpose)
                '                 : 12 - Shared Operation header fields

                If intHeaderType = 0 Or intHeaderType = 3 Then
                    objInMemory = HttpContext.Current.Session("gRHHeaderFields")
                ElseIf intHeaderType = 1 Or intHeaderType = 4 Then
                    objInMemory = HttpContext.Current.Session("gOPHeaderFields")
                End If

                For intK = 1 To UBound(objInMemory)
                    If objInMemory(intK).FieldType <> intHeaderType Then GoTo SkipThisRecord

                    intCount += 1
                    ReDim Preserve objSF(intCount)
                    objSF(intCount) = objInMemory(intK)
SkipThisRecord:
                Next
            Catch ex As Exception
                GeneralError("GetHeaderFieldsFromMemory ", ex, True)
            Finally
                objInMemory = Nothing
            End Try
            CheckForUserSessionErrors()


#If TRACE Then
            Log.OPERATION(String.Format("Exit Count:({0})", UBound(objSF)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return objSF
        End Function

        Public Function CheckForFileType(ByVal strFileName As String, ByVal objFE() As EASEClass7.stFileExtension, ByVal intDocType As Int16) As Boolean

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strFileName:({0}) intDocType:({1})", strFileName, intDocType), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim blnRtnValue As Boolean = False, intK As Int16 = 0
            Dim strTemp As String = ""

            strTemp = Ec.GeneralFunc.GetFileExtension(strFileName)
            For intK = 1 To UBound(objFE)
                If Ec.GeneralFunc.CompareString(objFE(intK).FileExtn, strTemp) Then
                    If intDocType > 0 Then
                        blnRtnValue = objFE(intK).DocType = intDocType
                    Else
                        blnRtnValue = True
                    End If
                End If
            Next

#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Sub DisplayPlanTypesGrid(ByVal intID As Int32, ByVal UGData As Infragistics.Web.UI.GridControls.WebDataGrid, _
                                   Optional ByVal blnRowSelector As Boolean = False, Optional ByVal strRectye As String = "0")


#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED(String.Format("Enter intID:({0}) blnRowSelector:({1}) strRectye:({2})", intID, blnRowSelector, strRectye), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim dataTable As New DataTable("tblData")

            dataTable.Clear()

            'GetType(Int16)
            Dim colDB As New DataColumn(Ec.AppConfig.GetWrd(3551), GetType(String))    'Plan Type
            Dim objPlanTypes(0) As EASEClass7.Parts.stPlanTypes
            Dim objRow As DataRow, intK As Int16 = 0, strRectype As String = ""
            Dim strPrevRecType As String = ""
            Try
                FormatUltraGrid(UGData, blnRowSelector)

                dataTable.Columns.Add(colDB)

                colDB = New DataColumn(Ec.AppConfig.GetWrd(25), GetType(String))         'Description
                dataTable.Columns.Add(colDB)

                colDB = New DataColumn(Ec.AppConfig.GetWrd(3582), GetType(String))         'Rev
                dataTable.Columns.Add(colDB)

                colDB = New DataColumn("rowcolor", GetType(String))         'row color/rectype, don't change this   
                dataTable.Columns.Add(colDB)

                colDB = New DataColumn("rectype", GetType(String))         'seq  don't change this (** used to open the plan)
                dataTable.Columns.Add(colDB)

                colDB = New DataColumn("seq", GetType(String))         'seq  don't change this (** used to open the plan)
                dataTable.Columns.Add(colDB)

                objPlanTypes = Ec.Parts.GetPlanTypes(intID, strRectye)
                For intK = 1 To UBound(objPlanTypes)
                    objRow = dataTable.NewRow()

                    strRectype = objPlanTypes(intK).Rectype
                    If Trim(strRectype) <> Trim(strPrevRecType) Then
                        objRow(0) = GetPlanTypeWord(strRectype)
                        strPrevRecType = strRectype
                    Else
                        objRow(0) = ""
                    End If

                    objRow(1) = objPlanTypes(intK).PartDesc
                    objRow(2) = objPlanTypes(intK).PartRev
                    objRow(3) = strRectype
                    objRow(4) = objPlanTypes(intK).Rectype
                    objRow(5) = objPlanTypes(intK).Seq


                    dataTable.Rows.Add(objRow)
                Next

                UGData.DataSource = dataTable
                UGData.DataBind()
                '**SS**
                'UGData.DisplayLayout.Bands(0).Columns(0).Width = New Unit("135px")         'plan type
                'UGData.DisplayLayout.Bands(0).Columns(1).Width = New Unit("100%")         'part desc
                'UGData.DisplayLayout.Bands(0).Columns(2).Width = New Unit("60px")          'rev
                'UGData.DisplayLayout.Bands(0).Columns(3).Hidden = True       'rowcolor
                'UGData.DisplayLayout.Bands(0).Columns(4).Hidden = True       'rectype
                'UGData.DisplayLayout.Bands(0).Columns(5).Hidden = True       'seq

                'If Ec.AppConfig.CAPP Then
                '    If EASEClass7.DBConfig.OperationSignoff Then
                '        UGData.DisplayLayout.Bands(0).Columns(1).Width = 0
                '        ' UGData.DisplayLayout.Bands(0).Columns(1).Width += UGData.DisplayLayout.Bands(0).Columns(2).Width
                '        UGData.DisplayLayout.Bands(0).Columns(2).Hidden = True     'Rev
                '    End If
                'Else
                '    UGData.DisplayLayout.Bands(0).Columns(1).Width = 0
                '    ' UGData.DisplayLayout.Bands(0).Columns(1).Width += UGData.DisplayLayout.Bands(0).Columns(2).Width
                '    UGData.DisplayLayout.Bands(0).Columns(2).Hidden = True     'Rev
                'End If

                'If UGData.Rows.Count > 0 Then
                '    For intK = 0 To UGData.Rows.Count - 1
                '        UGData.Rows(intK).CssClass = GetRectypeColor(UGData.Rows(intK).Cells(3).Text)
                '    Next
                'End If
                '**SS**
                SetPrimaryColumnWidth(UGData, 1)
            Catch ex As Exception
                GeneralError("DisplayPlanTypesGrid", ex, True)
            Finally
                dataTable = Nothing
                colDB = Nothing
                objRow = Nothing
            End Try
            CheckForUserSessionErrors()

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function CopyPlan(ByVal strPartNO As String, ByVal intID As Int32, ByVal strRectype As String, ByVal intSeq As Int16, _
                            ByVal strDestPartNO As String, ByVal strDestPartDesc As String, ByVal strDestRecType As String, _
                            ByVal intDestPlantID As Int32) As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strPartNO:({0}) intID:({1}) strRectype:({2}) intSeq:({3}) strDestPartNO:({4}) strDestPartDesc:({5}) strDestRecType:({6}) intDestPlantID:({7})",
                                                                 strPartNO, intID, strRectype, intSeq, strDestPartNO, strDestPartDesc, strDestRecType, intDestPlantID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim blnRtnValue As Boolean = False
            Dim strMsg As String = "" ', blnOverwrite As Boolean = False
            Dim objCopy As New EASEClass7.Parts.stCopyParams
            Dim intDestSeq As Int16 = 0
            Dim strErrorLocation As String = "Start"
            Try
                'If Not ValidateUserAccessRights(True) Then GoTo ExitThisSub

                If intID = 0 Then
                    strErrorLocation = "GetPartID"
                    intID = Ec.Parts.GetPartID(strPartNO)
                End If
                If intID = 0 Then
                    strMsg = Ec.AppConfig.GetWrd(4017) & " " & strPartNO & " " & Ec.AppConfig.GetWrd(3143)
                    ' MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                    GoTo ExitThisSub
                End If

                'TODO: Plant Access Rights/Tyson 

                If EASEClass7.DBConfig.SharePlan AndAlso Ec.Parts.CheckedOutPart(strPartNO) Then
                    strMsg = Ec.AppConfig.GetWrd(3970) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3136)
                    ' MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                    GoTo ExitThisSub
                End If

                If Ec.Parts.CheckPartNoExist(strDestPartNO) Then
                    If Trim(HttpContext.Current.Session("gFormParam")) = "restorefromrecyclebin" And strPartNO.Trim = strDestPartNO.Trim Then       '**RestoreFromRecycleBin - KM - 05/02/2012**
                        'Archive the active plan (to history) and delete active plan
                        If Not ArchiveActivePlan(intID, strRectype, intSeq, strPartNO, strDestRecType, intDestSeq) Then GoTo ExitThisSub
                    End If

                    'for active plans, if there are no operations, then allow copying (to cover restore from recyclebin option)
                    If strDestRecType.Trim = "0" AndAlso Ec.Parts.GetOperationsCount(intID, strDestRecType, intDestSeq) > 0 Then

                        'plan already exists, can not continue!
                        strMsg = Ec.AppConfig.GetWrd(24) & " " & strDestPartNO.Trim & " " & Ec.AppConfig.GetWrd(3030)
                        ' MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                        GoTo ExitThisSub

                        ''strMsg = Ec.AppConfig.GetWrd(24) & " " & strDestPartNO.Trim & " " & Ec.AppConfig.GetWrd(3136)
                        'If MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                        '    blnOverwrite = True
                        'Else
                        '    GoTo ExitThisSub
                        'End If
                    End If
                Else
                    If strDestRecType.Trim <> "0" Then
                        'active plan doesn't exist, change the plan type and try again.
                        strMsg = Ec.AppConfig.GetWrd(3554) & " " & Ec.AppConfig.GetWrd(3125) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3969)
                        '   MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                        GoTo ExitThisSub
                    End If
                End If

                'TODO: Is It a ChildPart
                'TODO: Shared OP Container

                Ec.Parts.ClearCopyParamsObject(objCopy)

                'copy the operation from shared part to active part 
                objCopy.FromID = intID
                objCopy.FromRectype = strRectype
                objCopy.FromSeq = intSeq


                objCopy.ToPartNo = strDestPartNO
                objCopy.ToPartDesc = strDestPartDesc
                objCopy.ToID = 0
                objCopy.ToRectype = strDestRecType
                objCopy.ToSeq = intDestSeq
                'get the plant id for the source part
                objCopy.ToPlantID = intDestPlantID 'Ec.Parts.GetPartPlantID(objCopy.FromID)
                strErrorLocation = "CopyPartOperation"
                If Not Ec.Parts.CopyPartOperation("copyplan", objCopy, Ec.AppConfig.EASETempPath, HttpContext.Current.Session("UserName")) Then
                    'error occurred while copying the operation into Active. exit out
                    GoTo ExitThisSub
                End If


                blnRtnValue = True
ExitThisSub:
            Catch ex As Exception
                GeneralError("CopyPlan: " & strErrorLocation.Trim, ex, True)
            Finally
                objCopy = Nothing
            End Try
            CheckForUserSessionErrors()


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Function ValidatePlanDelete(ByVal intID As Int32, ByVal strRectype As String, ByVal intSeq As Int16, _
                                  ByVal strPartNO As String, ByRef strValidationMsg As String, _
                                  Optional ByVal blnDeleteOneRecord As Boolean = False, _
                                  Optional ByRef intPartTypeCount As Int16 = 0, _
                                  Optional ByRef blnDelete62Plans As Boolean = False) As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRectype:({1}) intSeq:({2}) strPartNO:({3}) blnDeleteOneRecord:({4})",
                                                                 intID, strRectype, intSeq, strPartNO, blnDeleteOneRecord), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim strPartWrd As String = Ec.AppConfig.GetWrd(24)
            Dim blnRtnValue As Boolean = False, strUserID As String = ""
            Dim strTemp As String = ""

            strValidationMsg = "" : blnDelete62Plans = False
            'check whether the plan can be deleted or not


            Try

                'check part is in version 7 table
                If Not Ec.Parts.PartInVersion7Tables(intID, strRectype, intSeq) Then
                    'oops, plan is still in 6.2 table, just delete it

                    blnDelete62Plans = True

                    'The selected plans are in version 6.2 tables. The plans can not be retrieved, if deleted.

                    blnRtnValue = True
                    GoTo ExitThisFunctiton
                End If

                If EASEClass7.DBConfig.SharePlan AndAlso strRectype.Trim.Trim = "0" Then
                    If Ec.Parts.AnyOperationsCheckedOut(intID, strRectype, intSeq) Then
                        strValidationMsg = strPartWrd.Trim & " '" & strPartNO.Trim & "' " & Ec.AppConfig.GetWrd(3954)
                        'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        GoTo ExitThisFunctiton
                    End If
                End If

                If EASEClass7.DBConfig.SharePlan AndAlso strRectype.Trim.Trim = "0" Then
                    If Ec.Parts.AnyOperationsCheckedOut(intID, strRectype, intSeq) Then
                        strValidationMsg = strPartWrd.Trim & " '" & strPartNO.Trim & "' " & Ec.AppConfig.GetWrd(3954)
                        'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        GoTo ExitThisFunctiton
                    End If
                End If

                If Ec.Parts.CheckPartInUse(intID, strRectype, intSeq, strUserID) Then
                    'ok part is is use by another engineer, strUserID
                    strValidationMsg = Ec.AppConfig.GetWrd(24) & ": '" & Trim(strPartNO) & "' " & Ec.AppConfig.GetWrd(3903)
                    If Not Ec.GeneralFunc.CompareString(strUserID, HttpContext.Current.Session("UserID")) Then
                        strValidationMsg &= " by '" & Trim(strUserID) & "'"
                    End If
                    strValidationMsg &= vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3027)
                    'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    GoTo ExitThisFunctiton
                End If

                If Not blnDeleteOneRecord Then
                    'Do you want to delete Part#
                    strTemp = Ec.AppConfig.GetWrd(3002) & " " & Ec.AppConfig.GetWrd(4017) & " " & strPartNO.Trim & "?"
                    If strRectype <> "0" Then
                        strTemp &= vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3551) & ": " & GetPlanTypeWord(strRectype)
                    End If
                    '!@#$ have to do
                    'If MessageBox.Show(strTemp, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                    '    GoTo ExitThisFunctiton
                    'End If
                End If

                If strRectype.Trim.Trim = "0" Then

                    'for active parts, ensure there are no alternate or history plans exist
                    intPartTypeCount = Ec.Parts.GetPartTypeCount(intID)

                    If Not blnDeleteOneRecord Then
                        'strTemp = Ec.AppConfig.GetWrd(3965) & vbCrLf & Ec.AppConfig.GetWrd(3966)
                        strTemp = Ec.AppConfig.GetWrd(3965) & vbCrLf & Ec.AppConfig.GetWrd(3966) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3046)
                        '!@#$ have to do
                        'If MessageBox.Show(strTemp, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                        '    GoTo ExitThisFunctiton
                        'End If
                    End If
                End If


                blnRtnValue = True

ExitThisFunctiton:
            Catch ex As Exception
                GeneralError("ValidatePlanDelete", ex, True)
            End Try



#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function


        Public Function DeletePlan(ByVal intID As Int32, ByVal strRectype As String, ByVal intSeq As Int16, _
                               ByVal strPartNO As String, Optional ByVal blnDeleteOneRecord As Boolean = False, _
                               Optional ByVal blnForceDelete As Boolean = False, Optional ByVal blnNoRecycleBin As Boolean = False, _
                               Optional ByVal blnSkipValidation As Boolean = False, Optional ByVal blnDeletePartXrefRecord As Boolean = False, Optional ByVal blnRestoreFromRecycleBin As Boolean = False) As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRectype:({1}) intSeq:({2}) strPartNO:({3}) blnDeleteOneRecord:({4}) blnForceDelete:({5}) blnNoRecycleBin:({6}) blnSkipValidation:({7}) blnDeletePartXrefRecord:({8}) blnRestoreFromRecycleBin:({9})",
                                                                 intID, strRectype, intSeq, strPartNO, blnDeleteOneRecord, blnForceDelete, blnNoRecycleBin, blnSkipValidation, blnDeletePartXrefRecord, blnRestoreFromRecycleBin), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'blnDeleteOneRecord=true only when this sub is called from frmdelete plan
            'blnNOValidations - just delete this part.
            '**RestoreFromRecycleBin - KM - 05/02/2012**
            Dim blnRtnValue As Boolean = False
            Dim strRBPartNO As String = "", intPartTypeCount As Int16 = 0
            Dim strRBRectype As String = "3"      'recycle bin rectype
            Dim intRBSeq As Int16 = 0, strMsg As String = ""
            Dim objPlanTypes(0) As EASEClass7.Parts.stPlanTypes
            Dim intK As Int16 = 0, blnDeletePartXref As Boolean = False
            Dim objTextParams As New EASEClass7.Text.stTextParams

            Try
                blnDeletePartXref = blnDeletePartXrefRecord
                strRBPartNO = Ec.Parts.AddSpecialCharForRecycleBinPart(strPartNO)

                If blnForceDelete Then
                    blnDeletePartXref = True
                    GoTo JustDelete
                End If


                If blnSkipValidation Then GoTo SkipValidations '(used  in ProcessPlan->frmSearch )

                If Not ValidatePlanDelete(intID, strRectype, intSeq, strPartNO, strMsg, blnDeleteOneRecord, intPartTypeCount) Then
                    If strMsg.Trim <> "" Then
                        '  MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                    End If
                    GoTo ExitThisFunctiton
                End If
SkipValidations:

                '            If EASEClass7.DBConfig.SharePlan AndAlso strRectype.Trim.Trim = "0" Then
                '                If Ec.Parts.AnyOperationsCheckedOut(intID, strRectype, intSeq) Then
                '                    strMsg = Ec.AppConfig.GetWrd(3954)
                '                    MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                '                    GoTo ExitThisFunctiton
                '                End If
                '            End If


                '            If EASEClass7.DBConfig.SharedOperation Then
                '                'check, if the part no is a shared part container
                '                If Not Ec.Parts.CheckForSharedOpContainer(strPartNO) Then GoTo SkipSharedOp

                '                'OK, this is a shared op container, now check, if this op is being used anywhere
                '                If Ec.Parts.CheckSharedOperationUsage(intID) Then
                '                    strMsg = Ec.AppConfig.GetWrd(4254) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(4255) & Ec.AppConfig.GetWrd(4256)
                '                    MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                '                    GoTo ExitThisFunctiton
                '                End If
                'SkipSharedOp:
                '            End If
                '            'TODO: Hitachi Flag
                '            strMsg = ""
                '            If Ec.Parts.CheckPartInUse(intID, strRectype, intSeq, strUserID) Then
                '                'ok part is is use by another engineer, strUserID
                '                strMsg = Ec.AppConfig.GetWrd(24) & ": '" & Trim(strPartNO) & "' " & Ec.AppConfig.GetWrd(3903)
                '                If Not Ec.GeneralFunc.CompareString(strUserID, gUser.UserID) Then
                '                    strMsg &= " by '" & Trim(strUserID) & "'"
                '                End If
                '                strMsg &= vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3027)
                '                MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                '                GoTo ExitThisFunctiton
                '            End If

                '            If Not blnDeleteOneRecord Then
                '                'Do you want to delete Part#
                '                strMsg = Ec.AppConfig.GetWrd(3002) & " " & Ec.AppConfig.GetWrd(4017) & " " & strPartNO.Trim & "?"
                '                If strRectype <> "0" Then
                '                    strMsg &= vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3551) & ": " & GetPlanTypeWord(strRectype)
                '                End If
                '                If MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                '                    GoTo ExitThisFunctiton
                '                End If

                '            End If



                '            If strRectype.Trim.Trim = "0" AndAlso Not blnDeleteOneRecord Then
                '                'for active parts, ensure there are no alternate or history plans exist
                '                intPartTypeCount = Ec.Parts.GetPartTypeCount(intID)
                '                If intPartTypeCount > 1 Then
                '                    strMsg = Ec.AppConfig.GetWrd(3965) & vbCrLf & Ec.AppConfig.GetWrd(3966)
                '                    If MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                '                        GoTo ExitThisFunctiton
                '                    End If
                '                End If
                'End If


                'TODO: Check, if the Plan is locked



                If blnDeleteOneRecord Then  'deleting one record at a time
                    GoTo StartDeleting
                End If

                'intPartTypeCount = Ec.Parts.GetPartTypeCount(intID)
                If intPartTypeCount > 1 Then
                    objPlanTypes = Ec.Parts.GetPlanTypes(intID)

                    For intK = 1 To UBound(objPlanTypes)
                        If objPlanTypes(intK).Rectype.Trim = "0" Then GoTo SkipThisPlan 'don't delete the active plan, need to move this to recycle bin

                        'delete this plan
                        Ec.Text.ClearTextParamsObject(objTextParams)

                        objTextParams.Param1 = intID
                        objTextParams.Param2 = objPlanTypes(intK).Rectype
                        objTextParams.Param3 = objPlanTypes(intK).Seq

                        'all the operations are checked in, just delete the part header
                        Ec.Parts.DeleteData(1, objTextParams, HttpContext.Current.Session("UserName"), False, False)
SkipThisPlan:
                    Next
                End If

StartDeleting:
                If strRectype.Trim.Trim = "0" AndAlso blnNoRecycleBin = False Then
                    'deleting the active op, will move the data to recycle bin
                    Ec.Parts.MovePlan(intID, strRectype, intSeq, intID, strRBRectype, intRBSeq, strRBPartNO, HttpContext.Current.Session("UserName"))
                Else
                    'delete from the database
JustDelete:
                    'delete this plan
                    Ec.Text.ClearTextParamsObject(objTextParams)

                    objTextParams.Param1 = intID
                    objTextParams.Param2 = strRectype
                    objTextParams.Param3 = intSeq


                    '**RestoreFromRecycleBin - KM - 05/02/2012**
                    'when restroing the plan from recycle bin, need to delete the active pla
                    'but not the partxref record. All other times delete partxref record
                    If blnRestoreFromRecycleBin = False And strRectype.Trim = "0" Then
                        blnDeletePartXref = True
                    End If

                    'all the operations are checked in, just delete the part header
                    Ec.Parts.DeleteData(1, objTextParams, HttpContext.Current.Session("UserName"), False, blnDeletePartXref)
                End If

                blnRtnValue = True

ExitThisFunctiton:
            Catch ex As Exception
                GeneralError("DeletePlan", ex, True)
            Finally
                objTextParams = Nothing
                objPlanTypes = Nothing
            End Try



#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Function ImportPlanDetails(ByVal strImportPath As String, ByRef ObjPartsList() As EASEClass7.PartSearch.stPlanSearchResult) As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strImportPath:({0})", strImportPath), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim strErrLoc As String = "Start", blnRtnValue As Boolean = False
            Dim blnCopyPart As Boolean = False, blnCopyOP As Boolean = False
            Dim blnCopyPDFText As Boolean = True    'while exporting the data, export all the PDF docs too
            Dim blnCopyChangeHistory As Boolean = False
            Dim ArrSQL(0) As String, intK As Int32 = 0

            Dim objCopy As New EASEClass7.Parts.stCopyParams
            Dim strURL As String = ""
            Dim strMsg As String = ""
            Dim strHeaderFile As String = EaseCore.Extensions.Strings.AddSlash(strImportPath) & "easeoffline.xml"
            Dim objDocsList(0) As EASEClass7.Parts.stDocumentExport
            Dim strPartsExist As String = "", StrNoActivePlans As String = ""
            Dim strSQLFile As String = "", blnPartExist As Boolean = False
            Dim objTextParams As New EASEClass7.Text.stTextParams, intCount As Int16 = 0
            Dim strDocsPath As String = "", strCopyPartNo As String = ""
            Dim blnImportAllPlanTypes As Boolean = False, strOldPartNO As String = ""
            Dim intY As Int16 = 0
            Dim intIteration As Int16 = 0, intStatus As Int16 = 0

            Try
                If UBound(ObjPartsList) = 0 Then GoTo ExitThisFunction 'no parts inthe list, exit out, defensive coding, hi hi

                'check the import header file exists
                If Not Ec.GeneralFunc.FileExists(strHeaderFile) Then
                    strMsg = Ec.AppConfig.GetWrd(3978) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3979)
                    DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                    'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    GoTo ExitThisFunction
                End If

                'check user access rights, before importing the data into the database
                'EASE Access Rights: 1-All, 4-Process Plan RWD, 5-Process Plan RW, 6- Process Plan Read,
                '                    7-Read only Sign-off
                If HttpContext.Current.Session("UserAccessRights") = 1 Or HttpContext.Current.Session("UserAccessRights") = 4 Then
                Else
                    strMsg = Ec.AppConfig.GetWrd(3122) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3182) & _
                            " " & GetEASEAccessRightsValue(HttpContext.Current.Session("UserAccessRights")) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(4009)
                    DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                    ' MessageBox.Show(strMsg, Application.ProductName.Trim, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    GoTo ExitThisFunction
                End If


                'DisplayStatusForm("start", 1)


                'get which level need to copy: Plan/op and with or without pdf text
                Ec.Parts.GetCopyParamsFlags("exportplan", blnCopyPart, blnCopyOP, blnCopyPDFText)


                intIteration = 70 / UBound(ObjPartsList)
                intStatus = 20
                For intK = 1 To UBound(ObjPartsList)
                    'DisplayStatusForm("Processing Plan: " & ObjPartsList(intK).PartNO, intStatus)


                    If ObjPartsList(intK).OpDescX = "" Then GoTo SkipImporting
                    'importallPlantypes includes, alternate, history plans
                    blnImportAllPlanTypes = (ObjPartsList(intK).SubHeaderDescX.Trim = "**allplantypes**" AndAlso ObjPartsList(intK).Rectype <> 0)

                    blnPartExist = Ec.Parts.CheckPartNoExist(ObjPartsList(intK).PartNO)
                    strCopyPartNo = ObjPartsList(intK).PartNO

                    If ObjPartsList(intK).Rectype = "0" Then            'active plan
                        If blnPartExist Then

                            'DisplayStatusForm("", 1)            'to hide the status screen

                            HttpContext.Current.Session("gAN") = ObjPartsList(intK).PartNO  '!@#$ have to do

                            strOldPartNO = ObjPartsList(intK).PartNO.Trim
                            strURL = "frmPlanImportExist.aspx"
                            GoTo ExitThisFunction
                            ' ShowDialogForm(New frmPlanImportExist)
                            If HttpContext.Current.Session("gAN") = "*" Then GoTo SkipImporting

                            If HttpContext.Current.Session("gAN") = "overwrite" Then
                                'delete the plan

                                'delete this plan
                                Ec.Text.ClearTextParamsObject(objTextParams)

                                objTextParams.Param1 = Ec.Parts.GetPartID(ObjPartsList(intK).PartNO)
                                objTextParams.Param2 = ObjPartsList(intK).Rectype.Trim
                                objTextParams.Param3 = ObjPartsList(intK).Seq

                                'delete the active part.
                                If Not Ec.Parts.DeleteData(1, objTextParams, HttpContext.Current.Session("UserName"), False, False) Then GoTo SkipImporting
                            Else
                                strCopyPartNo = Trim(HttpContext.Current.Session("gAN"))
                            End If
                        End If


                        'ok, now, if the same part is in the list to be imported, we need to update those part# as well 
                        '(ex: importing alternate and history plans for the same plan)

                        For intY = intK To UBound(ObjPartsList)     'For intY = 1 To UBound(ObjPartsList)
                            If ObjPartsList(intY).PartNO.Trim.ToLower = strOldPartNO.Trim.ToLower AndAlso ObjPartsList(intY).Rectype.Trim <> "0" Then
                                ObjPartsList(intY).PartNO = strCopyPartNo.Trim 'reset the part no
                            End If
                        Next intY

                        ObjPartsList(intK).PartNO = strCopyPartNo.Trim

                    Else
                        If Not blnPartExist Then
                            If Trim(StrNoActivePlans) <> "" Then StrNoActivePlans &= vbCrLf

                            StrNoActivePlans &= ObjPartsList(intK).PartNO.Trim

                            'active plan doesn't exist. can not import
                            GoTo SkipImporting
                        End If
                    End If



                    '   DisplayStatusForm("Processing Plan: " & ObjPartsList(intK).PartNO, intStatus)


                    'reset the arrsql and the docs list
                    ReDim ArrSQL(0) : ReDim objDocsList(0)

                    strSQLFile = strImportPath & ObjPartsList(intK).ElementDescX.Trim & "plandetails.txt"

                    ' DisplayStatusForm("Processing Plan: " & ObjPartsList(intK).PartNO & "(Import SQL Scripts)", intStatus)

                    ArrSQL = GetSQLScriptsFromImportFile(strSQLFile)

                    strDocsPath = strImportPath & ObjPartsList(intK).ElementDescX.Trim 'data path is stored in elementdescx (DATA imported from XML file)

                    '  DisplayStatusForm("Processing Plan: " & ObjPartsList(intK).PartNO & "(Import documents details)", intStatus)
                    objDocsList = Ec.Parts.ImportTextDocFilesFromXML(strDocsPath)


                    objCopy.ToPlantID = HttpContext.Current.Session("UserPlantID")
                    objCopy.ToPartNo = strCopyPartNo
                    objCopy.ToOPEngineer = HttpContext.Current.Session("UserName")
                    objCopy.ToRectype = ObjPartsList(intK).Rectype.Trim
                    objCopy.ToSeq = ObjPartsList(intK).Seq


                    ' DisplayStatusForm("Processing Plan: " & ObjPartsList(intK).PartNO & "(Importing Plan...)", intStatus)
                    If Ec.Parts.ImportPlanOperationData(objCopy, "importplan", strImportPath, ArrSQL, objDocsList, HttpContext.Current.Session("UserName"), blnCopyPart, blnCopyOP, 0, blnImportAllPlanTypes) Then
                        ObjPartsList(intK).OpDescX = "success"
                        intCount += 1
                    End If

SkipImporting:
                    intStatus = intStatus + intIteration
                Next

                If Trim(StrNoActivePlans) <> "" Then
                    strMsg = Ec.AppConfig.GetWrd(3987) & vbCrLf & vbCrLf & StrNoActivePlans & _
                            vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3027)
                    'MessageBox.Show(strMsg, Application.ProductName.Trim, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                    GoTo ExitThisFunction
                End If

                If intCount > 0 Then blnRtnValue = True

ExitThisFunction:
            Catch ex As Exception
                GeneralError("ImportPlanDetails: " & strErrLoc, ex, True)
            Finally
                ArrSQL = Nothing
                objDocsList = Nothing
                objCopy = Nothing
                objTextParams = Nothing

            End Try

            CheckForUserSessionErrors()
            If strURL.Trim <> "" Then
                HttpContext.Current.Response.Redirect(strURL, True)
            End If


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Function GetSQLScriptsFromImportFile(ByVal strFileName As String) As String()

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strFileName:({0})", strFileName), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim arrSQL() As String
            ReDim arrSQL(0)

            'Dim intTempFileNo As Long = 0, int_tempNo As Int32 = 1
            Dim strSQL As String = ""
            Dim oRead As System.IO.StreamReader, blnFileOpen As Boolean = False
            Dim strLine As String = "", intCount As Int32 = 0


            Try
                'if file doesn't exist, exit out
                If Not Ec.GeneralFunc.FileExists(strFileName) Then GoTo ExitThisFunction


                oRead = System.IO.File.OpenText(strFileName)
                blnFileOpen = True
                ' Loop over each line in file, While list is Not Nothing.
                Do While (Not strLine Is Nothing)

                    ' Read in the next line.
                    strLine = oRead.ReadLine
                    If Not strLine Is Nothing Then
                        intCount += 1
                        ReDim Preserve arrSQL(intCount)
                        arrSQL(intCount) = strLine.Trim
                    End If
                Loop
                oRead.Close()
                blnFileOpen = False


ExitThisFunction:
            Catch ex As Exception
                If blnFileOpen Then oRead.Close()

                'If intTempFileNo > 0 Then FileClose(intTempFileNo)
                GeneralError("GetSQLScriptsFromImportFile: ", ex, True)
            Finally
                oRead = Nothing
            End Try



#If TRACE Then
            Log.OPERATION(String.Format("Exit Count:({0})", UBound(arrSQL)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return arrSQL
        End Function


        Public Sub ClearChangeHistoryRecord()

#If TRACE Then
            Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim arrTemp() As String ', strTemp As String =  EaseCore.Extensions.Dates.CurrentDate()
            ReDim arrTemp(0)

            Dim objCH As EASEClass7.Parts.stChangeHistory
            objCH = HttpContext.Current.Session("gobjCH")

            objCH.Comment = arrTemp
            objCH.CommentSeq = 1
            objCH.Engineer = ""
            objCH.EntryDate = EaseCore.Extensions.Dates.GetCurrentDateInNumFormat
            objCH.EntrySeq = 1
            objCH.ID = 0
            objCH.OPNO = ""
            objCH.OpRev = 1
            objCH.PCN = ""
            objCH.Reason = arrTemp
            objCH.Rev = 1
            objCH.Engineer = ""
            If Trim(HttpContext.Current.Session("UserID")) <> Trim(EASEClass7.LangConstants.NA) Then  '"N/A"
                objCH.Engineer = Trim(HttpContext.Current.Session("UserName"))
            End If
            HttpContext.Current.Session("gobjCH") = objCH


#If TRACE Then
            Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function GetPlanTypeWord(ByVal intRectype As Int16) As String


#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_LOW(String.Format("Enter intRectype:({0})", intRectype), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'README: Any changes in the Plan rectype should be changed in  GetPlanTypeWord, and FillPlanTypeCombo
            Dim strDisplayText As String = ""
            Select Case intRectype
                Case 2 ' "History", "History Plan"
                    strDisplayText = Ec.AppConfig.GetWrd(3556)
                Case 1          '"Alternate", "Alternate Plan"
                    strDisplayText = Ec.AppConfig.GetWrd(3555)
                Case 3            '"Recyclebin"
                    strDisplayText = Ec.AppConfig.GetWrd(249)
                Case 0            '"Active"
                    strDisplayText = Ec.AppConfig.GetWrd(3554)
                Case 4
                    strDisplayText = "Line Balance Backup"
                Case Else
                    strDisplayText = "Unknown"
            End Select


#If TRACE Then
            Log.UI_CONTROL_LOW(String.Format("Exit ({0})", strDisplayText), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return strDisplayText
        End Function

        Public Function CheckBUCodeExist(ByVal strBUCode As String,
            Optional ByVal blnFormula As Boolean = False,
            Optional ByVal blnTable As Boolean = False,
            Optional ByVal blnActivity As Boolean = False) As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strBUCode:({0}) blnFormula:({1}) blnTable:({2}) blnActivity:({3})", strBUCode, blnFormula, blnTable, blnActivity), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim blnRtnValue As Boolean = False
            blnRtnValue = Ec.TimeDB.ExistBuCode(strBUCode, "mft")

            If Not blnRtnValue Then
                If blnFormula Then
                    blnRtnValue = Ec.TimeDB.ExistBuCode(strBUCode, "formula")
                End If
                If blnTable Then
                    blnRtnValue = Ec.TimeDB.ExistBuCode(strBUCode, "tblhdr")
                End If
                If blnActivity Then
                    blnRtnValue = Ec.TimeDB.ExistBuCode(strBUCode, "actlin")
                End If
            End If


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Sub SetChangeDatabase(ByVal strDBKey As String)

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strDBKey:({0})", strDBKey), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Try
                Dim strINIFile As String = GetUserSettingsFile()
                WriteCookie("DB-db", Trim(strDBKey))
                'Application.DoEvents()
            Catch ex As Exception
                GeneralError("SetChangeDatabase", ex)
            End Try

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Private Function GetUserSettingsFile() As String

#If TRACE Then
            Dim startTicks As Long = Log.FILE_DIR_IO_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim strFileName As String = ""
            strFileName = Ec.AppConfig.EASEConfigDirectory & "easesettings.xml"


#If TRACE Then
            Log.FILE_DIR_IO_LOW(String.Format("Exit ({0})", strFileName), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return strFileName
        End Function



        Private Function ArchiveActivePlan(ByVal intID As Int32, ByVal strRectype As String, ByVal intSeq As Int16, ByVal strPartNO As String, _
                                       ByVal strDestRecType As String, ByVal intDestSeq As Int16) As Boolean

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRectype:({1}) intSeq:({2}) strPartNO:({3}) strDestRecType:({4}) intDestSeq:({5})", intID, strRectype, intSeq, strPartNO, strDestRecType, intDestSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            '**RestoreFromRecycleBin - KM - 05/02/2012**

            'used when restoring the plan from recycle bin.
            Dim strErrorLocation As String = "start", blnRtnValue As Boolean = False
            Dim objCopy As New EASEClass7.Parts.stCopyParams
            Try

                '1. backup existing active plan
                'Copy the Active Part to history
                Ec.Parts.ClearCopyParamsObject(objCopy)

                'copy 
                objCopy.FromID = intID
                objCopy.FromRectype = strRectype
                objCopy.FromSeq = intSeq

                objCopy.ToID = intID
                objCopy.ToRectype = "2"
                objCopy.ToSeq = Ec.Parts.GetPartSeqMax(objCopy.ToID, objCopy.ToRectype)
                objCopy.ToOPEngineer = HttpContext.Current.Session("UserName")
                If Not Ec.Parts.CopyPartOperation("unreleasepart", objCopy, Ec.AppConfig.EASETempPath, HttpContext.Current.Session("UserName")) Then
                    'error occurred while copying the operation into Active. exit out
                    GoTo ExitThisFunction
                End If

                '2. delete the active plan
                DeletePlan(intID, strDestRecType, intDestSeq, strPartNO, True, False, True, True, False, True)

                blnRtnValue = True
ExitThisFunction:
            Catch ex As Exception
                GeneralError("ArchieveActivePlan: " & strErrorLocation.Trim, ex, True)
            Finally
                objCopy = Nothing
            End Try



#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue

        End Function
    End Module
End Namespace
