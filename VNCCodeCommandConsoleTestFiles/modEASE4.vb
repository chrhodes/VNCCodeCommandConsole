Imports System.Drawing
Imports Microsoft.VisualBasic
Imports EaseCore


Namespace EASEWebApp
    Public Module modEASE4
        Private Const BASE_ERRORNUMBER As Integer = EaseCore.ErrorNumbers.EASEWEBAPP_EASE4
        'Private Const BASE_ERRORNUMBER As Integer = ErrorNumbers.CLIENTEDITOR_ABOUT
        Private Const LOG_APPNAME As String = "EASEWEBAPP"
#Const TRACE = 1

        Public Function CheckForSharedOperationFlag(ByVal objOPSumm() As EASEClass7.Parts.stOPSummary) As Boolean

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim blnRtnValue As Boolean = False
            If Not EASEClass7.DBConfig.OperationSignoff Then GoTo ExitThisFunction 'for part signoff
            If Not EASEClass7.DBConfig.SharedOperation Then GoTo ExitThisFunction

            Dim intK As Int16 = 0, strSharedOps As String = ""
            Dim strMsg As String = ""

            For intK = 1 To UBound(objOPSumm)
                If objOPSumm(intK).Selected AndAlso objOPSumm(intK).SharedOPID > 0 Then
                    strSharedOps &= objOPSumm(intK).OPNO & vbCrLf
                End If
            Next

            If Trim(strSharedOps) <> "" Then
                blnRtnValue = True
                strMsg = Ec.AppConfig.GetWrd(3948) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(4259) & vbCrLf & vbCrLf & strSharedOps
                '  MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                GoTo ExitThisFunction
            End If

ExitThisFunction:


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Function GetRectypeColor(ByVal strRectype As String, Optional ByVal blnReports As Boolean = False) As Color

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED(String.Format("Enter strRectype:({0}) blnReports:({1})", strRectype, blnReports), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim color As Color

            Select Case Trim(strRectype)
                Case "1"
                    color = Color.FromArgb(128, 0, 0)
                Case "2"
                    color = Color.FromArgb(113, 113, 0)
                Case "3"
                    color = Color.FromArgb(0, 102, 0)
                Case Else
                    If blnReports Then
                        color = Color.FromArgb(10, 10, 10)
                    End If

                    'Case Else
                    '    Return Color.FromArgb(255, 204, 0)

            End Select

#If TRACE Then
            Log.UI_CONTROL_MED(String.Format("Exit ({0})", color), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
        End Function


        Public Sub RefreshPart()

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim intID As Int32 = 0, strRectype As String = "0", intSeq As Int16 = 0
            intID = HttpContext.Current.Session("gID")
            strRectype = HttpContext.Current.Session("gRectype")
            intSeq = HttpContext.Current.Session("gSeq ")
            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Try
                'now need to retotal part
                HttpContext.Current.Session("gOPSumm") = Ec.Parts.GetOpSummary(intID, strRectype, intSeq)

                objRH.NOP = UBound(HttpContext.Current.Session("gOPSumm"))
                'HttpContext.Current.Session("gRH") = objRH
                'update plan total and update database
                RecalculateTimeTotals(False, False, True, True, "")
                '   HttpContext.Current.Session("gRH") = objRH
            Catch ex As Exception
                GeneralError("RefreshPart", ex, True)
            End Try
            Call CheckForUserSessionErrors()

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub FillPlanTypeCombo(ByVal objCBO As DropDownList, Optional ByVal intPlanType As Int16 = 0, Optional ByVal blnIncludeAll As Boolean = False)

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED(String.Format("Enter intPlanType:({0}) blnIncludeAll:({1})", intPlanType, blnIncludeAll), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'README: Any changes in the Plan rectype should be changed in  GetPlanTypeWord, and FillPlanTypeCombo

            Dim strTemp As String = ""
            Try
                objCBO.Items.Clear()

                If ViewEASEWebApplication() Then
                    strTemp = Ec.AppConfig.GetWrd(3554)
                    AddComboItem(objCBO, strTemp, 0)
                    If DisplayHistoryParts() Then
                        strTemp = Ec.AppConfig.GetWrd(3556)
                        AddComboItem(objCBO, strTemp, 2)
                    End If

                Else            'default for all applications

                    If blnIncludeAll Then
                        strTemp = Ec.AppConfig.GetWrd(134)
                        AddComboItem(objCBO, strTemp, -1)
                        'objCBO.Items.Add(New ListItemData(strTemp, -1))
                    End If

                    strTemp = Ec.AppConfig.GetWrd(3554)
                    AddComboItem(objCBO, strTemp, 0)

                    strTemp = Ec.AppConfig.GetWrd(3555)
                    AddComboItem(objCBO, strTemp, 1)

                    strTemp = Ec.AppConfig.GetWrd(3556)
                    AddComboItem(objCBO, strTemp, 2)

                    strTemp = Ec.AppConfig.GetWrd(249)
                    AddComboItem(objCBO, strTemp, 3)
                    objCBO.SelectedIndex = intPlanType
                End If
            Catch ex As Exception
                GeneralError("FillPlanTypeCombo", ex, True)
            End Try
            Call CheckForUserSessionErrors()

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub FlagSelectedRecords(ByVal objUGrid As Infragistics.Web.UI.GridControls.WebDataGrid,
                      ByRef objPartsList() As EASEClass7.PartSearch.stPlanSearchResult, Optional ByRef intStartPos As Int16 = 0)

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim intK As Int16 = 0, intCount As Int16 = 0
            Dim intPos As Int16 = 0, strMsg As String = ""
            Dim blnAddRecord As Boolean = False
            Dim intSLNO As Int16 = 0
            intStartPos = 0
            ' Try
            If GetSelectedRowsCount(objUGrid) = 0 Then
                strMsg = Ec.AppConfig.GetWrd(3003) & " " & Ec.AppConfig.GetWrd(3004)
                'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                DisplayUserMessage(HttpContext.Current.Session("StatusMessage"), strMsg)
                GoTo ExitThisFunciton
            End If

            'clear all existing record selection flags
            ClearRecordSelection(objPartsList)
            For intK = 0 To objUGrid.Rows.Count - 1
                '**S**
                'If objUGrid.Rows(intK).Selected Then
                '    'intPos = EaseCore.Extensions.Strings.ToInt16(objUGrid.Rows(intK).Cells(11).Text)            'Array Position
                '    intSLNO = EaseCore.Extensions.Strings.ToInt16(objUGrid.Rows(intK).Cells(0).Text)

                '    If intPos > UBound(objPartsList) Then GoTo SkipThisrecord

                '    If intStartPos = 0 Then intStartPos = intSLNO

                '    intCount += 1
                '    objPartsList(intSLNO).RecordSelected = True
                'End If
                '**S**
SkipThisrecord:
            Next
ExitThisFunciton:
            ' Catch ex As Exception
            'GeneralError("FlagSelectedRecords", ex, True)
            ' End Try
            '  CheckForUserSessionErrors()

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub ClearRecordSelection(ByRef objPartsList() As EASEClass7.PartSearch.stPlanSearchResult)

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim intK As Int16 = 0
            Try
                For intK = 1 To UBound(objPartsList)
                    objPartsList(intK).RecordSelected = False
                Next

            Catch ex As Exception
                GeneralError("ClearElementRecordSelection", ex, True)
            End Try
            CheckForUserSessionErrors()


#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub ClearSearchVars()

#If TRACE Then
            Dim startTicks As Long = Log.CLEAR_INITIALIZE("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'clear all Search objects
            Dim objPartsList(0) As EASEClass7.PartSearch.stPlanSearchResult
            HttpContext.Current.Session("gobjPartsList") = objPartsList
            HttpContext.Current.Session("gstrSearchType") = ""
            HttpContext.Current.Session("gSearchPointer") = 0
            HttpContext.Current.Session("gobjSearchObject") = Nothing           'holds search criteria

#If TRACE Then
            Log.CLEAR_INITIALIZE("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

    End Module

End Namespace
