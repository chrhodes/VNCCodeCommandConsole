Imports Microsoft.VisualBasic
Imports EaseCore


Namespace EASEWebApp
    Public Module modEASE2
        Private Const BASE_ERRORNUMBER As Integer = EaseCore.ErrorNumbers.EASEWEBAPP_EASE2
        'Private Const BASE_ERRORNUMBER As Integer = ErrorNumbers.CLIENTEDITOR_ABOUT
        Private Const LOG_APPNAME As String = "EASEWEBAPP"
#Const TRACE = 1

        Public Function VerifyUserAccessRights(ByVal blnPartInMemory As Boolean, _
                                            ByVal blnDisplayMessage As Boolean, _
                                            Optional ByVal blnDeletePlanOp As Boolean = False, _
                                            Optional ByVal blnSignoff As Boolean = False, _
                                            Optional ByVal blnRename As Boolean = False, _
                                            Optional ByVal blnAdminAccessOnly As Boolean = False, _
                                            Optional ByVal blnEditSharePlan As Boolean = False, _
                                            Optional ByVal blnCheckInCheckOut As Boolean = False, _
                                            Optional ByVal objLabel As System.Web.UI.WebControls.Label = Nothing) As Boolean



#If TRACE Then
            Dim startTicks As Long = Log.SECURITY(String.Format("Enter blnPartInMemory:({0}) blnDisplayMessage:({1}) blnDeletePlanOp:({2}) blnSignoff:({3}) blnRename:({4}) blnAdminAccessOnly:({5}) blnEditSharePlan:({6}) blnCheckInCheckOut:({7})",
                                                                blnPartInMemory, blnDisplayMessage, blnDeletePlanOp, blnSignoff, blnRename, blnAdminAccessOnly, blnEditSharePlan, blnCheckInCheckOut), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'check user has access right to edit plan
            'returns false, if the user has no rights 

            'blnDeletePlanOp - Whether the user can delete a plan or operation
            'blnSignoff - Release or unrelease an operation
            'blnRename - Rename a plan
            'blnAdminAccessOnly - Only EASE Admin can do this process
            'blnEditSharePlan - Edit Shared plan operation and subheader details

            Dim blnRtnValue As Boolean = True, intAccessRights As Int16 = 0      'used for displaying error message purpose

            'if the login control is off, then users has full access to the plan
            If Not EASEClass7.DBConfig.LoginControl Then GoTo ExitThisFunction

            'defensive coding, if the user id is N/A then user has full access rights
            If Trim(HttpContext.Current.Session("UserID")) = Trim(EASEClass7.LangConstants.NA) Then GoTo ExitThisFunction '"N/A"


            intAccessRights = HttpContext.Current.Session("gPartAccessRights")

            If Not blnPartInMemory Then  'OK, No Part is in memory, handle process such as, rename plan, create plan, delete plan and so on
                'EASE Access Rights: 1-All, 4-Process Plan RWD, 5-Process Plan RW, 6- Process Plan Read,
                '                    7-Read only Sign-off
                Select Case HttpContext.Current.Session("UserAccessRights")
                    Case 1                                              'All
                    Case 4                                              'Process Plan RWD, Signoff
                        If blnAdminAccessOnly Then blnRtnValue = False
                    Case 5                                              'Plan RW
                        If blnAdminAccessOnly Then blnRtnValue = False
                        If blnDeletePlanOp Then blnRtnValue = False
                        If blnSignoff Then blnRtnValue = False 'no access to signoff
                        If blnRename Then blnRtnValue = False 'no access to rename the plan
                    Case 6                                              'Plan-Read Only
                        blnRtnValue = False
                    Case 7                                              'Plan-Read Only Signoff
                        blnRtnValue = False
                    Case 9

                    Case Else
                        blnRtnValue = False
                End Select

            Else                        'Plan is in memory, check user access rights with plan status, brilliant huh!

                'gPartAccessRights: 1-Full Access, 3-Shared Plan,4-Alternate Plan,5-History Plan,6-Recycle Bin,7-LB Backup
                '                   8-Released Plan, 9-Released Op, 11-User Restriction Fields (Client Editor),12-Read only Plan
                '                   13-Read only Operation, 14-Plan is Locked, 15 - Master operation, 16-Shared Operation (linked to Master op)

                Select Case HttpContext.Current.Session("gPartAccessRights")
                    Case 1          'Full Access, either no login, or the plan is unreleased
                        Select Case HttpContext.Current.Session("UserAccessRights")
                            Case 1                                              'All
                            Case 4                                              'Process Plan RWD, Signoff
                            Case 5                                              'Plan RW
                                If blnSignoff Then blnRtnValue = False 'no access to signoff
                            Case 6                                              'Plan-Read Only
                                blnRtnValue = False
                            Case 7                                              'Plan-Read Only Signoff
                                If Not blnSignoff Then blnRtnValue = False 'Only signoff, no access to anything else
                            Case Else
                                blnRtnValue = False
                        End Select
                    Case 3          'Shared Plan
                        Select Case HttpContext.Current.Session("UserAccessRights")
                            Case 1                                              'All
                            Case 4                                              'Process Plan RWD, Signoff
                            Case 5                                              'Plan RW
                                If blnSignoff Then blnRtnValue = False 'no access to signoff
                            Case 6                                              'Plan-Read Only
                                blnRtnValue = False
                            Case 7                                              'Plan-Read Only Signoff
                                If Not blnSignoff Then blnRtnValue = False 'Only signoff, no access to anything else
                            Case Else
                                blnRtnValue = False
                        End Select
                    Case 4          'Alternate Plan

                        'no signoff for alternate plan. not sure, whether we should allow signoff, if we do, how do we track revision, history and so on.
                        Select Case HttpContext.Current.Session("UserAccessRights")
                            Case 1                                              'All
                            Case 4                                              'Process Plan RWD, Signoff
                            Case 5                                              'Plan RW
                            Case 6                                              'Plan-Read Only
                                blnRtnValue = False
                            Case 7                                              'Plan-Read Only Signoff
                            Case Else
                                blnRtnValue = False
                        End Select
                        If blnSignoff Then blnRtnValue = False 'no access to signoff

                    Case 5, 6, 7                'History, Backup, RecycleBin Plan
                        blnRtnValue = False     'users can not edit 

                    Case 8, 9                      '8-released plan, 9-Relesed Op

                        blnRtnValue = False

                        'allow the users to unrelease the plan
                        Select Case HttpContext.Current.Session("UserAccessRights")
                            Case 1                                              'All
                                If blnSignoff Then blnRtnValue = True
                                If HttpContext.Current.Session("gTypoChanges") Then blnRtnValue = True
                                If blnCheckInCheckOut Then blnRtnValue = True
                            Case 4                                              'Process Plan RWD, Signoff
                                If blnSignoff Then blnRtnValue = True
                                If HttpContext.Current.Session("gTypoChanges") Then blnRtnValue = True
                                If blnCheckInCheckOut Then blnRtnValue = True
                        End Select
                    Case 11                 'User Restriction Field
                        blnRtnValue = False
                    Case 12
                        blnRtnValue = False
                    Case 13
                        blnRtnValue = False
                    Case 14         'lock plan
                        blnRtnValue = False
                    Case Else
                        blnRtnValue = False
                End Select
            End If
            If blnRtnValue = False AndAlso blnDisplayMessage Then
                DisplayAccessRightsMessage(objLabel, intAccessRights)
            End If

ExitThisFunction:


#If TRACE Then
            Log.SECURITY(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue

        End Function

        Public Sub DisplayAccessRightsMessage(ByVal objLabel As System.Web.UI.WebControls.Label, Optional ByVal intAccessRights As Int16 = 0)

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intAccessRights:({0})", intAccessRights), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim objRH As EASEClass7.Parts.stRH
            Dim strMsg As String = ""
            objRH = HttpContext.Current.Session("gRH")

            Dim strAccess As String = ""
            strAccess = HttpContext.Current.Session("UserAccessRights")
            'gPartAccessRights: 1-Full Access, 3-Shared Plan,4-Alternate Plan,5-History Plan,6-Recycle Bin,7-LB Backup
            '                   8-Released Plan, 9-Released Op, 11-User Restriction Fields (Client Editor),12-Read only Plan
            '                   13-Read only Operation, 14-Plan is Locked, 15 - Master operation, 16-Shared Operation (linked to Master op)

            If intAccessRights = 0 Then
                intAccessRights = HttpContext.Current.Session("gPartAccessRights")
            End If

            Select Case intAccessRights
                Case 2, 10 ' Readonly Plan
                    'You can not edit this record, your access rights are
                    strMsg = Ec.AppConfig.GetWrd(3971) & vbCrLf & vbCrLf &
                    Ec.AppConfig.GetWrd(3182) & " " & GetEASEAccessRightsValue(strAccess)
                Case 3      'Share Plan
                    strMsg = Ec.AppConfig.GetWrd(3181) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 4      'Alternate Plan messages
                    strMsg = Ec.AppConfig.GetWrd(3183) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 5      ' History Plan
                    strMsg = Ec.AppConfig.GetWrd(3184) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 6      'Recycle Bin
                    strMsg = Ec.AppConfig.GetWrd(3185) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 7      'LB Backup
                    strMsg = Ec.AppConfig.GetWrd(3186) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 8      'Released part
                    strMsg = Ec.AppConfig.GetWrd(3187) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3190)
                Case 9      'Released Operation
                    strMsg = Ec.AppConfig.GetWrd(3188) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3191)
                Case 12
                    strMsg = Ec.AppConfig.GetWrd(3196)
                Case 13
                    strMsg = Ec.AppConfig.GetWrd(3198)
                Case 14
                    strMsg = Ec.AppConfig.GetWrd(3201) & " " & objRH.Engineer.Trim & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3189) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3203)
                Case 16         'shared operation
                    strMsg = Ec.AppConfig.GetWrd(3204) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3205)
                Case Else
                    strMsg = Ec.AppConfig.GetWrd(3122) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3182) &
                    " " & GetEASEAccessRightsValue(strAccess) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(4009)
            End Select

            If strMsg.Trim <> "" Then
                'MessageBox.Show(strMsg, Application.ProductName.Trim, MessageBoxButtons.OK, MessageBoxIcon.Information)
                DisplayUserMessage(objLabel, strMsg)
            End If

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub


        Public Sub WriteRecentlyViewedParts(ByVal strPartNO As String, ByVal strRectype As String, ByVal intSeq As Int16, Optional ByVal strSharedOP As String = "",
                                          Optional ByVal intPlantID As Int32 = 0)

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strPartNO:({0}) strRectype:({1}) intSeq:({2}) strSharedOP:({3}) intPlantID:({4})",
                                                                 strPartNO, strRectype, intSeq, strSharedOP, intPlantID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'recentparts
            Dim objRP() As EASEClass7.Parts.stRecentParts
            ReDim objRP(0)
            Dim intK As Int16 = 0
            Dim strPartKey As String = "", strRecTypeKey As String = "", strSeqKey As String = ""
            Dim strMisc1Key As String = "", strPlantKey As String = ""
            Dim strTemp As String = "", intCount As Int16 = 0
            Dim intRecentPartCount As Int16 = 20
            Dim intUserPlantID As Int16 = HttpContext.Current.Session("UserPlantID")

            Dim strRecentPartsList As String = "", strJoin As String = "ZyZ"

            Try
                If intPlantID = 0 Then
                    intPlantID = intUserPlantID
                End If
                objRP = GetRecentlyViewedParts()
                ReDim Preserve objRP(UBound(objRP) + 1)     'add one record

                For intK = UBound(objRP) To 1 Step -1
                    objRP(intK) = objRP(intK - 1)
                Next


                objRP(1).PartNo = strPartNO
                objRP(1).Rectype = strRectype
                objRP(1).Seq = intSeq
                objRP(1).Misc1 = strSharedOP.Trim
                objRP(1).PlantID = intPlantID


                'make sure the SAME partno doesn't exist in the array, if so reset the partno to ''
                For intK = 2 To UBound(objRP)
                    If Trim(objRP(intK).PartNo) = Trim(strPartNO) And Trim(objRP(intK).Rectype) = Trim(strRectype) _
                            And objRP(intK).Seq = intSeq Then
                        objRP(intK).PartNo = ""
                    End If
                Next


                For intK = 1 To UBound(objRP)
                    If objRP(intK).PartNo = "" Then GoTo SkipThisRecord

                    strPartKey = "PARTNO-" & intK

                    intCount += 1

                    strRecentPartsList = Trim(objRP(intK).PartNo) & strJoin & Trim(objRP(intK).Rectype) & strJoin & objRP(intK).Seq & strJoin & objRP(intK).Misc1 & strJoin & objRP(intK).PlantID
                    WriteCookie("RecentParts-" & strPartKey, strRecentPartsList)

                    If intK > intRecentPartCount Then Exit For 'hold upto 7 parts
SkipThisRecord:
                Next

            Catch ex As Exception
                GeneralError("WriteRecentlyViewedParts", ex, True)
            Finally
            End Try
            CheckForUserSessionErrors()

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function GetRecentlyViewedParts() As EASEClass7.Parts.stRecentParts()

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim objRP() As EASEClass7.Parts.stRecentParts
            ReDim objRP(0)

            'Dim strINIFile As String = GetUserSettingsFile() '"ease.ini"
            Dim intK As Int16 = 0
            Dim strPartKey As String = "", strRecTypeKey As String = "", strSeqKey As String = ""

            Dim strPartNo As String = "", strRectype As String = ""
            Dim intSeq As Int16 = 0, intCount As Int16 = 0, strMisc1 As String = ""
            Dim strTemp As String = "", strMisc1Key As String = ""
            Dim strPlantKey As String = "", intPlantID As Int16 = 0
            Dim intRecentPartCount As Int16 = 5
            Dim intUserPlantID As Int16 = HttpContext.Current.Session("UserPlantID")

            Dim strRecentPartsList As String = "", strJoin As String = "ZyZ"
            Dim ArrTemp(0) As String

            Try
                For intK = 1 To intRecentPartCount
                    strPartKey = "PARTNO-" & intK

                    strRecentPartsList = ReadCookie("RecentParts-" & strPartKey.Trim)

                    If strRecentPartsList.Trim = "" Then
                        GoTo SkipThisRecord
                    End If

                    ArrTemp = Split(strRecentPartsList, strJoin)

                    If UBound(ArrTemp) < 4 Then GoTo SkipThisRecord
                    If ArrTemp(0).Trim = "" Then GoTo SkipThisRecord 'nothing found.

                    strPartNo = ArrTemp(0).Trim
                    If strPartNo = "" Then
                        GoTo SkipThisRecord
                    End If
                    strRectype = ArrTemp(1).Trim
                    intSeq = EaseCore.Extensions.Strings.ToInt16(ArrTemp(2).Trim)
                    strMisc1 = ArrTemp(3).Trim
                    intPlantID = EaseCore.Extensions.Strings.ToInt16(ArrTemp(4).Trim)

                    If intUserPlantID <> intPlantID Then GoTo SkipThisRecord 'skip the plant if the plant id doesn't match

                    If CheckForDuplicateRecentlyViewedPart(objRP, strPartNo, strRectype, intSeq) Then GoTo SkipThisRecord


                    intCount += 1
                    ReDim Preserve objRP(intCount)

                    objRP(intCount).PartNo = strPartNo
                    objRP(intCount).Rectype = strRectype
                    objRP(intCount).Seq = intSeq
                    objRP(intCount).Misc1 = strMisc1
                    objRP(intCount).PlantID = intPlantID
SkipThisRecord:
                Next

            Catch ex As Exception
                ex = Nothing
            End Try
            CheckForUserSessionErrors()


#If TRACE Then
            Log.OPERATION(String.Format("Exit Count:({0})", UBound(objRP)), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return objRP
        End Function

        Private Function CheckForDuplicateRecentlyViewedPart(ByVal objRP() As EASEClass7.Parts.stRecentParts, _
           ByVal strPartNO As String, ByVal strRectype As String, ByVal intSeq As Int16) As Boolean

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strPartNO:({0}) strRectype:({1}) intSeq:({2})", strPartNO, strRectype, intSeq), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim intK As Int16 = 0, blnRtnValue As Boolean = False
            For intK = 1 To UBound(objRP)
                If Trim(objRP(intK).PartNo) = Trim(strPartNO) And Trim(objRP(intK).Rectype) = Trim(strRectype) _
                            And objRP(intK).Seq = intSeq Then
                    blnRtnValue = True
                End If
            Next


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Sub RecallPart(ByVal strPartNO As String, Optional ByVal strRectype As String = "", _
                    Optional ByVal intSeq As Int16 = -1, Optional ByRef strErrorMsg As String = "", _
                    Optional ByRef blnRedirectToPasswordPage As Boolean = False, _
                    Optional ByVal blnSkipPartInUseCheck As Boolean = False)  ' modEASE2


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strPartNO:({0}) strRectype:({1}) intSeq:({2}) blnSkipPartInUseCheck:({3})",
                                                                 strPartNO, strRectype, intSeq, blnSkipPartInUseCheck), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            strErrorMsg = "" : blnRedirectToPasswordPage = False

            'If the PartNo variable is passed, then use the part no to get the gID, gRectype and gSeq variable.
            'else, use the values in the gID, gRectype and gSeq variable to retrieve the Plan
            Dim objRH As EASEClass7.Parts.stRH
            Dim objOPSumm(0) As EASEClass7.Parts.stOPSummary
            Dim clsUpdate7 As New EASEClass7.Update7
            Dim strMsg As String = "", strPartUserID As String = "", intK As Int16 = 0
            Dim strTemp As String = "", blnOpenReadOnly As Boolean = False
            Dim blnResetOpEngineers As Boolean = False, blnUpdateMasterPartCheckOutEngieer As Boolean = False
            Dim intMasterID As Int32 = 0, objMasterOps(0) As EASEClass7.Parts.stOPSummary
            Dim blnTakeOver As Boolean = False, strMisc1 As String = ""
            Dim strName As String = HttpContext.Current.Session("UserName")
            Dim strURL As String = ""
            Dim blnUseSHLevel As Boolean = False
            Try

                If HttpContext.Current.Session("gAN") = "SH" Then
                    GoTo SKiptoSH
                Else
                End If

                blnOpenReadOnly = HttpContext.Current.Session("gReadOnlyPlan")

                'check part id, part no exist
                If Trim(strPartNO) <> "" Then

                    'if partno is passed, then get the ID for the Partno 
                    HttpContext.Current.Session("gID") = Ec.Parts.GetPartID(strPartNO)

                    If HttpContext.Current.Session("gID") = 0 Then     'part no doesn't exist,
NoPartMessage:
                        KillPart()      'kill the part variables

                        strMsg = Ec.AppConfig.GetWrd(3902)
                        If Trim(strPartNO) <> "" Then
                            strMsg &= "(" & GetPartNumberWord() & " " & strPartNO & ")"
                        End If
                        strErrorMsg = strMsg.Trim
                        GoTo ExitThisSub
                    End If


                    HttpContext.Current.Session("gRectype") = "0" : HttpContext.Current.Session("gSeq") = 0
                    If Trim(strRectype) <> "" Then
                        HttpContext.Current.Session("gRectype") = Trim(strRectype)
                    End If
                    If intSeq > 0 Then
                        HttpContext.Current.Session("gSeq") = intSeq
                    End If
                Else
                    strTemp = Ec.Parts.GetPartNO(HttpContext.Current.Session("gID"))
                    If Trim(strTemp) = "" Then
                        GoTo NoPartMessage
                    End If
                End If

                'user has full access rights
                '1-full access rights to access part, 2-open part as readonly
                HttpContext.Current.Session("gPartAccessRights") = 1
           
                HttpContext.Current.Session("gReadOnlyPlan") = False  'readonly plan
                'ok part exist, check part is not in use

                strPartUserID = HttpContext.Current.Session("UserID")

                'ok if User access rights for this user, is readonly in client editor,
                'why we need to check part is in use or not, just open the part as read-only, right!.
                If HttpContext.Current.Session("UserAccessRights") = 6 Then GoTo SkipCheckingPartsInUse
                If blnSkipPartInUseCheck Then GoTo SkipCheckingPartsInUse

                If Ec.Parts.CheckPartInUse(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), strPartUserID) Then
                    If strPartUserID.Trim = "" Then GoTo SkipCheckingPartsInUse
                    If strPartUserID.Trim.ToLower = Trim(LCase(HttpContext.Current.Session("UserName"))) Then
                        'same user, skip it
                        GoTo SkipCheckingPartsInUse
                    End If

                    strTemp = Ec.Parts.GetPartNO(HttpContext.Current.Session("gID"))
                    'ok part is is use by another engineer, strUserID
                    strMsg = GetPartNumberWord() & ": '" & Trim(strTemp) & "' " & Ec.AppConfig.GetWrd(3903) & ""
                    If Not Ec.GeneralFunc.CompareString(HttpContext.Current.Session("UserID"), strPartUserID) Then
                        strMsg &= " by '" & Trim(strPartUserID) & "'" & "<br><br>"
                    End If
                    strMsg &= vbCrLf & Ec.AppConfig.GetWrd(3904) & _
                                                " " & Ec.AppConfig.GetWrd(4500) '3905

                    HttpContext.Current.Session("gAN") = strMsg
                    blnRedirectToPasswordPage = True
                    GoTo ExitThisSub

                End If

SkipCheckingPartsInUse:

                Ec.Parts.ClearRHObject(objRH)     'clear rh object
                ReDim objOPSumm(0)                'clear the list of operations
                KillOP()                        'clear the operation level object

                If HttpContext.Current.Session("gID") = 0 Then GoTo ExitThisSub 'if no id is set, just exit out, hi hi!


                If Not Ec.Parts.PartRecordExist(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq")) Then

                    If Trim(strPartNO) <> "" Then
                        strMsg = Ec.AppConfig.GetWrd(24) & " '" & strPartNO & "' "
                    Else
                        strMsg = Ec.AppConfig.GetWrd(24)
                    End If
                    strMsg &= Ec.AppConfig.GetWrd(3125) & vbCrLf & vbCrLf & "Check Recycle-bin."
                    ' MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    strErrorMsg = strMsg.Trim
                    GoTo ExitThisSub

                    HttpContext.Current.Session("gblnUpgradePart") = False
                    GoTo ExitThisSub 'if part doesn't exist then exit out
                End If

                If HttpContext.Current.Session("gblnUpgradePart") Then GoTo UpdatePart

                If Not Ec.Parts.PartInVersion7Tables(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq")) Then
UpdatePart:
                    '**SHLevelFlag-In-Plan-Header**'
                    HttpContext.Current.Session("gAN") = "" : HttpContext.Current.Session("gLIN") = ""
                    'strTemp = "The selected plan is in EASE Version 6.2 format. The plan must be converted into Version 7 format and the process may take couple of minutes. Please stand-by."
                    If EASEClass7.DBConfig.UseSubHeaderLevel = True Then '**NonSHLevel Changes**
                        strTemp = Ec.AppConfig.GetWrd(4480) & " " & Ec.AppConfig.GetWrd(4481) & " " & Ec.AppConfig.GetWrd(4482) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(4494) & " " & Ec.AppConfig.GetWrd(4495)
                    Else
                        strTemp = Ec.AppConfig.GetWrd(4480) & " " & Ec.AppConfig.GetWrd(4481) & " " & Ec.AppConfig.GetWrd(4482) '& vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(4494) & " " & Ec.AppConfig.GetWrd(4495)
                    End If
                    HttpContext.Current.Session("gAN") = strTemp.Trim
                    HttpContext.Current.Session("gLIN") = strPartNO
                    strURL = "frmMessageBox.aspx"
                    '  ShowDialogForm(New frmMessageBox)
                    GoTo ExitThisSub

                    If Trim(HttpContext.Current.Session("gAN")) = "*" Then
                        HttpContext.Current.Session("gblnUpgradePart") = False
                        GoTo ExitThisSub 'if part doesn't exist then exit out
                    End If
SKiptoSH:
                    If HttpContext.Current.Session("gAN") = "SH" Then
                        blnUseSHLevel = True
                    End If
                    HttpContext.Current.Session("gAN") = "" : HttpContext.Current.Session("gLIN") = ""
                    strTemp = Trim(strPartNO) & " (Rectype: " & Trim(HttpContext.Current.Session("gRectype")) & ", Seq: " & HttpContext.Current.Session("gSeq") & ")"

                    'TODO: The subheader flag is turned off for NAVY, need to cover it, when implementing the Client changes into the Web apps, see '**SHLevelFlag-In-Plan-Header**' in client apps

                    If Not clsUpdate7.UpgradePartToVersion7Structure(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), Ec.AppConfig.EASETempPath, strName, strTemp, blnUseSHLevel) Then
                        strMsg = "Unable to upgrade the part data to Version 7."
                        'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        strErrorMsg = strMsg.Trim

                        'DisplayUserMessage(Label, strMsg)
                        HttpContext.Current.Session("gblnUpgradePart") = False
                        GoTo ExitThisSub 'if part doesn't exist then exit out
                    End If
                    HttpContext.Current.Session("gblnUpgradePart") = False
                End If

                objRH = Ec.Parts.GetRH(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"))

                If Trim(objRH.PartNo) = "" Then
                    strMsg = Ec.AppConfig.GetWrd(3902)
                    If Trim(strPartNO) <> "" Then
                        strMsg &= "(" & Ec.AppConfig.GetWrd(24) & " " & strPartNO & ")"
                    End If
                    strErrorMsg = strMsg.Trim
                    'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                    'DisplayUserMessage(lblUserMsg, strMsg)
                    GoTo ExitThisSub 'if part doesn't exist then exit out
                End If

                HttpContext.Current.Session("gRH") = objRH

                SetPlanAccessRights()       'Set the user access rights for this plan
              
                If blnOpenReadOnly Then     'plan is in use by another user
                    GoTo SkipTakeOver
                Else
                    ' HttpContext.Current.Session("gReadOnlyPlan") = False
                    If blnSkipPartInUseCheck Then GoTo SkipTakeOver

                    'Take over plan
                    If (Ec.AppConfig.CAPP Or Ec.AppConfig.CAPPOnly) AndAlso HttpContext.Current.Session("gPartAccessRights") = 1 Then          'take over part, user have full access right, hi hi
                        If Not EASEClass7.DBConfig.OperationSignoff Then    'For part signoff

                            'also confirm, the user has access rights to edit the plan
                            If Not objRH.Released Then
PlanTakeOver:
                                If LCase(objRH.Engineer.Trim) <> LCase(Trim(strName)) Then
                                    HttpContext.Current.Session("gParam") = "TakeOverPlan"
                                    If ViewEASEWebApplication() Then GoTo Skip 'skip the process plan  condition
                                    If HttpContext.Current.Session("gTemp5") = "PlanBalance" Then '  HttpContext.Current.Session("UserLog")
                                        HttpContext.Current.Session("gReadOnlyPlan") = True
                                        GoTo SkipTakeOver
                                    End If
Skip:
                                    strURL = "DialogBox.aspx"
                                    GoTo ExitThisSub
                                    '  strMsg = Ec.AppConfig.GetWrd(3194) & " " & objRH.Engineer.Trim & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3195)

                                    '!@#$ have to do
                                    ' If MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                    blnTakeOver = True
                                    'update the routeheader engineer
                                    objRH.Engineer = strName
                                    Ec.Parts.UpdatePlanEngineer(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), HttpContext.Current.Session("UserName"))
                                    ' Else
                                    HttpContext.Current.Session("gReadOnlyPlan") = True
                                    ' End If

                                End If

                            End If
                        Else  'if the part is operation signoff
                            If Ec.Parts.CheckedOutPart(objRH.PartNo) Then
                                blnResetOpEngineers = True
                                HttpContext.Current.Session("gblnTutor") = True
                                GoTo PlanTakeOver
                            End If
                        End If
                    End If
                End If
SkipTakeOver:

                If HttpContext.Current.Session("gReadOnlyPlan") Then
                    'reset the access rights, user is opening the plan as read only
                    SetPlanAccessRights()       'Set the user access rights for this plan
                End If

                If blnTakeOver Then     'take over plan (part signoff), shared plan (op signoff)
                    If blnResetOpEngineers Then         'update the operation engineer for the checkedout plan
                        Ec.Parts.UpdateOperationEngineer(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), "", HttpContext.Current.Session("UserName"))

                        ' we can't update the check out engieer here, because we don't know the operation numbers for the
                        'plan. so load the operation into memory and then update the checkout engineer in the master part. Brilliant, huh!.
                        blnUpdateMasterPartCheckOutEngieer = True
                    End If
                End If

                'partno exist, so load operations for this part
                objOPSumm = Ec.Parts.GetOpSummary(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"))
                HttpContext.Current.Session("gOPSumm") = objOPSumm

                If blnUpdateMasterPartCheckOutEngieer Then
                    'update the checkout engineer in the master part
                    intMasterID = Ec.Parts.GetPartID(objRH.MasterPart)

                    If intMasterID > 0 Then         'defensive coding
                        objMasterOps = HttpContext.Current.Session("gOPSumm")

                        For intK = 1 To UBound(objMasterOps)
                            objMasterOps(intK).ID = intMasterID
                            objMasterOps(intK).Selected = True
                        Next
                        Ec.Parts.UpdateCheckOutEngineer(objMasterOps, strName)
                    End If
                End If

                'Recalculate the totals
                RecalculateTimeTotals(False, False, True, False, "")

                strMisc1 = ""
                If EASEClass7.DBConfig.SharedOperation AndAlso objRH.Flag2 = 2 Then
                    'shared operation. display "sharedop-opno" instead of part#. Brilliant, huh!
                    strMisc1 = objRH.Desc.Trim
                    If Ec.Parts.CheckedOutPart(objRH.PartNo) Then
                        strMisc1 = strMisc1 & " " & Ec.Parts.GetCheckedOutCopyNumber(objRH.PartNo)
                    End If
                End If

                'write to recently viewed parts list
                WriteRecentlyViewedParts(objRH.PartNo, HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), strMisc1)

                '*** README: the below VerifyUserAccessRights function is used in three places (Recall Part, frmOPSumm:btnClose_Click, FrmMain: ClearNetworkLockForPlanInMemory.
                'Any Change must be updated in ALL THREE PLACES.
                If VerifyUserAccessRights(True, False, , , , , True, , ) Then     'if the user access rights are good, and the plan is writable then set the network lock
                    'set the network lock
                    Ec.Parts.SetNetworkLock(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), True, HttpContext.Current.Session("UserID"))
                End If

ExitThisSub:
            Catch ex As Exception
                GeneralError("RecallPart", ex, True)
            Finally
                clsUpdate7 = Nothing
                objMasterOps = Nothing
            End Try
            Call CheckForUserSessionErrors()
            If strURL.Trim <> "" Then
                System.Web.HttpContext.Current.Response.Redirect(strURL, True)
                ' Response.Redirect(strURL)
            End If


#If TRACE Then
            Log.OPERATION(String.Format("Exit strErrorMsg:({0}) blnRedirectToPasswordPage:({1})", strErrorMsg, blnRedirectToPasswordPage), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub KillPart()

#If TRACE Then
            Dim startTicks As Long = Log.APPLICATION_SESSION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim objOPSumm(0) As EASEClass7.Parts.stOPSummary
            Dim objRH As New EASEClass7.Parts.stRH

            Try
                HttpContext.Current.Session("gOPSumm") = objOPSumm
                HttpContext.Current.Session("gID") = 0
                HttpContext.Current.Session("gRectype") = "0"
                HttpContext.Current.Session("gSeq") = 0

                '0-All access rights, 2-Readonly user.
                HttpContext.Current.Session("gPartAccessRights") = 1  'user's access right for this part
                HttpContext.Current.Session("gReadOnlyPlan") = False

                Ec.Parts.ClearRHObject(objRH)     'clear part in memory
                HttpContext.Current.Session("gRH") = objRH

                HttpContext.Current.Session("gTypoChanges") = False

                HttpContext.Current.Session("ResetRectype") = False             'used in ViewEASE
                KillOP()
            Catch ex As Exception
                GeneralError("KillPart", ex, True)
            Finally
                objOPSumm = Nothing
                objRH = Nothing
            End Try
            Call CheckForUserSessionErrors()


#If TRACE Then
            Log.APPLICATION_SESSION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub KillOP()

#If TRACE Then
            Dim startTicks As Long = Log.APPLICATION_SESSION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim objOPH As New EASEClass7.Parts.stOPH
            Dim objSH(0) As EASEClass7.SubHeaders.stSubHeaders
            Try
                HttpContext.Current.Session("gReadOnlyOperation") = False

                Ec.Parts.ClearOPHObject(objOPH)   'clear the operation details
                HttpContext.Current.Session("gOPH") = objOPH


                'READ ME. the gSHSumm is used to hold the shsummary structure in Process Plan and ViewEASE. But for ViewEASE MES system, the object holds different structure
                HttpContext.Current.Session("gSHSumm") = objSH

                KillSubHeader()
            Catch ex As Exception
                GeneralError("KillOP", ex, True)
            End Try
            Call CheckForUserSessionErrors()

#If TRACE Then
            Log.APPLICATION_SESSION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub KillSubHeader()

#If TRACE Then
            Dim startTicks As Long = Log.APPLICATION_SESSION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            HttpContext.Current.Session("gSHID") = 0
            HttpContext.Current.Session("gSHDesc") = ""         'THIS VARIABLE IS USED ONLY IN VIEWEASE (NOT IN PROCESS PLAN OR ANY OTHER APPS)

            Dim objDocs(0) As EASEClass7.SubHeaders.stDocs
            HttpContext.Current.Session("gobjDocs") = objDocs           'This variable is used only in VIEWEASE (NOT IN PROCESS PLAN OR ANY OTHER APPS)

#If TRACE Then
            Log.APPLICATION_SESSION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function GetPartOpSearchClipboardFileName() As String

#If TRACE Then
            Dim startTicks As Long = Log.FILE_DIR_IO_LOW("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'Dim strFile As String = Ec.AppConfig.EASETempPath
            'If gUser.UserID <> "" AndAlso UCase(Trim(gUser.UserID)) <> UCase(Trim(EASEClass7.LangConstants.NA)) Then strFileName = LCase(gUser.UserID) & "-" & strFileName

            Dim strResult As String = AddUserIDToFileName(Ec.AppConfig.EASETempPath, "partsearch.xml")

#If TRACE Then
            Log.FILE_DIR_IO_LOW(String.Format("Exit ({0})", strResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return strResult
        End Function

        Public Function AddUserIDToFileName(ByVal strPath As String, ByVal strFileName As String) As String

#If TRACE Then
            Dim startTicks As Long = Log.FILE_DIR_IO_LOW(String.Format("Enter strPath:({0}) strFileName:({1})", strPath, strFileName), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim strTemp As String = Trim(strFileName)
            If HttpContext.Current.Session("UserID") <> "" AndAlso UCase(Trim(HttpContext.Current.Session("UserID"))) <> UCase(Trim(EASEClass7.LangConstants.NA)) Then strTemp = Trim(LCase(HttpContext.Current.Session("UserID"))) & "-" & Trim(strFileName)

            Dim strResult As String = strPath & strTemp

#If TRACE Then
            Log.FILE_DIR_IO_LOW(String.Format("Exit ({0})", strResult), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return strResult
        End Function

        Public Sub SaveSelectedPartOP(ByVal blnX As Boolean)

#If TRACE Then
            Dim startTicks As Long = Log.Trace9(String.Format("Enter blnX:({0})", blnX), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If



#If TRACE
            Log.Trace9("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
    #End If

        End Sub


        Public Sub DisplayPartFlagText(ByVal objLabel As Label, ByVal strRectype As String, _
                               Optional ByVal blnDisplayOnlyPartType As Boolean = False, _
                               Optional ByVal blnPlanLocked As Boolean = False, _
                               Optional ByVal blnMasterSharedOP As Boolean = False)


#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED(String.Format("Enter strRectype:({0}) blnDisplayOnlyPartType:({1}) blnPlanLocked:({2}) blnMasterSharedOP:({3})",
                                                          strRectype, blnDisplayOnlyPartType, blnPlanLocked, blnMasterSharedOP), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'Dim objFont As Font = New Font(GetDefaultFontName(), 8, FontStyle.Bold)
            Dim strDisplayText As String = GetPlanTypeWord(strRectype)
            Dim objOPH As EASEClass7.Parts.stOPH
            objOPH = HttpContext.Current.Session("gOPH")
            Try
                objLabel.Visible = False
                objLabel.Text = ""
                'objLabel.Font = objFont '
                objLabel.ForeColor = GetRectypeColor(1)     'always display on one color

                Select Case strRectype.Trim
                    Case "0", "1"
                        If strRectype.Trim = "1" Then
                            objLabel.Text = GetPlanTypeWord(strRectype)
                        End If
                        'gPartAccessRights: 1-Full Access, 3-Shared Plan,4-Alternate Plan,5-History Plan,6-Recycle Bin,7-LB Backup
                        '                   8-Released Plan, 9-Released Op, 11-User Restriction Fields (Client Editor),12-Read only Plan
                        '                   13-Read only Operation, 14-Plan is Locked, 15 - Master operation, 16-Shared Operation (linked to Master op)

                        Select Case HttpContext.Current.Session("gPartAccessRights")

                            'Case 2      'readonly
                            '** Display readonly message (frmopsumm) ** 
                            '    If Trim(objLabel.Text) <> "" Then objLabel.Text &= ", "
                            '    objLabel.Text &= " " & Ec.AppConfig.GetWrd(139).Trim
                            Case 3      'share plan
                                If Trim(objLabel.Text) <> "" Then objLabel.Text &= ", "
                                objLabel.Text &= " " & Ec.AppConfig.GetWrd(247).Trim
                                'Case 10     'read-only-signoff
                                '    If Trim(objLabel.Text) <> "" Then objLabel.Text &= ", "
                                '    objLabel.Text &= " " & Ec.AppConfig.GetWrd(140).Trim
                            Case 11     'user restriction fields
                                If Trim(objLabel.Text) <> "" Then objLabel.Text &= ", "
                                objLabel.Text &= " " & Ec.AppConfig.GetWrd(430).Trim
                        End Select

                        If blnPlanLocked Or HttpContext.Current.Session("gPartAccessRights") = 14 Then
                            If Trim(objLabel.Text) <> "" Then objLabel.Text &= ", "
                            objLabel.Text &= " " & Ec.AppConfig.GetWrd(3202).Trim
                        End If

                        If HttpContext.Current.Session("gTypoChanges") Then
                            If Trim(objLabel.Text) <> "" Then objLabel.Text &= ", "
                            objLabel.Text &= " " & Ec.AppConfig.GetWrd(3973).Trim
                        End If
                    Case Else       'for history, recyclebin, lb recyclebin plans
                        objLabel.Text = " " & GetPlanTypeWord(strRectype)
                End Select

                If HttpContext.Current.Session("gPartAccessRights") = 15 Then
                    If strRectype.Trim <> "0" Then
                        objLabel.Text = GetPlanTypeWord(strRectype)
                    End If
                    objLabel.Text &= " " & Ec.AppConfig.GetWrd(4246)
                ElseIf HttpContext.Current.Session("gPartAccessRights") = 16 Then
                    If strRectype.Trim <> "0" Then
                        objLabel.Text = GetPlanTypeWord(strRectype)
                    End If
                    objLabel.Text &= " " & Ec.AppConfig.GetWrd(4233)
                End If

                If blnMasterSharedOP Then
                    If objOPH.SharedOPID = -99 Then
                        objLabel.Text &= " " & Ec.AppConfig.GetWrd(4246)

                        If objOPH.CheckoutEngineer <> "" Then
                            objLabel.Text &= " (" & Ec.AppConfig.GetWrd(3591) & ")"
                        End If
                    End If

                End If

                'If blnDisplayOnlyPartType Then GoTo ExitThisSub

ExitThisSub:
                If Trim(objLabel.Text) <> "" Then
                    objLabel.Visible = True
                    ' objLabel.Refresh() 
                End If
            Catch ex As Exception
                GeneralError("DisplayPartFlagText", ex, True)
            Finally
            End Try
            Call CheckForUserSessionErrors()

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub RecalculateTimeTotals(ByVal blnAllSubheaders As Boolean, ByVal blnOperations As Boolean, ByVal blnPlan As Boolean, _
                         ByVal blnUpdateDB As Boolean, ByVal strOPNO As String, Optional ByVal blnUpdateOPCostRate As Boolean = False)

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnAllSubheaders:({0}) blnOperations:({1}) blnPlan:({2}) blnUpdateDB:({3}) strOPNO:({4}) blnUpdateOPCostRate:({5})",
                                                                 blnAllSubheaders, blnOperations, blnPlan, blnUpdateDB, strOPNO, blnUpdateOPCostRate), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'updateplantimetotal, updateoperationtimetotal, updatesubheadertimetotal
            'update the time total for the subheaders,operations and plan
            'Requirement: gid,grectype,gseq,goph, gopsumm,grh array should have values 


            'Param strOPNO: if the OPNo has value, then update the only one op
            Dim intK As Int16 = 0
            Dim sngCycleTime As Single = 0, sngSetupTime As Single = 0, sngStationTime As Single = 0
            Dim sngCycleTimeBasic As Single = 0, sngSetupTimeBasic As Single = 0
            Dim blnUpdateOPSumm As Boolean = False, blnExitFor As Boolean = False
            Dim objOPSumm(0) As EASEClass7.Parts.stOPSummary
            Dim objOPH As EASEClass7.Parts.stOPH
            Dim objRH As EASEClass7.Parts.stRH
            Dim blnSkipRecalcSHTotal As Boolean = False
            Try
                objOPH = HttpContext.Current.Session("gOPH")
                objRH = HttpContext.Current.Session("gRH")
                objOPSumm = HttpContext.Current.Session("gOPSumm")
                If blnAllSubheaders And Trim(strOPNO) <> "" Then
                    'now, Update the time total for every subheader and update the database
                    Ec.TimeTotals.RecalculateAllSubheadersTotal(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq "), strOPNO)
                    'we've already calculated the SH total for the op here, no need to do it again
                    blnSkipRecalcSHTotal = True     '**RETAILEASECHANGES - REV3 CHANGES**  (for better performance)
                End If

                If blnOperations Then
                    blnExitFor = False
                    For intK = 1 To UBound(objOPSumm)
                        blnUpdateOPSumm = False
                        If Trim(strOPNO) = "" Then  'recalculate total for all operations
                            blnUpdateOPSumm = True
                            'now recalculate the operation total and store it
                            Ec.TimeTotals.RecalculateOpTotal(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), objOPSumm(intK).OPNO, sngCycleTime, sngSetupTime, sngCycleTimeBasic, sngSetupTimeBasic, sngStationTime, blnUpdateDB, objOPSumm(intK).NoMen)
                        Else        'for a specific operation

                            If objOPSumm(intK).OPNO.Trim = strOPNO.Trim Then
                                'now recalculate the operation total and store it
                                Ec.TimeTotals.RecalculateOpTotal(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), objOPSumm(intK).OPNO, sngCycleTime, sngSetupTime, sngCycleTimeBasic, sngSetupTimeBasic, sngStationTime, blnUpdateDB, objOPSumm(intK).NoMen, blnSkipRecalcSHTotal)

                                objOPH.CycleTime = sngCycleTime
                                objOPH.SetupTime = sngSetupTime
                                objOPH.CycleTimeBasic = sngCycleTimeBasic
                                objOPH.SetupTimeBasic = sngSetupTimeBasic
                                objOPH.StationTime = sngStationTime
                                'update cost record
                                If blnUpdateOPCostRate Then objOPSumm(intK).CostRate = objOPH.CostRate

                                blnUpdateOPSumm = True
                                blnExitFor = True

                            End If
                        End If
                        If blnUpdateOPSumm Then
                            objOPSumm(intK).ACycTime = sngCycleTime
                            objOPSumm(intK).ASetupTime = sngSetupTime

                            objOPSumm(intK).ACycleTimeBasic = sngCycleTimeBasic
                            objOPSumm(intK).ASetupTimeBasic = sngSetupTimeBasic
                            objOPSumm(intK).StationTime = sngStationTime
                            HttpContext.Current.Session("gOPSumm") = objOPSumm
                        End If
                        If blnExitFor Then Exit For
                    Next

                End If

                If blnPlan Then
                    'recalculate the part total and update database
                    Ec.TimeTotals.ProcessPlanTotal(objOPSumm, objRH, blnUpdateDB)
                    HttpContext.Current.Session("gRH") = objRH
                    HttpContext.Current.Session("gOPSumm") = objOPSumm
                    HttpContext.Current.Session("gOPH") = objOPH
                End If
            Catch ex As Exception
                GeneralError("RecalculateTimeTotals", ex, True)
            End Try
            Call CheckForUserSessionErrors()

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function ConvertPlanDocumentsToPDF() As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim blnRtnValue As Boolean = True, blnDocExist As Boolean = False
            Dim intWP As Int16 = 0, strFileExtn As String = "", strFilename As String = ""
            Dim strErrLoc As String = "", strOPNO As String = "RHWITX"
            Dim objParams As New EASEClass7.Text.stTextParams, strPDFFileName As String = ""
            Dim objDocs() As EASEClass7.SubHeaders.stDocs, intK As Int16 = 0
            Dim intCompleteTTKey As Int32 = 0
            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")

            Try
                'Update the Plan Level work instructions
                If Not EASEClass7.DBConfig.ActivateWorkInstructions Then GoTo SkipToPlanDocs

                'check document exist
                blnDocExist = Ec.Text.CheckWorkInstructionsExist(objRH.ID, objRH.Rectype, objRH.Seq, strOPNO, intWP)
                If Not blnDocExist Then GoTo SkipToPlanDocs 'if document doesn't exist, then skip wI


                strFileExtn = Ec.GeneralFunc.GetDocumentType("numberToExtn", intWP)

                objParams.Param1 = objRH.ID
                objParams.Param2 = objRH.Rectype
                objParams.Param3 = objRH.Seq
                objParams.Param4 = strOPNO
                objParams.Param5 = 0     'ttkey
                objParams.Param6 = intWP

                strFilename = Ec.GeneralFunc.GetTempPath & Ec.GeneralFunc.GetRandomFileName(True) & strFileExtn
                strErrLoc = "WI: ReadTextFromDB"
                'extract the document from database
                If Not Ec.Text.ReadTextFromDB(7, objParams, strFilename) Then
                    strFilename = ""
                End If

                'if any error occurred while retriving the wi, then skipwi
                If strFilename = "" Then GoTo SkipToPlanDocs

                strErrLoc = "WI: UpdateWorkInstructionsTemplate"
                'update the Work Instructions Word, Excel or Powerpoint docs
                '!@#$  have to do 

                ' strPDFFileName = UpdateWorkInstructionsTemplate(strFilename, intWP, True, EASEClass7.DBConfig.ConvertToPDF, False, False, False, False)
                '  Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Ec.GeneralFunc.FileExists(strPDFFileName) Then
                    'After updating the document header and footer, update the database

                    strFileExtn = Ec.GeneralFunc.GetFileExtension(strPDFFileName)
                    intWP = Ec.GeneralFunc.GetDocumentType("extnTonumber", strFileExtn)

                    'write text to the database
                    objParams.Param1 = HttpContext.Current.Session("gID")
                    objParams.Param2 = HttpContext.Current.Session("gRectype")
                    objParams.Param3 = HttpContext.Current.Session("gSeq")
                    objParams.Param4 = strOPNO
                    objParams.Param5 = 0     'ttkey
                    objParams.Param6 = intWP

                    objParams.FileName = strPDFFileName
                    strErrLoc = "WI: WriteTextToDB"
                    If Not Ec.Text.WriteTextToDB(8, objParams, False) Then
                        'error occurred while writing text to database, exit out
                        GoTo SkipToPlanDocs
                    End If
                End If

SkipToPlanDocs:
                'update the plan level refdocs
                objDocs = Ec.Parts.GetPlanDocuments(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), 4)
                For intK = 1 To UBound(objDocs)
                    'skip mdm ref docs
                    If objDocs(intK).MDMDocID > 0 Then GoTo SkipThisRefDoc

                    If Not Ec.GeneralFunc.EASETemplate(objDocs(intK).Filename) Then GoTo SkipThisRefDoc


                    intWP = objDocs(intK).WPType

                    strFilename = GetDocumentFileName(objDocs(intK).TTKey, 0, _
                                       objDocs(intK).MDMDocID, intWP, _
                                  Trim(objDocs(intK).Filename), Trim(objDocs(intK).FilePathx), , , , True, False)

                    'defensive coding, if the file doesn't exist, just skip this record, hi hi
                    If Not Ec.GeneralFunc.FileExists(strFilename) Then GoTo SkipThisRefDoc


                    'update the document header and footer
                    '!@#$  have to do 
                    '  strPDFFileName = UpdateReferenceDocumentTemplate(strFilename, objDocs(intK).WPType, True, EASEClass7.DBConfig.ConvertToPDF, False, False, False, False)
                    'defensive coding, if the file doesn't exist, just skip this record, hi hi

                    If Not Ec.GeneralFunc.FileExists(strPDFFileName) Then GoTo SkipThisRefDoc


                    'get the new completettkey for the document
                    If intCompleteTTKey = 0 Then
                        intCompleteTTKey = Ec.Text.GetNewTTKeyForPDFTextTable(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"))
                    Else
                        intCompleteTTKey += 1
                    End If

                    strFileExtn = Ec.GeneralFunc.GetFileExtension(strPDFFileName)

                    objParams.Param1 = HttpContext.Current.Session("gID")
                    objParams.Param2 = HttpContext.Current.Session("gRectype")
                    objParams.Param3 = HttpContext.Current.Session("gSeq")
                    objParams.Param4 = strOPNO
                    objParams.Param5 = intCompleteTTKey
                    objParams.Param6 = Ec.GeneralFunc.GetDocumentType("extnTonumber", strFileExtn)          'wptype

                    objParams.FileName = strPDFFileName
                    strErrLoc = "WI: WriteTextToDB"
                    Ec.GeneralFunc.DelayProcess()
                    If Not Ec.Text.WriteTextToDB(8, objParams, False) Then
                        'error occurred while writing text to database, exit out
                        GoTo SkipThisRefDoc
                    End If


                    Ec.Parts.UpdateCompleteTTKey(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), strOPNO, objDocs(intK).MSeq, intCompleteTTKey, objDocs(intK).Type, True)

SkipThisRefDoc:
                Next
            Catch ex As Exception
                blnRtnValue = False
                GeneralError("ConvertPlanDocumentsToPDF", ex, True)
            Finally
                objParams = Nothing
                objDocs = Nothing
            End Try
            Call CheckForUserSessionErrors()


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue

        End Function

        Public Sub SetPlanAccessRights(Optional ByVal blnOperationLevel As Boolean = False)


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnOperationLevel:({0})", blnOperationLevel), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'When opening a plan, this sub is being called from  RecallPlan Sub and
            'set the user access rights

            'gPartAccessRights: 1-Full Access, 3-Shared Plan,4-Alternate Plan,5-History Plan,6-Recycle Bin,7-LB Backup
            '                   8-Released Plan, 9-Released Op, 11-User Restriction Fields (Client Editor),12-Read only Plan
            '                   13-Read only Operation, 14-Plan is Locked, 15 - Master operation, 16-Shared Operation (linked to Master op)

            'user restriction fields are user0,user1,user2, user3 (User Defined fields) and ALL Capp Fields
            Dim strFieldValue As String = ""

            Dim intField1 As Int16 = 0, strFieldValue1 As String = ""
            Dim intField2 As Int16 = 0, strFieldValue2 As String = "", intK As Int16 = 0
            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Dim objOPH As EASEClass7.Parts.stOPH
            objOPH = HttpContext.Current.Session("gOPH")
            Dim strName As String = ""
            strName = HttpContext.Current.Session("UserName")

            Try
                HttpContext.Current.Session("gPartAccessRights") = 1 ' user has full access
                Select Case HttpContext.Current.Session("gRectype")
                    Case 0, 1      'active part, alternate plan (allow users to edit)

                        If objRH.LockPlan AndAlso (objRH.Engineer.Trim.ToLower <> LCase(Trim(strName))) Then
                            HttpContext.Current.Session("gPartAccessRights") = 14
                        End If
                    Case 2  'history
                        HttpContext.Current.Session("gPartAccessRights") = 5
                    Case 3  'recycle bin
                        HttpContext.Current.Session("gPartAccessRights") = 6
                    Case 4 'lb backup
                        HttpContext.Current.Session("gPartAccessRights") = 7
                    Case Else       'for all other plan types, just open them as readonly
                        HttpContext.Current.Session("gPartAccessRights") = 12
                End Select

                'if the login control is off, then users has full access to the plan,  hi hi,  how is it?
                If Not EASEClass7.DBConfig.LoginControl Then GoTo ExitThisSub

                'defensive coding, if the user id is N/A then user has full access rights
                If Trim(HttpContext.Current.Session("UserID")) = Trim(EASEClass7.LangConstants.NA) Then GoTo ExitThisSub ' "N/A" Then GoTo ExitThisSub

                'if the user has only, readonly or readonly -signoff access rights, then reset the gPartAccessRights variable
                ''gUserAccessRights(EASE Access Rights): 1-All, 4-Process Plan RWD, 5-Process Plan RW, 6- Process Plan Read, 7-Read only Sign-off
                Select Case HttpContext.Current.Session("UserAccessRights")
                    Case 6      'user has only readonly access rights
                        HttpContext.Current.Session("gPartAccessRights") = 2
                        GoTo ExitThisSub
                    Case 7      'readonly-signoff
                        HttpContext.Current.Session("gPartAccessRights") = 10
                        GoTo ExitThisSub
                End Select

                If Ec.AppConfig.CAPP Or Ec.AppConfig.CAPPOnly Then
                    If Trim(HttpContext.Current.Session("gRectype")) = "0" Or Trim(HttpContext.Current.Session("gRectype")) = "1" Then
                        If Not EASEClass7.DBConfig.OperationSignoff Then
                            If objRH.Released Then
                                HttpContext.Current.Session("gPartAccessRights") = 8               'released plan
                            End If
                        Else
                            If blnOperationLevel Then
                                If objOPH.ReleasedFlag <> 0 Then
                                    HttpContext.Current.Session("gPartAccessRights") = 9               'released op
                                End If
                            End If
                        End If
                    End If
                End If

                If EASEClass7.DBConfig.SharePlan Then       'checked-out copy
                    'user does not have access rights to access this plan 
                    If Ec.Parts.CheckedOutPart(objRH.PartNo) Then
                        HttpContext.Current.Session("gPartAccessRights") = 3
                    End If
                End If

                ' check it ( gUser.RestrictedUser)
                If Not HttpContext.Current.Session("gUser.RestrictedUser") Then GoTo SkipRestrictedUser

                '**  restricted user flag is set, so the user has some restriction field set in his profile.   **

                'user the user restriction fields
                Ec.User.GetUserRestrictedFields(HttpContext.Current.Session("UserID"), HttpContext.Current.Session("UserPlantID"), intField1, strFieldValue1, intField2, strFieldValue2)

                'first restriction field, same code snippet is used for checking second field too (see below)
                If intField1 > 0 Then
                    strFieldValue = ""
                    If intField1 < 100 Then         'route header fields
                        Select Case intField1
                            Case 1              'user0
                                strFieldValue = objRH.User0
                            Case 2              'user1
                                strFieldValue = objRH.User1
                            Case 3              'user2
                                strFieldValue = objRH.User2
                            Case 4              'user3
                                strFieldValue = objRH.User3
                        End Select
                    Else                        'Route header capp fields
                        If Ec.AppConfig.CAPP AndAlso intField1 <= UBound(objRH.CAPPFields) Then
                            strFieldValue = objRH.CAPPFields(intField1)
                        End If
                    End If
                    If Ec.GeneralFunc.QPTrim(strFieldValue) = "" Or Ec.GeneralFunc.CompareString(strFieldValue, strFieldValue1) Then
                    Else
                        HttpContext.Current.Session("gPartAccessRights") = 12           'readonly access rights
                    End If
                End If

                'Second restriction field, same code snippet is used for checking second field too (see above)
                If intField2 > 0 Then
                    strFieldValue = ""
                    If intField2 < 100 Then         'route header fields
                        Select Case intField2
                            Case 1              'user0
                                strFieldValue = objRH.User0
                            Case 2              'user1
                                strFieldValue = objRH.User1
                            Case 3              'user2
                                strFieldValue = objRH.User2
                            Case 4              'user3
                                strFieldValue = objRH.User3
                        End Select
                    Else                        'Route header capp fields
                        If Ec.AppConfig.CAPP AndAlso intField2 <= UBound(objRH.CAPPFields) Then
                            strFieldValue = objRH.CAPPFields(intField2)
                        End If
                    End If
                    If Ec.GeneralFunc.QPTrim(strFieldValue) = "" Or Ec.GeneralFunc.CompareString(strFieldValue, strFieldValue1) Then
                    Else
                        HttpContext.Current.Session("gPartAccessRights") = 12           'readonly access rights
                    End If
                End If

SkipRestrictedUser:

                If HttpContext.Current.Session("gPartAccessRights") = 1 AndAlso HttpContext.Current.Session("gReadOnlyPlan") Then
                    HttpContext.Current.Session("gPartAccessRights") = 12           'readonly access rights
                End If

                If HttpContext.Current.Session("gPartAccessRights") = 1 AndAlso HttpContext.Current.Session("gReadOnlyOperation") Then
                    HttpContext.Current.Session("gPartAccessRights") = 13           'readonly access rights for the operation
                End If


                If blnOperationLevel AndAlso (EASEClass7.DBConfig.OperationSignoff And EASEClass7.DBConfig.SharedOperation) Then
                    If SharedOperation(objOPH.OPNO) Then
                        HttpContext.Current.Session("gPartAccessRights") = 16           'share plan
                    End If
                End If

ExitThisSub:
            Catch ex As Exception
                GeneralError("SetPlanAccessRights", ex, True)
            End Try
            Call CheckForUserSessionErrors()

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub OpenSearchForm(ByVal strFormParam As String, Optional ByVal strPartNO As String = "", Optional ByVal strParam1 As String = "")

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED(String.Format("Enter strFormParam:({0}) strPartNO:({1}) strParam1:({2})", strFormParam, strPartNO, strParam1), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            HttpContext.Current.Session("gAN") = strPartNO
            HttpContext.Current.Session("gFormParam") = strFormParam
            HttpContext.Current.Session("gLIN") = strParam1.Trim
            Log.REDIRECT_TRANSFER("Exit Redirect(frmSearch.aspx", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
            System.Web.HttpContext.Current.Response.Redirect("frmSearch.aspx", True)
            ' remaining covered in SEARCH Form

            'ShowDialogForm(New frmSearch)
            'If HttpContext.Current.Session("gFormParam") = "openplan" Then
            '    ShowThisPart()
            'ElseIf HttpContext.Current.Session("gFormParam") = "material-addpart" Then
            '    strAN = HttpContext.Current.Session("gAN")
            'ElseIf HttpContext.Current.Session("gFormParam") = "renameplan" Then
            '    GoTo FinalExit
            'ElseIf HttpContext.Current.Session("gFormParam") = "openhistory" Then
            '    GoTo FinalExit
            'ElseIf HttpContext.Current.Session("gFormParam") = "openhistory-op" Then
            '    GoTo FinalExit
            'ElseIf HttpContext.Current.Session("gFormParam") = "plansearch-general" Then
            '    GoTo FinalExit
            'ElseIf HttpContext.Current.Session("gFormParam") = "plansearch-general-rename" Then
            '    GoTo FinalExit
            'ElseIf HttpContext.Current.Session("gFormParam") = "sharedoperations-released" Then
            '    GoTo FinalExit
            'ElseIf HttpContext.Current.Session("gFormParam") = "sharedoperations-usage" Then
            '    GoTo FinalExit
            'ElseIf HttpContext.Current.Session("gFormParam") = "sharedoperations-delete" Then
            '    GoTo FinalExit
            'End If

            'HttpContext.Current.Session("gFormParam") = strTempParam
            'HttpContext.Current.Session("gFormParam") = strAN
            'HttpContext.Current.Session("gFormParam") = strLin

FinalExit:
            'frmMain.RefreshSearchButtons()

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub ShowThisPart(Optional ByVal strPartNO As String = "", _
            Optional ByVal strRectype As String = "0", Optional ByVal intSeq As Int16 = 0, _
            Optional ByVal blnSkipPartInUseCheck As Boolean = False)


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter strPartNO:({0}) strRectype:({1}) intSeq:({2}) blnSkipPartInUseCheck:({3})", strPartNO, strRectype, intSeq, blnSkipPartInUseCheck), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim blnRedirectToPasswordPage As Boolean = False
            Dim objRH As EASEClass7.Parts.stRH

            If ViewEASEWebApplication() Then        'viewEASE is a read-only app
                blnSkipPartInUseCheck = True
            End If


            RecallPart(strPartNO, strRectype, intSeq, , blnRedirectToPasswordPage)
            If blnRedirectToPasswordPage Then
                '** same code snippet exist in modEASE2 (ShowThisPart) and frmMain.aspx.
                'Any change should be updated in both places.

                Ec.Parts.ClearRHObject(objRH)
                objRH.PartNo = strPartNO
                objRH.Rectype = strRectype
                objRH.Seq = intSeq
                HttpContext.Current.Session("gRH") = objRH      'need it in password form


                HttpContext.Current.Session("gFormParam") = "Open-ShowThisPart"
                System.Web.HttpContext.Current.Response.Redirect("frmPassword.aspx", True)
            End If

            'frmSHSumm = Nothing 'disable the operaion screen in memory (to avoid displaying the prev op in the screen)

            ' we did this in frmOpSumm
            'If objRH.ID <> 0 Then 'part ID found
            '    frmOpSumm.RefreshOpSummaryScreen()
            'Else        'part doesn't exist
            '    frmOpSumm.LoadBlankScreen()
            'End If
            ' System.Web.HttpContext.Current.Response.Redirect("frmOpSumm.aspx", True)


#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function CheckPartInMemory() As Boolean

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")

            'check the part is in memory
            Dim blnRtnValue As Boolean = True
            If Trim(objRH.PartNo) = "" Then GoTo ExitThisSub
            Dim strMsg As String = "" 'Ec.AppConfig.GetWrd(3901)

            blnRtnValue = False
            HttpContext.Current.Session("gParam") = "CheckMemory"
            HttpContext.Current.Session("gTemp2") = objRH.PartNo.Trim
            System.Web.HttpContext.Current.Response.Redirect("DialogBox.aspx", True)

            ' remaining covered in above page

            ' strMsg = Ec.AppConfig.GetWrd(4017).Trim & " ('" & objRH.PartNo.Trim & "') " & Ec.AppConfig.GetWrd(3901)
            ' If MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            'part in memory, do you want to continue, YES
            'Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'Reset the network lock
            'Ec.Parts.SetNetworkLock(HttpContext.Current.Session("gID"), HttpContext.Current.Session("gRectype"), HttpContext.Current.Session("gSeq"), False, HttpContext.Current.Session("UserID"))

            'Windows.Forms.Cursor.Current = Cursors.Default

            ' System.Web.HttpContext.Current.Response.Redirect("frmOpSumm.aspx", True)
            'frmOpSumm.LoadBlankScreen()
            'blnRtnValue = True
            'End If
ExitThisSub:


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Sub SavePart()

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim lngID As Long = 0
            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Try

                If objRH.NewPart Then
                    lngID = Ec.Parts.GetNewPartID()
                Else
                    lngID = objRH.ID
                End If

                If lngID = 0 Then GoTo ExitThisSub 'never happens, just in case, hi hi

                If objRH.NewPart = True Then
                    objRH.ID = lngID
                    objRH.Rectype = "0"
                    objRH.Seq = 0

                    HttpContext.Current.Session("gID") = objRH.ID
                    HttpContext.Current.Session("gRectype") = objRH.Rectype
                    HttpContext.Current.Session("gSeq") = objRH.Seq

                    HttpContext.Current.Session("gRH") = objRH
                End If

                'update the database
                Ec.Parts.WriteRH(objRH)

                'reset the newpart flag
                If objRH.NewPart Then
                    objRH.NewPart = False 'set the new part, to get the new id and uses insert query
                    HttpContext.Current.Session("gRH") = objRH
                    Ec.Parts.WriteToEASELogTable(objRH.ID, objRH.PartNo, "", HttpContext.Current.Session("UserName"), 1, "")

                    'write to recently viewed parts list
                    WriteRecentlyViewedParts(objRH.PartNo, objRH.Rectype, objRH.Seq)
                End If

            Catch ex As Exception
                GeneralError("SavePart", ex, True)
            Finally
            End Try
            CheckForUserSessionErrors()

ExitThisSub:

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub RecallOperation(ByVal intID As Int32, ByVal strRectype As String, ByVal intSeq As Int16, _
                                ByVal strOPNO As String, Optional ByVal intOPSeq As Int16 = 0, _
                                Optional ByVal blnCheckTakeOver As Boolean = False, Optional ByVal intSharedOPID As Int32 = 0)


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intID:({0}) strRectype:({1}) intSeq:({2}) strOPNO:({3}) intOPSeq:({4}) blnCheckTakeOver:({5}) intSharedOPID:({6})",
                                                                 intID, strRectype, intSeq, strOPNO, intOPSeq, blnCheckTakeOver, intSharedOPID), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim strMsg As String = "", strURL As String = ""

            Dim objOPH As EASEClass7.Parts.stOPH

            Try

                KillOP()            'kill the op in memory
                objOPH = Ec.Parts.GetOPH(intID, strRectype, intSeq, strOPNO, intOPSeq)
                If EASEClass7.DBConfig.SharedOperation AndAlso intSharedOPID > 0 Then
                    objOPH.SharedOPID = intSharedOPID
                End If

                Select Case Trim(strRectype)
                    Case "0", "1"
                        'TODO: Get Cost Rate
                End Select
                HttpContext.Current.Session("gOPH") = objOPH
                If Ec.AppConfig.CAPP Or Ec.AppConfig.CAPPOnly Then
                    If EASEClass7.DBConfig.OperationSignoff Then

                        SetPlanAccessRights(True)       'Set the user access rights for this Operation

                        If blnCheckTakeOver Then
                            'If the if op is unreleased, check the engineer and ask for take over message?
                            'also confirm, the user has access rights to edit the plan
                            If HttpContext.Current.Session("gPartAccessRights") = 1 AndAlso objOPH.ReleasedFlag = 0 Then
                                'TODO: Custom For PCN/MCR Changes
                                If objOPH.Engineer.Trim <> "" AndAlso LCase(objOPH.Engineer.Trim) <> Trim(LCase(HttpContext.Current.Session("UserName"))) Then
                                    strMsg = Ec.AppConfig.GetWrd(3197) & " " & objOPH.Engineer.Trim & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3195)
                                    'GoTo ExitThisSub
                                    '!@#$ remaining covered in above page.
                                    'If MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                    '    'update the routeheader engineer
                                    '    gOPH.Engineer = gUser.UserName
                                    '    Ec.Parts.UpdateOperationEngineer(intID, strRectype, intSeq, gOPH.OPNO, gOPH.Engineer)
                                    'Else
                                    '    gReadOnlyOperation = True
                                    'End If
                                End If
                            End If

                            If HttpContext.Current.Session("gReadOnlyOperation") Then
                                SetPlanAccessRights()           'Reset the access rights
                            End If
                        End If
                    End If
                End If
ExitThisSub:

            Catch ex As Exception
                GeneralError("RecallOperation", ex, True)
            Finally
            End Try

            CheckForUserSessionErrors()
            If strURL.Trim <> "" Then
                HttpContext.Current.Response.Redirect(strURL, True)
            End If

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub SaveOP(Optional ByVal blnUpdateOpeationTimeTotal As Boolean = False)

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnUpdateOpeationTimeTotal:({0})", blnUpdateOpeationTimeTotal), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim objOPH As EASEClass7.Parts.stOPH
            objOPH = HttpContext.Current.Session("gOPH")

            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Dim blnGetOpSeq As Boolean = False
            Try
                If objOPH.NewOP Then        'get the new op seq
                    objOPH.OPSeq = Ec.Parts.GetOpSeqMax(objOPH.ID, objOPH.Rectype, objOPH.Seq)
                End If



                'write op header record
                Ec.Parts.WriteOPH(objOPH, , blnGetOpSeq)

                'reset the newpart flag
                If objOPH.NewOP Then
                    objOPH.NewOP = False 'uses insert query instead of update query
                    HttpContext.Current.Session("gOPH") = objOPH
                    Ec.Parts.WriteToEASELogTable(objRH.ID, objRH.PartNo, objOPH.OPNO, HttpContext.Current.Session("UserName"), 101, "")       'Write to EASELog table

                End If
                HttpContext.Current.Session("gOPH") = objOPH
                If blnUpdateOpeationTimeTotal Then
                    'update subheader time total

                    'update operation time total
                    RecalculateTimeTotals(True, True, True, True, objOPH.OPNO)
                End If


                RefreshPart()        'refresh gOPSumm Array, recalculate time totals and update database
            Catch ex As Exception
                GeneralError("SaveOP", ex, True)
            End Try
            CheckForUserSessionErrors()

#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub AddPartOperationInCaption(ByRef strCaption As String, Optional ByVal strPartNO As String = "", Optional ByVal strOPNO As String = "", Optional ByVal strSubHeader As String = "")

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED(String.Format("Enter strPartNO:({0}) strOPNO:({1}) strSubHeader:({2})", strPartNO, strOPNO, strSubHeader), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            If strPartNO.Trim <> "" Then
                strCaption &= " <span class=pageheader2_small>(" & Ec.AppConfig.GetWrd(4017) & " " & strPartNO.Trim

                If Trim(strOPNO) <> "" Then
                    strCaption &= ", " & Ec.AppConfig.GetWrd(4018) & " " & strOPNO.Trim
                End If
                If Trim(strSubHeader) <> "" Then
                    strCaption &= ", " & Ec.AppConfig.GetWrd(3698) & " " & strSubHeader.Trim
                End If

                strCaption &= ")</span>"
            End If

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function UnReleasePart(Optional ByVal blnManualUpdate As Boolean = False, Optional ByVal blnEnableTypoChange As Boolean = False, Optional ByVal objLabel As System.Web.UI.WebControls.Label = Nothing) As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnManualUpdate:({0}) blnEnableTypoChange:({1})", blnManualUpdate, blnEnableTypoChange), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'blnManualUpdate=true only in frmupdatereleasestatus form, to manually update the part/op rev
            'blnEnableTypoChange: ask the user, whether to do a typo change or complete unrelease op, this option is enabled only
            '                       when the sub is called from shsumm form (all the others, ignore it)

            Dim blnRtnValue As Boolean = False, intAuthGroups As Int16 = 0, strMsg As String = ""
            Dim strAuthRel As String = "", strFormParam As String = HttpContext.Current.Session("gFormParam")
            Dim strLIN As String = ""
            Dim strURL As String = ""
            'TODO: ChildPart Validation
            'TODO: Allow typo changes
            Dim objCH As EASEClass7.Parts.stChangeHistory
            objCH = HttpContext.Current.Session("gobjCH")
            'check User Access rights

            If Not VerifyUserAccessRights(True, True, , True, , , , , objLabel) Then GoTo FinalExit

            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Dim objCopy As New EASEClass7.Parts.stCopyParams, arrCH(1) As String
            Dim strName As String = ""
            strName = HttpContext.Current.Session("UserName")
            Try

                If HttpContext.Current.Session("gTypoChanges") Then            'the typo changes are in process.don't allow the user to do nothing.
                    HttpContext.Current.Session("gTypoChanges") = False
                    blnRtnValue = True
                    GoTo ExitThisFunction
                End If

                'Check number of signoff
                intAuthGroups = Ec.DBConfig.GetAuthGroupCount(objRH.PlantID)
                If intAuthGroups = 0 Then
                    strMsg = Ec.AppConfig.GetWrd(3177) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3123)
                    'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DisplayUserMessage(objLabel, strMsg)
                    GoTo ExitThisFunction
                End If

                If blnManualUpdate Then
                    'manual release
                    GoTo SkipcopyingToHistory
                End If

                'confirm the user to release the part
                strMsg = Ec.AppConfig.GetWrd(3938) & "<br><br>" & Ec.AppConfig.GetWrd(4493)
                HttpContext.Current.Session("gAN") = strMsg ' strLIN = HttpContext.Current.Session("gLIN") :
                HttpContext.Current.Session("gLIN") = ""

                HttpContext.Current.Session("gTypoChanges") = False
                If EASEClass7.DBConfig.AllowTypoChange Then HttpContext.Current.Session("gLIN") = "typochange"
                HttpContext.Current.Session("gFormParam") = "PlanSignOff"
                strURL = "frmPassword.aspx"
                GoTo ExitThisFunction
                ' remaining covered in password form

                If HttpContext.Current.Session("gLIN") = "processtypochange" Then HttpContext.Current.Session("gTypoChanges") = True
                HttpContext.Current.Session("gLIN") = strLIN  'reset the glin

                'user cancelled, do nothing
                If HttpContext.Current.Session("gAN") = "*" Then GoTo ExitThisFunction
                HttpContext.Current.Session("gAN") = ""

                If HttpContext.Current.Session("gTypoChanges") Then
                    'write to change history
                    objCH.ID = HttpContext.Current.Session("gID")
                    objCH.OPNO = ""
                    arrCH(1) = "Typo Changes"
                    objCH.Comment = arrCH
                    objCH.Reason = arrCH

                    objCH.PCN = ""
                    objCH.Rev = "1"
                    objCH.OpRev = 1

                    objCH.Engineer = HttpContext.Current.Session("UserName")
                    objCH.EntryDate =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat
                    objCH.CommentSeq = 1
                    objCH.EntrySeq = 0
                    Ec.Parts.WriteChangeHistory(objCH, True)


                    Ec.Parts.WriteToEASELogTable(objRH.ID, objRH.PartNo, "", strName, 7, "Plan Rev:" & objRH.RevNo)       'Write to EASELog table
                    HttpContext.Current.Session("gobjCH") = objCH
                    blnRtnValue = True
                    GoTo ExitThisFunction  'don't unrelease the PART
                End If

                'Copy the Active Part to history
                Ec.Parts.ClearCopyParamsObject(objCopy)

                'copy 
                objCopy.FromID = HttpContext.Current.Session("gID")
                objCopy.FromRectype = HttpContext.Current.Session("gRectype")
                objCopy.FromSeq = HttpContext.Current.Session("gSeq")

                objCopy.ToID = HttpContext.Current.Session("gID")
                objCopy.ToRectype = "2"
                objCopy.ToSeq = Ec.Parts.GetPartSeqMax(objCopy.ToID, objCopy.ToRectype)
                objCopy.ToOPEngineer = HttpContext.Current.Session("UserName")
                If Not Ec.Parts.CopyPartOperation("unreleasepart", objCopy, Ec.AppConfig.EASETempPath, strName) Then
                    'error occurred while copying the operation into Active. exit out
                    GoTo ExitThisFunction
                End If

SkipcopyingToHistory:
                'Update the Rev#, Auth status, Engineer

                strAuthRel = Ec.Parts.EncryptPartAuthRel(objRH.PlantID, False, intAuthGroups)

                objRH.RevNo += 1
                objRH.AuthRel = strAuthRel
                objRH.Engineer = Trim(HttpContext.Current.Session("UserName"))
                objRH.Released = False

                Ec.Parts.UpdatePartReleaseStatus(objRH.ID, objRH.Rectype, objRH.Seq, objRH.RevNo, Date.MinValue, objRH.AuthRel, objRH.PlantID, objRH.Engineer)


                'reset completettkey for plan, operation and subheader documents
                Ec.Parts.ResetCompleteTTKeyForAllDocuments(objRH.ID, objRH.Rectype, objRH.Seq.ToString, "")

                Ec.Parts.WriteToEASELogTable(objRH.ID, objRH.PartNo, "", strName, 5, "Plan Rev:" & objRH.RevNo)       'Write to EASELog table
                HttpContext.Current.Session("gRH") = objRH

                'Write to change history
                HttpContext.Current.Session("gFormParam") = "unreleasepart"
                ClearChangeHistoryRecord()

                If blnManualUpdate Then
                    'write to change history
                    objCH.ID = HttpContext.Current.Session("gID")
                    objCH.OPNO = ""
                    arrCH(1) = "Manual Un-release Plan "
                    objCH.Comment = arrCH
                    objCH.Reason = arrCH

                    objCH.PCN = ""
                    objCH.Rev = "1"
                    objCH.OpRev = 1

                    objCH.Engineer = HttpContext.Current.Session("UserName")
                    objCH.EntryDate =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat
                    objCH.CommentSeq = 1
                    objCH.EntrySeq = 0
                    Ec.Parts.WriteChangeHistory(objCH, True)
                Else
                    objCH.ID = objRH.ID
                    objCH.Engineer = objRH.Engineer
                    objCH.EntryDate =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat ' EaseCore.Extensions.Dates.GetCurrentDate
                    objCH.Rev = objRH.RevNo
                    HttpContext.Current.Session("gobjCH") = objCH
                    strURL = "frmChangeHistoryEdit.aspx"
                    GoTo ExitThisFunction
                End If

                blnRtnValue = True

ExitThisFunction:
            Catch ex As Exception
                GeneralError("UnReleasePart", ex, True)
            Finally
                objCopy = Nothing
            End Try
FinalExit:
            CheckForUserSessionErrors()
            If strURL.Trim <> "" Then
                System.Web.HttpContext.Current.Response.Redirect(strURL, True)
            End If


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Function ReleasePart(Optional ByVal blnManualUpdate As Boolean = False, Optional ByVal objLabel As System.Web.UI.WebControls.Label = Nothing) As Boolean

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnManualUpdate:({0})", blnManualUpdate), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'blnManualUpdate=true only in frmupdatereleasestatus form, to manually update the part/op rev

            'Requirment: Load gid,grectype,gseq,gopsumm vars
            Dim objRH As EASEClass7.Parts.stRH
            Dim objOPSumm(0) As EASEClass7.Parts.stOPSummary
            Dim blnRtnValue As Boolean = False
            Dim strMsg As String = "", intAuthGroups As Int16 = 0
            Dim intTemp As Int16 = 0
            Dim dtReleased As Date = Date.MinValue
            Dim objTempOPH As EASEClass7.Parts.stOPH = HttpContext.Current.Session("gOPH")
            Dim blnReleasePart As Boolean = False, intK As Int16 = 0
            Dim intTemp1 As Int16 = 0
            Dim strURL As String = ""
            Dim objCH As EASEClass7.Parts.stChangeHistory, arrCH(1) As String
            Dim strName As String = ""
            strName = HttpContext.Current.Session("UserName")
            HttpContext.Current.Session("gAN") = ""

            'check User Access rights
            If Not VerifyUserAccessRights(True, True, , True, , , , , objLabel) Then GoTo ExitThisSub
            'TODO: Check for Child part

            Try
                objRH = HttpContext.Current.Session("gRH")
                objOPSumm = HttpContext.Current.Session("gOPSumm")


                If EASEClass7.DBConfig.SharePlan Then
                    'if the plan is shared plan (checkout-part) then, no signoff
                    If Ec.Parts.CheckedOutPart(objRH.PartNo) Then
                        strMsg = Ec.AppConfig.GetWrd(3961) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3123)
                        ' MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        DisplayUserMessage(objLabel, strMsg)
                        GoTo ExitThisSub
                    End If

                    'make sure, no operations are checked out
                    strMsg = ""
                    For intK = 1 To UBound(objOPSumm)
                        If objOPSumm(intK).CheckoutEngineer.Trim <> "" Then
                            strMsg &= objOPSumm(intK).OPNO.Trim
                        End If
                    Next
                    If Trim(strMsg) <> "" Then
                        strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & strMsg & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3934)
                        'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        DisplayUserMessage(objLabel, strMsg)
                        GoTo ExitThisSub
                    End If

                End If

                'Check number of signoff
                intAuthGroups = Ec.DBConfig.GetAuthGroupCount(objRH.PlantID)
                If intAuthGroups = 0 Then
                    strMsg = Ec.AppConfig.GetWrd(3177) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3123) & _
                                vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(2724) & " " & Ec.MDM.GetPlantCode(objRH.PlantID)
                    ' MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DisplayUserMessage(objLabel, strMsg)
                    GoTo ExitThisSub
                End If

                'from fromupdatereleasestatus form, to manually release the part
                If blnManualUpdate Then
                    objRH.AuthRel = StrDup(intAuthGroups, "1")
                    HttpContext.Current.Session("gRH") = objRH
                    GoTo StartReleasing
                End If


                'HttpContext.Current.Session("gFormParam") = "Unrelease-part"
                HttpContext.Current.Session("gLIN") = "PlanSignoff"
                'Check the userid for the signoff and update the rhsign table
                strURL = "frmPlanSignoff.aspx"
                GoTo ExitThisSub
                '!@@#$ remaining is covered in frmPlanSignOff
StartReleasing:

                If HttpContext.Current.Session("gAN") = "*" Then GoTo ExitThisSub

                If Ec.Parts.CheckPartReleaseStatus(objRH.PlantID, objRH.AuthRel, intAuthGroups) Then

                    blnReleasePart = True

                    objRH.Released = True
                    If blnManualUpdate Then     'date for manual update is set in frmupdatereleasestatus form
                        dtReleased = CDate(objRH.ReleaseDate)
                    Else        'default selection
                        dtReleased = Date.Now
                        objRH.ReleaseDate = EaseCore.Extensions.Dates.CurrentDate
                    End If
                    HttpContext.Current.Session("gRH") = objRH
                End If

                'update the route header release status
                Ec.Parts.UpdatePartReleaseStatus(objRH.ID, objRH.Rectype, objRH.Seq, objRH.RevNo, dtReleased, objRH.AuthRel, objRH.PlantID)

                'TODO: TODO List ?
                'TODO: Hitachi Custom Stuff
                objRH = HttpContext.Current.Session("gRH")
                'update the reference document templates 
                If blnReleasePart Then
                    'convert the Plan Work Instructions and ref docs to pdf and update the completettkey
                    ConvertPlanDocumentsToPDF()

                    intTemp = 40 / UBound(objOPSumm)
                    intTemp1 = 45
                    'loop thro each operation and update refods and work instuctions header 
                    For intK = 1 To UBound(objOPSumm)
                        'get the operations and loop thro all documents and update the documents and work instructions
                        HttpContext.Current.Session("gOPH") = Ec.Parts.GetOPH(objRH.ID, objRH.Rectype, objRH.Seq, objOPSumm(intK).OPNO)
                        '!@#$  have to do
                        ConvertOperationDocumentsToPDF(HttpContext.Current.Session("gOPH"), True)
                        intTemp1 = 45 + intTemp

                    Next

                    'TODO: Pyxis custom program "WIEXPORT"

                    'TODO: MRPUPDATE?


                    Ec.Parts.WriteToEASELogTable(objRH.ID, objRH.PartNo, "", strName, 4, "Plan Rev:" & objRH.RevNo)       'Write to EASELog table
                End If


                If blnManualUpdate Then
                    'write to change history
                    objCH.ID = HttpContext.Current.Session("gID")
                    objCH.OPNO = ""
                    arrCH(1) = "Manual Release Plan "
                    objCH.Comment = arrCH
                    objCH.Reason = arrCH

                    objCH.PCN = ""
                    objCH.Rev = "1"
                    objCH.OpRev = 1

                    objCH.Engineer = HttpContext.Current.Session("UserName")
                    objCH.EntryDate =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat
                    objCH.CommentSeq = 1
                    objCH.EntrySeq = 0
                    Ec.Parts.WriteChangeHistory(objCH, True)
                End If

                blnRtnValue = True

            Catch ex As Exception
                GeneralError("ReleasePart", ex, True)
            Finally
                objCH = Nothing
            End Try
            CheckForUserSessionErrors()
ExitThisSub:
            HttpContext.Current.Session("gAN") = "" : HttpContext.Current.Session("gOPH") = objTempOPH
            If strURL.Trim <> "" Then
                System.Web.HttpContext.Current.Response.Redirect(strURL, True)
            End If


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Sub DisplayAccessRightsMessage(Optional ByVal intAccessRights As Int16 = 0, Optional ByVal objLabel As System.Web.UI.WebControls.Label = Nothing)

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter intAccessRights:({0})", intAccessRights), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Dim strMsg As String = ""

            'gPartAccessRights: 1-Full Access, 3-Shared Plan,4-Alternate Plan,5-History Plan,6-Recycle Bin,7-LB Backup
            '                   8-Released Plan, 9-Released Op, 11-User Restriction Fields (Client Editor),12-Read only Plan
            '                   13-Read only Operation, 14-Plan is Locked, 15 - Master operation, 16-Shared Operation (linked to Master op)

            If intAccessRights = 0 Then
                intAccessRights = HttpContext.Current.Session("gPartAccessRights")
            End If

            Select Case intAccessRights
                Case 2, 10 ' Readonly Plan
                    'You can not edit this record, your access rights are
                    strMsg = Ec.AppConfig.GetWrd(3971) & vbCrLf & vbCrLf &
                    Ec.AppConfig.GetWrd(3182) & " " & GetEASEAccessRightsValue(HttpContext.Current.Session("UserAccessRights"))
                Case 3      'Share Plan
                    strMsg = Ec.AppConfig.GetWrd(3181) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 4      'Alternate Plan messages
                    strMsg = Ec.AppConfig.GetWrd(3183) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 5      ' History Plan
                    strMsg = Ec.AppConfig.GetWrd(3184) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 6      'Recycle Bin
                    strMsg = Ec.AppConfig.GetWrd(3185) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 7      'LB Backup
                    strMsg = Ec.AppConfig.GetWrd(3186) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3971)
                Case 8      'Released part
                    strMsg = Ec.AppConfig.GetWrd(3187) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3190)
                Case 9      'Released Operation
                    strMsg = Ec.AppConfig.GetWrd(3188) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3191)
                Case 12
                    strMsg = Ec.AppConfig.GetWrd(3196)
                Case 13
                    strMsg = Ec.AppConfig.GetWrd(3198)
                Case 14
                    strMsg = Ec.AppConfig.GetWrd(3201) & " " & objRH.Engineer.Trim & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3189) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3203)
                Case 16         'shared operation
                    strMsg = Ec.AppConfig.GetWrd(3204) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3205)
                Case Else
                    strMsg = Ec.AppConfig.GetWrd(3122) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3182) &
                    " " & GetEASEAccessRightsValue(HttpContext.Current.Session("UserAccessRights")) & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(4009)
            End Select

            If strMsg.Trim <> "" Then
                DisplayUserMessage(objLabel, strMsg)
            End If


#If TRACE Then
            Log.OPERATION("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Function UnReleaseOperation(ByVal objOpSumm() As EASEClass7.Parts.stOPSummary, _
                                    Optional ByVal blnManualUpdate As Boolean = False, _
                                    Optional ByVal blnEnableTypoChange As Boolean = False, _
                                    Optional ByVal strCallFrom As String = "", Optional ByVal objLabel As System.Web.UI.WebControls.Label = Nothing) As Boolean


#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnManualUpdate:({0}) blnEnableTypoChange:({1}) strCallFrom:({2})", blnManualUpdate, blnEnableTypoChange, strCallFrom), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'blnManualUpdate=true only in frmupdatereleasestatus form, to manually update the part/op rev
            'blnEnableTypoChange: ask the user, whether to do a typo change or complete unrelease op, this option is enabled only
            '                       when the sub is called from shsumm form (all the others, ignore it)


            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Dim blnRtnValue As Boolean = False
            Dim strTempParam As String = HttpContext.Current.Session("gFormParam")

            Dim intK As Int16 = 0, strMsg As String = ""
            Dim objOPH As EASEClass7.Parts.stOPH, arrCH(1) As String

            Dim objCopyOP As New EASEClass7.Parts.stCopyParams
            Dim objTEMPOPH As EASEClass7.Parts.stOPH = HttpContext.Current.Session("gOPH")  'used in updating the ref docs template header
            Dim intIteration As Int16 = 0, intStatus As Int16 = 0, strLIN As String = ""
            Dim strURL As String = ""
            Dim objCH As EASEClass7.Parts.stChangeHistory
            objCH = HttpContext.Current.Session("gobjCH")
            Try

                If HttpContext.Current.Session("gTypoChanges") Then            'the typo changes are in process.don't allow the user to do nothing.
                    HttpContext.Current.Session("gTypoChanges") = False
                    blnRtnValue = True
                    GoTo FinalExit
                End If

                If Not VerifyUserAccessRights(True, True, , True, , , , , objLabel) Then        'check access rights message
                    GoTo ExitThisSub
                End If

                'if no operations in the object, defensive coding, hi hi
                If UBound(objOpSumm) = 0 Then GoTo ExitThisSub

                If EASEClass7.DBConfig.SharePlan AndAlso Trim(objRH.MasterPart) <> "" Then
                    'shared plan operations can not be unreleased
                    strMsg = Ec.AppConfig.GetWrd(3960)
                    DisplayUserMessage(objLabel, strMsg)
                    GoTo ExitThisSub
                End If

                'cann't release checked out op
                strMsg = ""
                For intK = 1 To UBound(objOpSumm)
                    If Trim(objOpSumm(intK).CheckoutEngineer) <> "" Then
                        objOpSumm(intK).Selected = False
                        If Trim(strMsg) <> "" Then strMsg &= " ,"
                        strMsg &= Trim(objOpSumm(intK).OPNO)
                    End If
                Next
                If Trim(strMsg) <> "" Then
                    strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3934)
                    DisplayUserMessage(objLabel, strMsg)
                    GoTo ExitThisSub
                End If

                'cann't un-release the released op
                strMsg = ""
                For intK = 1 To UBound(objOpSumm)
                    If objOpSumm(intK).ReleasedFlag <> 1 Then
                        objOpSumm(intK).Selected = False
                        If Trim(strMsg) <> "" Then strMsg &= " ,"
                        strMsg &= Trim(objOpSumm(intK).OPNO)
                    End If
                Next
                If Trim(strMsg) <> "" Then
                    strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & Ec.AppConfig.GetWrd(3940) & vbCrLf & vbCrLf & strMsg
                    'strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3934)
                    DisplayUserMessage(objLabel, strMsg)
                    GoTo ExitThisSub
                End If
                'check for shared operation and exit out, if the selected operation is a Shared OP
                If CheckForSharedOperationFlag(objOpSumm) Then GoTo ExitThisSub

                'If EASEClass7.DBConfig.SharePlan Then

                '    'cann't release checked out op
                '    strMsg = ""
                '    For intK = 1 To UBound(objOpSumm)
                '        If objOpSumm(intK).OPRev >= 1 Then
                '            objOpSumm(intK).Selected = False
                '            If Trim(strMsg) <> "" Then strMsg &= " ,"
                '            strMsg &= Trim(objOpSumm(intK).OPNO)
                '        End If
                '    Next
                '    If Trim(strMsg) <> "" Then
                '        strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3959)
                '        MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                '        GoTo ExitThisSub
                '    End If

                '    If Ec.Parts.CheckedOutPart(gRH.PartNo) Then
                '        strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3934)
                '        MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                '        GoTo ExitThisSub
                '    End If
                'End If


                'TODO: If the operation is in a shared part or checked-out part, don't do unrelease
                If blnManualUpdate Then
                    'manual unrelase, need no authorization
                    GoTo JustUnRelease
                End If

                'ask for password to un-release the operation
                strMsg = Ec.AppConfig.GetWrd(3170)
                HttpContext.Current.Session("gAN") = strMsg : strLIN = HttpContext.Current.Session("gLIN") : HttpContext.Current.Session("gLIN") = ""

                HttpContext.Current.Session("gTypoChanges") = False
                If strCallFrom.Trim <> "opsummarypage" AndAlso EASEClass7.DBConfig.AllowTypoChange Then
                    'don't allow the typo changes, when trying to un-release/update from operation summary page. because, we can't hold
                    'the typochanges flag for multiple operations
                    HttpContext.Current.Session("gLIN") = "typochange"
                End If
                HttpContext.Current.Session("gFormParam") = "UnreleaseOp"
                HttpContext.Current.Session("ObjInMemory") = objOpSumm
                strURL = "frmPassword.aspx"
                GoTo ExitThisSub

                '!@#$ remaining covered in frmPassword.aspx
                If HttpContext.Current.Session("gLIN") = "processtypochange" Then HttpContext.Current.Session("gTypoChanges") = True
                HttpContext.Current.Session("gLIN") = strLIN  'reset the glin
                If HttpContext.Current.Session("gAN") = "*" Then GoTo ExitThisSub

                If HttpContext.Current.Session("gTypoChanges") Then
                    'write to change history
                    objCH.ID = HttpContext.Current.Session("gID")
                    objCH.OPNO = ""
                    arrCH(1) = "Typo changes"
                    objCH.Comment = arrCH
                    objCH.Reason = arrCH

                    objCH.PCN = ""
                    objCH.Rev = "1"
                    objCH.OpRev = objOpSumm(1).OPRev
                    objCH.OPNO = objOpSumm(1).OPNO

                    objCH.Engineer = HttpContext.Current.Session("UserName")
                    objCH.EntryDate =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat
                    objCH.CommentSeq = 1
                    objCH.EntrySeq = 0
                    HttpContext.Current.Session("gobjCH") = objCH
                    Ec.Parts.WriteChangeHistory(objCH, True)

                    Ec.Parts.WriteToEASELogTable(HttpContext.Current.Session("gID"), objRH.PartNo, objOpSumm(1).OPNO, HttpContext.Current.Session("UserName"), 107, "OP.Rev:" & objOpSumm(1).OPRev)       'Write to EASELog table

                    GoTo GoodJob  'don't unrelease the operation
                End If


JustUnRelease:


                intIteration = 70 / UBound(objOpSumm)
                intStatus = 30

                For intK = 1 To UBound(objOpSumm)

                    intStatus = intStatus + intIteration

                    If Not objOpSumm(intK).Selected Then GoTo SkipThisOP


                    If blnManualUpdate Then
                        GoTo SkipCopyingOPToHistory
                    End If
                    Ec.Parts.ClearCopyParamsObject(objCopyOP)

                    objCopyOP.FromID = objOpSumm(intK).ID
                    objCopyOP.FromRectype = objOpSumm(intK).Rectype
                    objCopyOP.FromSeq = objOpSumm(intK).Seq
                    objCopyOP.FromOPNO = objOpSumm(intK).OPNO


                    objCopyOP.ToID = objOpSumm(intK).ID
                    objCopyOP.ToRectype = "2"
                    objCopyOP.ToSeq = Ec.Parts.GetPartSeqMax(objCopyOP.ToID, objCopyOP.ToRectype)
                    objCopyOP.ToOPNO = objOpSumm(intK).OPNO
                    objCopyOP.ToOPEngineer = objOpSumm(intK).Engineer

                    If Not Ec.Parts.CopyPartOperation("copyoptohistory", objCopyOP, Ec.AppConfig.EASETempPath, HttpContext.Current.Session("UserName")) Then
                        'error occurred while copying the operation into history. exit out
                        GoTo SkipThisOP
                    End If
SkipCopyingOPToHistory:

                    objOPH = Ec.Parts.GetOPH(objOpSumm(intK).ID, objOpSumm(intK).Rectype, objOpSumm(intK).Seq, objOpSumm(intK).OPNO)
                    If Trim(objOPH.OPNO) = "" Then GoTo SkipThisOP 'never happens, just in case


                    'update the operation header record
                    objOPH.NewOP = False
                    objOPH.ReleasedFlag = 0      ''ReleasedStatus: 0-UnReleased, 1-Released, 2-Pending, 3-Approved, other unknown
                    objOPH.Engineer = Trim(HttpContext.Current.Session("UserName"))

                    If objOPH.Rev <= 0 Then objOPH.Rev = 1 'defensiving coding, hi hi
                    objOPH.Rev += 1

                    If blnManualUpdate Then
                        objOPH.Rev = objOpSumm(intK).OPRev
                    End If

                    'update the status in database
                    If Not Ec.Parts.WriteOPH(objOPH, False) Then GoTo SkipThisOP


                    Ec.Parts.WriteToEASELogTable(HttpContext.Current.Session("gID"), objRH.PartNo, objOPH.OPNO, HttpContext.Current.Session("UserName"), 105, "OP.Rev:" & objOPH.Rev)       'Write to EASELog table

                    ClearChangeHistoryRecord()
                    If blnManualUpdate Then
                        'write to change history
                        objCH.ID = HttpContext.Current.Session("gID")
                        objCH.OPNO = ""
                        arrCH(1) = "Manual Un-release Operation"
                        objCH.Comment = arrCH
                        objCH.Reason = arrCH

                        objCH.PCN = ""
                        objCH.Rev = "1"
                        objCH.OpRev = objOPH.Rev
                        objCH.OPNO = objOPH.OPNO

                        objCH.Engineer = HttpContext.Current.Session("UserName")
                        objCH.EntryDate =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat
                        objCH.CommentSeq = 1
                        objCH.EntrySeq = 0
                        Ec.Parts.WriteChangeHistory(objCH, True)
                    Else
                        'write to change history
                        HttpContext.Current.Session("gFormParam") = "unreleaseop"


                        objCH.ID = objOpSumm(intK).ID
                        objCH.OPNO = objOpSumm(intK).OPNO
                        objCH.Engineer = objOPH.Engineer
                        objCH.EntryDate = objOPH.ReleasedDateX
                        objCH.OpRev = objOPH.Rev
                        objCH.PCN = objOPH.PCNNO

                        strURL = "frmChangeHistoryEdit.aspx"
                        GoTo ExitThisSub

                    End If


SkipThisOP:
                Next
GoodJob:
                blnRtnValue = True

ExitThisSub:
            Catch ex As Exception
                GeneralError("UnReleaseOperations", ex, True)
            Finally
                objOPH = Nothing
                'objCopy = Nothing
            End Try
            If strURL.Trim <> "" Then
                System.Web.HttpContext.Current.Response.Redirect(strURL, True)
                'Response.Redirect(strURL)
            End If
            CheckForUserSessionErrors()
FinalExit:
            HttpContext.Current.Session("gOPH") = objTEMPOPH           'used in updating the ref docs template header


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function


        Public Function ReleaseOperations(ByVal objOpSumm() As EASEClass7.Parts.stOPSummary, Optional ByVal blnManualUpdate As Boolean = False, Optional ByVal objLabel As System.Web.UI.WebControls.Label = Nothing) As Boolean

#If TRACE Then
            Dim startTicks As Long = Log.OPERATION(String.Format("Enter blnManualUpdate:({0})", blnManualUpdate), LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            'blnManualUpdate=true only in frmupdatereleasestatus form, to manually update the part/op rev

            'Requirment: Load gid,grectype,gseq,gopsumm vars

            Dim blnRtnValue As Boolean = False
            Dim intK As Int16 = 0, strMsg As String = ""
            Dim objOPH As EASEClass7.Parts.stOPH
            Dim strTempParam As String = HttpContext.Current.Session("gFormParam")
            Dim objCopyOP As New EASEClass7.Parts.stCopyParams
            Dim objTEMPOPH As EASEClass7.Parts.stOPH = HttpContext.Current.Session("gOPH")  'used in updating the ref docs template header
            Dim intSuccessCount As Int16 = 0, intIteration As Int16 = 0, intStatus As Int16 = 0
            Dim objCH As EASEClass7.Parts.stChangeHistory, arrCH(1) As String
            Dim objOPH1 As EASEClass7.Parts.stOPH
            Dim objRH As EASEClass7.Parts.stRH
            objRH = HttpContext.Current.Session("gRH")
            Dim strURL As String = ""

            Try
                If Not VerifyUserAccessRights(True, True, , True, , , True, , objLabel) Then        'check access rights message
                    GoTo ExitThisSub
                End If

                'if no operations in the object, defensive coding, hi hi
                If UBound(objOpSumm) = 0 Then GoTo ExitThisSub

                For intK = 1 To UBound(objOpSumm)
                    If objOpSumm(intK).Selected Then
                        HttpContext.Current.Session("ObjInMemory") = objOpSumm(intK)
                    End If
                Next
                'cann't unrelease checked out op

                strMsg = ""
                If EASEClass7.DBConfig.SharePlan Then
                    For intK = 1 To UBound(objOpSumm)
                        If Trim(objOpSumm(intK).CheckoutEngineer) <> "" Then
                            objOpSumm(intK).Selected = False
                            If Trim(strMsg) <> "" Then strMsg &= " ,"
                            strMsg &= Trim(objOpSumm(intK).OPNO)
                        End If
                    Next
                    If Trim(strMsg) <> "" Then
                        strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3934)
                        'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        DisplayUserMessage(objLabel, strMsg)
                        GoTo ExitThisSub
                    End If
                End If

                'check for shared operation and exit out, if the selected operation is a Shared OP
                If CheckForSharedOperationFlag(objOpSumm) Then GoTo ExitThisSub


                'cann't release the released op
                strMsg = ""
                For intK = 1 To UBound(objOpSumm)
                    If objOpSumm(intK).ReleasedFlag <> 0 Then
                        objOpSumm(intK).Selected = False
                        If Trim(strMsg) <> "" Then strMsg &= " ,"
                        strMsg &= Trim(objOpSumm(intK).OPNO)
                    End If
                Next
                If Trim(strMsg) <> "" Then
                    strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & Ec.AppConfig.GetWrd(3940) & vbCrLf & vbCrLf & strMsg
                    'strMsg = Ec.AppConfig.GetWrd(3939) & vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & Ec.AppConfig.GetWrd(3935)
                    'MessageBox.Show(strMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DisplayUserMessage(objLabel, strMsg)
                    GoTo ExitThisSub
                End If

                'TODO: IF PCN/MCR is available, check the pcn status
                'TODO: Harley custom changes for typo changes and shared op container

                If blnManualUpdate Then         'manual update, require not password or user confirmation
                    GoTo JustReleaseOp
                End If

                'ask for password to release the operation
                strMsg = Ec.AppConfig.GetWrd(3169)
                HttpContext.Current.Session("gAN") = strMsg
                HttpContext.Current.Session("gFormParam") = "ReleaseOp"
                strURL = "frmPassword.aspx"
                GoTo ExitThisSub

                '!@#$ Remaining is covered in frmPassword.aspx
                If HttpContext.Current.Session("gAN") = "*" Then GoTo ExitThisSub

                'IF PCN is available, then need to update the pcn record too


JustReleaseOp:

                '  DisplayStatusForm("Start", 1)
                ' DisplayStatusForm("Validating User rights", 5)


                intIteration = 70 / UBound(objOpSumm)
                intStatus = 30
                For intK = 1 To UBound(objOpSumm)
                    If Not objOpSumm(intK).Selected Then GoTo SkipThisOP
                    ' DisplayStatusForm("Processing Operation: " & objOpSumm(intK).OPNO, intStatus)



                    intStatus = intStatus + intIteration

                    objOPH = Ec.Parts.GetOPH(objOpSumm(intK).ID, objOpSumm(intK).Rectype, objOpSumm(intK).Seq, objOpSumm(intK).OPNO)
                    If Trim(objOPH.OPNO) = "" Then GoTo SkipThisOP 'never happens, just in case



                    'update the operation header record
                    objOPH.NewOP = False
                    objOPH.ReleasedFlag = 1      ''ReleasedStatus: 0-UnReleased, 1-Released, 2-Pending, 3-Approved, other unknown
                    'objOPH.ReleasedDate =  EaseCore.Extensions.Dates.GetCurrentDate
                    '  objOPH.ReleasedDateX =  EaseCore.Extensions.Dates.NumToDate( EaseCore.Extensions.Dates.GetCurrentDateinNumFormat, True)
                    objOPH.ReleasedDateX =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat
                    objOPH.Engineer = Trim(HttpContext.Current.Session("UserName"))

                    If blnManualUpdate Then         'manual update, require not password or user confirmation
                        objOPH.ReleasedDateX = objOpSumm(intK).ReleasedDate '
                        objOPH.Rev = objOpSumm(intK).OPRev
                    End If

                    'TODO: Need to update the Part Status

                    'update the status in database
                    If Not Ec.Parts.WriteOPH(objOPH, False) Then GoTo SkipThisOP

                    'Convert the Work Instructions/Reference Documents to PDF/JPG

                    HttpContext.Current.Session("gOPH") = objOPH       'used in updating the template header and footer
                    ConvertOperationDocumentsToPDF(objOPH)

                    intSuccessCount += 1

                    Ec.Parts.WriteToEASELogTable(HttpContext.Current.Session("gID"), objRH.PartNo, objOPH.OPNO, HttpContext.Current.Session("UserName"), 104, "OP.Rev:" & objOPH.Rev)       'Write to EASELog table

                    objOPH1 = HttpContext.Current.Session("gOPH")
                    If blnManualUpdate Then
                        'write to change history
                        objCH.ID = HttpContext.Current.Session("gID")
                        objCH.OPNO = ""
                        arrCH(1) = "Manual Release Operation"
                        objCH.Comment = arrCH
                        objCH.Reason = arrCH

                        objCH.PCN = ""
                        objCH.Rev = "1"
                        objCH.OpRev = objOPH1.Rev
                        objCH.OPNO = objOPH1.OPNO

                        objCH.Engineer = HttpContext.Current.Session("UserName")
                        objCH.EntryDate =  EaseCore.Extensions.Dates.GetCurrentDateinNumFormat
                        objCH.CommentSeq = 1
                        objCH.EntrySeq = 0
                        HttpContext.Current.Session("gobjCH") = objCH
                        Ec.Parts.WriteChangeHistory(objCH, True)
                    End If


                    If EASEClass7.DBConfig.SharedOperation AndAlso objOpSumm(intK).SharedOPID = -99 Then
                        'ok this operation is shared more couple of places. need to update the operation header in all places
                        Ec.Parts.UpdateSharedOperationLinks(objOpSumm(intK).ID, objOpSumm(intK).Rectype, objOpSumm(intK).Seq, objOpSumm(intK).OPNO)
                    End If

SkipThisOP:
                Next
                ' DisplayStatusForm("Refresh Part", 95)
                If intSuccessCount > 0 Then RefreshPart() 'refresh gOPSumm Array, recalculate time totals and update database

                blnRtnValue = True

ExitThisSub:
            Catch ex As Exception
                GeneralError("ReleaseOperations", ex, True)
            Finally
                objOPH = Nothing
                'objCopy = Nothing
            End Try
FinalExit:
            HttpContext.Current.Session("gOPH") = objTEMPOPH           'used in updating the ref docs template header
            '  DisplayStatusForm("", 100)

            CheckForUserSessionErrors()
            If strURL.Trim <> "" Then
                System.Web.HttpContext.Current.Response.Redirect(strURL, True)
            End If


#If TRACE Then
            Log.OPERATION(String.Format("Exit ({0})", blnRtnValue), LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If
            Return blnRtnValue
        End Function

        Public Sub FillElementValueTypeCombo(ByVal objCBO As DropDownList)

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            objCBO.Items.Clear()


            'Element Type-> 0,-Direct Input, 1-Element, 2-Formula, 3-Table


            ' objCBO.BeginUpdate()
            'don't change the order of the combo item
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(926), 0)       'direct input
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(1805), 1)      'Element
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(1851), 2)      'formula
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(1852), 3)      'table
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(134) & " " & Ec.AppConfig.GetWrd(1805), 4)       'all elements
            '  AddComboItem(objCBO, Ec.AppConfig.GetWrd(282), 4)       'all elements

            '  objCBO.EndUpdate()
            ' objCBO.DropDownStyle = ComboBoxStyle.DropDownList
            objCBO.SelectedIndex = 4

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

        Public Sub FillDocTypeOptionCombo(ByVal objCBO As DropDownList)

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            objCBO.Items.Clear()
            '  objCBO.BeginUpdate()

            'Condition: 0-equal to, 1 less than or equal to, 2-greater than or equal to,3-Range
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(2740), 0)          'ref docs
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(2741), 1)          'cad    
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(2742), 2)          'video

            '  objCBO.EndUpdate()
            ' objCBO.DropDownStyle = ComboBoxStyle.DropDownList
            objCBO.SelectedIndex = 0

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub
        Public Sub FillSaveOptionOptionCombo(ByVal objCBO As DropDownList)

#If TRACE Then
            Dim startTicks As Long = Log.UI_CONTROL_MED("Enter", LOG_APPNAME, BASE_ERRORNUMBER + 0)
#End If

            objCBO.Items.Clear()
            ' objCBO.BeginUpdate()

            'Condition: 0-equal to, 1 less than or equal to, 2-greater than or equal to,3-Range
            AddComboItem(objCBO, Ec.AppConfig.GetWrd(3815), 0)      'unique doc
            If Ec.AppConfig.MDM Then
                AddComboItem(objCBO, Ec.AppConfig.GetWrd(3802), 1)      'shared doc
            End If


            '  objCBO.EndUpdate()
            '  objCBO.DropDownStyle = ComboBoxStyle.DropDownList
            objCBO.SelectedIndex = 0

            objCBO.Enabled = False
            If Ec.AppConfig.MDM AndAlso EASEClass7.DBConfig.MDMPartDocs Then
                objCBO.Enabled = True
            End If

#If TRACE Then
            Log.UI_CONTROL_MED("Exit", LOG_APPNAME, BASE_ERRORNUMBER + 0, startTicks)
#End If

        End Sub

    End Module

End Namespace
