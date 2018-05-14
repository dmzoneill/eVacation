<%
    '****************************************
    'Collections:
    '============
    '   cColEmployees
    '   cColLeavePeriods
    '   cColLeaveTypes
    '   cColPublicHolidays

    'Objects:
    '========   
    '   cObjAnnualVacation
    '   cObjCarryOver
    '   cObjCollection
    '   cObjELPInstance    
    '   cObjELPRelief
    '   cObjCompTime        
    '   cObjLeavePeriod
    '   cObjLeaveType
    '   cObjOtherLeave
    '   cObjPublicHoliday
    '   cObjUser


    '****************************************


    '*** COLLECTION EMPLOYEES ***
    Class cColEmployees
        '====================================
        'Properties:                    Type:                       Perm:       Source:     Properties Required:
        'y  EE                          obj(User)                   RW          Virtual
        'y  CollectionType              lng                         RW          Virtual
        'y  SearchName                  str                         RW          Virtual
        'y  MaxSearchResultsExceeded    bln                         R           Virtual
        'y  MaxSearchResults            lng                         RW          Virtual
        '   (all other collection properties as per base collection object)
        '
        'Methods:
        '   Add
        '   Initialise_Collection
        '   DBLoad
        '
        '====================================
        Private m_objCollection
        Private m_objEE
        Private m_lngCollectionType
        Private m_strSearchName
        Private m_blnMaxSearchResultsExceeded
        Private m_lngMaxSearchResults
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cColEmployees (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            SetNotLoaded
        End Sub
        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cColEmployees (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            Set m_objCollection = Nothing
        End Sub


        Private Sub SetEmpty
            m_strSearchName = ""
            m_blnMaxSearchResultsExceeded = False
            m_lngMaxSearchResults = 100
            Set m_objEE = nothing
            If typename(m_objCollection) = "cObjCollection" then
                m_objCollection.Clear
            End If
        End Sub
        

        Private Sub SetNotLoaded
            m_blnLoaded = False
        End Sub
        
        
        Public Property Set EE(ByRef objUser)
            SetEmpty
            SetNotLoaded
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID         
        End Property
        
        Public Property Get EE()
            Set EE = m_objEE
        End Property

        
        Public Property Let CollectionType(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            m_lngCollectionType = lngValue
        End Property
        Public Property Get CollectionType
            CollectionType = m_lngCollectionType
        End Property


        Public Property Let SearchName(ByVal strValue)
            SetEmpty
            SetNotLoaded
            m_strSearchName = strValue
            if instr(m_strSearchName," ") > 0 then
                m_strSearchName = replace(m_strSearchName," ","")
            end if
        End Property
        Public Property Get SearchName()
            SearchName = m_strSearchName
        End Property
        

        Public Property Get MaxSearchResultsExceeded()
            DBLoad
            MaxSearchResultsExceeded = m_blnMaxSearchResultsExceeded
        End Property
        

        Public Property Let MaxSearchResults(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            if lngValue = 0 then
                exit Property
            End If
            m_lngMaxSearchResults = lngValue
        End Property
        Public Property Get MaxSearchResults()
            MaxSearchResults = m_lngMaxSearchResults
        End Property
        
        
        Public Property Get Count
            DBLoad
            If typename(m_objCollection) <> "cObjCollection" then
                Count = 0
            Else
                Count = m_objCollection.Count
            End If
        End Property


        Public Sub FreeFromMemory(ByVal loclngIndex)
            If typename(m_objCollection) = "cObjCollection" then
                m_objCollection.FreeFromMemory(loclngIndex)
            End If
        End Sub
        
            
        Public Default Property Get Item(ByVal loclngIndex)
            DBLoad
            Set Item = m_objCollection(loclngIndex)
        End Property
        

        Public Sub Add(ByVal locObject)
            Initialise_Collection
            m_blnLoaded = True
            m_objCollection.Add(locObject)
        End Sub


        Private Sub Initialise_Collection()
            If typename(m_objCollection) <> "cObjCollection" then
                Set m_objCollection = new cObjCollection
            End If
        End Sub
                

        Private Sub DBLoad
            Dim locCmd
            Dim locCmdRevoke
            Dim locParam
            Dim locRS
            Dim blnAddEmployee
            Dim loclngCounter
            Dim newloclngCounter
            
            Dim fldstrWWID
            Dim fldstrIDSID
            Dim fldstrFirstNm
            Dim fldstrLastNm
            Dim fldstrNextLevelWWID
            Dim fldstrEmail
            Dim fldstrCompanyCd
            Dim fldstrMailStopTxt
            Dim flddatLDOH
            Dim flddatODOH
            Dim fldblnIsExempt
            Dim fldblnIsBlueBadge
            Dim fldblnIsPartTimer
            Dim flddatTerminationDate
            Dim fldstrActiveStatus
            Dim fldblnIsEELeaveTracked
            Dim flddatDOB
            Dim fldendDate 
            Dim fldblnIsAdmin
            Dim fldblnIsExemptStatusChanged
            Dim fldblnIsException
            Dim fldstrExceptionComments
            Dim fldstrActiveDelegateWWID
            
            Dim locobjEmployee
            
            If m_blnLoaded then
                Exit Sub
            End If


            '*** CHECK TO SEE - have we got enough info for the type of employees collection we are loading?
            Select Case m_lngCollectionType
                Case CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS
                    '*** If we have no WWID set up, we can't limit our employees - so exit the routine.
                    If EE.WWID = "" then
                        Exit Sub
                    End If
                Case CONST_EMPLOYEE_COLLECTION_TYPE_NAME_SEARCH
                    '*** If we have no SearchName set up we can't limit our employees - so exit the routine.
                    If SearchName = "" then
                        Exit Sub
                    End If
                Case CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT
                    '*** All ok - nothing needed for this one.
                Case Else
                    Exit Sub
            End Select

            m_blnLoaded = True
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            '*** SET UP THE STORED PROCEDURE - depending on the type of leave types collection we are loading.
            mDebugPrint "debug " &  Abs(m_lngCollectionType) & "<br>"
           
          
                if Abs(m_lngCollectionType) = CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS then
                    mDebugPrint "<br>debug " &  CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS
                    locCmd.CommandText = "usp_eedirectreports"
                    mDebugPrint locCmd.CommandText
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                elseif Abs(m_lngCollectionType)  = CONST_EMPLOYEE_COLLECTION_TYPE_NAME_SEARCH then
                    mDebugPrint "<br>debug " &  CONST_EMPLOYEE_COLLECTION_TYPE_NAME_SEARCH
                    locCmd.CommandText = "usp_eesearchbyname"
                    mDebugPrint locCmd.CommandText
                    locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
                    locCmd.Parameters.Append locCmd.CreateParameter("strName", adWChar, adParamInput, 50, SearchName)
                    locCmd.Parameters.Append locCmd.CreateParameter("lngMaxResults", adInteger, adParamInput, , MaxSearchResults)
                elseif Abs(m_lngCollectionType) =  CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT then
                    locCmd.CommandText = "usp_ee_payrollreport"
                    mDebugPrint "<br>debug " &  CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT
                    mDebugPrint locCmd.CommandText
                Else
                    Set locCmd = nothing
                    m_blnLoaded = False
                    Exit Sub
            End if
            
            Set locRS = locCmd.Execute

            Set fldstrWWID = locRS("WWID")
            Set fldstrFirstNm = locRS("FirstNm")
            Set fldstrLastNm = locRS("LastNm")
            Set fldstrMailStopTxt = locRS("MailStopTxt")
           
            If m_lngCollectionType = CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS or _
                m_lngCollectionType = CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT then
                Set fldstrIDSID = locRS("Idsid")
                Set fldstrNextLevelWWID = locRS("NextLevelWWID")
                Set fldstrCompanyCd = locRS("CompanyCd")
                Set fldstrEmail = locRS("CorporateEmailTxt")
                Set flddatLDOH = locRS("StartDt")
                Set flddatODOH = locRS("OriginalStartDt")
                Set fldblnIsExempt = locRS("FLSACd")
                Set fldblnIsBlueBadge = locRS("blnBlueBadge")
                Set fldblnIsPartTimer = locRS("FullTmPartTmCd")
                Set flddatTerminationDate = locRS("EndDt")
                Set fldstrActiveStatus = locRS("EmployeeStatusCd")
                Set fldblnIsEELeaveTracked = locRS("blnIsEELeaveTracked")
                Set flddatDOB = locRS("datDOB")
                Set fldendDate = locRS("endDate")
                Set fldblnIsAdmin = locRS("blnIsAdmin")
                Set fldblnIsExemptStatusChanged = locRS("blnIsExemptStatusChanged")
                Set fldblnIsException = locRS("blnIsException")
                Set fldstrExceptionComments = locRS("strExceptionComments")
                Set fldstrActiveDelegateWWID = locRS("strDelegateWWID")
            End If
            
            loclngCounter = 0
            
            m_blnMaxSearchResultsExceeded = False

            If not locRS.eof then
                Initialise_Collection
            End If
            
            While not locRS.eof and (not m_blnMaxSearchResultsExceeded)
                '*** FOLLOWING LINE FOR TESTING PURPOSES ONLY
                'If (not m_lngCollectionType = CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT) or _
                '   (m_lngCollectionType = CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT and _
                '   (fldstrWWID = "10778445" or fldstrWWID="10637768")) then
                    
                    Set locobjEmployee = new cObjUser
                    locobjEmployee.WWID = fldstrWWID
                    locobjEmployee.FirstNm = fldstrFirstNm
                    locobjEmployee.LastNm = fldstrLastNm
                    locobjEmployee.MailStopTxt = fldstrMailStopTxt
                    If m_lngCollectionType= CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS or _
                        m_lngCollectionType = CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT then
                        locobjEmployee.IDSID = fldstrIDSID
                        locobjEmployee.ManagerWWID = fldstrNextLevelWWID
                        locobjEmployee.CompanyCd = fldstrCompanyCd
                        locobjEmployee.Email = fldstrEmail
                        locobjEmployee.LDOH = flddatLDOH
                        locobjEmployee.ODOH = flddatODOH
                        locobjEmployee.IsExempt = mIf(fldblnIsExempt="E", True, False)
                        locobjEmployee.IsBlueBadge = mIf(fldblnIsBlueBadge=True, True, False)
                        locobjEmployee.IsPartTimer = mIf(fldblnIsPartTimer="PT", True, False)
                        locobjEmployee.TerminationDate = flddatTerminationDate
                        locobjEmployee.ActiveStatus = fldstrActiveStatus
                        locobjEmployee.IsEELeaveTracked = fldblnIsEELeaveTracked
                        mDebugPrint " locobjEmployee.IsEELeaveTracked " & locobjEmployee.IsEELeaveTracked
                        locobjEmployee.DOB = flddatDOB
                        locobjEmployee.endDate = fldendDate
                        locobjEmployee.IsAdmin = mIf(fldblnIsAdmin=True, True, False)
                        locobjEmployee.IsExemptStatusChanged = mIf(fldblnIsExemptStatusChanged=True, True, False)
                        locobjEmployee.IsException = mIf(fldblnIsException=True, True, False)
                        locobjEmployee.ExceptionComments = fldstrExceptionComments
                        locobjEmployee.ActiveDelegateWWID = fldstrActiveDelegateWWID
                    end if

                    blnAddEmployee = True
                
                    Select Case CollectionType
                        Case CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS
                            If not locobjEmployee.ActiveThisYear then
                                blnAddEmployee = False
                            End If
                        Case CONST_EMPLOYEE_COLLECTION_TYPE_NAME_SEARCH
                            loclngCounter = loclngCounter + 1
                            if loclngCounter > MaxSearchResults then
                                m_blnMaxSearchResultsExceeded = True
                                blnAddEmployee = False
                            End If
                    End Select

                    If blnAddEmployee then
                        Add locobjEmployee
                        mDebugPrint "   - Added new employee object to collection: " & fldstrFirstNm &  " " & fldstrLastNm & ".<br>"
                    End If
                    
                    Set locobjEmployee = nothing
                '*** FOLLOWING LINE FOR TESTING PURPOSES ONLY
                'End If
                locRS.movenext
            Wend

            Set fldstrWWID = nothing
            Set fldstrIDSID = nothing
            Set fldstrFirstNm = nothing
            Set fldstrLastNm = nothing
            Set fldstrNextLevelWWID = nothing
            Set fldstrEmail = nothing
            Set fldstrCompanyCd = nothing
            Set fldstrMailStopTxt = nothing
            Set flddatLDOH = nothing
            Set flddatODOH = nothing
            Set fldblnIsExempt = nothing
            Set fldblnIsBlueBadge = nothing
            Set fldblnIsPartTimer = nothing
            Set flddatTerminationDate = nothing
            Set fldstrActiveStatus = nothing
            Set fldblnIsEELeaveTracked = nothing
            Set flddatDOB = nothing
            Set fldblnIsAdmin = nothing
            Set fldblnIsExemptStatusChanged = nothing
            Set fldblnIsException = nothing
            Set fldstrExceptionComments = nothing
            Set fldstrActiveDelegateWWID = nothing

            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing

        End Sub
                    
    End Class
	
	
	'*** COLLECTION EMPLOYEES BASIC ***
    Class cColEmployeesAll
        Private m_objCollection
        Private m_objEE
        Private m_lngCollectionType
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cColEmployees (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            SetNotLoaded
        End Sub
        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cColEmployees (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            Set m_objCollection = Nothing
        End Sub


        Private Sub SetEmpty
            Set m_objEE = nothing
            If typename(m_objCollection) = "cObjCollection" then
                m_objCollection.Clear
            End If
        End Sub
        

        Private Sub SetNotLoaded
            m_blnLoaded = False
        End Sub
        
        
        Public Property Set EE(ByRef objUser)
            SetEmpty
            SetNotLoaded
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID         
        End Property
        
        Public Property Get EE()
            Set EE = m_objEE
        End Property

        
        Public Property Let CollectionType(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            m_lngCollectionType = lngValue
        End Property
        Public Property Get CollectionType
            CollectionType = m_lngCollectionType
        End Property
               
        
        Public Property Get Count
            DBLoad
            If typename(m_objCollection) <> "cObjCollection" then
               mDebugPrint "it is 0"  
               Count = 0
            Else
                mDebugPrint "it is NOT 0"
                Count = m_objCollection.Count
            End If
        End Property


        Public Sub FreeFromMemory(ByVal loclngIndex)
            If typename(m_objCollection) = "cObjCollection" then
                m_objCollection.FreeFromMemory(loclngIndex)
            End If
        End Sub
        
            
        Public Default Property Get Item(ByVal loclngIndex)
            DBLoad
            Set Item = m_objCollection(loclngIndex)
        End Property
        

        Public Sub Add(ByVal locObject)
            Initialise_Collection
            m_blnLoaded = True
            m_objCollection.Add(locObject)
        End Sub


        Private Sub Initialise_Collection()
            If typename(m_objCollection) <> "cObjCollection" then
                Set m_objCollection = new cObjCollection
            End If
        End Sub
                

        Private Sub DBLoad
            Dim locCmd
            Dim locCmdRevoke
            Dim locParam
            Dim locRS
            Dim blnAddEmployee
            Dim loclngCounter
            Dim newloclngCounter
            
            Dim fldstrWWID
            Dim fldstrIDSID
            Dim fldstrFirstNm
            Dim fldstrLastNm
            Dim fldstrNextLevelWWID
            Dim fldstrEmail
            Dim fldstrCompanyCd
            Dim fldstrMailStopTxt
            Dim flddatLDOH
            Dim flddatODOH
            Dim fldblnIsExempt
            Dim fldblnIsBlueBadge
            Dim fldblnIsPartTimer
            Dim flddatTerminationDate
            Dim fldstrActiveStatus
            Dim fldblnIsEELeaveTracked
            Dim flddatDOB
            Dim fldendDate 
            Dim fldblnIsAdmin
            Dim fldblnIsExemptStatusChanged
            Dim fldblnIsException
            Dim fldstrExceptionComments
            Dim fldstrActiveDelegateWWID
            
            Dim locobjEmployee
            
            If m_blnLoaded then
                Exit Sub
            End If

            m_blnLoaded = True
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc
			locCmd.CommandText = "usp_eeshowall"
                                 
            Set locRS = locCmd.Execute

            Set fldstrWWID = locRS("WWID")
            Set fldstrFirstNm = locRS("FirstNm")
            Set fldstrLastNm = locRS("LastNm")
            Set fldstrMailStopTxt = locRS("MailStopTxt")
			           
            loclngCounter = 0

            If not locRS.eof then
                Initialise_Collection
            End If
            
            While not locRS.eof
                    
                    Set locobjEmployee = new cObjUser
                    locobjEmployee.WWID = Trim(fldstrWWID)
                    locobjEmployee.FirstNm = Trim(fldstrFirstNm)
                    locobjEmployee.LastNm = Trim(fldstrLastNm)
                    locobjEmployee.MailStopTxt = fldstrMailStopTxt					      

                    Add locobjEmployee
                    
                    Set locobjEmployee = nothing
                '*** FOLLOWING LINE FOR TESTING PURPOSES ONLY
                'End If
                locRS.movenext
            Wend

            Set fldstrWWID = nothing
            Set fldstrIDSID = nothing
            Set fldstrFirstNm = nothing
            Set fldstrLastNm = nothing
            Set fldstrNextLevelWWID = nothing
            Set fldstrEmail = nothing
            Set fldstrCompanyCd = nothing
            Set fldstrMailStopTxt = nothing
            Set flddatLDOH = nothing
            Set flddatODOH = nothing
            Set fldblnIsExempt = nothing
            Set fldblnIsBlueBadge = nothing
            Set fldblnIsPartTimer = nothing
            Set flddatTerminationDate = nothing
            Set fldstrActiveStatus = nothing
            Set fldblnIsEELeaveTracked = nothing
            Set flddatDOB = nothing
            Set fldblnIsAdmin = nothing
            Set fldblnIsExemptStatusChanged = nothing
            Set fldblnIsException = nothing
            Set fldstrExceptionComments = nothing
            Set fldstrActiveDelegateWWID = nothing

            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing

        End Sub
                    
    End Class


    '*** COLLECTION LEAVE PERIODS ***
    Class cColLeavePeriods
        '====================================
        'Properties:                    Type:                       Perm:       Source:     Properties Required:
        'y  EE                          obj(User)                   RW          Virtual
        'y  CollectionType              lng                         RW          Virtual
        '   CollectionLeaveType         str                         RW          Virtual
        '   CompareLeavePeriod          obj(LeavePeriod)            RW          Virtual
        'y  LeaveDaysInPeriod()         dbl                         R           Calc
        '   (all other collection properties as per base collection object)
        '
        'Methods:
        '   Add
        '   DBLoad
        '
        '====================================
        Private m_objCollection
        Private m_objEE
        Private m_lngCollectionType
        Private m_strCollectionLeaveType
        Private m_objCompareLeavePeriod
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cColLeavePeriods (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            m_lngCollectionType = 0
            SetEmpty
            SetNotLoaded
        End Sub
        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cColLeavePeriods (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            Set m_objCompareLeavePeriod = nothing
            Set m_objCollection = Nothing
        End Sub


        Private Sub SetEmpty
            m_strCollectionLeaveType = ""
            Set m_objCompareLeavePeriod = nothing
            Set m_objCollection = nothing
        End Sub
        

        Private Sub SetNotLoaded
            m_blnLoaded = False
        End Sub
        
        
        Public Property Set EE(ByRef objUser)
            SetEmpty
            SetNotLoaded
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID
            m_objEE.YearToView = objUser.YearToView
        End Property
        Public Property Get EE()
            Set EE = m_objEE
        End Property

        
        Public Property Let CollectionType(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            If lngValue <> 0 and lngValue <> m_lngCollectionType then
                m_lngCollectionType = lngValue
                SetEmpty
                SetNotLoaded
            End If
        End Property
        
        Public Property Get CollectionType
            CollectionType = m_lngCollectionType
        End Property


        Public Property Let CollectionLeaveType(ByVal strValue)
            If strValue <> "" and strValue <> m_strCollectionLeaveType then
                SetEmpty
                SetNotLoaded
                m_strCollectionLeaveType = strValue
            End If
        End Property
        Public Property Get CollectionLeaveType
            CollectionLeaveType = m_strCollectionLeaveType
        End Property
        
        
        Public Property Set CompareLeavePeriod(ByRef objLeavePeriod)
            Set m_objCompareLeavePeriod = objLeavePeriod
        End Property
        Public Property Get CompareLeavePeriod()
            Set CompareLeavePeriod = m_objCompareLeavePeriod
        End Property
        
        
        Public Property Get LeaveDaysInPeriod(ByVal datStart, ByVal datEnd, ByVal blnOnlyAffectingLegalAllowance, ByVal blnOnlyActivePeriods)
            Dim locdblDays
            Dim loclngIndex
            locdblDays = 0
            loclngIndex = 0
                        
            While loclngIndex < Count       'Forces DBLoad
                loclngIndex = loclngIndex + 1
                
                If (not blnOnlyAffectingLegalAllowance) or _
                    (blnOnlyAffectingLegalAllowance and Item(loclngIndex).StopsLegalAdjAccrual) then
                    
                    'So it doesn't pick up cancelled leave periods
                    If blnOnlyActivePeriods then
                        If (not (Item(loclngIndex).Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED or Item(loclngIndex).Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED or Item(loclngIndex).Status = CONST_LEAVE_PERIOD_STATUS_REJECTED)) then
                            locdblDays = locdblDays + Item(loclngIndex).LeaveDaysInPeriod(datStart, datEnd)
                        End If
                    Else
                        '******** CA - Keep for debuggin only *****
     '                   Response.Write "obj 510 datStart:" & datStart & ", " & "datEnd:" & datEnd & "<BR>" 'CA Keep for debug
                        '***********
                        locdblDays = locdblDays + Item(loclngIndex).LeaveDaysInPeriod(datStart, datEnd)
                    End If
                End If
            Wend
            LeaveDaysInPeriod = locdblDays
        End Property
                
        
        Public Property Get Count
            DBLoad
			
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - begin Count" ) 
			'Response.Write( "<br>" ) 
			
            If typename(m_objCollection) <> "cObjCollection" then
                Count = 0
            Else
                Count = m_objCollection.Count
            End If
		
			'endtime = timer()
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - end Count" ) 
			'Response.Write( "<br>" ) 
        End Property


        Public Default Property Get Item(ByVal loclngIndex)
            DBLoad
            Initialise_Collection
            Set Item = m_objCollection(loclngIndex)
        End Property
        

        Public Sub Add(ByVal locObject)
            m_blnLoaded = True
            Initialise_Collection
            m_objCollection.Add(locObject)
        End Sub
        

        Private Sub DBLoad
            Dim locCmd
            Dim locCmdRevoke
            Dim locParam
            Dim locRS
            Dim locblnAddObject
            
            Dim fldlngID
            Dim fldstrEEWWID
            Dim fldstrApproverWWID
            Dim fldintLeaveTypeID
            Dim flddatStartDate
            Dim fldstrStartTime
            Dim flddatEndDate
            Dim fldstrEndTime
            Dim flddatDateRaised
            Dim flddatDateApproved
            Dim flddatDateRejected
            Dim flddatDateCancelRequested
            Dim flddatDateCancelApproved
            Dim flddatDateCancelRejected
            Dim fldstrResponseComments
            Dim fldstrRequestComments
            Dim fldlngELPID
            
            Dim locobjLeavePeriod
			
			mDebugPrint "In DBLoad of leave requests collection, EE=" & EE.WWID & "<br>"
            
            If m_blnLoaded then
                Exit Sub
            End If


            '*** CHECK TO SEE - have we got enough info for the type of leave period collection we are loading?
            Select Case m_lngCollectionType
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_REQUESTS, _
                     CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_APPROVALS, _
                     CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE, _
                     CONST_LEAVE_PERIOD_COLLECTION_TYPE_OTHER_LEAVE, _
                     CONST_LEAVE_PERIOD_COLLECTION_TYPE_ADMIN_VIEW
                    '*** If we have no WWID set up, we can't find our leave requests - so exit the routine.
                    If EE.WWID = "" then
                        Exit Sub
                    End If
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_ALL_EE_LEAVE_IN_PERIOD
                    '**** We need a WWID and we need a CompareLeavePeriod.StartDate and CompareLeavePeriod.EndDate
                    If EE.WWID = "" then
                        Exit Sub
                    End If
                    If (not isdate(CompareLeavePeriod.StartDate)) or (not isdate(CompareLeavePeriod.EndDate)) then
                        Exit Sub
                    End If
                Case Else
                    Exit Sub
            End Select

            m_blnLoaded = True
            
			mDebugPrint "We have enough info to load...<br>"
			
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            '*** SET UP THE STORED PROCEDURE - depending on the type of leave period collection we are loading.
            
			mDebugPrint "WWID=" & EE.WWID & ", YearToView=" & EE.YearToView & "<br>"
						
			'  response.write "614 Obj  " & locobjLeavePeriod  & " - " &      "WWID=" & EE.WWID & ", YearToView=" & EE.YearToView & "<br>"  
						
			if EE.YearToView < "1901" or EE.YearToView > "2078" then
                response.redirect CONST_APPLICATION_PATH & "/default.asp"
            end if
    
            Select Case m_lngCollectionType
                                
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_REQUESTS
                    locCmd.CommandText = "usp_eeleaverequests"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datStartDate", adDBTimeStamp, adParamInput, , mFirstDayOfYear(EE.YearToView))
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datEndDate", adDBTimeStamp, adParamInput, , mLastDayOfYear(EE.YearToView))
                    locCmd.Parameters.Append locParam
					mDebugPrint "using usp_eeleaverequests command<br>" 
					
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_APPROVALS
                    locCmd.CommandText = "usp_eeapprovalspending"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
    
				Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE
                    locCmd.CommandText = "usp_eeleaveperiods_by_type_for_year"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("strLeaveTypeName", adWChar, adParamInput, 30, CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("strYear", adInteger, adParamInput, , EE.YearToView)
                    locCmd.Parameters.Append locParam    
     
                 
                  Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE
                    locCmd.CommandText = "usp_eeleaveperiods_by_type_for_year"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("strLeaveTypeName", adWChar, adParamInput, 30, CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datStartDate", adDBTimeStamp, adParamInput, , mFirstDayOfYear(EE.YearToView))
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datEndDate", adDBTimeStamp, adParamInput, , mLastDayOfYear(EE.YearToView))
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("strYear", adInteger, adParamInput, , EE.YearToView)
                    locCmd.Parameters.Append locParam
                    
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE
                    locCmd.CommandText = "usp_eeleaveperiods_by_type_for_year2"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("strLeaveTypeName", adWChar, adParamInput, 30, CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datStartDate", adDBTimeStamp, adParamInput, , EE.StartDate)
                    locCmd.Parameters.Append locParam   
                    Set locParam = locCmd.CreateParameter("datEndDate", adDBTimeStamp, adParamInput, , EE.EndDate)
                    locCmd.Parameters.Append locParam   
                     Set locParam = locCmd.CreateParameter("strYear", adInteger, adParamInput, 8, EE.YearToView)
                    locCmd.Parameters.Append locParam
                    
                    
                    Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE
                    locCmd.CommandText = "usp_save_new_leave_request_standard_Comp"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("strLeaveTypeName", adWChar, adParamInput, 30, CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datStartDate", adDBTimeStamp, adParamInput, , EE.StartDate)
                    locCmd.Parameters.Append locParam   
                    Set locParam = locCmd.CreateParameter("datEndDate", adDBTimeStamp, adParamInput, , EE.EndDate)
                    locCmd.Parameters.Append locParam   
                    Set locParam = locCmd.CreateParameter("strYear", adInteger, adParamInput, 8, EE.YearToView)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("lngCompTimeID", adInteger, adParamInput, 4, EE.ComptimeID)
                    locCmd.Parameters.Append locParam
                    
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_ALL_EE_LEAVE_IN_PERIOD
                    locCmd.CommandText = "usp_eeleaverequests_overlapping"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datStartDate", adDBTimeStamp, adParamInput, , CompareLeavePeriod.StartDate)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datEndDate", adDBTimeStamp, adParamInput, , CompareLeavePeriod.EndDate)
                    locCmd.Parameters.Append locParam
                    
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_OTHER_LEAVE
                    locCmd.CommandText = "usp_eeleaveperiods_by_type_for_year2"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("strLeaveTypeName", adWChar, adParamInput, 30, CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION)
                    locCmd.Parameters.Append locParam
                    Set locParam = locCmd.CreateParameter("datStartDate", adDBTimeStamp, adParamInput, , EE.StartDate)
                    locCmd.Parameters.Append locParam   
                    Set locParam = locCmd.CreateParameter("datEndDate", adDBTimeStamp, adParamInput, , EE.EndDate)
                    locCmd.Parameters.Append locParam   
                    Set locParam = locCmd.CreateParameter("strYear", adInteger, adParamInput, 8, EE.YearToView)
                    locCmd.Parameters.Append locParam
                    
                Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_ADMIN_VIEW
                    locCmd.CommandText = "usp_eeleaverequests_adminview"
                    Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
                    locCmd.Parameters.Append locParam
                    
            End Select      
            Set locRS = locCmd.Execute
			
			Dim counter
			counter = 0

            Set fldlngID = locRS("lngID")
            Set fldstrEEWWID = locRS("strEEWWID")
            Set fldstrApproverWWID = locRS("strApproverWWID")
            Set fldintLeaveTypeID = locRS("lngLeaveTypeID")
            
   '  response.write "obj  680  ID  " & CONST_LEAVE_TYPE_NAME_COMP_TIME &  "</br>"
     
            Set flddatStartDate = locRS("datStartDate")
 
            Set fldstrStartTime = locRS("strStartTime")
            Set flddatEndDate = locRS("datEndDate")
            Set fldstrEndTime = locRS("strEndTime")
            Set flddatDateRaised = locRS("datRaised")
            Set flddatDateApproved = locRS("datApproved")
            Set flddatDateRejected = locRS("datRejected")
            Set flddatDateCancelRequested = locRS("datCancelRequested")
            Set flddatDateCancelApproved = locRS("datCancelApproved")
            Set flddatDateCancelRejected = locRS("datCancelRejected")
            Set fldstrResponseComments = locRS("strResponseComments")
            Set fldstrRequestComments = locRS("strRequestComments")
            Set fldlngELPID = locRS("lngELPID")
           
			mDebugPrint "executed query to get all leave periods...<br>"
			
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - DBLoad aa" ) 
			'Response.Write( "<br>" ) 
			
            While not locRS.eof
			    counter = counter + 1
                Set locobjLeavePeriod = new cObjLeavePeriod
                locobjLeavePeriod.ID = fldlngID
                Select Case m_lngCollectionType
                    Case CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_APPROVALS
                        locobjLeavePeriod.EE.WWID = trim(fldstrEEWWID)
                    Case Else
                        Set locobjLeavePeriod.EE = EE
                End Select
				
				'endtime = timer()
				'benchmark = endtime - starttime
				'Response.Write( benchmark ) 
				'Response.Write( " after select" ) 
				'Response.Write( "<br>" ) 
				
                locobjLeavePeriod.Approver.WWID = trim(fldstrApproverWWID)
                locobjLeavePeriod.LeaveType.ID = fldintLeaveTypeID
                locobjLeavePeriod.StartDate = flddatStartDate
                locobjLeavePeriod.StartTime = fldstrStartTime
                locobjLeavePeriod.EndDate = flddatEndDate
                locobjLeavePeriod.EndTime = fldstrEndTime
                locobjLeavePeriod.DateRaised = flddatDateRaised
                locobjLeavePeriod.DateApproved = flddatDateApproved
                locobjLeavePeriod.DateRejected = flddatDateRejected
                locobjLeavePeriod.DateCancelRequested = flddatDateCancelRequested
                locobjLeavePeriod.DateCancelApproved = flddatDateCancelApproved
                locobjLeavePeriod.DateCancelRejected = flddatDateCancelRejected
                locobjLeavePeriod.ResponseComments = fldstrResponseComments
                locobjLeavePeriod.RequestComments = fldstrRequestComments
                locobjLeavePeriod.ELPID = fldlngELPID
				
				'endtime = timer()
				'benchmark = endtime - starttime
				'Response.Write( benchmark ) 
				'Response.Write( " after cObjLeavePeriod popluation" ) 
				'Response.Write( "<br>" ) 

                locblnAddObject = True
                
                If m_lngCollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_ALL_EE_LEAVE_IN_PERIOD then
                    '***Check if leave period is active.
                    If not locobjLeavePeriod.IsActive then
                        locblnAddObject = False
                    '*** We know this leave period overlaps somewhere based on dates, but
                    '*** now we need to check whether it REALLY overlaps - one could end
                    '*** in the morning while the next starts in the afternoon.
                    ElseIf CDate(CompareLeavePeriod.StartDate) = CDate(locobjLeavePeriod.EndDate) And _
                        CompareLeavePeriod.StartTime = "PM" And _
                        locobjLeavePeriod.EndTime = "AM" then
                        locblnAddObject = False
                    ElseIf CDate(CompareLeavePeriod.EndDate) = CDate(locobjLeavePeriod.StartDate) And _
                        CompareLeavePeriod.EndTime = "AM" And _
                        locobjLeavePeriod.StartTime = "PM" then
                        locblnAddObject = False
                    End If
                End If

                if m_lngCollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE then
                    if not locobjLeavePeriod.IsActive then locblnAddObject = False
                end if
                
                if m_lngCollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_APPROVALS then
                    if not(locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_RAISED or _
                        locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED) then
                        locblnAddObject = False
                    End if
                end if

				mDebugPrint "Adding period (" & locblnAddObject & ")...<br>"
				
                if locblnAddObject then             
                    Add locobjLeavePeriod
                end if
                Set locobjLeavePeriod = nothing
                locRS.movenext
				
				'endtime = timer()
				'benchmark = endtime - starttime
				'Response.Write( benchmark ) 
				'Response.Write( " - counter " & counter & " iteration" ) 
				'Response.Write( "<br>" ) 
				
            Wend
						
            Set fldlngID = nothing
            Set fldstrEEWWID = nothing
            Set fldstrApproverWWID = nothing
            Set fldintLeaveTypeID = nothing
            Set flddatStartDate = nothing
            Set fldstrStartTime = nothing
            Set flddatEndDate = nothing
            Set fldstrEndTime = nothing
            Set flddatDateRaised = nothing
            Set flddatDateApproved = nothing
            Set flddatDateRejected = nothing
            Set flddatDateCancelRequested = nothing
            Set flddatDateCancelApproved = nothing
            Set flddatDateCancelRejected = nothing
            Set fldstrResponseComments = nothing
            Set fldstrRequestComments = nothing
            Set fldlngELPID = nothing

            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
			'endtime = timer()
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - DBLoad end" ) 
			'Response.Write( "<br>" ) 

        End Sub
        
        
        '*** INITIALISE COLLECTION ***  
        Private Sub Initialise_Collection
            If typename(m_objCollection) <> "cObjCollection" then
                Set m_objCollection = new cObjCollection
            End If
        End Sub     
        
    End Class
    
        
    '*** COLLECTION LEAVE TYPES ***
    Class cColLeaveTypes
        '====================================
        'Properties:                    Type:                       Perm:       Source:     Properties Required:
        'y  EE                          obj(User)                   RW          Virtual
        'y  CollectionType              lng                         RW          Virtual
        '   (all other collection properties as per base collection object)
        '
        'Methods:
        '   Add
        '   DBLoad
        '
        '====================================
        Private m_objCollection
        Private m_objEE
        Private m_lngCollectionType
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing ColLeaveTypes (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            SetNotLoaded
        End Sub
        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating ColLeaveTypes (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            Set m_objCollection = Nothing
        End Sub


        Private Sub SetEmpty
            Set m_objEE = nothing
            Set m_objCollection = nothing
        End Sub
        

        Private Sub SetNotLoaded
            m_blnLoaded = False
        End Sub
        
        
        Public Property Set EE(ByRef objUser)
            SetEmpty
            SetNotLoaded
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID
        End Property
        Public Property Get EE()
            Set EE = m_objEE
        End Property

        
        Public Property Let CollectionType(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            m_lngCollectionType = lngValue
        End Property
        Public Property Get CollectionType
            CollectionType = m_lngCollectionType
        End Property

        
        Public Property Get Count
            DBLoad
            If typename(m_objCollection) <> "cObjCollection" then
                Count = 0
            Else
                Count = m_objCollection.Count
            End If
        End Property


        Public Default Property Get Item(ByVal loclngIndex)
            DBLoad
            Initialise_Collection
            Set Item = m_objCollection(loclngIndex)
        End Property
        

        Public Sub Add(ByVal locObject)
            m_blnLoaded = True
            Initialise_Collection
            m_objCollection.Add(locObject)
        End Sub
        

        Private Sub DBLoad
        
            Dim locCmd
            Dim locCmdRevoke
            Dim locRS
            Dim locblnAddObject
            
            Dim fldlngID
            Dim fldstrName
            Dim fldblnEERequests
            Dim fldblnAdminEnters
            Dim fldblnRequestBeforeAccrued
            Dim flddblMinDays
            Dim flddblEntitlementAmount
            Dim flddblDaysBeforeStopsLegalAdjAccrual
            Dim fldblnDaysBeforeStopsLegalAdjAccrualIsConsecutive
            
            Dim locobjLeaveType
            
            If m_blnLoaded then
                Exit Sub
            End If

            '*** CHECK TO SEE - have we got enough info for the type of leave types collection we are loading?
            Select Case m_lngCollectionType
                Case CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_EE, _
                        CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_ADMIN, _
                        CONST_LEAVE_TYPE_COLLECTION_TYPE_OTHER_LEAVE_TYPES_FOR_EE
                    '*** If we have no WWID set up, we can't limit our leave types - so exit the routine.
                    If EE.WWID = "" then
                        Exit Sub
                    End If
                Case Else
                    Exit Sub
            End Select

            m_blnLoaded = True
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            '*** SET UP THE STORED PROCEDURE - depending on the type of leave types collection we are loading.
            Select Case m_lngCollectionType
                Case CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_EE
                    locCmd.CommandText = "usp_leavetypes_ee_requests"
                Case CONST_LEAVE_TYPE_COLLECTION_TYPE_OTHER_LEAVE_TYPES_FOR_EE
                    locCmd.CommandText = "usp_leavetypes_ee_other_leave"
                Case CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_ADMIN
                    locCmd.CommandText = "usp_leavetypes_admin"
            End Select
                        
            Set locRS = locCmd.Execute

            Set fldlngID = locRS("lngID")
            Set fldstrName = locRS("strLeaveTypeName")
            Set fldblnEERequests = locRS("blnEERequests")
            Set fldblnAdminEnters = locRS("blnAdminRequests")
            Set fldblnRequestBeforeAccrued = locRS("blnRequestBeforeAccrued")
            Set flddblMinDays = locRS("dblMinimumDays")
            Set flddblEntitlementAmount = locRS("dblEntitlement")
            Set flddblDaysBeforeStopsLegalAdjAccrual = locRS("dblDaysBeforeStopsLegalAdjAccrual")
            Set fldblnDaysBeforeStopsLegalAdjAccrualIsConsecutive = locRS("blnDaysBeforeStopsIsConsecutive")
            
            While not locRS.eof
                Set locobjLeaveType = new cObjLeaveType
                locobjLeaveType.ID = fldlngID
                locobjLeaveType.Name = fldstrName
                locobjLeaveType.EERequests = fldblnEERequests
                locobjLeaveType.AdminEnters = fldblnAdminEnters
                locobjLeaveType.RequestBeforeAccrued = fldblnRequestBeforeAccrued
                locobjLeaveType.MinDays = flddblMinDays
                locobjLeaveType.EntitlementAmount = flddblEntitlementAmount
                locobjLeaveType.DaysBeforeStopsLegalAdjAccrual = flddblDaysBeforeStopsLegalAdjAccrual
                locobjLeaveType.DaysBeforeStopsLegalAdjAccrualIsConsecutive = fldblnDaysBeforeStopsLegalAdjAccrualIsConsecutive

                locblnAddObject = True
                
                if (m_lngCollectionType = CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_EE OR _
                    m_lngCollectionType = CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_ADMIN)AND _
                    locobjLeaveType.Name = CONST_LEAVE_TYPE_NAME_ELP then
                    'This have been added so that ELP is never an option on the leave screen.
                    'if not EE.AnnualVacation.HasMaturedELP then
                        locblnAddObject = False
                    'end if
                end if
                    
                If locblnAddObject then Add locobjLeaveType
                
                Set locobjLeaveType = nothing
                locRS.movenext
            Wend

            Set fldlngID = nothing
            Set fldstrName = nothing
            Set fldblnEERequests = nothing
            Set fldblnAdminEnters = nothing
            Set fldblnRequestBeforeAccrued = nothing
            Set flddblMinDays = nothing
            Set flddblEntitlementAmount = nothing
            Set flddblDaysBeforeStopsLegalAdjAccrual = nothing
            Set fldblnDaysBeforeStopsLegalAdjAccrualIsConsecutive = nothing

            locRS.Close
            Set locRS = nothing
            Set locCmd = nothing

        End Sub
            
            
        '**** INITIALISE COLLECTION ****
        Private Sub Initialise_Collection()
            If typename(m_objCollection) <> "cObjCollection" then
                Set m_objCollection = new cObjCollection
            End If
        End Sub
                    
    End Class


    '**** PUBLIC HOLIDAYS COLLECTION ****
    Class cColPublicHolidays
        '====================================
        'Should only be used as global object! (Preventing repeated calls to DB)
        'Allows the retrieval of all public holidays.
        'y  Contains            lng                 R           Virtual
        'y  Item                objPublicHoliday    R           Virtual
        'y  CountInPeriod()     lng                 R           Calc        (StartDate, EndDate, IncludeWeekendDays)
        '
        '   (all other collection properties as per base collection object)
        '
        'Methods:
        '   Add
        '   DBLoad
        '
        '====================================
        Private m_objCollection
        Private m_blnLoaded

        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cColPublicHolidays (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objCollection = new cObjCollection
            m_blnLoaded = False
        End Sub
        
        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cColPublicHolidays (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objCollection = Nothing
        End Sub


        Public Property Get Contains(ByVal datValue, ByVal blnIncludeWeekends)
            Dim loclngCounter
            Dim loclngCount
            '(DBLoad forced on Get Item)
            If IsDate(datValue) then
                datValue = CDate(datValue)
                loclngCounter = 0
                loclngCount = Count
                While loclngCounter < loclngCount
                    loclngCounter = loclngCounter + 1
                    If Item(loclngCounter).Date = datValue then
                        If blnIncludeWeekends then
                            Contains = loclngCounter
                            Exit Property
                        ElseIf not mIsWeekendDay(Item(loclngCounter).Date) then
                            Contains = loclngCounter
                            Exit Property
                        End If
                    End If
                Wend
            End If
            Contains = False
        End Property
        
                
        Public Property Get Count
            DBLoad
            If typename(m_objCollection) <> "cObjCollection" then
                Count = 0
            Else
                Count = m_objCollection.Count
            End If
        End Property


        Public Default Property Get Item(ByVal loclngIndex)
            DBLoad
            Set Item = m_objCollection(loclngIndex)
        End Property


        Public Property Get CountInPeriod(ByVal locdatStartDate, ByVal locdatEndDate, ByVal locblnIncludeWeekends)
            Dim loclngCounter
            Dim loclngCount
            Dim loclngCountInPeriod
            If not (IsDate(locdatStartDate) and IsDate(locdatEndDate)) then
                CountInPeriod = 0
            Else
            '(DBLoad forced on Get Item)
                locdatStartDate = CDate(locdatStartDate)
                locdatEndDate = CDate(locdatEndDate)
                loclngCountInPeriod = 0
                loclngCounter = 0
                loclngCount = Count
                While loclngCounter < loclngCount
                    loclngCounter = loclngCounter + 1
                    If Item(loclngCounter).Date >= locdatStartDate and Item(loclngCounter).Date <= locdatEndDate then
                        If locblnIncludeWeekends then
                            loclngCountInPeriod = loclngCountInPeriod + 1
                        ElseIf not mIsWeekendDay(Item(loclngCounter).Date) then
                            loclngCountInPeriod = loclngCountInPeriod + 1
                        End If
                    End If
                Wend
            End If
            CountInPeriod = loclngCountInPeriod
        End Property
                
        
        Public Sub Add(ByVal locObject)
            m_objCollection.Add(locObject)
        End Sub
        
        
        Public Sub DBLoad
            Dim locCmd
            Dim locCmdRevoke
            Dim locRS
            Dim fldDate
            Dim fldDescription
            Dim locobjPublicHoliday

            '*** Already loaded - then exit ***
            if m_blnLoaded then
                exit sub
            end if
            
            m_blnLoaded = True

            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_publicholidays"
            locCmd.CommandType = adCmdStoredProc
                        
            Set locRS = locCmd.Execute

            Set fldDate = locRS("datDate")
            Set fldDescription = locRS("strDescription")
            
            While not locRS.eof
                Set locobjPublicHoliday = new cObjPublicHoliday
                locobjPublicHoliday.Date = fldDate
                locobjPublicHoliday.Description = fldDescription
                Add locobjPublicHoliday
                Set locobjPublicHoliday = nothing
                locRS.movenext
            Wend

            Set fldDate = nothing
            Set fldDescription = nothing

            locRS.Close
            Set locRS = nothing
            Set locCmd = nothing
        
        End Sub
            
    End Class
    

    '*** AnnualVacation Object ***
    Class cObjAnnualVacation
        '=================================================================
        'Description:   Contains definition Annual Vacation.
        'Properties:                    Type:                       Perm:       Source:     Properties Required:
        'y  EE                          obj(User)                   RW          DB (User)
        'y  Year                        lng                         R           Virtual     (EE.YearToView)
        'y  BalanceAtDate()             dat                         RW          Virtual
        'y  IsBalanceAtDate             bln                         R           Calc        (BalanceAtDate)
        'y  TotalEntitlement            dbl                         R           Calc
		'y  TotalEntitlementPre2016     dbl                         R           Calc
        'y  LeaveTypeRules              obj(LeaveType)              R(private)  DB
        'y  SeniorityEntitlement        dbl                         R           Calc
		'y  BasicEntitlement            dbl                         R           Calc        (LeaveTypeRules.EntitlementAmount)
        'y  CarryOverEOY                obj(CarryOver)              R           DB
        'y  CarryOverPreArranged        obj(CarryOver)              R           DB          (EE, Year)
        'y  AvailableELPDays            lng                         R           Calc        (HasActiveELP,ELPActive)
        'y  LegalAdjustmentType         str(Const)                  R           Calc        (EE.IsExempt)
        'y  LegalAdjustmentAccrued      dbl                         R           Calc        (LegalAdjustmentAccrued)
        'y  LegalAdjustmentAccruedNotRounded    dbl                 R           Calc        (LegalAdjustmentType,RestDaysAccrued,JRTTDaysAccrued)
        'y  LegalAdjustmentLost         dbl                         R           Calc
        'y  RestDaysAccruedPerMonth     dbl                         R           Calc
        'y  RestDaysAccruedPerYear      dbl                         R           Calc
        'y  RestDaysAccrued             dbl                         R           Calc
        'y  RestDaysLost                dbl                         R           Calc
        'y  JRTTDaysAccrued             dbl                         R           Calc
        'y  JRTTDaysLost                dbl                         R           Calc
        'y  JRTTDaysAccruedPerMonth     dbl                         R           Calc
        'y  JRTTDaysAccruedPerYear      dbl                         R           Calc
        'y  Leave                       col(LeavePeriod)            R           DB          (EE, Year)
        'y  CompTime                    col(CompTime)               R           DB          (EE) ADDED BY [MOF 1/09]
        'y  TotalCompTimeGranted        dbl                         R           Calc        (EE) ADDED BY [MOF 1/09]
        'y  NextCompLeaveDaysGranted    lng                         R           Calc        (EE) ADDED BY [MOF 1/09]
        'y  AvailableCompTime           lng                         R           Calc        (EE) ADDED BY [MOF 1/09]
        'y  DaysBooked                  dbl                         R           Calc        (Leave, Year)
        'y  DaysBookedNotExpired        dbl                         R           Calc        (ADDED BY [MOF 1/09])
        'y  DaysPendingApproval         dbl                         R           Calc        (ADDED BY [MOF 1/09])
        'y  Balance                     dbl                         R           Calc
        'y  ApprovedLeave               dbl                         R           Calc        (ADDED by [MOF 1/09])
        'y  HasActiveELP                bln                         R           Calc        (ELPActive)
        'y  HasMaturedELP               bln                         R           Calc        (ELPMatured)
        'y  HasUsedELP              bln
        'y  ELPActive                   obj(ELPInstance)            R           DB
        'y  ELPMatured                  obj(ELPInstance)            R           DB
        'y  ELPUsed                     obj(ELPInstance)
        'y  IsELPAvailable              bln                         R           Calc        (HasMaturedELP, ELPMatured)
        '
        '   ??? Projected ????
        '
        'Methods:
        '   DBLoadLeaveTypeRules
        '   DBLoadELP
        '   TestForChangedYear
        '
        '=================================================================
        Private m_objEE
        Private m_lngYear
        Private m_datBalanceAtDate
        Private m_objLeaveTypeRules
        Private m_strLegalAdjustmentType
        Private m_colLeave
        Private m_colCompTime
        Private m_objELPActive
        Private m_objELPMatured
        Private m_objELPUsed
        
        Private m_blnLeaveTypeRulesLoaded
        Private m_blnELPLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjAnnualVacation (" & m_lngtempobjid & "): " & timer & "<br>"
            m_datBalanceAtDate = ""
            SetEmpty
            SetNotLoaded
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjAnnualVacation (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            Set m_objLeaveTypeRules = nothing
            Set m_colLeave = nothing
            Set m_colCompTime = nothing ' Added by [MOF 5-09]
            Set m_objELPActive = nothing
            Set m_objELPMatured = nothing
            Set m_objELPUsed = nothing
        End Sub


        Private Sub SetNotLoaded()
            m_blnLeaveTypeRulesLoaded = False
        End Sub
        
        
        Private Sub SetEmpty()
            Set m_colLeave = nothing
            Set m_colCompTime = nothing ' Added by [MOF 05-09]
            Set m_objLeaveTypeRules = nothing
            m_strLegalAdjustmentType = ""
        End Sub
        
        
        Public Property Set EE(ByRef objUser)
            SetEmpty
            SetNotLoaded
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID
        End Property
        Public Property Get EE()
            Set EE = m_objEE
        End Property

        ' [MOF 1/09] NOTE: This property conflicts with the function Year() provided by ASP default library. 
        Public Property Get Year()
            
            TestForChangedYear
            Year = m_lngYear
        End Property

        
        Public Property Let BalanceAtDate(ByVal datValue)
            If datValue = "" then
                m_datBalanceAtDate = ""
            ElseIf IsDate(datValue) then
                m_datBalanceAtDate = datValue
            End If
        End Property
        Public Property Get BalanceAtDate()
            BalanceAtDate = m_datBalanceAtDate
        End Property


        Public Property Get IsBalanceAtDate()
            If not IsDate(m_datBalanceAtDate) then
                IsBalanceAtDate = False
            Else
                IsBalanceAtDate = True
            End If
        End Property
                
        
        Public Property Get TotalEntitlement()
            TestForChangedYear
            TotalEntitlement = BasicEntitlement + CarryOverEOY + CarryOverPreArranged + LegalAdjustmentAccrued
        End Property
		
		Public Property Get TotalEntitlementPre2016()
            TestForChangedYear
            TotalEntitlementPre2016 = BasicEntitlementPre2016 + SeniorityEntitlement + CarryOverEOY + CarryOverPreArranged + LegalAdjustmentAccrued
        End Property

        Private Property Get LeaveTypeRules()
            TestForChangedYear
            DBLoadLeaveTypeRules
            Set LeaveTypeRules = m_objLeaveTypeRules
        End Property
        
        ' return the basic entitlement for the current user this year      
        ' calculated if user has started or will finish this year (not working all the year) 
        Public Property Get BasicEntitlement()
            Dim locdblEntitlement
            Dim loclngEEStartDateDay
            Dim loclngEEEndDateDay
            Dim loclngFirstMonthEntitlement 
            Dim loclngLastMonthEntitlement
            
            Dim loclngDaysInEEsFirstMonth
            Dim loclngDaysInEEsLastMonth
            Dim loclngEEWholeMonthsInYear
            Dim loclngMonthlyEntitlement

            TestForChangedYear
            
            If EE.IsBlueBadge then
				if (IsDate(EE.EndDate) and  datepart("yyyy",EE.EndDate)< Year) then
					locdblEntitlement = 0
                elseif (EE.StartDate >= mFirstDayOfYear(Year)) or (IsDate(EE.EndDate) and  datepart("yyyy",EE.EndDate)= Year) then
                'specific basic entitlement
					loclngMonthlyEntitlement = (int((LeaveTypeRules.EntitlementAmount / 12)*100))/100
					
					If EE.StartDate >= mFirstDayOfYear(Year) then
						' starting working this year
						loclngEEStartDateDay = EE.StartDate
						loclngDaysInEEsFirstMonth = mGetDaysInMonthForDate(EE.StartDate)
						loclngFirstMonthEntitlement = (loclngDaysInEEsFirstMonth - datepart("d",EE.StartDate)) + 1 'nb of days
                    	loclngFirstMonthEntitlement = ((loclngFirstMonthEntitlement * loclngMonthlyEntitlement) / loclngDaysInEEsFirstMonth)
					else
						loclngFirstMonthEntitlement = loclngMonthlyEntitlement
						loclngEEStartDateDay = mFirstDayOfYear(Year)
					end if
					if (IsDate(EE.EndDate) and  datepart("yyyy",EE.EndDate)= Year) then
						' finishing working this year
						loclngEEEndDateDay = EE.EndDate
						loclngDaysInEEsLastMonth = mGetDaysInMonthForDate(EE.EndDate)
						loclngLastMonthEntitlement = (datepart("d",EE.EndDate) * loclngMonthlyEntitlement) / loclngDaysInEEsLastMonth	
	
					else 
						loclngLastMonthEntitlement = loclngMonthlyEntitlement
						loclngEEEndDateDay = mLastDayOfYear(Year)
					end if
					
					' special case if the same month for both
                    if ((DatePart("m", loclngEEEndDateDay)= DatePart("m", loclngEEStartDateDay)) and (DatePart("yyyy", loclngEEEndDateDay)= DatePart("yyyy", loclngEEStartDateDay))) then
						loclngFirstMonthEntitlement = (loclngEEEndDateDay - loclngEEStartDateDay) + 1 'nb of days
						locdblEntitlement = (loclngFirstMonthEntitlement * loclngMonthlyEntitlement)/mGetDaysInMonthForDate(loclngEEStartDateDay)
                    else
	                    
						loclngEEWholeMonthsInYear = ((DatePart("m", loclngEEEndDateDay)-1)) - (DatePart("m", loclngEEStartDateDay)) 
						
						locdblEntitlement = (loclngLastMonthEntitlement + loclngFirstMonthEntitlement + (loclngEEWholeMonthsInYear * loclngMonthlyEntitlement))                   
					end if
            			locdblEntitlement = mRoundToNextHalf(locdblEntitlement)
			    else ' working all the year
                   locdblEntitlement = LeaveTypeRules.EntitlementAmount
                end if


            Else
                locdblEntitlement = 0
            End If
                
            BasicEntitlement = locdblEntitlement
        End Property
		
		' return the basic entitlement for the current user this year      
        ' calculated if user has started or will finish this year (not working all the year) 
        Public Property Get BasicEntitlementPre2016()
            Dim locdblEntitlement
            Dim loclngEEStartDateDay
            Dim loclngEEEndDateDay
            Dim loclngFirstMonthEntitlement 
            Dim loclngLastMonthEntitlement
            
            Dim loclngDaysInEEsFirstMonth
            Dim loclngDaysInEEsLastMonth
            Dim loclngEEWholeMonthsInYear
            Dim loclngMonthlyEntitlement

            TestForChangedYear
            
            If EE.IsBlueBadge then
				if (IsDate(EE.EndDate) and  datepart("yyyy",EE.EndDate)< Year) then
					locdblEntitlement = 0
                elseif (EE.StartDate >= mFirstDayOfYear(Year)) or (IsDate(EE.EndDate) and  datepart("yyyy",EE.EndDate)= Year) then
                'specific basic entitlement
					loclngMonthlyEntitlement = (int((21 / 12)*100))/100
					
					If EE.StartDate >= mFirstDayOfYear(Year) then
						' starting working this year
						loclngEEStartDateDay = EE.StartDate
						loclngDaysInEEsFirstMonth = mGetDaysInMonthForDate(EE.StartDate)
						loclngFirstMonthEntitlement = (loclngDaysInEEsFirstMonth - datepart("d",EE.StartDate)) + 1 'nb of days
                    	loclngFirstMonthEntitlement = ((loclngFirstMonthEntitlement * loclngMonthlyEntitlement) / loclngDaysInEEsFirstMonth)
					else
						loclngFirstMonthEntitlement = loclngMonthlyEntitlement
						loclngEEStartDateDay = mFirstDayOfYear(Year)
					end if
					if (IsDate(EE.EndDate) and  datepart("yyyy",EE.EndDate)= Year) then
						' finishing working this year
						loclngEEEndDateDay = EE.EndDate
						loclngDaysInEEsLastMonth = mGetDaysInMonthForDate(EE.EndDate)
						loclngLastMonthEntitlement = (datepart("d",EE.EndDate) * loclngMonthlyEntitlement) / loclngDaysInEEsLastMonth	
	
					else 
						loclngLastMonthEntitlement = loclngMonthlyEntitlement
						loclngEEEndDateDay = mLastDayOfYear(Year)
					end if
					
					' special case if the same month for both
                    if ((DatePart("m", loclngEEEndDateDay)= DatePart("m", loclngEEStartDateDay)) and (DatePart("yyyy", loclngEEEndDateDay)= DatePart("yyyy", loclngEEStartDateDay))) then
						loclngFirstMonthEntitlement = (loclngEEEndDateDay - loclngEEStartDateDay) + 1 'nb of days
						locdblEntitlement = (loclngFirstMonthEntitlement * loclngMonthlyEntitlement)/mGetDaysInMonthForDate(loclngEEStartDateDay)
                    else
	                    
						loclngEEWholeMonthsInYear = ((DatePart("m", loclngEEEndDateDay)-1)) - (DatePart("m", loclngEEStartDateDay)) 
						
						locdblEntitlement = (loclngLastMonthEntitlement + loclngFirstMonthEntitlement + (loclngEEWholeMonthsInYear * loclngMonthlyEntitlement))                   
					end if
            			locdblEntitlement = mRoundToNextHalf(locdblEntitlement)
			    else ' working all the year
                   locdblEntitlement = 21
                end if


            Else
                locdblEntitlement = 0
            End If
                
            BasicEntitlementPre2016 = locdblEntitlement
        End Property
		
        Public Property Get SeniorityEntitlement()
            Dim loclngQualLOS
           
            TestForChangedYear
            
            If EE.IsBlueBadge then
                ' If EndDate is set (coops will have this) then check if they will be employed more than a year when they leave (Added by MOF 1/09)
                If isDate(EE.EndDate) then
                    ' Check if the end date is this year
                    'response.Write "EndDate(yr) = " & DatePart("yyyy", EE.EndDate) & ", Year = " & DatePart("yyyy", Now) & "<br>"
                    If DatePart("yyyy", EE.EndDate) = DatePart("yyyy", Now) then
                        'Response.Write "End Date = " & EE.EndDate & "<br>"
                        loclngQualLOS = mCountWholeYears(EE.StartDate, EE.EndDate)
                    End If
                Else 
                    'Response.Write "No End Date Given. Using last day of year = " & mLastDayOfYear(Year) & "<br>"
                    loclngQualLOS = mCountWholeYears(EE.StartDate, mLastDayOfYear(Year))'[MFILLAST 08-2006] modif : year-1 => year see HR guideline
                End If
	
				If loclngQualLOS < 1 then
				 SeniorityEntitlement = 0
				ElseIf loclngQualLOS <=2 then
				 SeniorityEntitlement = 1
				ElseIf loclngQualLOS <=4 then
				 SeniorityEntitlement = 2
				ElseIf loclngQualLOS >4 then
				 SeniorityEntitlement = 3
				Else
				 SeniorityEntitlement = 0
				End If
            Else
                SeniorityEntitlement = 0
            End If
            
        End Property
        
        
        Public Property Get CarryOverEOY()
            TestForChangedYear
            Set CarryOverEOY = EE.CarryOverEOYForYear(Year)
        End Property
        
        
        Public Property Get CarryOverPreArranged()
            TestForChangedYear
            Set CarryOverPreArranged = EE.CarryOverPreArrangedForYear(Year)
        End Property        
        
        
        Public Property Get AvailableELPDays()
            If HasMaturedELP then
                If ELPMatured.IsUsed then
                    AvailableELPDays = 0
                Else
                    AvailableELPDays = ELPMatured.DaysBanked
                End If
            Else
                AvailableELPDays = 0
            End If
        
        End Property
        
        
        Public Property Get LegalAdjustmentType()		
            LegalAdjustmentType = ""
        End Property        

        
        Public Property Get LegalAdjustmentAccrued()
            LegalAdjustmentAccrued = mRoundToNextHalf(LegalAdjustmentAccruedNotRounded)
        End Property


        Public Property Get LegalAdjustmentAccruedNotRounded()
            LegalAdjustmentAccruedNotRounded = LegalAdjustmentAccruedBeforeLost - LegalAdjustmentLost
        End Property


        Public Property Get LegalAdjustmentAccruedBeforeLost()
            TestForChangedYear
            Select Case LegalAdjustmentType
                Case CONST_LEGAL_ADJUSTMENT_TYPE_NAME_REST_DAYS
                    LegalAdjustmentAccruedBeforeLost = RestDaysAccrued
                Case CONST_LEGAL_ADJUSTMENT_TYPE_NAME_JRTT
                    LegalAdjustmentAccruedBeforeLost = JRTTDaysAccrued
                Case Else
                    LegalAdjustmentAccruedBeforeLost = 0
            End Select
        End Property
                    

        Public Property Get LegalAdjustmentLost()
            Select Case LegalAdjustmentType
                Case CONST_LEGAL_ADJUSTMENT_TYPE_NAME_REST_DAYS
                    LegalAdjustmentLost = RestDaysLost
                    
                Case CONST_LEGAL_ADJUSTMENT_TYPE_NAME_JRTT
                    LegalAdjustmentLost = JRTTDaysLost
                Case Else
                    LegalAdjustmentLost = 0
            End Select
        End Property

		Private Property Get RestDaysAccruedPerYear()
            Dim loclngRestDaysAccrued
            loclngRestDaysAccrued = mGetDaysInYear(Year) _
                - glbPublicHolidays.CountInPeriod(mFirstDayOfYear(Year),mLastDayOfYear(Year), False) _
                - mCountWeekendDays(mFirstDayOfYear(Year),mLastDayOfYear(Year)) _
                - CONST_LEGAL_ADJUSTMENT_MAX_WORKING_DAYS _
                - LeaveTypeRules.EntitlementAmount
            RestDaysAccruedPerYear = loclngRestDaysAccrued
            RestDaysAccruedPerYear =0
        End Property
        
		Private Property Get RestDaysAccruedPerMonth()
            Dim loclngRestDays
            loclngRestDays = RestDaysAccruedPerYear()
            loclngRestDays = loclngRestDays / 12
            RestDaysAccruedPerMonth = loclngRestDays
        End Property


		Private Property Get RestDaysAccrued
            Dim locdatCurrentDate
            Dim loclngCurrentMonth
            Dim loclngFirstMonth
            Dim loclngMonths
			
            loclngCurrentMonth = 12
            If DatePart("yyyy", EE.StartDate) < Year then
                loclngFirstMonth = 1
            Else
                loclngFirstMonth = month(EE.StartDate)
                'loclngFirstMonth = 12
                If day(EE.StartDate) > 15 then
                    loclngFirstMonth = loclngFirstMonth + 1
                End If
            End If
            
            loclngMonths = loclngCurrentMonth - loclngFirstMonth + 1
            if loclngMonths < 0 then loclngMonths = 0
            
            RestDaysAccrued = (loclngMonths * RestDaysAccruedPerMonth)
            
        End Property
        
        
       Private Property Get RestDaysLost()
            Dim loclngApplicableLeaveDays
            Dim loclngRestDaysLost
            Dim loclngEndDate
            Dim loclngWorkingDaysInYear
            
			loclngEndDate = BalanceAtDate
			loclngEndDate = mLastDayOfYear(Year)
             
            
            loclngApplicableLeaveDays = EE.LeaveRequests.LeaveDaysInPeriod(mFirstDayOfYear(Year),loclngEndDate, True, True)
            
            loclngWorkingDaysInYear = mGetDaysInYear(Year) _
                - glbPublicHolidays.CountInPeriod(mFirstDayOfYear(Year),mLastDayOfYear(Year), False) _
                - mCountWeekendDays(mFirstDayOfYear(Year),mLastDayOfYear(Year))

            loclngRestDaysLost = (loclngApplicableLeaveDays / loclngWorkingDaysInYear) * RestDaysAccruedPerYear
            
            
            RestDaysLost = loclngRestDaysLost
            
        End Property
        
        
		Private Property Get JRTTDaysAccrued()
            Dim locdatCurrentDate
            Dim loclngCurrentMonth
            Dim loclngFirstMonth
            Dim loclngMonths

            loclngCurrentMonth = 12
            If DatePart("yyyy", EE.StartDate) < Year then
                loclngFirstMonth = 1
            Else
                loclngFirstMonth = month(EE.StartDate)
                If day(EE.StartDate) > 15 then
                    loclngFirstMonth = loclngFirstMonth + 1
                End If
            End If
            
            loclngMonths = loclngCurrentMonth - loclngFirstMonth + 1
            if loclngMonths < 0 then loclngMonths = 0
            
            JRTTDaysAccrued = loclngMonths * JRTTDaysAccruedPerMonth
            
        End Property

		Property Get JRTTDaysAccruedPerMonth
            JRTTDaysAccruedPerMonth = 0
        End Property

		Property Get JRTTDaysAccruedPerYear
            JRTTDaysAccruedPerYear = JRTTDaysAccruedPerMonth * 12
        End Property
        
 
		Private Property Get JRTTDaysLost()
            Dim loclngApplicableLeaveDays
            Dim loclngJRTTDaysLost
            Dim loclngEndDate
            Dim loclngWorkingDaysInYear
            
            loclngEndDate = mLastDayOfYear(Year)
            
            loclngApplicableLeaveDays = EE.LeaveRequests.LeaveDaysInPeriod(mFirstDayOfYear(Year),loclngEndDate, True, True) 'CA - do the same as in RestDaysLost

            loclngWorkingDaysInYear = mGetDaysInYear(Year) _
                - glbPublicHolidays.CountInPeriod(mFirstDayOfYear(Year),mLastDayOfYear(Year), False) _
                - mCountWeekendDays(mFirstDayOfYear(Year),mLastDayOfYear(Year)) 
            
            loclngJRTTDaysLost = (loclngApplicableLeaveDays / loclngWorkingDaysInYear) * JRTTDaysAccruedPerYear
            
            JRTTDaysLost = loclngJRTTDaysLost
            
        End Property


        Public Property Get Leave()
            TestForChangedYear
            Initialise_Leave
            Set Leave = m_colLeave
        End Property
        
        Public Property Get CompTime()
            Initialise_CompTime
            Set CompTime = m_colCompTime
        End Property

        Public Property Get TotalCompTimeGranted()
            Dim loclngCompTimeGranted
            Dim loclngCount
            loclngCompTimeGranted = 0
            
            Initialise_CompTime
            
            ' Count the days of comp time available
            For loclngCount = 1 To m_colCompTime.Count
            
                ' TO DO: Don't include expired comp time
                
                loclngCompTimeGranted = loclngCompTimeGranted + m_colCompTime.Item(loclngCount).DaysGranted + m_colCompTime.Item(loclngCount).DaysBooked2 + m_colCompTime.Item(loclngCount).Taken
            Next
            
            TotalCompTimeGranted = loclngCompTimeGranted
        End Property
        
        Public Property Get NextCompLeaveDaysGranted
            Initialise_CompTime
    
            If CompTime.Count = 0 then
                NextCompLeaveDaysGranted = 0
            Else
                NextCompLeaveDaysGranted = CompTime.Item(1).DaysGranted          
            End If
        End Property
        
        Public Property Get NextCompLeaveDaysRevoked
            Initialise_CompTime
    
            If CompTime.Count = 0 then
                NextCompLeaveDaysGranted = 0
            Else
                NextCompLeaveDaysRevoked = CompTime.Item(1).DaysRevoked          
            End If
        End Property

        Public Property Get AvailableCompTime
            Dim loclngAvailableCompDays
            Dim loclngCounter
            
            Initialise_CompTime
            
		    loclngAvailableCompDays = 0
		    For loclngCounter = 1 To CompTime.Count
		        loclngAvailableCompDays = loclngAvailableCompDays + _
		                                  CompTime.Item(loclngCounter).DaysGranted - _
		                                  CompTime.Item(loclngCounter).DaysBooked
		    Next
		    
		    AvailableCompTime = loclngAvailableCompDays
        End Property
        
        Public Property Get NewAvailableCompTime
            Dim loclngAvailableCompDays
            Dim loclngCounter
            
            Initialise_CompTime
            
		    loclngAvailableCompDays = 0
		    For loclngCounter = 1 To CompTime.Count
		        loclngAvailableCompDays = loclngAvailableCompDays - _
		                                  CompTime.Item(loclngCounter).DaysRevoked - _
		                                  CompTime.Item(loclngCounter).DaysBooked
		    Response.write "Str 1898  " & CompTime.Item(loclngCounter).DaysRevoked
		    
		    Next
		    
		    NewAvailableCompTime = loclngAvailableCompDays
		    
		    Response.write "Str 1898  " &   NewAvailableCompTime
        End Property

        Public Property Get DaysBooked()
            Dim loclngEndDate
            TestForChangedYear
            If IsBalanceAtDate then
                loclngEndDate = BalanceAtDate
            Else
                loclngEndDate = mLastDayOfYear(Year)
            End If
            DaysBooked = Leave.LeaveDaysInPeriod(mFirstDayOfYear(Year),loclngEndDate, False, False)
        End Property
        
        
        Public Property Get DaysBookedNotExpired() 
            Dim locdblDaysBookedNotExpired
            Dim locdblLeaveDays
            Dim loclngCount
            locdblLeaveDays = 0
            locdblDaysBookedNotExpired = 0
            
            TestForChangedYear
            'response.Write "Count = " & Leave.Count & "<br>"
            For loclngCount = 1 To Leave.Count
                'response.Write "Status: " & Leave.Item(loclngCount).Status & "<br>"
                If Leave.Item(loclngCount).Status = CONST_LEAVE_PERIOD_STATUS_APPROVED and _
                   not Leave.Item(loclngCount).CanConfirmLeave Then
                    locdblLeaveDays = Leave.Item(loclngCount).LeaveDaysInPeriod(Leave.Item(loclngCount).StartDate, _
                                                                                Leave.Item(loclngCount).EndDate)

                    locdblDaysBookedNotExpired = locdblDaysBookedNotExpired + locdblLeaveDays
                End If
            Next
            
            DaysBookedNotExpired = locdblDaysBookedNotExpired
        End Property
        
        
        Public Property Get DaysPendingApproval() 
            Dim locdblDaysPendingApproval
            Dim locdblLeaveDays
            Dim loclngCount
            locdblLeaveDays = 0
            locdblDaysPendingApproval = 0
            
            TestForChangedYear
            'response.Write "Count = " & Leave.Count & "<br>"
            For loclngCount = 1 To Leave.Count

            '    response.Write "ID " & Leave.Item(loclngCount).ID & ", " & Leave.Item(loclngCount).Status
                
             '   response.Write "Status: '" & Leave.Item(loclngCount).Status & "', '" & CONST_LEAVE_PERIOD_STATUS_RAISED & "'<br>"
                
                If Leave.Item(loclngCount).Status = CONST_LEAVE_PERIOD_STATUS_RAISED Then
                    locdblLeaveDays = Leave.Item(loclngCount).LeaveDaysInPeriod(Leave.Item(loclngCount).StartDate, _
                                                                                Leave.Item(loclngCount).EndDate)

                    locdblDaysPendingApproval = locdblDaysPendingApproval + locdblLeaveDays
                End If
            Next
            
            DaysPendingApproval = locdblDaysPendingApproval
        End Property
        
        Public Property Get Balance()
            TestForChangedYear
            
            Balance = TotalEntitlement - _
                        (ELPActive.DaysBankedInCurrentYear + _
                        ELPMatured.DaysBankedInCurrentYear + _
                        ELPUsed.DaysBankedInCurrentYear) - _
                        DaysBooked 'CA - added ELPUsed.
            
        End Property   

		Public Property Get BalancePre2016()
            TestForChangedYear
            
            BalancePre2016 = TotalEntitlementPre2016 - _
                        (ELPActive.DaysBankedInCurrentYear + _
                        ELPMatured.DaysBankedInCurrentYear + _
                        ELPUsed.DaysBankedInCurrentYear) - _
                        DaysBooked 'CA - added ELPUsed.
            
        End Property  		
        
      
        ' Note: This is required because Balance() returns a value inclusive of Available Comp Days.
        ' In the Team Leave Summary, Managers need to see the balance without Comp Time as it is in another column
        Public Property Get BalanceForMgrView()
            BalanceForMgrView = Balance - AvailableCompTime
        End Property
        
        Public Property Get DaysApproved()
            Dim locdblTotalDaysApproved
            Dim locdblLeaveDays
            Dim loclngCount
            locdblLeaveDays = 0
            locdblTotalDaysApproved = 0
            
            TestForChangedYear
            'response.Write "Count = " & Leave.Count & "<br>"
            For loclngCount = 1 To Leave.Count
                'response.Write "Status: " & Leave.Item(loclngCount).Status & "<br>"
                If Leave.Item(loclngCount).Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
                   Leave.Item(loclngCount).Status = CONST_LEAVE_PERIOD_STATUS_CONFIRMED or _
                   Leave.Item(loclngCount).Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED or _
                   Leave.Item(loclngCount).Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED Then
                    locdblLeaveDays = Leave.Item(loclngCount).LeaveDaysInPeriod(Leave.Item(loclngCount).StartDate, _
                                                                                Leave.Item(loclngCount).EndDate)

                    locdblTotalDaysApproved = locdblTotalDaysApproved + locdblLeaveDays
                End If
            Next
            
            DaysApproved = locdblTotalDaysApproved
        End Property
        
        
        Public Property Get HasActiveELP()
            If ELPActive.ELPID <> 0 then
                HasActiveELP = True
            Else
                HasActiveELP = False
            End If
        End Property


        Public Property Get HasMaturedELP()
            If ELPMatured.ELPID <> 0 then
                HasMaturedELP = True
            Else
                HasMaturedELP = False
            End If
        End Property
        
        Public Property Get HasUsedELP()
            If ELPUsed.ELPID <> 0 then
                HasUsedELP = True
            Else
                HasUsedELP = False
            End If
        End Property
        

        Public Property Get ELPActive()
            TestForChangedYear
            Initialise_ELPActive
            DBLoadELP
            Set ELPActive = m_objELPActive                        
        End Property

        
        Public Property Get ELPMatured()
            TestForChangedYear
            Initialise_ELPMatured   
            DBLoadELP
            Set ELPMatured = m_objELPMatured                      
        End Property    
        
        Public Property Get ELPUsed()
            TestForChangedYear
            Initialise_ELPUsed
            DBLoadELP
            Set ELPUsed = m_objELPUsed                    
        End Property    
        
        
        Public Property Get IsELPAvailable()
            If HasMaturedELP then
                If (Not ELPMatured.IsUsed) then
                    IsELPAvailable = True
                Else
                    IsELPAvailable = False
                End If
            Else
                IsELPAvailable = False
            End If
        End Property
        
        
        Private Sub DBLoadLeaveTypeRules()
            If m_blnLeaveTypeRulesLoaded then
                Exit Sub
            End If
            
            m_blnLeaveTypeRulesLoaded = True
            
            Set m_objLeaveTypeRules = new cObjLeaveType
            
            m_objLeaveTypeRules.Name = CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION
            
        End Sub
        
        
        Private Sub DBLoadELP()
            
            Dim locCmd
            Dim locCmdRevoke
            Dim locParam
            Dim locRS

            Dim fldlngID
            
            Dim locobjELPInstance
            
            If m_blnELPLoaded Then
                Exit Sub
            End If
            
            '*** If we have no EE.WWID set up, we can't find our leave requests for this EE - so exit the routine.
            If EE.WWID = "" then
                exit sub
            End if
            
            m_blnELPLoaded = True
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_eeelp"
            locCmd.CommandType = adCmdStoredProc
                        
            Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, EE.WWID)
            locCmd.Parameters.Append locParam

            Set locRS = locCmd.Execute

            Set fldlngID = locRS("ELPID")
            
            While not locRS.eof
                
                Set locobjELPInstance = new cObjELPInstance
                locobjELPInstance.ELPID = fldlngID
                locobjELPInstance.CurrentYear = Year
                
                If locobjELPInstance.Status = CONST_ELP_STATUS_ACTIVE then
                    Set m_objELPActive = locobjELPInstance
                ElseIf locobjELPInstance.Status = CONST_ELP_STATUS_MATURED then
                    'response.write("Still here!")
                    Set m_objELPMatured = locobjELPInstance
                '*** Added for Used ELP
                ElseIf locobjELPInstance.Status = CONST_ELP_STATUS_USED then
                    Set m_objELPUsed = locobjELPInstance
            
                End If
                Set locobjELPInstance = nothing
                locRS.movenext
            Wend

            Set fldlngID = nothing

            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
        End Sub


        Private Sub TestForChangedYear()
            If IsBalanceAtDate Then
                If datepart("yyyy",BalanceAtDate) <> m_lngYear then
                    m_lngYear = datepart("yyyy",BalanceAtDate)
                    EE.YearToView = m_lngYear
                End If
            Else
                'Response.Write "EE.YearToView: " & EE.YearToView & "  m_lngYear: " & m_lngYear & "<BR>"    
                If EE.YearToView <> m_lngYear then
                    SetEmpty
                    SetNotLoaded
                    
                    m_lngYear = EE.YearToView
                    'Response.Write "m_lngYear: " & m_lngYear & "<BR>"  
                End If
            End If
        End Sub
        
        
        '**** INITIALISE LEAVE ****
        Private Sub Initialise_Leave
            If typename(m_colLeave) <> "cColLeavePeriods" then
                mDebugPrint "   Creating cColLeavePeriods - 1 <br>"
                Set m_colLeave = new cColLeavePeriods
                m_colLeave.CollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE
                Set m_colLeave.EE = EE
            End If
        End Sub


        '**** INITIALISE ELP ACTIVE ****
        Private Sub Initialise_ELPActive
            If typename(m_objELPActive) <> "cObjELPInstance" then
                Set m_objELPActive = new cObjELPInstance
            End If
        End Sub
        
        
        '**** INITIALISE ELP MATURED ****
        Private Sub Initialise_ELPMatured
	            If typename(m_objELPMatured) <> "cObjELPInstance" then
                Set m_objELPMatured = new cObjELPInstance
            End If
        End Sub
            
        '**** INITIALISE ELP USED ****
        Private Sub Initialise_ELPUSed
            If typename(m_objELPUsed) <> "cObjELPInstance" then
                Set m_objELPUsed = new cObjELPInstance
            End If
        End Sub
        
        '**** INITIALISE COMPENSATORY TIME ****
        Public Sub Initialise_CompTime 
            
            Dim Booked_Days
            Dim Revoked_Days
            Dim Granted_Days
            Dim Taken_Days
            
            If typename(m_colCompTime) <> "cObjCollection" Then
            
                mDebugPrint "   Creating cObjCollection <br>" 
                Set m_colCompTime = new cObjCollection
                
                ' Load the Comp Times for this year 
                Dim locstrCompTimesSQL
                Dim locRSCompTime
                locstrCompTimesSQL = "SELECT * FROM tblCompTime WHERE lngWWID = " & EE.WWID
                Set locRSCompTime = Server.CreateObject("ADODB.RecordSet")
                locRSCompTime.Open locstrCompTimesSQL, _
			                       glbConnection, _
			                       adOpenStatic
               
                While not locRSCompTime.EOF
                    Dim locobjCompTime
                    Set locobjCompTime = new cObjCompTime
           
                    locobjCompTime.ID = locRSCompTime("lngID")
                    Booked_Days = locRSCompTime("lngDaysBooked")
                    Revoked_Days = locRSCompTime("IngDaysRevoked")
                    Granted_Days = locRSCompTime("lngDaysGranted")
                    Taken_Days = locRSCompTime("Taken")
                    
                    m_colCompTime.Add locobjCompTime
                    locRSCompTime.MoveNext
                    Set locobjCompTime = nothing
                Wend
            End If
        End Sub
        
          '**** INITIALISE COMPENSATORY TIME ****
        Public Sub DaysRevoked        
            If typename(m_colCompTime) <> "cObjCollection" Then
            
                mDebugPrint "   Creating cObjCollection <br>" 
                Set m_colCompTime = new cObjCollection
                
                ' Load the Comp Times for this year 
                Dim locstrCompTimesSQL
                Dim locRSCompTime
                locstrCompTimesSQL = "SELECT * FROM tblCompTime WHERE lngWWID=" & EE.WWID
                Set locRSCompTime = Server.CreateObject("ADODB.RecordSet")
                locRSCompTime.Open locstrCompTimesSQL, _
			                       glbConnection, _
			                       adOpenStatic
               
                While not locRSCompTime.EOF
                    Dim locobjRevokeCompTime
                    Set locobjRevokeCompTime = new cObjCompTime
           
                    locobjRevokeCompTime.ID = locRSCompTime("IngDaysRevoke")
                    
                    m_colRevokeCompTime.Add locobjRevokeCompTime
                    locRSCompTime.MoveNext
                    Set locobjCompTime = nothing
                Wend
            End If
        End Sub
        
    End Class
    
    
    '*** CarryOver Object ***
    Class cObjCarryOver
        '=================================================================
        '***
        '       For CarryOver the YEAR refers to the year in which the
        '       leave days are added to the Annual Leave entitlement, NOT the year
        '       that they were 'borrowed' from.
        '***
        '
        'Description:   Contains definition of a CarryOver.
        'Properties:                Type:   Perm:       Source:     Properties Required:
        'y      ID                  lng         RW          DB
        'y      EE                  objUser     RW          DB
        'y      Year                lng         RW          DB
        'y      Days (DEFAULT)      lng         R           DB
        'y      EnteredBy           objUser     R           DB (WWID)
        'y      DateEntered         dat         R           DB
        'y      Comments            text        R           DB
        'y      IsPreArranged       bln         RW          DB
        'y      AdminCarryOverFormErrorMessage  str R       Calc
        'y      AdminCarryOverFormIsValid   bln R           Calc
        'y      IsInvalidField      bln         R           Calc
        '
        'Methods:
        '   DBLoad
        '   LoadAdminCarryOverFromForm
        '   Save
        '
        '=================================================================
        Private m_lngID
        Private m_objEE
        Private m_lngYear
        Private m_lngDays 
        Private m_objEnteredBy 
        Private m_datEntered 
        Private m_strComments
        Private m_blnIsPreArranged
        Private m_strErrorMessage
        Private m_strErrorFieldList
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjCarryOver (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = new cObjUser
            m_lngID = 0
            m_lngYear = DatePart("yyyy",now())
            SetEmpty
            SetNotLoaded
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjCarryOver (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            Set m_objEnteredBy = nothing
        End Sub

        
        Private Sub SetEmpty()
            m_lngDays = 0
            m_strErrorMessage = ""
            m_strErrorFieldList = ""
            Set m_objEnteredBy = nothing
            Set m_objEnteredBy = new cObjUser
            m_datEntered = ""
            m_strComments = ""
            m_blnIsPreArranged = False
        End Sub


        Private Sub SetNotLoaded()
            m_blnLoaded = False
        End Sub
        
        
        Public Property Let ID(ByVal varValue)
            varValue = mGetSafeLongInteger(varValue,0)
            If varValue <> 0 then
                If m_lngID <> varValue then
                    SetEmpty
                    SetNotLoaded
                    m_lngID = varValue
                End If
            End If
        End Property
        Public Property Get ID()
            DBLoad
            ID = m_lngID
        End Property

        
        Public Property Set EE(ByRef objUser)
            If m_objEE.WWID <> objUser.WWID then
                SetEmpty
                SetNotLoaded
                Set m_objEE = new cObjUser
                m_objEE.WWID = objUser.WWID
            End If
        End Property
        Public Property Get EE()
            Set EE = m_objEE
        End Property


        Public Property Let Year(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            If m_lngYear <> lngValue then
                SetEmpty
                SetNotLoaded
                m_lngYear = lngValue
            End If
        End Property
        Public Property Get Year() 
            DBLoad       
            Year = m_lngYear
        End Property


        Public Property Let Days(ByVal dblValue)
            m_blnLoaded = True
            m_lngDays = dblValue        
        End Property
        Public Default Property Get Days() 
            DBLoad               
            Days = m_lngDays
        End Property


        Public Property Get EnteredBy() 
            DBLoad
            Set EnteredBy = m_objEnteredBy                         
        End Property

                        
        Public Property Let DateEntered(ByVal datValue)
            m_blnLoaded = True
            m_datEntered = datValue
        End Property
        Public Property Get DateEntered()
            DBLoad             
            DateEntered = m_datEntered
        End Property        

        
        Public Property Let Comments(ByVal strValue)
            m_blnLoaded = True
            m_strComments = strValue
        End Property
        Public Property Get Comments() 
            DBLoad                
            Comments = m_strComments
        End Property
        
        
        Public Property Let IsPreArranged(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsPreArranged = blnValue
        End Property
        Public Property Get IsPreArranged()
            DBLoad
            IsPreArranged = m_blnIsPreArranged
        End Property
        

        Public Property Get AdminCarryOverFormErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage

            m_strErrorFieldList = ""
            locblnValid = True

            '*** Check for valid year ***
            If Year = "" then
                locstrErrorMessage = locstrErrorMessage & "  - The Year must be entered.\n"
                m_strErrorFieldList = m_strErrorFieldList & "Year;"
                locblnValid = False
            ElseIf (len(Year) <> 4) or (mGetSafeLongInteger(Year,0) = 0) then
                locstrErrorMessage = locstrErrorMessage & "  - The Year must be in yyyy format (e.g. 2001).\n"
                m_strErrorFieldList = m_strErrorFieldList & "Year;"
                locblnValid = False
            ElseIf Year < CONST_FIRST_YEAR_SYSTEM_ACTIVE then
                locstrErrorMessage = locstrErrorMessage & "  - The Year cannot be before " & CONST_FIRST_YEAR_SYSTEM_ACTIVE & ".\n"
                m_strErrorFieldList = m_strErrorFieldList & "Year;"
                locblnValid = False
            End If
                
            '*** Check for valid days ***
            If Days = "" then
                locstrErrorMessage = locstrErrorMessage & "  - The Days must be entered.\n"
                m_strErrorFieldList = m_strErrorFieldList & "Days;"
                locblnValid = False
            ElseIf isnull(mGetSafeLongInteger(Days,null)) then
                locstrErrorMessage = locstrErrorMessage & "  - The Days entered must be a number.\n"
                m_strErrorFieldList = m_strErrorFieldList & "Days;"
                locblnValid = False
            End If

            '*** Check for comments ***
            If len(Comments) > 100 then
                locstrErrorMessage = locstrErrorMessage & "  - Comments can be up to a maximum of 100 characters (" & len(Comments) & " entered).\n"
                m_strErrorFieldList = m_strErrorFieldList & "Comments;"
                locblnValid = False
            End If
            
            If locblnValid = True then
                AdminCarryOverFormErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct all errors and re-submit the form."
                AdminCarryOverFormErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        
        End Property

        
        Public Property Get AdminCarryOverFormIsValid()
            If AdminCarryOverFormErrorMessage = "" then
                AdminCarryOverFormIsValid = True
            Else
                AdminCarryOverFormIsValid = False
            End If
        End Property
        

        Public Property Get IsInvalidField(ByVal strFieldName)
            If instr(m_strErrorFieldList,strFieldName & ";") then
                IsInvalidField = True
            Else
                IsInvalidField = False
            End If
        End Property


        Private Sub DBLoad()
            Dim locCmd
            Dim locCmdRevoke           
            Dim locParam
            Dim locRS
            Dim locblnLoadFromID
			            
            If m_blnLoaded then
                Exit Sub
            End If
            
            '*** If we have no EE.WWID then we can't find our CarryOver object ***
            If typename(m_objEE) <> "cObjUser" and m_lngID = 0 then
                Exit Sub
            End If
            
            If typename(m_objEE) <> "cObjUser" then
                locblnLoadFromID = True
            ElseIf m_objEE.WWID <> "" then
                locblnLoadFromID = False
            Else
                Exit Sub
            End If

            m_blnLoaded = True

            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            If locblnLoadFromID then
                locCmd.CommandText = "usp_carryover_by_ID"
                Set locParam = locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)
                locCmd.Parameters.Append locParam
            Else
                locCmd.CommandText = "usp_carryover_by_WWID_Year"
                Set locParam = locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, EE.WWID)
                locCmd.Parameters.Append locParam
                Set locParam = locCmd.CreateParameter("lngYear", adInteger, adParamInput, , Year)
                locCmd.Parameters.Append locParam
                
            End If
            
            Set locRS = locCmd.Execute			

            If not locRS.eof then
                If typename(m_objEE) <> "cObjUser" then
                    Set m_objEE = new cObjUser
                    m_objEE.WWID = locRS("strEEWWID")
                    m_lngYear = locRS("lngYear")
                Else
                    m_lngID = locRS("lngID")
                End If
                m_lngDays = locRS("lngDays")
                m_objEnteredBy.WWID = locRS("strEnteredByWWID")
                m_datEntered = locRS("datEntered")
                m_strComments = locRS("strComments")
                m_blnIsPreArranged = locRS("blnIsPreArranged")
            Else
                SetEmpty
            End If
            
            locRS.Close
            
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
			
        End Sub


        '*** LOAD ADMIN CARRY OVER FROM FORM ***
        Public Function LoadAdminCarryOverFromForm()
            ID = request("itemid")
            EE.WWID = request("ee")
            Year = request("fldlngyear")
            Days = request("flddbldays")
            EnteredBy.WWID = objCurrentUser.WWID
            DateEntered = date()
            Comments = request("fldstrcomments")
            If request("fldcbotype") = "Pre-Arranged" then
                IsPreArranged = True
            Else
                IsPreArranged = False
            End If
        End Function
        
        
        Public Function Save
            Dim loclngResult
            Dim locCmdRevoke
            Dim locCmd
            Dim loclngReturnValue

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc
            loclngReturnValue = 0

            locCmd.CommandText = "usp_save_carryover"
            locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
            locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, trim(EE.WWID))
            locCmd.Parameters.Append locCmd.CreateParameter("lngYear", adInteger, adParamInput, , Year)
            locCmd.Parameters.Append locCmd.CreateParameter("lngDays", adDouble, adParamInput, , Days)
            locCmd.Parameters.Append locCmd.CreateParameter("strEnteredByWWID", adWChar, adParamInput, 8, EnteredBy.WWID)
            locCmd.Parameters.Append locCmd.CreateParameter("datEntered", adDBTimeStamp, adParamInput, , DateEntered)
            locCmd.Parameters.Append locCmd.CreateParameter("strComments", adWChar, adParamInput, 100, trim(Comments))
            locCmd.Parameters.Append locCmd.CreateParameter("blnPreArranged", adBoolean, adParamInput, , IsPreArranged)

            on error resume next
            
            locCmd.Execute
            
            loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
            
            on error goto 0

            Set locCmd = nothing

            Save = loclngReturnValue
                        
        End Function

    End Class


        '*** OBJECT COLLECTION ***
        Class cObjCollection
                '=================================================================
                'Description:   Base collection object
                'Properties:
                '       Count
                '       Item(ByVal loclngIndex)
                '
                'Methods:
                '       Add(ByVal locObject)            Public
                '       Clear
                '=================================================================

                Private m_Objects()
                Public Count

                Public m_lngtempobjid
                Private Sub Class_Initialize()
                    glbObjectCounter = glbObjectCounter + 1
                    m_lngtempobjid = glbObjectCounter
                    mDebugPrint "Initializing cObjCollection (" & m_lngtempobjid & "): " & timer & "<br>"
                    Count = 0
                    ReDim m_Objects(0)
                End Sub
                
                
                Private Sub Class_Terminate()
                    glbObjectTerminateCounter = glbObjectTerminateCounter + 1
                    'mDebugPrint "Terminating cObjCollection (" & m_lngtempobjid & ") (Count = " & Count & "): " & timer & "<br>"
                    Dim loclngCounter
                    If Count > 0 Then
                        For loclngCounter = 0 To Count
                            FreeFromMemory loclngCounter
                        Next
                    End If
                    erase m_Objects
                    'Set m_Objects = Nothing
                End Sub


                Public Default Property Get Item(ByVal loclngIndex)
                        If Count = 0 Then
                                Set Item = Nothing
                                Exit Property
                        End If
                        loclngIndex = mGetSafeLongInteger(loclngIndex, 0)
                        If loclngIndex > 0 And loclngIndex <= Count Then
                                Set Item = m_Objects(loclngIndex)
                        Else
                                Set Item = Nothing
                        End If
                End Property
                
                
                Public Sub Add(ByVal locObject)
                        Count = Count + 1
                        ReDim Preserve m_Objects(Count)
                        Set m_Objects(Count) = locObject
                        
                        if isObject(m_Objects(0)) then
                            if not isNULL(m_Objects(0).ID) then
                          '      response.Write "Obj[0].ID = " & m_Objects(0).ID & "<br>"
                            end if
                        End if
                End Sub
                
                
                Public Sub Clear()
                    Count = 0
                    Redim m_Objects(Count)
                End Sub
                
                
                Public Sub FreeFromMemory(ByVal loclngIndex)
                    Set m_Objects(loclngIndex) = Nothing
                End Sub

        End Class
        
        
    '*** ELP Instance Object ***
    Class cObjELPInstance
        '=================================================================
        'Description:   Contains definition of an Instance of ELP.
        '               (details belonging to a single ELP "policy" from
        '               Activation to Completion - i.e. ELP had been taken
        '
        'Properties:                        Type:           Perm:       Source:     Properties Required:
        'y      ELPID                       lng             RW          DB
        '       EE                          objUser         R           DB (WWID)
        'y      DateActivated               dat             R           DB
        'y      ActivatedBy                 objUser         R           DB(WWID)
        'y      DaysBankedOnActivation      lng             R           DB
        'y      Reliefs                     col(ELPRelief)  R           DB
        'y      LeavePeriod                 objLeavePeriod  R           DB
        'y      CurrentYear                 dat             RW          Virtual
        'y      ReliefsToCurrentYear        lng             R           Calc        (Reliefs, CurrentYear)
        'y      ReliefInYear()              bln             R           Calc        (Year)
        'y      BankYearsToCurrentYear      lng             R           Calc        (ReliefsToCurrentYear)
        'y      DaysBanked                  lng             R           Calc        (DaysBankedOnActivation,BankYearsToCurrentYear,DaysBankedPerYear)
        'y      DaysBankedPerYear           lng             R           Calc        (CONSTANT)      
        'y      DaysBankedInCurrentYear     lng             R           Calc
        '       DaysBankedToCurrentYear     lng             R           Calc        (FORMULA: DaysBankedOnActivation+(NumOfYears-NumOfReliefYears)*DaysBankedPerYear)
        'y      IsUsed                      bln             R           Calc        (LeavePeriod.StartDate)
        'y      MaturityDate                dat             R           Calc        (DateActivated, ExpectedYearsToMature)
        'y      ExpiryDate                  dat             R           Calc        (DateActivated, ExpectedYearsToExpire)
        'y      IsMatured                   bln             R           Calc        (MaturityDate, CurrentYear)
        'y      IsExpired                   bln             R           Calc        (ExpiryDate, IsUsed, CurrentYear)
        'y      DaysBankedPerYear           lng             R           Calc        (CONST_ELP_DAYS_BANKED_PER_YEAR)
        'y      TargetDays                  lng             R           Calc        (CONST_ELP_TARGET_DAYS)
        'y      StandardYearsToExpire       lng             R           Calc        (TargetDays,DaysBankedPerYear)
        'y      StandardYearsToMature       lng             R           Calc        (TargetDays,DaysBankedOnActivation,DaysBankedPerYear)
        'y      ExpectedYearsToExpire       lng             R           Calc        (StandardYearsToExpire, Reliefs.Count)
        'y      ExpectedYearsToMature       lng             R           Calc        (StabdardYearsToMature, Reliefs.Count)
        'y      Status                      str (CONST)     R           Calc
        'y      AdminActivateELPFormIsValid bln             R           Calc
        'y      IsInvalidField()            bln             R           Calc
        'y      AdminActivateELPFormErrorMessage    str     R           Calc
        '
        'Methods:
        '   DBLoad
        '   DBLoadELPReliefs
        '   LoadAdminActivateELPFromForm
        '
        '=================================================================
        Private m_lngELPID
        Private m_objEE
        Private m_datDateActivated
        Private m_dblTargetDays
        Private m_objActivatedBy
        Private m_dblDaysBankedOnActivation
        Private m_colReliefs
        Private m_objLeavePeriod
        Private m_lngCurrentYear

        Private m_blnLoaded
        Private m_blnReliefsLoaded
        Private m_blnLeaveLoaded
        Private m_strErrorFieldList
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjELPInstance (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            SetNotLoaded
            m_lngCurrentYear = datepart("yyyy",Date)
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjELPInstance (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = nothing
            Set m_objActivatedBy = nothing
            Set m_colReliefs = nothing
            Set m_objLeavePeriod = nothing
        End Sub


        Private Sub SetNotLoaded()
            m_blnLoaded = False
            m_blnReliefsLoaded = False
            m_blnLeaveLoaded = False
        End Sub
        
        
        Private Sub SetEmpty()
            m_lngELPID = 0  
            'm_datDateActivated  = mFormatDate(Date(),"medium")
            'm_dblTargetDays = CONST_ELP_TARGET_DAYS
            Set m_colReliefs = nothing
            Set m_objEE = nothing
            Set m_objEE = new cObjUser
            Set m_objActivatedBy = nothing
            Set m_objActivatedBy = new cObjUser
            'm_dblDaysBankedOnActivation = 5
        End Sub

        
        Public Property Let ELPID(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            SetEmpty
            SetNotLoaded
            m_lngELPID = lngValue 
        End Property
        
        Public Property Get ELPID()
            DBLoad
            'Response.Write "m_lngELPID(Class ELPInstance Let ELPID): " & m_lngELPID & "<BR>"          
            ELPID = m_lngELPID
        End Property


        Public Property Set EE(ByRef objUser)
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID
        End Property
        Public Property Get EE()
            DBLoad
            Set EE = m_objEE
        End Property        


        Public Property Let DateActivated(ByVal datValue)
            m_blnLoaded = True
            m_datDateActivated = datValue
        End Property
        Public Property Get DateActivated() 
            DBLoad           
            DateActivated = m_datDateActivated
        End Property        

        
        Public Property Get ActivatedBy() 
            DBLoad
            Set ActivatedBy = m_objActivatedBy
        End Property        


        Public Property Let TargetDays(ByVal dblValue)
            m_blnLoaded = True
            m_dblTargetDays = dblValue
        End Property
        Public Property Get TargetDays()
            DBLoad
            'response.write ("ELP Target Days: " & m_dblTargetDays)
            TargetDays = m_dblTargetDays
        End Property


        Public Property Let DaysBankedOnActivation(ByVal lngValue)
            m_blnLoaded = True
            m_dblDaysBankedOnActivation = mGetSafeLongInteger(lngValue,0)
        End Property
        Public Property Get DaysBankedOnActivation() 
            DBLoad               
            DaysBankedOnActivation = m_dblDaysBankedOnActivation
        End Property    


        Public Property Get Reliefs() 
            DBLoadELPReliefs
            Initialise_Reliefs
            Set Reliefs = m_colReliefs      
        End Property


        Public Property Get LeavePeriod()
            DBLoad
            Initialise_LeavePeriod
            Set LeavePeriod = m_objLeavePeriod  
        End Property


        Public Property Let CurrentYear(ByVal lngValue)
            If mGetSafeLongInteger(lngValue,-1) <> -1 then
                m_lngCurrentYear = lngValue
            Else
                m_lngCurrentYear = null
            End If
        End Property
        Public Property Get CurrentYear
            CurrentYear = m_lngCurrentYear
        End Property
        

        Public Property Get ReliefsToCurrentYear()
            Dim loclngIndex
            Dim loclngCounter
            loclngIndex = 0
            loclngCounter = 0
            While loclngIndex < Reliefs.Count
                loclngIndex = loclngIndex + 1
                If Reliefs.Item(loclngIndex).Year <= CurrentYear then
                    loclngCounter = loclngCounter + 1
                End If
            Wend
            ReliefsToCurrentYear = loclngCounter
        End Property
        

        Public Property Get ReliefInYear(ByVal lngYear)
            Dim loclngIndex
            loclngIndex = 0
            While loclngIndex < Reliefs.Count
                loclngIndex = loclngIndex + 1
                If Reliefs.Item(loclngIndex).Year = lngYear then
                    ReliefInYear = True
                    Exit Property
                End If
            Wend
            ReliefInYear = False
        End Property
                
                
        Public Property Get BankYearsToCurrentYear()
            Dim loclngYears
            Dim loclngYear
            If CurrentYear > datepart("yyyy",MaturityDate) then
                loclngYear = datepart("yyyy",MaturityDate)
            Else
                loclngYear = CurrentYear
            End If
            If IsDate(DateActivated) then
                loclngYears = mCountWholeYears(mFirstDayOfYear(DatePart("yyyy",DateActivated)),mFirstDayOfYear(loclngYear))
                loclngYears = loclngYears - ReliefsToCurrentYear
            Else
                loclngYears = 0
            End If
            if loclngYears > ExpectedYearsToMature then loclngYears = ExpectedYearsToMature
            BankYearsToCurrentYear = loclngYears
        End Property
        

        Public Property Get DaysBanked()
            DaysBanked = DaysBankedOnActivation + (BankYearsToCurrentYear * DaysBankedPerYear)
        End Property

        
        Public Property Get DaysBankedInCurrentYear()
            If IsDate(DateActivated) then
                If CurrentYear = DatePart("yyyy",DateActivated) then
                    'Response.Write ("CurrentYear:" & CurrentYear)
                    'Response.Write ("DateActivated:" & DateActivated)
                    DaysBankedInCurrentYear = DaysBankedOnActivation
                    'Response.Write ("DaysBankedOnActivation:" & DaysBankedOnActivation)
                ElseIf ReliefInYear(CurrentYear) then
                    DaysBankedInCurrentYear = 0
                ElseIf CurrentYear <= (DatePart("yyyy",DateActivated) + ExpectedYearsToMature) _
                    AND CurrentYear > DatePart("yyyy",DateActivated) then
                    DaysBankedInCurrentYear = DaysBankedPerYear
                Else
                    DaysBankedInCurrentYear = 0
                End If
            Else
                DaysBankedInCurrentYear = 0
            End If
        End Property
            
        Public Property Get DaysBankedToCurrentYear()
            
                If CurrentYear < DatePart("yyyy",DateActivated) then
                    DaysBankedToCurrentYear = 0
                    'response.write ("1") & "<BR>"
                    Exit Property
                End If
                
                
                If CurrentYear = DatePart("yyyy",DateActivated) then 
                    DaysBankedToCurrentYear = DaysBankedOnActivation
                    'response.write ("2") & "<BR>"
                    Exit Property
                End If
                
                
                If CurrentYear >=  DatePart("yyyy",MaturityDate) then 
                    DaysBankedToCurrentYear = TargetDays
                    'response.write ("3") & "<BR>"
                    Exit Property
                End If
                                                
                
                Dim NumOfYears
                NumOfYears =  Clng(CurrentYear) - CLng(DatePart("yyyy",DateActivated))
    
                Dim NumOfReliefYears
                NumOfReliefYears = 0
                
                Dim i
                For i = Clng(DatePart("yyyy",DateActivated)) to CLng(CurrentYear)
                    If ReliefInYear(i) then
                        NumOfReliefYears = NumOfReliefYears + 1
                    End If
                Next
                            
                DaysBankedToCurrentYear = DaysBankedOnActivation + (NumOfYears - NumOfReliefYears) * DaysBankedPerYear
                                
        End Property
            
        Public Property Get IsUsed()
            If LeavePeriod.StartDate <> "" then
                IsUsed = True
            Else
                IsUsed = False
            End If
        End Property


        Public Property Get MaturityDate()
            Dim locdatDate
            if IsDate(DateActivated) then               
                locdatDate = DateActivated
                locdatDate = dateadd("d",1,locdatDate)
                locdatDate = dateadd("yyyy",ExpectedYearsToMature,locdatDate)                                       
            else
                locdatDate = ""
            end if
            MaturityDate = locdatDate
        End Property


        Public Property Get ExpiryDate() 
            Dim locdatDate
            Dim loclngYears
            if IsDate(DateActivated) then
                locdatDate = DateActivated
                locdatDate = dateadd("d",1,locdatDate)
                locdatDate = dateadd("yyyy",ExpectedYearsToExpire,locdatDate)
            else
                locdatDate = ""
            end if
            ExpiryDate = locdatDate
        End Property


        Public Property Get IsMatured()
            If mFirstDayOfYear(CurrentYear) >= MaturityDate then
                IsMatured = True
            Else
                IsMatured = False
            End If
        End Property 


        Public Property Get IsExpired()
            If (Not IsUsed) And mLastDayOfYear(CurrentYear) > ExpiryDate then
                IsExpired = True
            Else
                IsExpired = False
            End If
        End Property
        

        Public Property Get DaysBankedPerYear()
            DaysBankedPerYear = TargetDays / (TargetDays / DaysBankedOnActivation)
        End Property
        
        
        Public Property Get StandardYearsToMature()
            StandardYearsToMature = (TargetDays - DaysBankedOnActivation) / DaysBankedPerYear
        End Property
        
        
        Public Property Get StandardYearsToExpire()         
                StandardYearsToExpire = StandardYearsToMature + 5
        End Property
        

        Public Property Get ExpectedYearsToMature()
            ExpectedYearsToMature = StandardYearsToMature + Reliefs.Count
        End Property
        

        Public Property Get ExpectedYearsToExpire()
            ExpectedYearsToExpire = StandardYearsToExpire + Reliefs.Count
        End Property


        Public Property Get Status()
            If IsExpired then
                Status = CONST_ELP_STATUS_EXPIRED
            ElseIf IsUsed then
                Status = CONST_ELP_STATUS_USED
            ElseIf IsMatured then
                Status = CONST_ELP_STATUS_MATURED
            Else
                Status = CONST_ELP_STATUS_ACTIVE
            End If
        End Property

        
        Public Property Get AdminActivateELPFormErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage

            m_strErrorFieldList = ""
            locblnValid = True

            '*** Check to see if the employee alrady has an active ELP instance ***
            If ELPID = 0 and EE.AnnualVacation.HasActiveELP then
                locstrErrorMessage = locstrErrorMessage & "  - The employee already has an active ELP.\n"
                locblnValid = False
            Else

                '*** Check DateActivated exists, is a valid date and is in the right format.***
                If DateActivated = "" then
                    locstrErrorMessage = locstrErrorMessage & "  - The Activation Date must be entered.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DateActivated;"
                    locblnValid = False
                ElseIf not IsDate(DateActivated) then
                    locstrErrorMessage = locstrErrorMessage & "  - The Activation Date is invalid.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DateActivated;"
                    locblnValid = False
                ElseIf not mIsFormattedDate(DateActivated) then
                    locstrErrorMessage = locstrErrorMessage & "  - The Activation Date should be in (dd mmm yyyy) format - (e.g. 01 Jan 2001).\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DateActivated;"
                    locblnValid = False
                End If
                
                '*** Check that the DaysBankedOnActivation is more than 0. ***
                If DaysBankedOnActivation <= 0 then
                    locstrErrorMessage = locstrErrorMessage & "  - You must enter a positive value for Days Banked On Activation.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DaysBankedOnActivation;"
                    locblnValid = False
                End If

                '*** Check that the Target Days figure is more than 0. ***
                If TargetDays <= 0 then
                    locstrErrorMessage = locstrErrorMessage & "  - You must enter a positive value for Total Days to be Banked.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "TargetDays;"
                    locblnValid = False
                End If

                '*** Check that the Target Days figure is more than or equal to DaysBankedOnActivation. ***
                If TargetDays <= DaysBankedOnActivation then
                    locstrErrorMessage = locstrErrorMessage & "  - The Total Days to be Banked figure must be greater than the number of Days Banked On Activation.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DaysBankedOnActivation;TargetDays;"
                    locblnValid = False
                End If

                '*** Check that the DaysBankedOnActivation is more than 0. ***
                If DaysBankedOnActivation > 0 and DaysBankedOnActivation <=5 then
                    If ((TargetDays - DaysBankedOnActivation) / DaysBankedOnActivation) <> Int((TargetDays - DaysBankedOnActivation) / DaysBankedOnActivation) then
                        locstrErrorMessage = locstrErrorMessage & "  - Total Days To Be Banked must be a multiple of Days Banked On Activation" 'divisible by " & CONST_ELP_DAYS_BANKED_PER_YEAR & ".\n"
                        m_strErrorFieldList = m_strErrorFieldList & "DaysBankedOnActivation;TargetDays;"
                        locblnValid = False
                    End If
                else
                    locstrErrorMessage = locstrErrorMessage & "  - Days Banked on Activation must be from 1 up to 5.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DaysBankedOnActivation;TargetDays;"
                    locblnValid = False
                End if
            End If

            If locblnValid = True then
                AdminActivateELPFormErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct all errors and re-submit the form."
                AdminActivateELPFormErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        End Property


        Public Property Get AdminActivateELPFormIsValid()
            If AdminActivateELPFormErrorMessage = "" then
                AdminActivateELPFormIsValid = True
            Else
                AdminActivateELPFormIsValid = False
            End If
        End Property
        

        Public Property Get IsInvalidField(ByVal strFieldName)
            If instr(m_strErrorFieldList,strFieldName & ";") then
                IsInvalidField = True
            Else
                IsInvalidField = False
            End If
        End Property

        
        Private Sub DBLoad()
            Dim locCmdRevoke
            Dim locCmd
            Dim locParam
            Dim locRS

            If m_blnLoaded then
                
                Exit Sub
            End If              

            '*** If we have no m_lngELPID set up, we can't find our ee details - so exit the routine.
            If m_lngELPID = 0 then
                exit sub
            End if
            
            m_blnLoaded = True

            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_elpinstance"
            locCmd.CommandType = adCmdStoredProc
            Set locParam = locCmd.CreateParameter("lngID", adInteger, adParamInput, , ELPID)
            locCmd.Parameters.Append locParam
            
            Set locRS = locCmd.Execute

            If not locRS.eof then
                
                m_datDateActivated = locRS("datActivated")
                ActivatedBy.WWID = locRS("strActivatedByWWID")
                m_dblDaysBankedOnActivation = locRS("dblInitialDaysBanked")
                m_dblTargetDays = mGetSafeLongInteger(locRS("dblTargetDays"),CONST_ELP_TARGET_DAYS)
                '*** LOAD LEAVE PERIOD IF IT EXISTS ***
                If mGetSafeLongInteger(locRS("LeavePeriodID"),0) <> 0 then
                    
                    Initialise_LeavePeriod
                    m_objLeavePeriod.ID = locRS("LeavePeriodID")
                    m_objLeavePeriod.EE.WWID = locRS("strEEWWID")
                    m_objLeavePeriod.Approver.WWID = locRS("strApproverWWID")
                    m_objLeavePeriod.LeaveType.ID = locRS("lngLeaveTypeID")
                    m_objLeavePeriod.StartDate = locRS("datStartDate")
                    m_objLeavePeriod.StartTime = locRS("strStartTime")
                    m_objLeavePeriod.EndDate = locRS("datEndDate")
                    m_objLeavePeriod.EndTime = locRS("strEndTime")
                    m_objLeavePeriod.DateRaised = locRS("datRaised")
                    m_objLeavePeriod.DateApproved = locRS("datApproved")
                    m_objLeavePeriod.DateRejected = locRS("datRejected")
                    m_objLeavePeriod.DateCancelRequested = locRS("datCancelRequested")
                    m_objLeavePeriod.DateCancelApproved = locRS("datCancelApproved")
                    m_objLeavePeriod.DateCancelRejected = locRS("datCancelRejected")
                    m_objLeavePeriod.RequestComments = locRS("strRequestComments")
                    m_objLeavePeriod.ResponseComments = locRS("strResponseComments")
                    m_objLeavePeriod.ELPID = locRS("ELPID")
                End If
            Else
                SetEmpty
            End If
            
            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
        End Sub
        
        
        Private Sub DBLoadELPReliefs()
            Dim locCmd
            Dim locParam
            Dim locRS
            Dim locCmdRevoke

            Dim fldlngID
            Dim fldlngYear
            Dim fldstrEnteredByWWID
            Dim flddatEntered
            Dim fldstrComments
            
            Dim locobjELPRelief
            
            '*** Already loaded - then exit ***
            If m_blnReliefsLoaded then
                Exit Sub
            End If

            m_blnReliefsLoaded = True
            
            '*** If we have no WWID set up, we can't find our direct reports - so exit the routine.
            If ELPID = 0 then
                exit sub
            End if
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_elpinstance_elpreliefs"
            locCmd.CommandType = adCmdStoredProc
                        
            Set locParam = locCmd.CreateParameter("lngELPID", adInteger, adParamInput, , ELPID)
            locCmd.Parameters.Append locParam
            
            Set locRS = locCmd.Execute

            Set fldlngID = locRS("lngID")
            Set fldlngYear = locRS("lngYear")
            Set fldstrEnteredByWWID = locRS("strEnteredByWWID")
            Set flddatEntered = locRS("datEntered")
            Set fldstrComments = locRS("strComments")
            
            While not locRS.eof
                Set locobjELPRelief = new cObjELPRelief
                locobjELPRelief.ID = fldlngID
                locobjELPRelief.ELPID = ELPID
                locobjELPRelief.Year = fldlngYear
                locobjELPRelief.EnteredBy.WWID = fldstrEnteredByWWID
                locobjELPRelief.DateEntered = flddatEntered
                locobjELPRelief.Comments = fldstrComments
                Reliefs.Add locobjELPRelief
                Set locobjELPRelief = nothing
                locRS.movenext
            Wend

            Set fldlngID = nothing
            Set fldlngYear = nothing
            Set fldstrEnteredByWWID = nothing
            Set flddatEntered = nothing
            Set fldstrComments = nothing

            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
        End Sub
        
        
        Public Sub LoadAdminActivateELPFromForm()
            ELPID = mGetSafeLongInteger(request.form("itemid"),0)
            if ELPID <> 0 then
                EE.WWID = request.form("ee")
            End If
            DateActivated = request.form("fldstrActivationDate")
            DaysBankedOnActivation = mGetSafeLongInteger(request.form("flddblDaysBankedOnActivation"),0)
            ActivatedBy.WWID = objCurrentUser.WWID
            TargetDays = mGetSafeLongInteger(request.form("flddblTargetDays"),0)
        End Sub
        
        Public Sub LoadAdminSQLAddForm()
			
        End Sub


        Public Function Save()
            Dim locCmd
            Dim loclngReturnValue

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            loclngReturnValue = 0
            
            Response.Write ("WWID:" & EE.WWID)
            Response.Write ("WWID:" & Request.Form("ee"))
            

            locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
            locCmd.Parameters.Append locCmd.CreateParameter("lngELPID", adInteger, adParamInput, , mGetSafeLongInteger(ELPID,0))            
            'locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, trim(EE.WWID))
            locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, trim(Request.Form("ee")))
            locCmd.Parameters.Append locCmd.CreateParameter("datActivated", adDBTimeStamp, adParamInput, , DateActivated)
            locCmd.Parameters.Append locCmd.CreateParameter("strActivatedBy", adWChar, adParamInput, 8, trim(ActivatedBy.WWID))
            locCmd.Parameters.Append locCmd.CreateParameter("dblInitialDaysBanked", adDouble, adParamInput, , DaysBankedOnActivation)
            locCmd.Parameters.Append locCmd.CreateParameter("dblTargetDays", adDouble, adParamInput, , TargetDays)

            locCmd.CommandText = "usp_save_elp_activation"

            on error resume next
                
            locCmd.Execute
                
            loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
                
            on error goto 0
                
            if loclngReturnValue <> 0 then
                m_lngELPID = loclngReturnValue
            End If

            Set locCmd = nothing

            Save = loclngReturnValue
        
        End Function
        
        
        '**** INITIALISE RELIEFS ****
        Private Sub Initialise_Reliefs
            If typename(m_colReliefs) <> "cObjCollection" then
                Set m_colReliefs = new cObjCollection
            End If
        End Sub
        
        
        '**** INITIALISE LEAVE PERIOD ****
        Private Sub Initialise_LeavePeriod
            If typename(m_objLeavePeriod) <> "cObjLeavePeriod" then
                Set m_objLeavePeriod = new cObjLeavePeriod
            End If
        End Sub
    End Class


    '*** ELP Relief Object ***
    Class cObjELPRelief
        '=================================================================
        'Description:   Contains definition of an ELP Relief.
        'Properties:                Type:   Perm:       Source:     Properties Required:
        'y      ID                  lng     RW          DB
        'y      ELPID               lng     RW          DB
        'y      Year                lng     RW          DB
        'y      EnteredBy           objUser RW          DB(WWID)
        'y      DateEntered         dat     RW          DB
        'y      Comments            str     RW          DB
        '
        'Methods:
        '
        '=================================================================
        Private m_lngID
        Private m_lngELPID 
        Private m_lngYear
        Private m_objEnteredBy
        Private m_datEntered 
        Private m_strComments
        Private m_lngAction

        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjELPRelief (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            Set m_objEnteredBy = new cObjUser
            m_blnLoaded = False
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjELPRelief (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEnteredBy = nothing                        
        End Sub


        Private Sub SetEmpty()
            m_lngID = 0
            m_lngELPID = 0
            m_lngYear = 0
            m_datEntered = ""
            m_strComments = ""
            m_lngAction = 0
        End Sub


        Public Property Let ID(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            If lngValue <> m_lngID then
                m_blnLoaded = False
                SetEmpty
                m_lngID = lngValue
            End If
        End Property
        Public Property Get ID()
            DBLoad
            ID = m_lngID
        End Property


        Public Property Let ELPID(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            If lngValue <> 0 then
                m_blnLoaded = True
                m_lngELPID = lngValue
            End If
        End Property
        Public Property Get ELPID() 
            DBLoad                 
            ELPID = m_lngELPID
        End Property

        
        Public Property Let Year(ByVal lngValue)
            lngValue = mGetSafeLongInteger(lngValue,0)
            m_blnLoaded = True
            m_lngYear = lngValue
        End Property
        Public Property Get Year()
            DBLoad               
            Year = m_lngYear
        End Property                

        
        Public Property Get EnteredBy()
            DBLoad
            Set EnteredBy = m_objEnteredBy
        End Property        

        
        Public Property Let DateEntered(ByVal datValue)
            If IsDate(datValue) then
                m_datEntered = datValue
            Else
                m_datEntered = ""
            End If
            m_blnLoaded = True
        End Property
        Public Property Get DateEntered()
            DBLoad                
            DateEntered = m_datEntered
        End Property


        Public Property Let Comments(ByVal strValue)
            m_strComments = strValue
            m_blnLoaded = True
        End Property                
        Public Property Get Comments()
            DBLoad       
            Comments = m_strComments
        End Property        
        

        Private Sub DBLoad()
            Dim locCmd
            Dim locParam
            Dim locCmdRevoke
            Dim locRS               
                        
            'We will check m_blnLoaded and load in the property values here.
            If m_blnLoaded then
                Exit Sub
            End If

            m_blnLoaded = true

            '*** If we have no ID then we can't find our ELPRelief object ***
            If ID = 0 then
                exit sub
            End if
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_elprelief"
            locCmd.CommandType = adCmdStoredProc
                        
            Set locParam = locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)
            locCmd.Parameters.Append locParam
            
            Set locRS = locCmd.Execute

            If not locRS.eof then
                Set m_objEnteredBy = new cObjUser
    
                m_lngELPID = locRS("lngELPID")
                m_lngYear = locRS("lngYear")
                m_datEntered = locRS("datEntered")
                m_objEnteredBy.WWID = locRS("strEnteredByWWID")

            End If
                                
            locRS.Close
            
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing                    
                                                        
        End Sub


        Public Sub AdminELPReliefLoadFromForm
            If request("btnSubmit") = "Add Relief" then
                Year = mGetSafeLongInteger(request("fldcboAddYear"),0)
                m_lngAction = CONST_ELP_RELIEF_ACTION_ADD
            Else
                Year = mGetSafeLongInteger(request("fldcboRemoveYear"),0)
                m_lngAction = CONST_ELP_RELIEF_ACTION_REMOVE
            End If
            m_lngELPID = mGetSafeLongInteger(request("elpid"),0)
            m_datEntered = Date()
            m_objEnteredBy.WWID = objCurrentUser.WWID
        End Sub
        
        
        Public Function Save()
            Dim loclngResult
            Dim locCmd
            Dim locCmdRevoke
            Dim loclngReturnValue

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc
            loclngReturnValue = 0

            locCmd.CommandText = "usp_save_elprelief"
            locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
            locCmd.Parameters.Append locCmd.CreateParameter("blnAdd", adBoolean, adParamInput, , m_lngAction)
            locCmd.Parameters.Append locCmd.CreateParameter("lngYear", adInteger, adParamInput, , Year)
            locCmd.Parameters.Append locCmd.CreateParameter("lngELPID", adInteger, adParamInput, , ELPID)
            locCmd.Parameters.Append locCmd.CreateParameter("datDateEntered", adDBTimeStamp, adParamInput, , DateEntered)
            locCmd.Parameters.Append locCmd.CreateParameter("strEnteredByWWID", adWChar, adParamInput, 8, trim(EnteredBy.WWID))

            on error resume next
            
            locCmd.Execute
            
            loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
            on error goto 0

            Set locCmd = nothing

            Save = loclngReturnValue
        End Function
        
        
    End Class
    
 
    Class cObjCompTime 
        '=================================================================
        'Description:   Contains definition of a Leave Period.
        'Properties:                    Type:           Perm:       Source:     Properties Required:
        'y  ID                          lng             RW          DB
        'y  EE                          objUser         RW          DB
        'y  DaysGranted                 lng             RW          DB
        'y  DaysTaken                   lng             R           Calc        EE (Uses the LeavePeriods)
        'y  DateGranted                 dat             RW          DB
        'y  DateRevoked                 dat             RW          DB
        'y  ExpiryDate                  dat             R           Calc
        'y  Reason                      str             RW          DB
        'y  Status                      str             RW          Calc        (Granted, Revoked, Expired, Used)
        '
        ' Methods:
        '   
        '   
        '
        '=================================================================
        
        Private m_lngID
        Private m_objEE
        Public m_lngDaysGranted
        Public m_lngDaysRevoked
        Public m_lngDaysBooked
        Public m_lngDaysTaken
        public Days_Booked2
        Private m_datDateGranted
        Private m_datDateRevoked
        Private m_strReason
        
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjCompTime (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            SetNotLoaded
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjCompTime (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE  = nothing
        End Sub
        
        Private Sub SetEmpty()
            m_lngID = 0
            Set m_objEE  = nothing
            Set m_objEE = new cObjUser
            m_lngDaysGranted = 0
            m_lngDaysRevoked = 0
            m_datDateGranted = ""
            m_datDateRevoked = 1
            Days_Booked2 = 0
            m_strReason = ""
            m_blnLoaded = False
        End Sub

        Private Sub SetNotLoaded()
            m_blnLoaded = False
        End Sub
        
        Public Property Let ID(ByVal varValue)
            varValue = mGetSafeLongInteger(varValue,0)
            If varValue <> 0 then
                If m_lngID <> varValue then
                    SetEmpty
                    SetNotLoaded
                    m_lngID = varValue
                End If
            End If
        End Property
        Public Property Get ID()
            DBLoad
            ID = m_lngID
        End Property
        
        Public Property Let EE(ByRef objUser)
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID
        End Property
        Public Property Get EE()
            DBLoad
            Set EE = m_objEE
        End Property      
                
          Public Property Get setDays()
            Dim loclngDaysBooked
            Dim locstrSQL
            Dim locRS
            loclngDaysBooked = 0
            Dim locobjLeavePeriod
            
            DBLoad
            
            ' TO DO: Calculate the Days Booked for this Comp Time
            
            locstrSQL = "SELECT * FROM tblCompTime " & _
                        "WHERE lngID=" & ID
                        'lngWWID=" & EE.WWID
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, glbConnection, adOpenStatic
            
            While Not locRS.EOF
            Set locobjLeavePeriod = new cObjLeavePeriod
            locobjLeavePeriod.ID = locRS("lngID")
            
         m_lngDaysGranted = locRS("lngDaysGranted")
         m_lngDaysRevoked = locRS("IngDaysRevoked")
         m_lngDaysBooked = locRS("lngDaysBooked")
         m_lngDaysTaken = locRS("Taken") 
                locRS.MoveNext
            Wend
       
            
            ' For now:
    '        DaysBooked = loclngDaysBooked
        End Property   
                
                
             
        Public Property Let DateGranted(ByVal datValue)
            m_datDateGranted = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateGranted() 
            DBLoad
            DateGranted = m_datDateGranted
        End Property
        
        Public Property Let DateRevoked(ByVal datValue)
            m_datDateRevoked = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateRevoked() 
            DBLoad
            DateRevoked = m_datDateRevoked
        End Property
        
        Public Property Get ExpiryDate()
            DBLoad
            ExpiryDate = DateAdd("d", CONST_COMP_TIME_EXPIRY_DAYS, DateGranted)
        End Property
        
        Public Property Let DaysGranted(ByVal lngDays)
            m_lngDaysGranted = lngDays
            m_blnLoaded = True
        End Property
        
        Public Property Get DaysGranted()
            DBLoad
            setDays
            DaysGranted = m_lngDaysGranted
         '   response.write " </br> Str 3745 Gra  ID " & ID & " = " & m_lngDaysGranted & " * R" & m_lngDaysRevoked  & " * " & m_lngDaysBooked  & " * " &  m_lngDaysTaken    
            
        End Property
        
        Public Property Get TotalGranted()
            DBLoad
            setDays
            TotalGranted = m_lngDaysGranted + m_lngDaysBooked + m_lngDaysTaken
            response.write " </br> Str 3795 " & TotalGranted
            
        End Property
        
        Public Property Let DaysRevoked(ByVal lngDays)
            m_lngDaysRevoked = lngDays
            m_blnLoaded = True
        End Property
        
        Public Property Get DaysRevoked2()
            DBLoad
             setDays
            DaysRevoked2 = m_lngDaysRevoked
        '    response.write " </br> obj " & m_lngDaysRevoked
           
        End Property
        
         Public Property Get DaysBooked2()
            DBLoad
             setDays
            DaysBooked2 = m_lngDaysBooked
        '     response.write " </br> Days booked " & m_lngDaysBooked
        End Property
     
        
         Public Property Get Taken()
            DBLoad
            setDays
            Taken = m_lngDaysTaken
        End Property
        
        Public Property Get DaysBooked()
            Dim loclngDaysBooked
            Dim locstrSQL
            Dim locRS
 '     response.write "obj 3706  "             
            loclngDaysBooked = 0
            
            DBLoad
            
            ' TO DO: Calculate the Days Booked for this Comp Time
            
            locstrSQL = "SELECT * FROM tblLeavePeriod " & _
                        "WHERE lngCompTimeID=" & ID
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, glbConnection, adOpenStatic
            
            While Not locRS.EOF
                Dim locobjLeavePeriod
                Set locobjLeavePeriod = new cObjLeavePeriod
                locobjLeavePeriod.ID = locRS("lngID")
                
             '   response.Write "[ID=" & locobjLeavePeriod.ID
                
                ' Check if this leave was confirmed as taken or not
                If locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_RAISED or _
                   locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
                   locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED Then
                    loclngDaysBooked = loclngDaysBooked + locobjLeavePeriod.Days
                    
            '        response.Write " Booked"
                End If
                
            '    response.Write "]<br>"
                
                locRS.MoveNext
            Wend
   '  response.write "obj 3736  "         
            
            ' For now:
            DaysBooked = loclngDaysBooked
        End Property
        
        Public Property Get DaysTaken()
            Dim loclngDaysTaken
            Dim locstrSQL
            Dim locRS
            
            loclngDaysTaken = 0
            
            DBLoad
 'response.write "obj 3750  "           
            '**** Calculate the days confirmed as taken from this Comp. Time
            ' 1. Load the leave periods this is associated with... 
            ' 2. Add up the days taken from each of them
            
            locstrSQL = "SELECT * FROM tblLeavePeriod " & _
                        "WHERE lngCompTimeID=" & ID
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, glbConnection, adOpenStatic
            
            While Not locRS.EOF
                Dim locobjLeavePeriod
                Set locobjLeavePeriod = new cObjLeavePeriod
                locobjLeavePeriod.ID = locRS("lngID")
                
             '   response.Write "[ID=" & locobjLeavePeriod.ID
                
                ' Check if this leave was confirmed as taken or not
                If locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CONFIRMED Then
                    loclngDaysTaken = loclngDaysTaken + locobjLeavePeriod.Days
                    
             '       response.Write " Taken"
                End If
                
             '   response.Write "]<br>"
                
                locRS.MoveNext
            Wend
            
            DaysTaken = loclngDaysTaken
        End Property
        
        Public Property Get DaysAvailable()
            DBLoad
            
            DaysAvailable = DaysGranted - DaysTaken - DaysBooked
        End Property
        
        Public Property Get Days_Booked()
            DBLoad
            
            Days_Booked =  DaysBooked
        End Property
        
         Public Property Get DaysRevoked()
            DBLoad
            DaysBooked
            DaysRevoked =  m_lngDaysRevoked
        End Property
        
        Public Property Let Reason(ByRef strValue)
            m_strReason = strValue
        End Property
        Public Property Get Reason()
            DBLoad
            Reason = m_strReason
        End Property      
        
        Public Property Get IsLoaded()
            DBLoad   
        
            IsLoaded = m_blnLoaded
        End Property
        
        Public Property Get Status()
            If DaysTaken = DaysGranted and DaysGranted <> 0 Then
                Status = CONST_COMP_TIME_STATUS_USED
            ElseIf IsDate(DateRevoked) Then
                Status = CONST_COMP_TIME_STATUS_REVOKED
            ElseIf IsDate(DateGranted) and (Date > DateAdd("d", CONST_COMP_TIME_EXPIRY_DAYS, DateGranted)) and DaysAvailable > 0 Then
                Status = CONST_COMP_TIME_STATUS_EXPIRED
            ElseIf DaysBooked2 <> 0 Then
                Status = CONST_COMP_TIME_STATUS_BOOKED
            ElseIf Taken <> 0 Then
                Status = CONST_COMP_TIME_STATUS_USED
                 ElseIf DaysRevoked2 <> 0 Then
                Status = CONST_COMP_TIME_STATUS_REVOKED
            ElseIf IsDate(DateGranted) Then
                Status = CONST_COMP_TIME_STATUS_GRANTED
            Else
                response.write "<h3>UNKNOWN STATUS</h3>"
                response.Write ""
            End If
        End Property
        
        Public Property Get CompTimeEmailBody()
            Dim locstrHTML
            
            locstrHTML = ""
            locstrHTML = locstrHTML & "<!DOCTYPE HTML><HTML>" & vbcrlf
			    locstrHTML = locstrHTML & "<head><title>E-vacation Email</title><style type='text/css'>.pageContentHeader,.pageContentTable,table{border-collapse:collapse;border-spacing:0}.pageContentHeader,.pageContentTable{width:800px;border:0;padding:3px;margin-right:auto}.th,th{font-size:16px;text-align:left}.th{color:#0062a8}.pageContentHeader{margin-left:0;margin-bottom:10px}body,h1,h2,h3,h4,h5,h6,html,td,th{font-family:intel-clear,tahoma,Helvetica,helvetica,Arial,sans-serif}body,html{color:#111}.pageContentTable{margin-left:15px;margin-bottom:0}td,th{padding:4px;color:#404040}th{color:#0062a8}h1,h2,h3,h4,h5,h6{-webkit-font-smoothing:antialiased;-webkit-user-select:none;margin-top:0;font-weight:100;clear:both;margin-bottom:10px}"
                locstrHTML = locstrHTML & "</style></head><BODY>" & vbcrlf
					locstrHTML = locstrHTML & "<img src='email.png' style='border: 1px solid #000000' alt='e-vacation'><br><br>" & vbcrlf
                    locstrHTML = locstrHTML & "<table class='pageContentHeader'><tr><td><h3>" & vbcrlf
                            locstrHTML = locstrHTML & "Compensatory Time Granted</h3></td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "Please do <u>not</u> reply to this e-mail. It was sent by the e-Vacation Application, which cannot receive e-mails." & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>This is a notification that you have been granted compensatory leave. It will expire in " 
                                locstrHTML = locstrHTML & CONST_COMP_TIME_EXPIRY_DAYS & " days.<br>"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<br>" & vbcrlf

                    locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan=2>" & vbcrlf
                                locstrHTML = locstrHTML & "<span class='th'>Details</span>"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Days Granted :</b>&nbsp;"
                                locstrHTML = locstrHTML & DaysGranted
                                 locstrHTML = locstrHTML & "<b>Days Revoked :</b>&nbsp;"
                                locstrHTML = locstrHTML & DaysRevoked
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "<td align=left>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Expiry Date:</b>&nbsp;"
                            locstrHTML = locstrHTML & mFormatDate(ExpiryDate, "medium with day") & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan='2'>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Manager:</b>&nbsp;"
                                locstrHTML = locstrHTML & "<a href=""mailto:"
                                locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.Email)
                                locstrHTML = locstrHTML & """ title=""Click here to send an e-mail to "
                                locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.FullName)
                                locstrHTML = locstrHTML & " now."">"
                                    locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.FullName)
                                locstrHTML = locstrHTML & "</a>" & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan=2>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Reason:</b><br>" & vbcrlf
                                if trim(Reason) <> "" then
                                    locstrHTML = locstrHTML & mHTMLEncode(Reason)
                                else
                                    locstrHTML = locstrHTML & "(none)"
                                end if
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>"
                            locstrHTML = locstrHTML & "<td colspan=2>"
                                locstrHTML = locstrHTML & "Please click "
								locstrHTML = locstrHTML & "<a href='https://" & Request.ServerVariables("http_host")
                            	locstrHTML =  locstrHTML & CONST_APPLICATION_PATH & "/requestleave.asp?m=ct&itemid=" & ID & "'>here</a>"
                                locstrHTML = locstrHTML & " to book this leave."
                            locstrHTML = locstrHTML & "</td>"
                        locstrHTML = locstrHTML & "<tr>"
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<br>" & vbcrlf
    
                    locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & CONST_APPLICATION_FOOTER & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                locstrHTML = locstrHTML & "</body>" & vbcrlf
            locstrHTML = locstrHTML & "</html>" & vbcrlf
            CompTimeEmailBody = locstrHTML
        End Property        
        
          Public Property Get CompTimeEmailBodyRevoke()
            Dim locstrHTML
            
            locstrHTML = ""
            locstrHTML = locstrHTML & "<!DOCTYPE HTML><HTML>" & vbcrlf
			    locstrHTML = locstrHTML & "<head><title>E-vacation Email</title><style type='text/css'>.pageContentHeader,.pageContentTable,table{border-collapse:collapse;border-spacing:0}.pageContentHeader,.pageContentTable{width:800px;border:0;padding:3px;margin-right:auto}.th,th{font-size:16px;text-align:left}.th{color:#0062a8}.pageContentHeader{margin-left:0;margin-bottom:10px}body,h1,h2,h3,h4,h5,h6,html,td,th{font-family:intel-clear,tahoma,Helvetica,helvetica,Arial,sans-serif}body,html{color:#111}.pageContentTable{margin-left:15px;margin-bottom:0}td,th{padding:4px;color:#404040}th{color:#0062a8}h1,h2,h3,h4,h5,h6{-webkit-font-smoothing:antialiased;-webkit-user-select:none;margin-top:0;font-weight:100;clear:both;margin-bottom:10px}"
                locstrHTML = locstrHTML & "</style></head><BODY>" & vbcrlf
					locstrHTML = locstrHTML & "<img src='email.png' style='border: 1px solid #000000' alt='e-vacation'><br><br>" & vbcrlf
                    locstrHTML = locstrHTML & "<table class='pageContentHeader'><tr><td><h3>" & vbcrlf
                            locstrHTML = locstrHTML & "Compensatory Time Revoked</h3></td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "Please do <u>not</u> reply to this e-mail. It was sent by the e-Vacation Application, which cannot receive e-mails." & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "This is a notification that your compensatory leave has been revoked. " 
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<br>" & vbcrlf

                    locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan=2>" & vbcrlf
                                locstrHTML = locstrHTML & "<span class='th'>Details</span>"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Days Revoked :</b>&nbsp;"
                                locstrHTML = locstrHTML & 1
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "<td align=left>" & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td align=left>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Manager:</b>&nbsp;"
                                locstrHTML = locstrHTML & "<a href=""mailto:"
                                locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.Email)
                                locstrHTML = locstrHTML & """ title=""Click here to send an e-mail to "
                                locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.FullName)
                                locstrHTML = locstrHTML & " now."">"
                                    locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.FullName)
                                locstrHTML = locstrHTML & "</a>" & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan=2>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Reason:</b><br>" & vbcrlf
                                if trim(Reason) <> "" then
                                    locstrHTML = locstrHTML & mHTMLEncode(Reason)
                                else
                                    locstrHTML = locstrHTML & "(none)"
                                end if
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>"
                            locstrHTML = locstrHTML & "<td colspan=2>"
                                locstrHTML = locstrHTML & "Please click "
								locstrHTML = locstrHTML & "<a href='https://" & Request.ServerVariables("http_host")
                            	locstrHTML =  locstrHTML & CONST_APPLICATION_PATH & "/requestleave.asp?m=ct&itemid=" & ID & "'>here</a>"
                                locstrHTML = locstrHTML & " to book this leave."
                            locstrHTML = locstrHTML & "</td>"
                        locstrHTML = locstrHTML & "<tr>"
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<br>" & vbcrlf
    
                    locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & CONST_APPLICATION_FOOTER & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                locstrHTML = locstrHTML & "</body>" & vbcrlf
            locstrHTML = locstrHTML & "</html>" & vbcrlf
            CompTimeEmailBodyRevoke = locstrHTML
        End Property        
        
        Public Property Get CompTimeEmailBodyGrant()
            Dim locstrHTML
            
            locstrHTML = ""
            locstrHTML = locstrHTML & "<!DOCTYPE HTML><HTML>" & vbcrlf
			    locstrHTML = locstrHTML & "<head><title>E-vacation Email</title><style type='text/css'>.pageContentHeader,.pageContentTable,table{border-collapse:collapse;border-spacing:0}.pageContentHeader,.pageContentTable{width:800px;border:0;padding:3px;margin-right:auto}.th,th{font-size:16px;text-align:left}.th{color:#0062a8}.pageContentHeader{margin-left:0;margin-bottom:10px}body,h1,h2,h3,h4,h5,h6,html,td,th{font-family:intel-clear,tahoma,Helvetica,helvetica,Arial,sans-serif}body,html{color:#111}.pageContentTable{margin-left:15px;margin-bottom:0}td,th{padding:4px;color:#404040}th{color:#0062a8}h1,h2,h3,h4,h5,h6{-webkit-font-smoothing:antialiased;-webkit-user-select:none;margin-top:0;font-weight:100;clear:both;margin-bottom:10px}"
                locstrHTML = locstrHTML & "</style></head><BODY>" & vbcrlf
					locstrHTML = locstrHTML & "<img src='email.png' style='border: 1px solid #000000' alt='e-vacation'><br><br>" & vbcrlf
                    locstrHTML = locstrHTML & "<table class='pageContentHeader'><tr><td><h3>" & vbcrlf
                            locstrHTML = locstrHTML & "Compensatory Time Granted</h3></td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<table>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "Please do <u>not</u> reply to this e-mail. It was sent by the e-Vacation Application, which cannot receive e-mails." & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "This is a notification that you have been granted compensatory leave. It will expire in " 
                                locstrHTML = locstrHTML & CONST_COMP_TIME_EXPIRY_DAYS & " days.<br>"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<br>" & vbcrlf

                    locstrHTML = locstrHTML & "<table>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td nowrap colspan=2>" & vbcrlf
                                locstrHTML = locstrHTML & "<span>Details</span>"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Days Granted :</b>&nbsp;"
                                locstrHTML = locstrHTML & 1
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Expiry Date:</b>&nbsp;"
                            locstrHTML = locstrHTML & mFormatDate(ExpiryDate, "medium with day") & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan='2'>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Manager:</b>&nbsp;"
                                locstrHTML = locstrHTML & "<a href=""mailto:"
                                locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.Email)
                                locstrHTML = locstrHTML & """ title=""Click here to send an e-mail to "
                                locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.FullName)
                                locstrHTML = locstrHTML & " now."">"
                                    locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.FullName)
                                locstrHTML = locstrHTML & "</a>" & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan=2>" & vbcrlf
                                locstrHTML = locstrHTML & "<b>Reason:</b><br>" & vbcrlf
                                if trim(Reason) <> "" then
                                    locstrHTML = locstrHTML & mHTMLEncode(Reason)
                                else
                                    locstrHTML = locstrHTML & "(none)"
                                end if
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>"
                            locstrHTML = locstrHTML & "<td colspan=2>"
                                locstrHTML = locstrHTML & "Please click "
								locstrHTML = locstrHTML & "<a href='https://" & Request.ServerVariables("http_host")
                            	locstrHTML =  locstrHTML & CONST_APPLICATION_PATH & "/requestleave.asp?m=ct&itemid=" & ID & "'>here</a>"
                                locstrHTML = locstrHTML & " to book this leave."
                            locstrHTML = locstrHTML & "</td>"
                        locstrHTML = locstrHTML & "<tr>"
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                    locstrHTML = locstrHTML & "<br>" & vbcrlf
    
                    locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td>" & vbcrlf
                                locstrHTML = locstrHTML & CONST_APPLICATION_FOOTER & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                    locstrHTML = locstrHTML & "</table>" & vbcrlf

                locstrHTML = locstrHTML & "</body>" & vbcrlf
            locstrHTML = locstrHTML & "</html>" & vbcrlf
            CompTimeEmailBodyGrant = locstrHTML
        End Property        
        
                
        Private Sub DBLoad()
            Dim locstrSQL
            Dim locstrCondition
            Dim locRS

            If m_blnLoaded then
                Exit Sub
            End If
            
            '*** If we have no ID then we can't find our LeavePeriod object ***
            If m_lngID = 0 and m_objEE.WWID = "" then
                Exit Sub
            Else 
                If m_lngID <> 0 Then
                    locstrCondition = "WHERE lngID=" & m_lngID
                End If
            End If
            
            
            m_blnLoaded = True
            
            '*** Load the Comp Time 
            locstrSQL = "SELECT * FROM tblCompTime " & locstrCondition
            'response.Write "cObjCompTime::DBLoad() SQL=" & locstrSQL & "<br>"
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, _
			           glbConnection, _
			           adOpenStatic
            
            If locRS.RecordCount > 0 Then
                m_lngID          = locRS("lngID")
                m_lngDaysGranted = locRS("lngDaysGranted")
        '        m_lngDaysRevoked = locRS("lngDaysRevoked")
                m_datDateGranted = locRS("datDateGranted")
               m_datDateRevoked = locRS("datDateRevoked")
                m_strReason      = locRS("strReason")
            Else
                SetEmpty                
            End If
            
        End Sub
        
        Public Sub LoadFromGrantForm()
            ID = 0
            EE.WWID = strRequestEEWWID
            DateGranted = Now()
            DaysGranted = request.Form("fldlngCompDays")
            Reason = request.form("fldstrReason")
        End Sub

         Public Sub LoadEE()
                EE.WWID = strRequestEEWWID
            End Sub
        
         Public Sub LoadFromRevokeForm()
            ID = 0
            EE.WWID = strRequestEEWWID
            DateGranted = Now()
            DaysRevoked = request.Form("fldlngCompDays")
            Reason = request.form("fldstrReason")
        End Sub
        
        Public Function Save()
            Dim loclngReturnValue
            Dim locstrSQL
        '    Dim locCmdRevoke
            Dim locCmd
            Dim locstrEmailBody
            Dim locstrEmailBody2
            Dim locWWID
            Dim locstrReason
            Dim locDateGranted
            Dim locDaysGranted
            Dim m_lngTaken
            Dim m_lngDaysBooked
 '   response.write "Test " &  m_lngDaysGranted  & " "  &   m_strReason     
        while m_lngDaysGranted  > 0
           
			Set locCmd = Server.CreateObject("ADODB.Command")
			Set locCmd.ActiveConnection = glbConnection
			locCmd.CommandText = "usp_save_new_comp_time"
				
				locCmd.CommandType = adCmdStoredProc
				locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
				locCmd.Parameters.Append locCmd.CreateParameter("lngWWID", adInteger, adParamInput, 8, EE.WWID)
				locCmd.Parameters.Append locCmd.CreateParameter("lngDaysGranted", adInteger, adParamInput,2 , 1)
				locCmd.Parameters.Append locCmd.CreateParameter("strReason", adWChar, adParamInput, 100, trim(m_strReason))
				
			   locCmd.Execute 
			   m_lngDaysGranted = m_lngDaysGranted -1
        wend     
   
                
			' Send a mail to the employee informing them of the granted leave
			locstrEmailBody = CompTimeEmailBodyGrant
			mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Email, "e-Vacation - Compensatory Leave Granted", locstrEmailBody, True
                
     Set  locCmd = nothing
      
        End Function
        
     Public Function Revoke()
            Dim loclngReturnValue
            Dim locstrSQL
        '    Dim locCmdRevoke
            Dim locCmd
            Dim locstrEmailBody
            Dim locWWID
            Dim m_lngWWID
            Dim locstrReason
            Dim locDateGranted
            Dim locDaysGranted
            Dim m_lngTaken
            Dim m_lngDaysBooked
            Dim locRevokeAmount
            
            Dim loclngDaysBooked 
            Dim locobjComptime
            Dim locRS
  '   response.write "Test " &  m_lngDaysGranted  & " "  &   m_strReason   
            locstrSQL = "SELECT * FROM tblcomptime " & _
                        "WHERE lngWWID=" & EE.WWID
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, _
			           glbConnection, _
			           adOpenStatic

         locRevokeAmount = m_lngDaysGranted
               
         while not locRS.EOF
                m_lngID    = locRS("lngID")
                m_lngWWID = locRS("lngWWID")
                m_lngDaysGranted = locRS("lngDaysGranted")
                m_datDateGranted = locRS("datDateGranted")
               m_datDateRevoked = locRS("datDateRevoked")
       
     
            if locRevokeAmount > 0  AND m_lngDaysGranted > 0 then                 
	            Set locCmd = Server.CreateObject("ADODB.Command")
	            Set locCmd.ActiveConnection = glbConnection
				locCmd.CommandText = "UPDATE tblCompTime " &_
									 "SET strReason='" & m_strReason & "', lngDaysGranted=0, IngDaysRevoked=1 " &_
									 "WHERE lngID=" & m_lngID							
				locCmd.Execute()
				
         
                   locRevokeAmount = locRevokeAmount -1
                  
              end if  
                locRS.movenext      
            Wend
            
            
     Set   locCmd = nothing
                
                ' Send a mail to the employee informing them of the granted leave
                locstrEmailBody = CompTimeEmailBodyRevoke
                mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Email, "e-Vacation - Compensatory Leave Revoked", locstrEmailBody, True
                
     
      
        End Function
        
           Public Function BookCompTime()
            Dim loclngReturnValue
            Dim locstrSQL
        '    Dim locCmdRevoke
            Dim locCmd
            Dim locstrEmailBody
            Dim locWWID
            Dim m_lngWWID
            Dim locstrReason
            Dim locDateGranted
            Dim locDaysGranted
            Dim m_lngTaken
            Dim m_lngDaysBooked
            Dim locRevokeAmount
            
            Dim loclngDaysBooked 
            Dim locobjComptime
            Dim locRS
     response.write "obj 4572 " &  m_lngDaysGranted  & " "  &   m_strReason   
            locstrSQL = "SELECT * FROM tblcomptime " & _
                        "WHERE lngWWID=" & EE.WWID
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, _
			           glbConnection, _
			           adOpenStatic

         locRevokeAmount = m_lngDaysGranted
               
         while not locRS.EOF
                m_lngID    = locRS("lngID")
                m_lngWWID = locRS("lngWWID")
                m_lngDaysGranted = locRS("lngDaysGranted")
                m_datDateGranted = locRS("datDateGranted")
               m_datDateRevoked = locRS("datDateRevoked")
       
     
            if locRevokeAmount > 0  AND m_lngDaysGranted > 0 then                 
	            Set locCmd = Server.CreateObject("ADODB.Command")
	            Set locCmd.ActiveConnection = glbConnection
				locCmd.CommandText = "UPDATE tblCompTime " &_
									 "SET strReason='" & m_strReason & "', lngDaysGranted=1, IngDaysBooked=1 " &_
									 "WHERE lngID=" & m_lngID							
				locCmd.Execute()
				
         
                   locRevokeAmount = locRevokeAmount -1
                  
              end if  
                locRS.movenext      
            Wend
            
            
     Set   locCmd = nothing
                
                ' Send a mail to the employee informing them of the granted leave
                locstrEmailBody = CompTimeEmailBody
                mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Email, "e-Vacation - Compensatory Leave Booked", locstrEmailBody, True
                
     
      
        End Function
        
    End Class


    '*** LeavePeriod Object ***
    Class cObjLeavePeriod
        '=================================================================
        'Description:   Contains definition of a Leave Period.
        'Properties:                    Type:           Perm:       Source:     Properties Required:
        'y  ID                          lng             RW          DB
        'y  EE                          objUser         RW          DB
        'y  Approver                    objUser         R           DB
        'y  LeaveType                   objLeaveType    R           DB
        'y  StartDate                   dat             RW          DB
        'y  StartTime                   str             RW          DB  e.g. AM or PM
        'y  EndDate                     dat             RW          DB
        'y  EndTime                     str             RW          DB  e.g. AM or PM
        'y  DateRaised                  dat             RW          DB
        'y  DateApproved                dat             RW          DB
        'y  DateRejected                dat             RW          DB
        'y  DateCancelRequested         dat             RW          DB
        'y  DateCancelApproved          dat             RW          DB
        'y  DateCancelRejected          dat             RW          DB
		'y 	DateConfirmed				dat				RW			DB		
		'y  DateConfirmEmailSent        dat             R           DB  (Added [MOF 12/08])
        'y  ResponseComments            str             RW          DB
        'y  RequestComments             str             RW          DB
        'y  ELPID                       lng             RW          DB
        'y  CompTimeID                  lng             RW          DB  (Added [MOF 1/09])
        'y  EmailConfOfRequestReq       bln             R           Virtual
        'y  WeekendDays                 lng             R           Calc        (StartDate, EndDate)
        'y  Days                        dbl             R           Calc        (PublicHolidays, WeekendDays, DaysConsecutive)
        'y  DaysConsecutive             dbl             R           Calc        (StartDate, EndDate)
        'y  LeaveDaysInPeriod()         dbl             R           Calc        (StartDate, EndDate)
        'y  ELPEndDate()                dat             R           Calc        (StartDate, objELPInstance.TargetDays)
        'y  Status                      str(CONST)      R           Calc        (DateCancelApproved, DateCancelRequested, DateRejected, DateApproved)
        'y  IsActive                    bln             R           Calc        (Status)
        'y  IsELP                       bln             R           Calc        (ELPID)
        'y  FormErrorMessage            str             R           Calc
        'y  FormIsValid                 bln             R           Calc        (FormErrorMessage)
        'y  NewRequestErrorMessage      str             R           Calc
        'y  NewRequestIsValid           bln             R           Calc
        'y  ApprovalErrorMessage        str             R           Calc
        'y  ApprovalFormIsValid         bln             R           Calc
        'y  AdminLeavePeriodFormIsValid bln             R           Calc
        'y  IsInvalidField(strField)    bln             R           Calc        (m_strErrorFieldList)
        'y  AppointedApprover           objUser         R           Calc        (Approver, EE)
        'y  AwaitingApproval            bln             R           Calc        (Status)
        'y  LeaveYear                   lng             R           Calc        (StartDate)
        'y  OverlappingLeavePeriods     col(LeavePeriods)R          Virtual
        'y  LeaveRequestEmailBody()     str             R           Calc        
        'y  IsValidApprover()           bln             R           Calc        (Approver, EE, EE.Manager, EE.Manager.ActiveDelegate)
        'y  IsAdminUpdating             bln             RW          Virtual
        'y  CancelRequiresApproval      bln             R           Calc
        'y  StopsLegalAdjAccrual        bln             R           Calc
		'y  CanConfirmLeave				bln				R			Calc		
        '
        'Methods:
        '   DBLoad
        '   LoadLeaveApprovalFromForm
        '   LoadNewRequestFromForm
        '   LoadAdminLeavePeriodFromForm
        '   Save
        '   Cancel
        '
        '=================================================================
        Private m_lngID
        Private m_objEE
        Private m_objApprover
        Private m_objLeaveType 
        Private m_datStartDate 
        Private m_strStartTime 
        Private m_datEndDate 
        Private m_strEndTime 
        Private m_datDateRaised 
        Private m_datDateApproved 
        Private m_datDateRejected 
        Private m_datDateCancelRequested 
        Private m_datDateCancelApproved 
        Private m_datDateCancelRejected
		Private m_datDateConfirmed	' [MOF 11/17/80] Added column
		Private m_datDateConfirmEmailSent
        Private m_strResponseComments 
        Private m_strRequestComments 
        Private m_blnEmailConfOfRequestReq
        Private m_blnShareLeaveWithTeamCalendar
        Private m_lngELPID 
        Private m_lngCompTimeID
        Private m_colOverlappingLeavePeriods
        Private m_blnIsAdminUpdating
        Public m_strErrorFieldList
        
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjLeavePeriod (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            SetNotLoaded
            m_blnIsAdminUpdating = False
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjLeavePeriod (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE  = nothing
            Set m_objApprover = nothing
            Set m_objLeaveType = nothing
            Set m_colOverlappingLeavePeriods = nothing          
        End Sub

    
        Private Sub SetEmpty()
            m_lngID = 0
            Set m_objEE  = nothing
            Set m_objEE = new cObjUser
            Set m_objApprover = nothing
            Set m_objApprover = new cObjUser
            Set m_objLeaveType = nothing
            Set m_objLeaveType = new cObjLeaveType
            m_datStartDate = ""
            m_strStartTime = "AM"
            m_datEndDate = ""
            m_strEndTime = "PM"
            m_datDateRaised = "" 
            m_datDateApproved = ""
            m_datDateRejected = ""
            m_datDateCancelRequested = ""
            m_datDateCancelApproved = "" 
            m_datDateCancelRejected = ""
			m_datDateConfirmed = ""	' [MOF 11/17/08] Added column
			m_datDateConfirmEmailSent = "" ' [MOF 30/12/08] Added column
            m_strResponseComments = "" 
            m_strRequestComments = "" 
            m_blnEmailConfOfRequestReq = False
            m_lngELPID = 0 
            m_blnShareLeaveWithTeamCalendar = 0
            Set m_colOverlappingLeavePeriods = nothing

        End Sub

        Private Sub SetNotLoaded()
            m_blnLoaded = False
        End Sub
            
        Public Property Let ID(ByVal varValue)
            varValue = mGetSafeLongInteger(varValue,0)
            If varValue <> 0 then
                If m_lngID <> varValue then
                    SetEmpty
                    SetNotLoaded
                    m_lngID = varValue
                End If
            End If
        End Property
        Public Property Get ID()
            DBLoad
            ID = m_lngID
        End Property
        

        Public Property Set EE(ByRef objUser)
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID
        End Property
        Public Property Get EE()
            DBLoad
            Set EE = m_objEE
        End Property        


        Public Property Get Approver() 
            DBLoad
            Set Approver = m_objApprover
        End Property


        Public Property Get LeaveType()
            DBLoad
            Set LeaveType = m_objLeaveType
        End Property


        Public Property Let StartDate(ByVal datValue)
            m_datStartDate = datValue
            m_blnLoaded = True
        End Property
        Public Property Get StartDate() 
            DBLoad
            StartDate = m_datStartDate
        End Property


        Public Property Let StartTime(ByVal strValue)
            If strValue <> m_strStartTime then
                m_strStartTime = strValue
            End If
            m_blnLoaded = True
        End Property
        Public Property Get StartTime() 
            DBLoad
            StartTime = m_strStartTime
        End Property


        Public Property Let EndDate(ByVal datValue)
            m_datEndDate = datValue
            m_blnLoaded = True
        End Property
        Public Property Get EndDate() 
            DBLoad
            EndDate = m_datEndDate
        End Property


        Public Property Let EndTime(ByVal strValue)
            If strValue <> m_strEndTime then
                m_strEndTime = strValue
            End If
            m_blnLoaded = True
        End Property
        Public Property Get EndTime() 
            DBLoad
            EndTime = m_strEndTime
        End Property


        Public Property Let DateRaised(ByVal datValue)
            m_datDateRaised = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateRaised() 
            DBLoad
            DateRaised = m_datDateRaised
        End Property


        Public Property Let DateApproved(ByVal datValue)
            m_datDateApproved = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateApproved() 
            DBLoad
            DateApproved = m_datDateApproved
        End Property
        
        
        Public Property Let DateRejected(ByVal datValue)
            m_datDateRejected = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateRejected()
            DBLoad
            DateRejected = m_datDateRejected
        End Property

        
        Public Property Let DateCancelRequested(ByVal datValue)
            m_datDateCancelRequested = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateCancelRequested() 
            DBLoad
            DateCancelRequested = m_datDateCancelRequested
        End Property
        

        Public Property Let DateCancelApproved(ByVal datValue)
            m_datDateCancelApproved = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateCancelApproved() 
            DBLoad
            DateCancelApproved = m_datDateCancelApproved
        End Property
        
        
        Public Property Let DateCancelRejected(ByVal datValue)
            m_datDateCancelRejected = datValue
            m_blnLoaded = True
        End Property
        Public Property Get DateCancelRejected() 
            DBLoad
            If not isDate(m_datDateCancelRejected) then
                m_datDateCancelRejected = null
            End If
            DateCancelRejected = m_datDateCancelRejected
        End Property
		
		Public Property Let DateConfirmed(ByVal datValue)	' [MOF 11/17/80] Added Let/Get properties for m_datDateConfirmed
			 m_datDateConfirmed = datValue
			 m_blnLoaded = True
		End Property
		Public Property Get DateConfirmed()
			DBLoad
			If not isDate(m_datDateConfirmed) then
				m_datDateConfirmed = null
			End If
			DateConfirmed = m_datDateConfirmed
		End Property
        
        Public Property Get DateConfirmEmailSent() ' [MOF 30/12/08] Added Get properties for m_datDateConfirmEmailSent
            DBLoad
            If not isDate(m_datDateConfirmEmailSent) then
                m_datDateConfirmEmailSent = null
            End If
            
            DateConfirmEmailSent = m_datDateConfirmEmailSent
        End Property
        
        Public Property Let ResponseComments(ByVal strValue)
            m_strResponseComments = trim(strValue)
            m_blnLoaded = True
        End Property
        Public Property Get ResponseComments() 
            DBLoad
            If IsNull(m_strResponseComments) then m_strResponseComments = ""
            ResponseComments = m_strResponseComments
        End Property
        

        Public Property Let RequestComments(ByVal strValue)
            m_strRequestComments = trim(strValue)
            m_blnLoaded = True
        End Property
        Public Property Get RequestComments() 
            DBLoad
            If IsNull(m_strRequestComments) then m_strRequestComments = ""
            RequestComments = m_strRequestComments
        End Property


        Public Property Let ELPID(ByVal varValue)
            varValue = mGetSafeLongInteger(varValue,0)
            m_lngELPID = varValue
            m_blnLoaded = True
        End Property
        Public Property Get ELPID() 
            DBLoad
            ELPID = m_lngELPID
        End Property


        Public Property Let CompTimeID(ByVal varValue) ' Added by MOF [1/09]
            varValue = mGetSafeLongInteger(varValue, 0)
            m_lngCompTimeID = varValue
            m_blnLoaded = True
        End Property
        Public Property Get CompTimeID() ' Added by MOF [1/09]
            DBLoad
            CompTimeID = m_lngCompTimeID
        End Property


        Public Property Let EmailConfOfRequestReq(ByVal blnValue)
            m_blnLoaded = True
            If blnValue = True then
                m_blnEmailConfOfRequestReq = True
            Else
                m_blnEmailConfOfRequestReq = False
            End If
        End Property
        Public Property Get EmailConfOfRequestReq() 
            DBLoad
            EmailConfOfRequestReq = m_blnEmailConfOfRequestReq
        End Property

        Public Property Let ShareLeaveWithTeamCalendar(ByVal blnValue)
            m_blnLoaded = True
            If blnValue = True then
                m_blnShareLeaveWithTeamCalendar = True
            Else
                m_blnShareLeaveWithTeamCalendar = False
            End If
        End Property
        Public Property Get ShareLeaveWithTeamCalendar() 
            DBLoad
            ShareLeaveWithTeamCalendar = m_blnShareLeaveWithTeamCalendar
        End Property    

        Public Property Get WeekendDays
            WeekendDays = mCountWeekendDays(StartDate,EndDate)
        End Property
                    
            
        Public Property Get Days()
            Dim locdblDays
            
            If (not isdate(StartDate)) or not isdate(EndDate) then
                Days = 0
                Exit Property
            End if

            locdblDays = DaysConsecutive
            
            locdblDays = locdblDays - glbPublicHolidays.CountInPeriod(StartDate,EndDate,False)
            locdblDays = locdblDays - WeekendDays
            Days = locdblDays
            
        End Property


        Public Property Get DaysConsecutive()
            Dim locdblDays
            
            If (not isdate(StartDate)) or not isdate(EndDate) then
                DaysConsecutive = 0
                Exit Property
            End if
            
            locdblDays = DateDiff("y",StartDate,EndDate) + 1

            If StartTime = "PM" then
                if (not glbPublicHolidays.Contains(StartDate, False)) and (not mIsWeekendDay(StartDate)) then
                    locdblDays = locdblDays - 0.5
                end if
            end if
            
            if EndTime = "AM" then
                if (not glbPublicHolidays.Contains(EndDate, False)) and (not mIsWeekendDay(EndDate)) then
                    locdblDays = locdblDays - 0.5
                end if
            end if

            DaysConsecutive = locdblDays
        End Property


        Public Property Get LeaveDaysInPeriod(ByVal datStart, ByVal datEnd)
        
            Dim locdatStartDate
            Dim locdatEndDate
            Dim locdblDays
            Dim locLeaveType

            If StartDate > datEnd or EndDate < datStart then
                LeaveDaysInPeriod = 0
                Exit Property
            End If
            
            If StartDate < datStart then
                locdatStartDate = datStart
            Else
                locdatStartDate = StartDate
            End If

            If EndDate > datEnd then
                locdatEndDate = datEnd
            Else
                locdatEndDate = EndDate
            End If
            
            '******** CA - Keep for debuggin only *****
     '       Response.Write "Obj 4555 CA - Here in cObjLeavePeriod.LeaveDaysInPeriod " & "<BR>"  'CA
     '       Response.Write  "locdatStartDate:"  & locdatStartDate & ", locdatEndDate:" & locdatEndDate &  "<BR>"  'CA
     '       Response.Write  "Type :"  & locLeaveType &  "</BR>"
            '*********
            
            locdblDays = DateDiff("y",locdatStartDate,locdatEndDate) + 1

            '*** If the start date is in the requested period, and is PM and not a weekend day or public holiday, take off half a day.
            If StartDate >= datStart and StartDate <= datEnd and StartTime = "PM" then
                if (not glbPublicHolidays.Contains(StartDate, False)) and (not mIsWeekendDay(StartDate)) then
                    locdblDays = locdblDays - 0.5
                end if
            end if

            If EndDate >= datStart and EndDate <= datEnd and EndTime = "AM" then                
                if (not glbPublicHolidays.Contains(EndDate, False)) and (not mIsWeekendDay(EndDate)) then
                    locdblDays = locdblDays - 0.5
                end if
            end if
            
            '******** CA - Keep for debuggin only *****
            'Response.Write "locdblDays:" & locdblDays & "<BR>"  'CA
            'Response.Write "glbPublicHolidays:" & glbPublicHolidays.CountInPeriod(locdatStartDate,locdatEndDate,False) & "<BR>"  'CA
            'Response.Write "mCountWeekendDays:" & mCountWeekendDays(locdatStartDate, locdatEndDate) & "<BR>"  'CA
            '*********
            
            LeaveDaysInPeriod = locdblDays - glbPublicHolidays.CountInPeriod(locdatStartDate,locdatEndDate,False) - mCountWeekendDays(locdatStartDate, locdatEndDate)
            
            
        End Property
        
        
        Public Property Get Status() 
            If IsDate(DateCancelRejected) and not IsDate(DateConfirmed) then
                Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED
            ElseIf IsDate(DateCancelApproved) then
                Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED
            ElseIf IsDate(DateCancelRequested) and not IsDate(DateConfirmed) then
                Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED
            ElseIf IsDate(DateRejected) then
                Status = CONST_LEAVE_PERIOD_STATUS_REJECTED
			ElseIf IsDate(DateConfirmed) then ' [MOF 11/18/2008] New 'Confirmed' state added
				Status = CONST_LEAVE_PERIOD_STATUS_CONFIRMED
            ElseIf IsDate(DateApproved) then                
                Status = CONST_LEAVE_PERIOD_STATUS_APPROVED
            Else
                Status = CONST_LEAVE_PERIOD_STATUS_RAISED
            End If
        End Property


        Public Property Get IsActive()
            If Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED or _
                Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED or _
                Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
                Status = CONST_LEAVE_PERIOD_STATUS_RAISED or _
                Status = CONST_LEAVE_PERIOD_STATUS_CONFIRMED then 
                IsActive = True
            End If
        End Property
        
        
        Public Property Get IsELP() 
            If ELPID <> 0 then
                IsELP = True
            Else
                IsELP = False
            End If
        End Property


        Public Property Get FormErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage
            Dim locblnStartDateValid
            Dim locblnEndDateValid

            m_strErrorFieldList = ""
            locblnValid = True

            '*** Check for valid Approver (if entered at all) ***
            If Approver.WWID <> "" then
                '**** Is the Approver WWID a valid WWID? ****
                If not mIsValidWWID(Approver.WWID) then
                    locstrErrorMessage = locstrErrorMessage & " The Approver WWID is invalid.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                    
                '**** Was the approver's WWID found in WDS? ****
                Elseif Approver.IDSID = "" then
                    locstrErrorMessage = locstrErrorMessage & " The Approver WWID could not be found.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                    
                '**** Check that the user isn't trying to appoint him or herself as the approver ****
                Elseif trim(lcase(Approver.IDSID)) = trim(lcase(objCurrentUser.IDSID)) then
                    locstrErrorMessage = locstrErrorMessage & " You cannot appoint yourself as an approver when raising leave requests.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                    
                '**** Check that the user isn't trying to appoint the employee (for which the leave request is being raised) as the approver ****
                Elseif trim(lcase(Approver.IDSID)) = trim(lcase(EE.IDSID)) then
                    locstrErrorMessage = locstrErrorMessage & " You cannot appoint this employee as an approver as the leave request applies to him or her.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                End if


            Else
                '*** An approver hasn't been entered - so check wether an approver is required.
                '*** Is the EE also his or her manager's delegate?
                If trim(lcase(EE.WWID)) = trim(lcase(EE.Manager.ActiveDelegate.WWID)) then
                    '*** Is the leave request for the current employee
                    If trim(lcase(EE.WWID)) = trim(lcase(objCurrentUser.WWID)) then
                        locstrErrorMessage = locstrErrorMessage & " You must supply an alternative approver as you are your manager\'s appointed delegate.\n"
                    ELse
                        locstrErrorMessage = locstrErrorMessage & " You must supply an alternative approver as the employee is also his or her manager\'s appointed delegate.\n"
                    End If
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                End If
                
            End If
            
            '*** Check Start Date exists, is a valid date and is in the right format, not a weekend day and not a public holiday ***
            If StartDate = "" then
                locstrErrorMessage = locstrErrorMessage & " The Start Date must be entered.\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf not IsDate(StartDate) then
                locstrErrorMessage = locstrErrorMessage & " The Start Date is invalid.\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            elseif LeaveYear < "1901" or LeaveYear > "2078" then
                locstrErrorMessage = locstrErrorMessage & " The Start Date is invalid.\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf not mIsFormattedDate(StartDate) then
                locstrErrorMessage = locstrErrorMessage & " The Start Date should be in (dd mmm yyyy) format - (e.g. 01 Jan 2001).\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf CDate(StartDate) < mFirstDayOfYear(CONST_FIRST_YEAR_SYSTEM_ACTIVE) then
                locstrErrorMessage = locstrErrorMessage & "  - e-Vacation can only be used to work with leave during or after the year " & CONST_FIRST_YEAR_SYSTEM_ACTIVE & ".\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf mIsWeekendDay(StartDate) then
                locstrErrorMessage = locstrErrorMessage & " The Start Date must not be a weekend day ("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(StartDate,"medium") & " is a " & WeekdayName(Weekday(StartDate)) & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf glbPublicHolidays.Contains(StartDate, False) then
                locstrErrorMessage = locstrErrorMessage & " The Start Date must not be a public holiday ("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(StartDate,"medium") & " is " & glbPublicHolidays.Item(glbPublicHolidays.Contains(StartDate, False)).Description & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            Else
                locblnStartDateValid = True
            End If

            '*** Check End Date exists, is a valid date and is in the right format, not a weekend day and not a public holiday ***
            If EndDate = "" then
                locstrErrorMessage = locstrErrorMessage & " The End Date must be entered.\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            ElseIf not IsDate(EndDate) then
                locstrErrorMessage = locstrErrorMessage & "  - The End Date is invalid.\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            elseif  DatePart("yyyy",EndDate) < "1901" or DatePart("yyyy",EndDate) > "2078" then
                locstrErrorMessage = locstrErrorMessage & " The EndDate is invalid.\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf not mIsFormattedDate(EndDate) then
                locstrErrorMessage = locstrErrorMessage & " The End Date should be in (dd mmm yyyy) format - (e.g. 01 Jan 2001).\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            ElseIf mIsWeekendDay(EndDate) then
                locstrErrorMessage = locstrErrorMessage & " The End Date must not be a weekend day("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(EndDate,"medium") & " is a " & WeekdayName(Weekday(EndDate)) & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            ElseIf glbPublicHolidays.Contains(EndDate, False) then
                locstrErrorMessage = locstrErrorMessage & " The End Date must not be a public holiday ("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(EndDate,"medium") & " is " & glbPublicHolidays.Item(glbPublicHolidays.Contains(EndDate, False)).Description & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            Elseif LeaveType.Name = "Comp Leave" then
           
       '        response.write "obj 4750 Leave type " & CONST_LEAVE_TYPE_NAME_COMP_TIME & "<BR>"
               
                ' Check if End Date > Granted Comp Time End Date
                'If EndDate > EE.AnnualLeave then
                '    locblnEndDateValid = False
                'EndIf
            Else
                locblnEndDateValid = True
            End If
 ' response.write "obj 4758 Leave type " & CONST_LEAVE_TYPE_NAME_COMP_TIME & "<BR>"
            If locblnStartDateValid and locblnEndDateValid then

                '*** Check that the StartDate and EndDate are in the same year.
                '*** Except for ELP
                If not LeaveType.Name = CONST_LEAVE_TYPE_NAME_ELP then
                    If DatePart("yyyy",StartDate) <> DatePart("yyyy",EndDate) then
                        locstrErrorMessage = locstrErrorMessage & " The End Date must be in the same year as the Start Date.\n"
                        m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                        locblnValid = False
                        locblnEndDateValid = False
                        locblnStartDateValid = False
                    End If
                End If
                
                '*** Check Start Date and End Date do not cause the period to be less than half a day.***
                If Days < 0.5 then
                    locstrErrorMessage = locstrErrorMessage & " The Start Date and time must be before the End Date and time.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "StartDate;EndDate"
                    locblnValid = False
                End If
                
                '*** For Comp Time 
                '*** Check if Days requested equal the days granted
                If LeaveType.Name = CONST_LEAVE_TYPE_NAME_COMP_TIME Then
                    ' Check if Days > Days Granted
                    
              response.write "obj 4891" & CompLeaveDaysGranted
                    If Days > EE.AnnualVacation.NextCompLeaveDaysGranted Then
                        locstrErrorMessage = locstrErrorMessage & "  - The number of days requested(" & _
                                             Days & ") must be equal to or less than the number of comp days\n      granted(" & _
                                             EE.AnnualVacation.NextCompLeaveDaysGranted & ") on you oldest unexpired comp time.\n"
                        locblnValid = False
                    End If 'ElseIf
                    
                    ' 
                    'If EndDate
                    
                    ' TO DO: Check if StartTime = AM and EndTime = PM - cannot have a half day for Comp Time
                    
                End If
            End If
            
            If len(RequestComments) > 100 then
                locstrErrorMessage = locstrErrorMessage & " Comments can be up to a maximum of 100 characters (" & len(RequestComments) & " entered).\n"
                m_strErrorFieldList = m_strErrorFieldList & "RequestComments;"
                locblnValid = False
            End If
            
            If locblnValid = True then
                FormErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct all errors and re-submit the form."
                FormErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        End Property
        

        Public Property Get FormIsValid()
            If FormErrorMessage = "" then
                FormIsValid = True
            Else
                FormIsValid = False
            End If
        End Property


        Public Property Get NewRequestErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage
            Dim loclngCounter
            Dim loclngCount
            Dim loclngDays
            Dim loclngMinDays
            
            m_strErrorFieldList = ""
            locblnValid = True

            '*** Make Sure the EE.YearToView is set correctly.
            EE.YearToView = LeaveYear
            'response.write ("EE.WWID: " & EE.AnnualVacation.HasMaturedELP)
            '*** We have checked the validity of the form itself, so can assume all properties of the leave object are now available.
            
            '*** Does this leave period overlap with any other for the same user? ***
            loclngCount = OverlappingLeavePeriods.Count
            If  loclngCount > 0 then
                loclngCounter = 0
                locstrErrorMessage = locstrErrorMessage & "  - This leave period overlaps the following leave period"
                if loclngCount <> 1 then
                    locstrErrorMessage = locstrErrorMessage & "s, which have"
                else
                    locstrErrorMessage = locstrErrorMessage & ", which has"
                end if
                locstrErrorMessage = locstrErrorMessage & " already been requested:\n"
                while loclngCounter < loclngCount
                    loclngCounter = loclngCounter + 1
                    With OverlappingLeavePeriods.Item(loclngCounter)
                        locstrErrorMessage = locstrErrorMessage & "             " & _
                            loclngCounter & ". " & _
                            .LeaveType.Name & _
                            " from " & _
                            mFormatDate(.StartDate, "medium with day") & " " & _
                            .StartTime & " to " & _
                            mFormatDate(.EndDate, "medium with day") & " " & _
                            .EndTime & "\n"
                    End WIth
                wend
                locblnValid = False
            End If

            '**** Is the period requested long enough according to the leave type
            loclngDays = Days
            loclngMinDays = LeaveType.MinDays
            If LeaveType.Name = CONST_LEAVE_TYPE_NAME_ELP and _
                EE.AnnualVacation.ELPMatured.TargetDays <> CONST_ELP_TARGET_DAYS then
                loclngMinDays = EE.AnnualVacation.ELPMatured.TargetDays
            End If
            If loclngDays < loclngMinDays then
                locstrErrorMessage = locstrErrorMessage & "  - The leave request only has a duration of " & loclngDays & " day"
                if loclngDays <> 1 then
                    locstrErrorMessage = locstrErrorMessage & "s"
                end if
                locstrErrorMessage = locstrErrorMessage & " - this leave type has a minimum duration of " & loclngMinDays & " day"
                if loclngMinDays <> 1 then
                    locstrErrorMessage = locstrErrorMessage & "s"
                end if
                locstrErrorMessage = locstrErrorMessage & ".\n"
                locblnValid = False
            end if

            '**** AND ANY OTHER VALIDATION NEEDED ****
            Select Case LeaveType.Name
                Case CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION
                    If loclngDays > EE.AnnualVacation.Balance then
                        locstrErrorMessage = locstrErrorMessage & "  - Not enough leave left to book this leave period.\n"
                        locblnValid = False
                    End If
        '        Case CONST_LEAVE_TYPE_NAME_ELP
        '            If loclngDays > EE.AnnualVacation.ELPMatured.TargetDays then
        '                locstrErrorMessage = locstrErrorMessage & "  - The leave period must begin before the ELP expiry date.\n" & loclngDays & " ..." & EE.AnnualVacation.ELPMatured.TargetDays
        '                locblnValid = False
        '            End If
            End Select


            If locblnValid = True then
                NewRequestErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct all errors and re-submit the form."
                NewRequestErrorMessage = replace(locstrErrorMessage,"'","")
            End If
            'response.write locstrErrorMessage
        End Property


        Public Property Get NewRequestIsValid()
            If NewRequestErrorMessage = "" then
                NewRequestIsValid = True
            Else
                NewRequestIsValid = False
            End If
        End Property


        Public Property Get ApprovalErrorMessage()
            Dim locstrErrorMessage
            Dim locblnValid
            
            m_strErrorFieldList = ""
            locstrErrorMessage = ""
            locblnValid = True
            
            If len(ResponseComments) > 100 then
                locstrErrorMessage = locstrErrorMessage & "  - Comments can be up to a maximum of 100 characters (" & len(ResponseComments) & " entered).\n"
                m_strErrorFieldList = m_strErrorFieldList & "ResponseComments;"
                locblnValid = False
            End If

            If IsDate(DateRejected) And len(ResponseComments) = 0 then
                locstrErrorMessage = locstrErrorMessage & "  - When rejecting a leave request you must give your reason in the comments section.\n"
                m_strErrorFieldList = m_strErrorFieldList & "ResponseComments;"
                locblnValid = False
            End If
            
            If locblnValid = True then
                ApprovalErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct all errors and re-submit the form."
                ApprovalErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        End Property


        Public Property Get AdminLeavePeriodErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage
            Dim locblnStartDateValid
            Dim locblnEndDateValid

            m_strErrorFieldList = ""
            locblnValid = True

            
            '*** Check Start Date exists, is a valid date and is in the right format, not a weekend day and not a public holiday ***
            If StartDate = "" then
                locstrErrorMessage = locstrErrorMessage & "  - The Start Date must be entered.\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf not IsDate(StartDate) then
                locstrErrorMessage = locstrErrorMessage & "  - The Start Date is invalid.\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf not mIsFormattedDate(StartDate) then
                locstrErrorMessage = locstrErrorMessage & "  - The Start Date should be in (dd mmm yyyy) format - (e.g. 01 Jan 2001).\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            ElseIf CDate(StartDate) < mFirstDayOfYear(CONST_FIRST_YEAR_SYSTEM_ACTIVE) then
                locstrErrorMessage = locstrErrorMessage & "  - e-Vacation can only be used to work with leave during or after the year " & CONST_FIRST_YEAR_SYSTEM_ACTIVE & ".\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            
            ElseIf mIsWeekendDay(StartDate) and (LeaveType.Name <> "Sick Leave" and LeaveType.Name <> "Sick Leave (job related)") then
                locstrErrorMessage = locstrErrorMessage & "  - The Start Date must not be a weekend day ("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(StartDate,"medium") & " is a " & WeekdayName(Weekday(StartDate)) & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            
            ElseIf glbPublicHolidays.Contains(StartDate, False) and (LeaveType.Name <> "Sick Leave" and LeaveType.Name <> "Sick Leave (job related)") then
                locstrErrorMessage = locstrErrorMessage & "  - The Start Date must not be a public holiday ("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(StartDate,"medium") & " is " & glbPublicHolidays.Item(glbPublicHolidays.Contains(StartDate, False)).Description & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "StartDate;"
                locblnValid = False
                locblnStartDateValid = False
            Else
                locblnStartDateValid = True
            End If

            '*** Check End Date exists, is a valid date and is in the right format, not a weekend day and not a public holiday ***
            If EndDate = "" then
                locstrErrorMessage = locstrErrorMessage & "  - The End Date must be entered.\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            ElseIf not IsDate(EndDate) then
                locstrErrorMessage = locstrErrorMessage & "  - The End Date is invalid.\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            ElseIf not mIsFormattedDate(EndDate) then
                locstrErrorMessage = locstrErrorMessage & "  - The End Date should be in (dd mmm yyyy) format - (e.g. 01 Jan 2001).\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            ElseIf mIsWeekendDay(EndDate)  and (LeaveType.Name <> "Sick Leave" and LeaveType.Name <> "Sick Leave (job related)") then
                locstrErrorMessage = locstrErrorMessage & "  - The End Date must not be a weekend day("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(EndDate,"medium") & " is a " & WeekdayName(Weekday(EndDate)) & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            ElseIf glbPublicHolidays.Contains(EndDate, False)  and (LeaveType.Name <> "Sick Leave" and LeaveType.Name <> "Sick Leave (job related)") then
                locstrErrorMessage = locstrErrorMessage & "  - The End Date must not be a public holiday ("
                locstrErrorMessage = locstrErrorMessage & mFormatDate(EndDate,"medium") & " is " & glbPublicHolidays.Item(glbPublicHolidays.Contains(EndDate, False)).Description & ").\n"
                m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                locblnValid = False
                locblnEndDateValid = False
            Else
                locblnEndDateValid = True
            End If

            If locblnStartDateValid and locblnEndDateValid then

                '*** Check that the StartDate and EndDate are in the same year.
                If DatePart("yyyy",StartDate) <> DatePart("yyyy",EndDate) then
                    locstrErrorMessage = locstrErrorMessage & "  - The End Date must be in the same year as the Start Date.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                    locblnValid = False
                    locblnEndDateValid = False
                    locblnStartDateValid = False
                End If

                '*** Check Start Date and End Date do not cause the period to be less than half a day.***
                If Days < 0.5 then
                    locstrErrorMessage = locstrErrorMessage & "  - The Start Date and time must be before the End Date and time.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "StartDate;EndDate"
                    locblnValid = False
                End If
            End If
            
            If len(RequestComments) > 100 then
                locstrErrorMessage = locstrErrorMessage & "  - Comments can be up to a maximum of 100 characters (" & len(RequestComments) & " entered).\n"
                m_strErrorFieldList = m_strErrorFieldList & "RequestComments;"
                locblnValid = False
            End If
            
            If locblnValid = True then
                AdminLeavePeriodErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct all errors and re-submit the form."
                AdminLeavePeriodErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        End Property
        
        
        Public Property Get ApprovalFormIsValid()
            If ApprovalErrorMessage = "" then
                ApprovalFormIsValid = True
            Else
                ApprovalFormIsValid = False
            End If
        End Property

        
        Public Property Get AdminLeavePeriodFormIsValid()
            If AdminLeavePeriodErrorMessage = "" then
                AdminLeavePeriodFormIsValid = True
            Else
                AdminLeavePeriodFormIsValid = False
            End If
        End Property


        Public Property Get NewAdminRequestIsValid()
            NewAdminRequestIsValid = NewRequestIsValid
        End Property
        
        
        Public Property Get NewAdminRequestErrorMessage()
            NewAdminRequestErrorMessage = NewRequestErrorMessage
        End Property


        Public Property Get CancelRequestFormIsValid()
            If CancelRequestFormErrorMessage = "" then
                CancelRequestFormIsValid = True
            Else
                CancelRequestFormIsValid = False
            End If
        End Property
        
        
        Public Property Get CancelRequestFormErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage

            m_strErrorFieldList = ""
            locblnValid = True
        
            '*** Check for valid Approver (if entered at all) ***
            If Approver.WWID <> "" then
                '**** Is the Approver WWID a valid WWID? ****
                If not mIsValidWWID(Approver.WWID) then
                    locstrErrorMessage = locstrErrorMessage & "  - The Approver WWID is invalid.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                    
                '**** Was the approver's WWID found in WDS? ****
                Elseif Approver.IDSID = "" then
                    locstrErrorMessage = locstrErrorMessage & "  - The Approver WWID could not be found.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                    
                '**** Check that the user isn't trying to appoint him or herself as the approver ****
                Elseif trim(lcase(Approver.IDSID)) = trim(lcase(objCurrentUser.IDSID)) then
                    locstrErrorMessage = locstrErrorMessage & "  - You cannot appoint yourself as an approver when raising cancellation requests.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                    
                '**** Check that the user isn't trying to appoint the employee (for which the cancellation request is being raised) as the approver ****
                Elseif trim(lcase(Approver.IDSID)) = trim(lcase(EE.IDSID)) then
                    locstrErrorMessage = locstrErrorMessage & "  - You cannot appoint this employee as an approver as the cancellation request applies to him or her.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                End if
            Else
                '*** An approver hasn't been entered - so check wether an approver is required.
                '*** Is the EE also his or her manager's delegate?
                If trim(lcase(EE.WWID)) = trim(lcase(EE.Manager.ActiveDelegate.WWID)) then
                    '*** Is the leave request for the current employee
                    If trim(lcase(EE.WWID)) = trim(lcase(objCurrentUser.WWID)) then
                        locstrErrorMessage = locstrErrorMessage & "  - You must supply an alternative approver as you are your manager\'s appointed delegate.\n"
                    ELse
                        locstrErrorMessage = locstrErrorMessage & "  - You must supply an alternative approver as the employee is also his or her manager\'s appointed delegate.\n"
                    End If
                    m_strErrorFieldList = m_strErrorFieldList & "Approver.WWID;"
                    locblnValid = False
                End If
                
            End If

            If len(RequestComments) > 100 then
                locstrErrorMessage = locstrErrorMessage & "  - Comments can be up to a maximum of 100 characters (" & len(RequestComments) & " entered).\n"
                m_strErrorFieldList = m_strErrorFieldList & "RequestComments;"
                locblnValid = False
            End If

            If locblnValid = True then
                CancelRequestFormErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct all errors and re-submit the form."
                CancelRequestFormErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        
        End Property

        
        
        Public Property Get IsInvalidField(ByVal strFieldName)
            If instr(m_strErrorFieldList,strFieldName & ";") then
                IsInvalidField = True
            Else
                IsInvalidField = False
            End If
        End Property
        

        Public Property Get AppointedApprover()
            DBLoad
            If Approver.WWID <> "" then
                Set AppointedApprover = Approver
            ElseIf EE.Manager.HasDelegate then
                Set AppointedApprover = EE.Manager.ActiveDelegate
            Else
                Set AppointedApprover = EE.Manager
            End If
        End Property
        

        Public Property Get AwaitingApproval()
            If Status = CONST_LEAVE_PERIOD_STATUS_RAISED OR _
                Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED then
                AwaitingApproval = True
            Else
                AwaitingApproval = False
            End If
        End Property


        Public Property Get LeaveYear()
            If IsDate(StartDate) then
                LeaveYear = DatePart("yyyy",StartDate)
            Else
                LeaveYear = DatePart("yyyy",Date)
            End If
        End Property
        
        
        Public Property Get OverlappingLeavePeriods()
            If typename(m_colOverlappingLeavePeriods) <> "cColLeavePeriods" then
                mDebugPrint "   Creating cColLeavePeriods - 2 <br>"
                Set m_colOverlappingLeavePeriods = new cColLeavePeriods
            End If
            If m_colOverlappingLeavePeriods.CollectionType = 0 then
                m_colOverlappingLeavePeriods.CollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_ALL_EE_LEAVE_IN_PERIOD
                Set m_colOverlappingLeavePeriods.EE = EE
                Set m_colOverlappingLeavePeriods.CompareLeavePeriod = Me
            End If
            Set OverlappingLeavePeriods = m_colOverlappingLeavePeriods
        End Property
        

        Public Property Get LeaveRequestEmailBody(ByVal loclngEmailType)
            Dim locstrHTML
            locstrHTML = ""
            locstrHTML = locstrHTML & "<!DOCTYPE HTML><HTML>" & vbcrlf
			    locstrHTML = locstrHTML & "<head><title>E-vacation Email</title><style type='text/css'>.pageContentHeader,.pageContentTable,table{border-collapse:collapse;border-spacing:0}.pageContentHeader,.pageContentTable{width:800px;border:0;padding:3px;margin-right:auto}.th,th{font-size:16px;text-align:left}.th{color:#0062a8}.pageContentHeader{margin-left:0;margin-bottom:10px}body,h1,h2,h3,h4,h5,h6,html,td,th{font-family:intel-clear,tahoma,Helvetica,helvetica,Arial,sans-serif}body,html{color:#111}.pageContentTable{margin-left:15px;margin-bottom:0}td,th{padding:4px;color:#404040}th{color:#0062a8}h1,h2,h3,h4,h5,h6{-webkit-font-smoothing:antialiased;-webkit-user-select:none;margin-top:0;font-weight:100;clear:both;margin-bottom:10px}"
                locstrHTML = locstrHTML & "</style></head><BODY>" & vbcrlf
					locstrHTML = locstrHTML & "<img src='email.png' style='border: 1px solid #000000' alt='e-vacation'><br><br>" & vbcrlf
                    locstrHTML = locstrHTML & "<table class='pageContentHeader'><tr><td><h3>" & vbcrlf
                            
                            Select Case loclngEmailType
                                Case CONST_EMAIL_TYPE_EE_NOTIFICATION
                                    locstrHTML = locstrHTML & "Leave Request Raised - Confirmation"
                                Case CONST_EMAIL_TYPE_APPOINTED_APPROVER
                                    locstrHTML = locstrHTML & "Leave Request For Your Approval"
                                Case CONST_EMAIL_TYPE_MANAGER_INFORMATION
                                    locstrHTML = locstrHTML & "Leave Request - For Information Only"
                                Case CONST_EMAIL_TYPE_REQUEST_RESPONSE
                                    locstrHTML = locstrHTML & "Request Response - "
                                    locstrHTML = locstrHTML & mHTMLEncode(Status)
                                Case CONST_EMAIL_TYPE_EE_CANCELLATION_REQUEST_NOTIFICATION
                                    locstrHTML = locstrHTML & "Cancellation Request Raised - Confirmation"
                                Case CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER
                                    locstrHTML = locstrHTML & "Cancellation Request For Your Approval"
                                Case CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION  
                                    locstrHTML = locstrHTML & "Cancellation Request - For Information Only"
								Case CONST_EMAIL_TYPE_EE_CONFIRM_LEAVE_TAKEN ' [MOFLYNN 11-2008] Added new e-mail
									locstrHTML = locstrHTML & "Please Confirm Leave Was Taken"
                            End Select
                                
                    locstrHTML = locstrHTML & "</h3></td></tr></table><table class='pageContentTable'><tr><td colspan='2'>" & vbcrlf

                                locstrHTML = locstrHTML & "Please do <b>not</b> reply to this e-mail.  It was sent by the e-Vacation Application, which cannot receive e-mails.<br><br>" & vbcrlf
                            
                            Select Case loclngEmailType
                                Case CONST_EMAIL_TYPE_EE_NOTIFICATION
                                    locstrHTML = locstrHTML & "This is confirmation that the following leave request has been raised successfully, and is awaiting approval."
                                Case CONST_EMAIL_TYPE_APPOINTED_APPROVER
                                    locstrHTML = locstrHTML & "This is notification that the following leave request has been raised and is awaiting your approval."
                                Case CONST_EMAIL_TYPE_MANAGER_INFORMATION
                                    locstrHTML = locstrHTML & "This is notification that the following leave request has been raised by one of your team."
                                Case CONST_EMAIL_TYPE_REQUEST_RESPONSE
                                    locstrHTML = locstrHTML & "This is notification of the response to your request."
                                Case CONST_EMAIL_TYPE_EE_CANCELLATION_REQUEST_NOTIFICATION
                                    locstrHTML = locstrHTML & "This is confirmation that the following leave cancellation request has been raised successfully, and is awaiting approval."
                                Case CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER
                                    locstrHTML = locstrHTML & "This is notification that the following leave cancellation request has been raised and is awaiting your approval."
                                Case CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION  
                                    locstrHTML = locstrHTML & "This is notification that the following leave cancellation request has been raised by one of your team."
								Case CONST_EMAIL_TYPE_EE_CONFIRM_LEAVE_TAKEN ' [MOFLYNN 11-2008] Added new e-mail
									locstrHTML = locstrHTML & "This is notification that the following leave period has concluded. Please confirm/cancel this leave as appropriate below."
									
                            End Select
							
                    locstrHTML = locstrHTML & "<br><br></td></tr>" & vbcrlf

                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<th colspan='2'>" & vbcrlf
                                If Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED or _
                                    Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED or _
                                    Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED then
                                    locstrHTML = locstrHTML & "Cancellation"
                                Else
                                    locstrHTML = locstrHTML & "Leave"
                                End If
                                    locstrHTML = locstrHTML & " Request"
                                If loclngEmailType = CONST_EMAIL_TYPE_APPOINTED_APPROVER or _
                                    loclngEmailType = CONST_EMAIL_TYPE_MANAGER_INFORMATION or _
                                    loclngEmailType = CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER or _
                                    loclngEmailType = CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION then
                                        locstrHTML = locstrHTML & " for " & mHTMLEncode(EE.FullNAME)
                                End If
                            locstrHTML = locstrHTML & "</th>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td style='width:150px'>" & vbcrlf
                                locstrHTML = locstrHTML & "<span class='th'>Leave Type</span> "							
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "<td style='width:650px'>" & vbcrlf
                                locstrHTML = locstrHTML & LeaveType.Name
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
						locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf						
                            locstrHTML = locstrHTML & "<td style='width:150px'>" & vbcrlf
                                locstrHTML = locstrHTML & "<span class='th'>Start Date</span>"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "<td style='width:650px'>" & vbcrlf
                                locstrHTML = locstrHTML & mFormatDate(StartDate,"medium with day") & " (" & StartTime & ")"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td style='width:150px'>" & vbcrlf
                                locstrHTML = locstrHTML & "<span class='th'>Appointed Approver</span>"
								locstrHTML = locstrHTML & "</td>" & vbcrlf
								locstrHTML = locstrHTML & "<td style='width:650px'>" & vbcrlf
                                Select Case loclngEmailType
                                    Case CONST_EMAIL_TYPE_EE_NOTIFICATION, _
                                         CONST_EMAIL_TYPE_MANAGER_INFORMATION, _
                                         CONST_EMAIL_TYPE_REQUEST_RESPONSE, _
                                         CONST_EMAIL_TYPE_EE_CANCELLATION_REQUEST_NOTIFICATION, _
                                         CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION, _
                                         CONST_EMAIL_TYPE_EE_CONFIRM_LEAVE_TAKEN 
                                        locstrHTML = locstrHTML & "<a href=""mailto:"
                                        locstrHTML = locstrHTML & mHTMLEncode(AppointedApprover.Email)
                                        locstrHTML = locstrHTML & """ title=""Click here to send an e-mail to "
                                        locstrHTML = locstrHTML & mHTMLEncode(AppointedApprover.FullName)
                                        locstrHTML = locstrHTML & " now."">"
                                        locstrHTML = locstrHTML & mHTMLEncode(AppointedApprover.FullName)
                                        locstrHTML = locstrHTML & "</a>" & vbcrlf
                                    Case CONST_EMAIL_TYPE_APPOINTED_APPROVER, _
                                        CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER
                                        locstrHTML = locstrHTML & mHTMLEncode(AppointedApprover.FullName)
                                End Select
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
						locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf	
                            locstrHTML = locstrHTML & "<td style='width:150px'>" & vbcrlf
                                locstrHTML = locstrHTML & "<span class='th'>End Date</span>" & vbcrlf
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "<td style='width:650px'>" & vbcrlf
                                locstrHTML = locstrHTML & mFormatDate(EndDate,"medium with day") & " (" & EndTime & ")"
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "<tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<td colspan='2'>" & vbcrlf
                                locstrHTML = locstrHTML & "<br><span class='th'>Request Comments</span><br><br>" & vbcrlf
                                if trim(RequestComments) <> "" then
                                    locstrHTML = locstrHTML & mHTMLEncode(RequestComments)
                                else
                                    locstrHTML = locstrHTML & "(none)"
                                end if
                            locstrHTML = locstrHTML & "</td>" & vbcrlf
                        locstrHTML = locstrHTML & "</tr>" & vbcrlf
						locstrHTML = locstrHTML & "</table>" & vbcrlf
						
                        if loclngEmailType = CONST_EMAIL_TYPE_REQUEST_RESPONSE then
                            
                            locstrHTML = locstrHTML & "<br>" & vbcrlf
                            locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                            locstrHTML = locstrHTML & "<tr>" & vbcrlf
                                locstrHTML = locstrHTML & "<td>" & vbcrlf
                                    locstrHTML = locstrHTML & "<span class='th'>Response</span>"
                                locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "</tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<tr>" & vbcrlf
                                locstrHTML = locstrHTML & "<td>" & vbcrlf
                                    locstrHTML = locstrHTML & "<span class='th'>Response</span>" & vbcrlf
                                    locstrHTML = locstrHTML & mHTMLEncode(Status)
                                locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "</tr>" & vbcrlf
                            locstrHTML = locstrHTML & "<tr>" & vbcrlf
                                locstrHTML = locstrHTML & "<td>" & vbcrlf
                                    locstrHTML = locstrHTML & "<span class='th'>Response Comments</span><br>" & vbcrlf
                                    if trim(ResponseComments) <> "" then
                                        locstrHTML = locstrHTML & mHTMLEncode(ResponseComments)
                                    else
                                        locstrHTML = locstrHTML & "(none)"
                                    end if
                                locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "</tr>" & vbcrlf
							locstrHTML = locstrHTML & "</table>" & vbcrlf
                        End if
                        
                        ' [MOF 16/12/08] Added Link to confirm the given leave request
                        If loclngEmailType = CONST_EMAIL_TYPE_EE_CONFIRM_LEAVE_TAKEN Then
							locstrHTML = locstrHTML & "<br>" & vbcrlf
                            locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                            locstrHTML = locstrHTML & "<tr>"
                                locstrHTML = locstrHTML & "<td>"
                                    locstrHTML = locstrHTML & "<span class='th'>Please either</span>"
                                locstrHTML = locstrHTML & "</td>"
                                locstrHTML = locstrHTML & "<td>"	
									locstrHTML = locstrHTML & "<a href='https://" & Request.ServerVariables("http_host")
                                	locstrHTML =  locstrHTML & CONST_APPLICATION_PATH & "/leavesummary.asp?m=cf&amp;itemid=" & ID & "'>Click here to confirm this leave was taken</a>"
                                locstrHTML = locstrHTML & "</td>"
                            locstrHTML = locstrHTML & "</tr>"
                            locstrHTML = locstrHTML & "<tr>"
                                locstrHTML = locstrHTML & "<td>"	
									locstrHTML = locstrHTML & "<a href='https://" & Request.ServerVariables("http_host")
                                	locstrHTML =  locstrHTML & CONST_APPLICATION_PATH & "/leavesummary.asp?m=cl&amp;itemid=" & ID & "'>Click here to submit a cancellation request for this leave</a>"
                                locstrHTML = locstrHTML & "</td>"
                            locstrHTML = locstrHTML & "</tr>"
							locstrHTML = locstrHTML & "</table>" & vbcrlf
                        End if
                        
                        if loclngEmailType = CONST_EMAIL_TYPE_APPOINTED_APPROVER or _
                            loclngEmailType = CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER then
								locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                                locstrHTML = locstrHTML & "<tr>"
                                    locstrHTML = locstrHTML & "<td>"	
										locstrHTML = locstrHTML & "<a href='https://" & Request.ServerVariables("http_host")
                                    	locstrHTML =  locstrHTML & CONST_APPLICATION_PATH & "/approverequests.asp'>Click here to approve or reject</a>"
                                    locstrHTML = locstrHTML & "</td>"
                                locstrHTML = locstrHTML & "</tr>" 
								locstrHTML = locstrHTML & "</table>" & vbcrlf
                        end if

                    If loclngEmailType = CONST_EMAIL_TYPE_EE_NOTIFICATION or _
                        loclngEmailType = CONST_EMAIL_TYPE_MANAGER_INFORMATION or _
                        loclngEmailType = CONST_EMAIL_TYPE_EE_CANCELLATION_REQUEST_NOTIFICATION or _
                        loclngEmailType = CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION then
                        locstrHTML = locstrHTML & "<br>" & vbcrlf
                        locstrHTML = locstrHTML & "<table class='pageContentTable'>" & vbcrlf
                            locstrHTML = locstrHTML & "<tr>" & vbcrlf
                                locstrHTML = locstrHTML & "<td>" & vbcrlf
                                    locstrHTML = locstrHTML & mHTMLEncode(AppointedApprover.FullName)
                                    locstrHTML = locstrHTML & " has been sent an e-mail asking for a response to this leave request.<br>" & vbcrlf
                                    If loclngEmailType = CONST_EMAIL_TYPE_EE_NOTIFICATION or _
                                       loclngEmailType = CONST_EMAIL_TYPE_EE_CANCELLATION_REQUEST_NOTIFICATION then
                                        if trim(EE.Manager.WWID) <> trim(AppointedApprover.WWID) then
                                            locstrHTML = locstrHTML & "Your Manager, "
                                            locstrHTML = locstrHTML & mHTMLEncode(EE.Manager.FullName)
                                            locstrHTML = locstrHTML & " has also been sent notification of this leave request for informational purposes.<br>" & vbcrlf
                                        end if
                                    End If
                                locstrHTML = locstrHTML & "</td>" & vbcrlf
                            locstrHTML = locstrHTML & "</tr>" & vbcrlf
                        locstrHTML = locstrHTML & "</table>" & vbcrlf
                    end if
                    
                    '**VARIABLES USED TO DISPLAY CALENDAR
                    dim myEndDate
                    dim myStartDate
                    dim myBgStartDate
                    dim myBgEndDate
                    dim counter
                    dim myBgColor
                    counter = 0
                    myBgStartDate = DateAdd("d",0,StartDate)
                    myStartDate = DateAdd("d",-7,StartDate)
                    myEndDate = DateAdd("d",7,EndDate)
                    myBgEndDate = DateAdd("d",0,EndDate)
                    
                    if loclngEmailType = CONST_EMAIL_TYPE_APPOINTED_APPROVER or _
						 loclngEmailType = CONST_EMAIL_TYPE_MANAGER_INFORMATION or _
							 loclngEmailType = CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER or _
								 loclngEmailType = CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION then 
                    
						'***DISPLAY OF LEAVE PERIOD
                        locstrHTML = locstrHTML & "<br>" & vbcrlf
						locstrHTML = locstrHTML & "<br><h3>Week Schedule</h3><table>" & vbcrlf
							locstrHTML = locstrHTML & "<tr>" & vbcrlf
	                        
								dim i
								dim x
								dim start
								x = 0
								start = true
								
									Dim rstLeaveHols
									Dim cmGetEmployeeHolidayData
									Dim m_cnDB
									Dim tdColor
									
									'create a new connection because need to modify cursor location(error otherwise)
									set m_cnDB = Server.CreateObject("ADODB.Connection")
									m_cnDB.ConnectionString = CONST_ADO_EVACATION_CONNECTION_STRING	
									m_cnDB.CursorLocation = adUseClient
									m_cnDB.Open
									
									'*********************************************
									'**				
									'** retrieve the employees leave for the year
									'**
									'*********************************************
									Set cmGetEmployeeHolidayData = Server.CreateObject("ADODB.Command")
									Set cmGetEmployeeHolidayData.ActiveConnection =  m_cnDB
									cmGetEmployeeHolidayData.CommandType = 4
									cmGetEmployeeHolidayData.CommandText = "dbo.mail_cal_display"
									cmGetEmployeeHolidayData.Parameters.Append cmGetEmployeeHolidayData.CreateParameter("@vWWID", adChar, adParamInput, 8, EE.Manager.WWID)
									cmGetEmployeeHolidayData.Parameters.Append cmGetEmployeeHolidayData.CreateParameter("@vMyStartDate", adDBTimeStamp, adParamInput, , myStartDate)
									cmGetEmployeeHolidayData.Parameters.Append cmGetEmployeeHolidayData.CreateParameter("@vMyEndDate", adDBTimeStamp, adParamInput, , myEndDate)
									Set rstLeaveHols = cmGetEmployeeHolidayData.Execute
									
								'**CYCLES THROUGH FROM START DATE OF TIME-SPAN TO END DATE
								for i = myStartDate to myEndDate
									
									'**FOR DISPLAY PURPOSES, TO HAVE 7 DAY WEEK DISPLAYED AND THEN ONTO NEW LINE
									if x = 7 then
										locstrHTML = locstrHTML & "</tr>" & vbcrlf
										locstrHTML = locstrHTML & "<tr>" & vbcrlf
										x = 1
									else
										x = x + 1
									end if
									
									'**SETS THE BACKGROUND COLOUR OF THE CELLS TO BE DIFFERENT FOR THE HOLIDAY REQUESTED THAN THE TIME AROUND IT
									if myStartDate => myBgStartDate AND myStartDate =< myBgEndDate then
										myBgColor = CONST_EMAIL_LEAVE_BACKGROUND
									else
										myBgColor = "white"
									end if
				
				
									'**SELECT FOR FORMATTING OF CALENDAR WITH 7 CELLS IN EACH ROW
									if start = true then	
										Select Case WeekDayName(WeekDay(i))
											Case "Monday"
												start = false
											Case "Tuesday"
												locstrHTML = locstrHTML & "<td>" & vbcrlf
												locstrHTML = locstrHTML & "</td>" & vbcrlf
												start = false
												x = x + 1
											Case "Wednesday"
												locstrHTML = locstrHTML & "<td colspan=2>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
												start = false
												x = x + 2
											Case "Thursday"
												locstrHTML = locstrHTML & "<td colspan=3>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
												start = false
												x = x + 3
											Case "Friday"
												locstrHTML = locstrHTML & "<td colspan=4>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
												start = false
												x = x + 4
										end select
									end if
									
									locstrHTML = locstrHTML & "<td valign=top nowrap bgcolor="
									locstrHTML = locstrHTML & myBgColor
									locstrHTML = locstrHTML & ">" & vbcrlf
									
										'**TABLE FOR DISPLAY OF EMPLOYEES IN DAY
										locstrHTML = locstrHTML & "<table border=0 cellspacing=0 cellpadding=0>" & vbcrlf
										locstrHTML = locstrHTML & "<tr>" & vbcrlf
										locstrHTML = locstrHTML & "<td>" & vbcrlf
										locstrHTML = locstrHTML & trim(WeekDayName(WeekDay(i))) & vbcrlf
										locstrHTML = locstrHTML & "</br>" & vbcrlf
										locstrHTML = locstrHTML & trim(i) & vbcrlf
										locstrHTML = locstrHTML & "</td>" & vbcrlf
										locstrHTML = locstrHTML & "</tr>" & vbcrlf
										'counter = counter + 1
										
										'**DOES NOTHING IF THE WEEKDAY IS A SATURDAY OR SUNDAY AS THESE DAYS ARE NOT WORK DAYS
										if WeekDayName(weekDay(i)) = "Saturday" or WeekDayName(weekDay(i)) = "Sunday" then
										
										else
											'**ADDING EMPLOYEES ABSENT TO DAY DISPLAY
											do while not rstLeaveHols.eof
														if rstLeaveHols.fields.item(2).value <= i AND rstLeaveHols.fields.item(3).value >= i then
															if not rstLeaveHols.fields.item(5).value = "" then
																locstrHTML = locstrHTML & "<tr>" & vbcrlf
																locstrHTML = locstrHTML & "<td nowrap bgcolor="
																locstrHTML = locstrHTML & trim(Mid(rstLeaveHols.fields.item(4).value,1,6))
																locstrHTML = locstrHTML & ">" & vbcrlf
																locstrHTML = locstrHTML & "<a href=https://shannon.intel.com/" & CONST_APPLICATION_PATH & "/adminhome.asp?m=lv&ee=" & rstLeaveHols.fields.item(4).value & ">"
																locstrHTML = locstrHTML & "<font color=white>" & vbcrlf
																locstrHTML = locstrHTML & mHTMLEncode(rstLeaveHols.fields.item(0).value) & "<br>" & mHTMLEncode(rstLeaveHols.fields.item(1).value) & vbcrlf
																locstrHTML = locstrHTML & "</a>"
																locstrHTML = locstrHTML & "</font>" & vbcrlf
																locstrHTML = locstrHTML & "</td>" & vbcrlf
																locstrHTML = locstrHTML & "</tr>" & vbcrlf
															end if
														end if
												
												'**MOVES TO THE NEXT ELEMENT
												rstLeaveHols.movenext
											loop
											
											'**MOVES BACK TO THE FIRST ELEMENT AFTER CYCLING THROUGH EVERYTHING FOR EACH DAY
											rstLeaveHols.moveFirst
										end if
										
										locstrHTML = locstrHTML & "</table>" & vbcrlf
										
									locstrHTML = locstrHTML & "</td>" & vbcrlf
									
									'**SELECT FOR FORMATTING OF CALENDAR WITH 7 CELLS IN EACH ROW
									if i = myEndDate then
										Select Case WeekDayName(WeekDay(i))
											Case "Monday"
												locstrHTML = locstrHTML & "<td colspan=6>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
											Case "Tuesday"
												locstrHTML = locstrHTML & "<td colspan=5>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
											Case "Wednesday"
												locstrHTML = locstrHTML & "<td colspan=4>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
											Case "Thursday"
												locstrHTML = locstrHTML & "<td colspan=3>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
											Case "Friday"
												locstrHTML = locstrHTML & "<td colspan=2>"
												locstrHTML = locstrHTML & "</td>" & vbcrlf
											Case "Saturday"
												locstrHTML = locstrHTML & "<td>" & vbcrlf
												locstrHTML = locstrHTML & "</td>" & vbcrlf
											Case "Sunday"
										end select
									end if
									
									'**INCREMENTS THE STARTDATE BY 1 TO MOVE TO THE NEXT DAY
									myStartDate = DateAdd("d",1,myStartDate)
									
									
								next
	                            
							locstrHTML = locstrHTML & "</tr>" & vbcrlf
						locstrHTML = locstrHTML & "</table>" & vbcrlf
						
					end if
                locstrHTML = locstrHTML & "</body>" & vbcrlf
            locstrHTML = locstrHTML & "</html>" & vbcrlf
            LeaveRequestEmailBody = locstrHTML
        End Property


        Public Property Get IsValidApprover(ByRef locobjUser)
            If locobjUser.WWID = EE.WWID then
                IsValidApprover = False
            ElseIf locobjUser.WWID = Approver.WWID then
                IsValidApprover = True
            ElseIf locobjUser.WWID = EE.Manager.WWID then
                IsValidApprover = True
            ElseIf locobjUser.WWID = EE.Manager.ActiveDelegate.WWID then
                IsValidApprover = True
            Else
                IsValidApprover = False
            End If
        End Property


        Public Property Let IsAdminUpdating(ByVal blnValue)
            m_blnIsAdminUpdating = blnValue
        End Property
        Public Property Get IsAdminUpdating()
            IsAdminUpdating = m_blnIsAdminUpdating
        End Property


        Public Property Get CancelRequiresApproval()
            If StartDate <= Date() then
                CancelRequiresApproval = True
            Else
                CancelRequiresApproval = False
            End if
        End Property
        
        
        Public Property Get StopsLegalAdjAccrual()
            If LeaveType.StopsLegalAdjAccrual then
                If LeaveType.DaysBeforeStopsLegalAdjAccrualIsConsecutive then
                    If DaysConsecutive > LeaveType.DaysBeforeStopsLegalAdjAccrual then
                        StopsLegalAdjAccrual = True
                    Else
                        StopsLegalAdjAccrual = False
                    End If
                Else
                    If Days > LeaveType.DaysBeforeStopsLegalAdjAccrual then
                        StopsLegalAdjAccrual = True
                    Else
                        StopsLegalAdjAccrual = False
                    End If
                End If
            Else
                StopsLegalAdjAccrual = False
            End If
        End Property
		
		' [MOF 11/24/08] Added Confirm logic
		Public Property Get CanConfirmLeave()
			Dim blnReturnValue
			blnReturnValue = False

            ' Only allow leave periods that end after a certain day be confirmed
            If EndDate > cDate(CONST_DATE_CONFIRMING_LEAVE_BEGINS) Then
			    If EndTime = "AM" Then
				    ' Now must be > EndDate at 12PM	before user can confirm (i.e. It must be in the afternoon onwards.)
				    If Now() > DateAdd("h", 12, EndDate) Then
					    blnReturnValue = True
				    End If
			    Else
				    ' Now must be > EndDate + 1 Day at 12AM before user can confirm (i.e. It must be in the morning time of the next day)
				    If Now() > DateAdd("d", 1, EndDate) Then
					    blnReturnValue = True
				    End If
			    End If
			End If
			
			CanConfirmLeave = blnReturnValue
		End Property
		' [/MOF]
		
        Private Property Get ELPEndDate()
            Dim locobjELP
            Dim locdatOldDate
            Dim locdatNewDate
            Dim loclngDaysOff   '***Weekends and Public Holidays
            
            If len(StartDate) > 0 then
                locdatOldDate = StartDate
                'response.write "StartDate: " & locdatOldDate & "<BR>"
                Set locobjELP = new cobjELPInstance
                locobjELP.ELPID = ELPID
                'response.write "ELP Days to take: " & locobjELP.TargetDays & "<BR>"
                locdatNewDate = DateAdd("w", locobjELP.TargetDays, CDate(locdatOldDate)-1)
                'response.write "EndDate (excluding holidays): " & locdatNewDate & "<BR>"
                Set locobjELP = Nothing
                
                loclngDaysOff = glbPublicHolidays.CountInPeriod(CDate(locdatOldDate)+1,locdatNewDate,False) + mCountWeekendDays(CDate(locdatOldDate)+1, locdatNewDate)
                'response.write "Holidays ("& locdatOldDate & " - " & locdatNewDate & "): " & loclngDaysOff & "<BR>"
                While loclngDaysOff > 0
                    locdatOldDate = locdatNewDate
                    locdatNewDate = DateAdd("w", loclngDaysOff, locdatNewDate)
                    loclngDaysOff = glbPublicHolidays.CountInPeriod(CDate(locdatOldDate)+1,locdatNewDate,False) + mCountWeekendDays(CDate(locdatOldDate)+1, locdatNewDate)
                    'response.write "OldDate: " & locdatOldDate & "<BR>"
                    'response.write "NewDate: " & locdatNewDate & "<BR>"
                    'response.write "Holidays ("& locdatOldDate & " - " & locdatNewDate & "): " & loclngDaysOff & "<BR>"
                Wend    
                'Response.write locdatNewDate
                ELPEndDate = locdatNewDate
            Else
                
                ELPEndDate = date() 'Just so the user doesn't see the 'End date must be entered' 
                            'error which could confue them
            End if
        End Property
		
        Private Sub DBLoad()
            Dim locCmd
            Dim locCmdRevoke
            Dim locParam
            Dim locRS

            If m_blnLoaded then
                Exit Sub
            End If
            
            '*** If we have no ID then we can't find our LeavePeriod object ***
            If m_lngID = 0 then
                exit sub
            End if

            m_blnLoaded = True

            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_leaveperiod"
            locCmd.CommandType = adCmdStoredProc
                        
            Set locParam = locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)
            locCmd.Parameters.Append locParam
            
            Set locRS = locCmd.Execute

            If not locRS.eof then
                m_objEE.WWID = trim(locRS("strEEWWID"))
                m_objApprover.WWID = trim(locRS("strApproverWWID"))
                m_objLeaveType.ID = locRS("lngLeaveTypeID")
                m_datStartDate = locRS("datStartDate")
                m_strStartTime = trim(locRS("strStartTime"))
                m_datEndDate = locRS("datEndDate")
                m_strEndTime = trim(locRS("strEndTime"))
                m_datDateRaised = locRS("datRaised")
                m_datDateApproved = locRS("datApproved")
                m_datDateRejected = locRS("datRejected")
                m_datDateCancelRequested = locRS("datCancelRequested")
                m_datDateCancelApproved = locRS("datCancelApproved")
                m_datDateCancelRejected = locRS("datCancelRejected")
				m_datDateConfirmed = locRS("datConfirmed") ' [MOF 11/17/08] Added new column '
				m_datDateConfirmEmailSent = locRS("datConfirmEmailSent") '
                m_strResponseComments = trim(locRS("strResponseComments"))
                m_strRequestComments = trim(locRS("strRequestComments"))
                m_blnShareLeaveWithTeamCalendar = locRS("shareLeaveWithTeamCalendar")
                m_lngELPID = locRS("lngELPID")
                m_lngCompTimeID = locRS("lngCompTimeID") ' MOF - Added new column (1/09) '
            Else
                SetEmpty
            End If
            
            locRS.Close
            
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
			
			'Response.Write( " - DBLoad cObjLeavePeriod" ) 
			'Response.Write( "<br>" ) 
            
        End Sub


        Public Sub LoadLeaveApprovalFromForm
            If ID = 0 then exit sub
            If locobjLeaveRequest.AwaitingApproval then
                ResponseComments = trim(request.form("fldstrComments"))
            
                '*** The date is stored only to return the correct status of the object,
                '*** the SQL Stored Procedure stores the date / time in the relevant field as it runs.
                If request.form("btnSubmit") = "Approve" then
                    If Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED then
                        DateCancelApproved = Now()
                    Else
                        DateApproved = Now()
                    End If
                Else
                    If Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED then
                        DateCancelRejected = Now()
                    Else
                        DateRejected = Now()
                    End If
                End If
            End If
        End Sub


        Public Sub LoadNewRequestFromForm
            ID = 0
            EE.WWID = request.form("ee")
            Approver.WWID = request.form("fldstrApproverWWID")
            LeaveType.Name = request.form("fldstrLeaveType")
            StartDate = request.form("fldstrStartDate")
            StartTime = request.form("fldstrStartTime")
            EndDate = request.form("fldstrEndDate")
            EndTime = request.form("fldstrEndTime")
            RequestComments = request.form("fldstrComments")
            If request.form("fldblnNotify") = "True" then
                EmailConfOfRequestReq = True
            Else
                EmailConfOfRequestReq = False
            End If

            If request.form("fldblnShare") = "True" then
                ShareLeaveWithTeamCalendar = True
            Else
                ShareLeaveWithTeamCalendar = False
            End If
            
            '*** Modifications for ELP Leave Requests
            If request.form("formname") = "frmRequestELPLeave" then
                StartTime = "AM"
                EndTime = "PM"
                LeaveType.Name = CONST_LEAVE_TYPE_NAME_ELP
                ELPID = request.form("fldlngELPID")
                
                EndDate = mFormatDate(ELPEndDate, "medium")
            End If
            
        End Sub


        Public Sub LoadAdminLeavePeriodFromForm
            ID = mGetSafeLongInteger(request.form("itemid"),0)
            EE.WWID = request.form("ee")
            LeaveType.Name = request.form("fldstrLeaveType")
            StartDate = request.form("fldstrStartDate")
            StartTime = request.form("fldstrStartTime")
            EndDate = request.form("fldstrEndDate")
            EndTime = request.form("fldstrEndTime")
            RequestComments = request.form("fldstrComments")
            ResponseComments = "Administrator Updated (" & objCurrentUser.FullName & " - " & objCurrentUser.WWID & ")"
        End Sub
        

        Public Function Delete()
           
            Dim locCmdRevoke
            Dim locCmd
            Dim loclngReturnValue

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            loclngReturnValue = 0
            
            '*** If we have no ID set up then we can't delete - return a 1 to indicate failure ***
            If m_lngID = 0 then
                Delete = 1
                Exit Function
            End If
            
            locCmd.CommandText = "usp_delete_leave_request"
            locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
            locCmd.Parameters.Append locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)

            on error resume next
            
            locCmd.Execute
            
            loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
            
            on error goto 0
                
            Set locCmd = nothing

            Delete = loclngReturnValue
        End Function

        
        Public Function Save()
            Dim locCmd
            Dim locCmdRevoke
            Dim loclngReturnValue
            Dim locstrBody
            Dim DB
            Dim lngID

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            loclngReturnValue = 0
            
'    response.write " obj6108  " & LeaveType.ID  & "</br>"
 '    response.write "obj 6109 " & lngID  & "</br>"
  
            '*** If we have no ID set up then this is a new leave request ***
            If m_lngID = 0 or IsAdminUpdating then
                locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
                If IsAdminUpdating then
                    locCmd.Parameters.Append locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)
                End If
                locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, trim(EE.WWID))
                locCmd.Parameters.Append locCmd.CreateParameter("strApproverWWID", adWChar, adParamInput, 8, trim(Approver.WWID))
                locCmd.Parameters.Append locCmd.CreateParameter("lngLeaveTypeID", adInteger, adParamInput, , LeaveType.ID)
                locCmd.Parameters.Append locCmd.CreateParameter("datStartDate", adDBTimeStamp, adParamInput, , StartDate)
                locCmd.Parameters.Append locCmd.CreateParameter("strStartTime", adWChar, adParamInput, 2, trim(StartTime))
                locCmd.Parameters.Append locCmd.CreateParameter("datEndDate", adDBTimeStamp, adParamInput, , EndDate)
                locCmd.Parameters.Append locCmd.CreateParameter("strEndTime", adWChar, adParamInput, 2, trim(EndTime))
                locCmd.Parameters.Append locCmd.CreateParameter("strRequestComments", adWChar, adParamInput, 200, trim(RequestComments))
                If IsAdminUpdating then
                    locCmd.Parameters.Append locCmd.CreateParameter("strResponseComments", adWChar, adParamInput, 100, trim(ResponseComments))
                End If
                locCmd.Parameters.Append locCmd.CreateParameter("lngELPID", adInteger, adParamInput, , mIf((ELPID=0),Null,ELPID))
                locCmd.Parameters.Append locCmd.CreateParameter("lngCompTimeID", adInteger, adParamInput, , mIf((CompTimeID=0),Null,CompTimeID))
                locCmd.Parameters.Append locCmd.CreateParameter("shareLeaveWithTeamCalendar", adInteger, adParamInput, , mIf((ShareLeaveWithTeamCalendar=0),Null,ShareLeaveWithTeamCalendar))            
                
'response.write "obj 6129 " &  CompTimeID & "</br>"

                If LeaveType.Name = CONST_LEAVE_TYPE_NAME_ELP then
                    ELPID = EE.AnnualVacation.ELPMatured.ELPID
                End If

                If IsAdminUpdating then
                    locCmd.CommandText = "usp_save_leave_request_admin"
                Else
                    locCmd.CommandText = "usp_save_new_leave_request_standard"
                End If

                on error resume next
                
                locCmd.Execute
                
                loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
                
                on error goto 0
                
                if loclngReturnValue <> 0 then
                    m_lngID = loclngReturnValue

                    If not IsAdminUpdating then
                        '**** IF THE USER REQUESTED A CONFIRMATION EMAIL, SEND THE USER A CONFIRMATION E-MAIL OF THIS REQUEST ****
                        If EmailConfOfRequestReq then
                            locstrBody = LeaveRequestEmailBody(CONST_EMAIL_TYPE_EE_NOTIFICATION)
                            mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Email, "e-Vacation - Leave Request Raised - Confirmation", locstrBody, True
                        
                        End If
                        
                        '**** SEND THE APPOINTED APPROVER AN E-MAIL OF THIS REQUEST ****
                        locstrBody = LeaveRequestEmailBody(CONST_EMAIL_TYPE_APPOINTED_APPROVER)
                        mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, AppointedApprover.Email, "e-Vacation - " & EE.FullName & " - Leave Request For Approval", locstrBody, True
                        
                        '**** IF THE MANAGER IS NOT THE APPOINTED APPROVER, SEND THE MANAGER AN INFORMATION ONLY E-MAIL OF THIS REQUEST ****
                        If EE.Manager.WWID <> AppointedApprover.WWID then
                            locstrBody = LeaveRequestEmailBody(CONST_EMAIL_TYPE_MANAGER_INFORMATION)
                            mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Manager.Email, "e-Vacation - For Information Only - Leave Request Raised by " & EE.FullName, locstrBody, True
                        End If
                    End If
                end if
            Else
                '*** Set up a return parameter, and specify which record we're about to update.
                locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
                locCmd.Parameters.Append locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)
                locCmd.Parameters.Append locCmd.CreateParameter("strResponseComments", adWChar, adParamInput, 100, trim(ResponseComments))
                '*** Set up the ADO command, depending on the status of the leave period object.
                Select Case Status
                    Case CONST_LEAVE_PERIOD_STATUS_REJECTED
                        locCmd.CommandText = "usp_reject_leave_request"
                    Case CONST_LEAVE_PERIOD_STATUS_APPROVED
                        locCmd.CommandText = "usp_approve_leave_request"
                    Case CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED
                        locCmd.CommandText = "usp_reject_cancel_request"
                    Case CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED
                        locCmd.CommandText = "usp_approve_cancel_request"
                    Case Else
                        Set locCmd = nothing
                        Save = 0
                        Exit Function
                End Select
 '    response.write "obj 6209 " & ID & "</br>"              
                on error resume next
                
                locCmd.Execute
                
                loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
                
                on error goto 0
                
                if loclngReturnValue <> 0 then
                    locstrBody = LeaveRequestEmailBody(CONST_EMAIL_TYPE_REQUEST_RESPONSE)
                    if Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED or _
                        Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED then
                        mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Email, "e-Vacation - Cancellation Request Response", locstrBody, True
                    Else
                        mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Email, "e-Vacation - Leave Request Response", locstrBody, True
                    End if
                end if
                
            End If

            Set locCmd = nothing

            Save = loclngReturnValue
 '************************************************   
     if LeaveType.ID = 25 then
  '   if LeaveType.ID = 25 then
     BookCompTime2()
     
          '  Dim loclngReturnValue
            Dim locstrSQL
        '    Dim locCmdRevoke
   '         Dim locCmd
            Dim locstrEmailBody
            Dim locWWID
            Dim m_lngWWID
            Dim locstrReason
            Dim locDateGranted
            Dim locDaysGranted
            Dim m_lngTaken
            Dim m_lngDaysBooked
            Dim locRevokeAmount
            Dim m_lngDaysGranted
    ' Revoke       
            Dim loclngDaysBooked 
            Dim locobjComptime
            Dim locRS
   
            locstrSQL = "SELECT * FROM tblcomptime " & _
                        "WHERE lngWWID=" & EE.WWID
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, _
			           glbConnection, _
			           adOpenStatic
' response.write "Obj 6243 " &  CompTimeID  & "</br>"
   '      locRevokeAmount = m_lngDaysGranted
             
         while not locRS.EOF
                m_lngID    = locRS("lngID")
                m_lngWWID = locRS("lngWWID")
                m_lngDaysGranted = locRS("lngDaysGranted")
        '        m_datDateGranted = locRS("datDateGranted")
        '       m_datDateRevoked = locRS("datDateRevoked")
       
     
            if locRevokeAmount > 0  AND m_lngDaysGranted > 0 then                 
	            Set locCmd = Server.CreateObject("ADODB.Command")
	            Set locCmd.ActiveConnection = glbConnection
				locCmd.CommandText = "UPDATE tblCompTime " &_
									 "SET strReason='" & m_strReason & "', lngDaysGranted=0, IngDaysBooked=1 " &_
									 "WHERE lngID=" & m_lngID							
				locCmd.Execute()
				
         
                   locRevokeAmount = locRevokeAmount -1
                  
              end if  
                locRS.movenext      
            Wend
            
            
     Set   locCmd = nothing
    end if 
 '  response.write "Obj 6280 " &  CompTimeID     & "</br>"
                        
    End Function
    
  Public Function BookCompTime2()
            Dim loclngReturnValue
            Dim locstrSQL
        '    Dim locCmdRevoke
            Dim locCmd
       Dim locstrEmailBody
            Dim locWWID
            Dim m_lngWWID
            Dim locstrReason
            Dim locDateGranted
            Dim locDaysGranted
            Dim m_lngTaken
            Dim m_lngDaysBooked
            Dim m_IngDaysRevoked
            Dim locRevokeAmount
            Dim m_lngDaysGranted
            Dim m_strReason
            Dim count
            Dim m_lngDaysRevoked
            Dim loclngDaysBooked 
            Dim locobjComptime
            Dim locRS
 '    response.write " obj 6724 " &  m_lngDaysGranted  & " "  &   m_strReason   
            locstrSQL = "SELECT * FROM tblcomptime " & _
                        "WHERE lngWWID=" & EE.WWID
            Set locRS = Server.CreateObject("ADODB.RecordSet")
            locRS.Open locstrSQL, _
			           glbConnection, _
			           adOpenStatic
               
         while not locRS.EOF AND count < 1
                m_lngID    = locRS("lngID")
                m_lngWWID = locRS("lngWWID")
                m_lngDaysGranted = locRS("lngDaysGranted")
                m_lngDaysBooked = locRS("lngDaysBooked")
                 m_lngDaysRevoked = locRS("IngDaysRevoked")
    
            if m_lngDaysGranted  > 0  AND count < 1 Then
         
	            Set locCmd = Server.CreateObject("ADODB.Command")
	            Set locCmd.ActiveConnection = glbConnection
					locCmd.CommandText = "UPDATE tblCompTime " &_
									 "SET strReason='" & m_strReason & "', lngDaysGranted=0, lngDaysBooked=1 " &_
									 "WHERE lngID=" & m_lngID							
				locCmd.Execute()
	
			end if
			
			locRS.movenext      
		Wend
            
            
     Set   locCmd = nothing
                
        End Function

        Public Function Cancel()
            Dim locCmdRevoke
            Dim locCmd
            Dim loclngReturnValue
            Dim locstrBody
            Dim locblnCancelRequiresApproval

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            loclngReturnValue = 0
            
            '*** If we have no ID set up then exit with a false return value.***
            If m_lngID = 0 then
                loclngReturnValue = 1
            Else
                '*** Store state of CancelRequiresApproval as this will change ***
                locblnCancelRequiresApproval = CancelRequiresApproval
                
                If locblnCancelRequiresApproval then
                    DateCancelApproved = Null
                Else
                    DateCancelApproved = Now()
                End If

                DateCancelRequested = Now()
                DateCancelRejected = Null

                locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
                locCmd.Parameters.Append locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)
                locCmd.Parameters.Append locCmd.CreateParameter("strApproverWWID", adWChar, adParamInput, 8, trim(Approver.WWID))
                locCmd.Parameters.Append locCmd.CreateParameter("strRequestComments", adWChar, adParamInput, 200, trim(RequestComments))
                locCmd.Parameters.Append locCmd.CreateParameter("datCancelRaised", adDBTimeStamp, adParamInput, , DateCancelRequested)
                locCmd.Parameters.Append locCmd.CreateParameter("datCancelApproved", adDBTimeStamp, adParamInput, , DateCancelApproved)
                locCmd.Parameters.Append locCmd.CreateParameter("datCancelRejected", adDBTimeStamp, adParamInput, , DateCancelRejected)
                                
                locCmd.CommandText = "usp_cancel_leave_request"

                on error resume next
                
                locCmd.Execute
                
                loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
                
                on error goto 0
                    
                if loclngReturnValue = 0 and locblnCancelRequiresApproval then
                    '**** IF THE USER REQUESTED A CONFIRMATION EMAIL, SEND THE USER A CONFIRMATION E-MAIL OF THIS REQUEST ****
                    If EmailConfOfRequestReq then
                        locstrBody = LeaveRequestEmailBody(CONST_EMAIL_TYPE_EE_CANCELLATION_REQUEST_NOTIFICATION)
                        mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Email, "e-Vacation - Leave Cancellation Request Raised - Confirmation", locstrBody, True
                    End If
                
                    '**** SEND THE APPOINTED APPROVER AN E-MAIL OF THIS CANCELLATION REQUEST ****
                    locstrBody = LeaveRequestEmailBody(CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER)
                    mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, AppointedApprover.Email, "e-Vacation - " & EE.FullName & " - Leave Cancellation Request For Approval", locstrBody, True
                    
                    '**** IF THE MANAGER IS NOT THE APPOINTED APPROVER, SEND THE MANAGER AN INFORMATION ONLY E-MAIL OF THIS CANCELLATION REQUEST ****
                    If EE.Manager.WWID <> AppointedApprover.WWID then
                        locstrBody = LeaveRequestEmailBody(CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION)
                        mSendEmail CONST_EMAIL_SYSTEM_EMAIL_FROM, EE.Manager.Email, "e-Vacation - For Information Only - Leave Cancellation Request Raised by " & EE.FullName, locstrBody, True
                    End If
                end if

            End If

            Set locCmd = nothing

            Cancel = loclngReturnValue
                        
        End Function
		
		' [MOF 21/11/08] Added Confirm function to confirm the day was taken
		Public Function Confirm()
            Dim locCmd
            Dim locCmdRevoke
			Dim locCmdSQL
            Dim loclngReturnValue
			
			'*** If we have no ID set up then exit with a false return value.***
            If m_lngID = 0 then
                loclngReturnValue = 1
			Else				
				' ** Store the state of 
				DateConfirmed = Now()
			
				' *** Confirm the leave period
	            Set locCmd = Server.CreateObject("ADODB.Command")
	            Set locCmd.ActiveConnection = glbConnection
				locCmd.CommandText = "UPDATE tblLeavePeriod " &_
									 "SET datConfirmed='" & Now() & "', datCancelRejected=NULL, datCancelRequested=NULL " &_
									 "WHERE lngID=" & ID								
				locCmd.Execute()
				
				loclngReturnValue = 0
			End If	

			Confirm = loclngReturnValue
		End Function
		' [MOF] End Function
        
        Public Sub LoadCancelRequestFromForm()
            Approver.WWID = request.form("fldstrApproverWWID")
            RequestComments = request.form("fldstrComments")
            If request.form("fldblnNotify") = "True" then
                EmailConfOfRequestReq = True
            Else
                EmailConfOfRequestReq = False
            End If
        End Sub
        
    End Class
    
    
    '*** LeaveType Object ***
    Class cObjLeaveType
        '=================================================================
        'Description:   Contains definition of a specific type of leave.
        'Properties:                                Type:   Perm:       Source:     Properties Required:
        'y  ID                                          lng     RW          DB
        'y  Name                                        str     RW          DB
        'y  EERequests                                  bln     RW          DB
        'y  AdminEnters                                 bln     RW          DB
        'y  RequestBeforeAccrued                        bln     RW          DB
        'y  MinDays                                     dbl     RW          DB
        'y  EntitlementAmount                           dbl     RW          DB
        'y  DaysBeforeStopsLegalAdjAccrual              dbl     RW          DB
        'y  DaysBeforeStopsLegalAdjAccrualIsConsecutive bln     RW          DB
        'y  StopsLegalAdjAccrual                        bln     R           Calc        (DaysBeforeStopsLegalAdjustmentAccrual)
        'y  HasEntitlement                              bln     R           Calc
        '
        'Methods:
        '   DBLoad
        '   
        '=================================================================
        Private m_lngID
        Private m_strName
        Private m_blnEERequests
        Private m_blnAdminEnters
        Private m_blnRequestBeforeAccrued
        Private m_dblMinDays
        Private m_dblEntitlementAmount
        Private m_dblDaysBeforeStopsLegalAdjAccrual
        Private m_blnDaysBeforeStopsLegalAdjAccrualIsConsecutive
        Private m_blnLoaded
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjLeaveType (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
            m_blnLoaded = False
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjLeaveType (" & m_lngtempobjid & "): " & timer & "<br>"
            '(nothing needed)
        End Sub

    
        Private Sub SetEmpty
            m_lngID = 0
            m_strName = ""
            m_blnEERequests = False
            m_blnAdminEnters = False
            m_blnRequestBeforeAccrued = False
            m_dblMinDays = 0
            m_dblEntitlementAmount = 0
            m_dblDaysBeforeStopsLegalAdjAccrual = -1
            m_blnDaysBeforeStopsLegalAdjAccrualIsConsecutive = False
        End Sub


        Public Property Let ID(varValue)
            If m_lngID <> 0 then
                If m_lngID <> varValue then
                    m_blnLoaded = False
                    SetEmpty
                    m_lngID = varValue
                End If
            Else
                m_blnLoaded = False
                SetEmpty
                m_lngID = varValue
            End If          
        End Property
        Public Property Get ID
            ID = m_lngID
        End Property

        
        Public Property Let Name(ByVal strValue)
            If m_strName <> "" then
                If m_strName <> strValue then
                    m_blnLoaded = False
                    SetEmpty
                    m_strName = strValue
                    DBLoad
                End If
            Else
                m_blnLoaded = False
                SetEmpty
                m_strName = strValue
                DBLoad
            End If
        End Property
        Public Property Get Name
            DBLoad
            Name = m_strName
        End Property

    
        Public Property Let EERequests(ByVal blnValue)
            m_blnLoaded = True
            m_blnEERequests = blnValue
        End Property
        Public Property Get EERequests
            DBLoad
            EERequests = m_blnEERequests
        End Property


        Public Property Let AdminEnters(ByVal blnValue)
            m_blnLoaded = True
            m_blnAdminEnters = blnValue
        End Property
        Public Property Get AdminEnters
            DBLoad
            AdminEnters = m_blnAdminEnters
        End Property


        Public Property Let RequestBeforeAccrued(ByVal blnValue)
            m_blnLoaded = True
            m_blnRequestBeforeAccrued = blnValue
        End Property
        Public Property Get RequestBeforeAccrued
            DBLoad
            RequestBeforeAccrued = m_blnRequestBeforeAccrued
        End Property


        Public Property Let MinDays(ByVal dblValue)
            m_blnLoaded = True
            m_dblMinDays = dblValue
        End Property
        Public Property Get MinDays
            DBLoad
            MinDays = m_dblMinDays
        End Property


        Public Property Let EntitlementAmount(ByVal dblValue)
            m_blnLoaded = True
            m_dblEntitlementAmount = dblValue
        End Property
        Public Property Get EntitlementAmount
            DBLoad
            EntitlementAmount = m_dblEntitlementAmount
        End Property


        Public Property Let DaysBeforeStopsLegalAdjAccrual(ByVal dblValue)
            m_blnLoaded = True
            m_dblDaysBeforeStopsLegalAdjAccrual = dblValue
        End Property
        Public Property Get DaysBeforeStopsLegalAdjAccrual
            DBLoad
            DaysBeforeStopsLegalAdjAccrual = m_dblDaysBeforeStopsLegalAdjAccrual
        End Property
        
        
        Public Property Let DaysBeforeStopsLegalAdjAccrualIsConsecutive(ByVal blnValue)
            m_blnLoaded = True
            m_blnDaysBeforeStopsLegalAdjAccrualIsConsecutive = blnValue
        End Property
        Public Property Get DaysBeforeStopsLegalAdjAccrualIsConsecutive
            DBLoad
            DaysBeforeStopsLegalAdjAccrualIsConsecutive = m_blnDaysBeforeStopsLegalAdjAccrualIsConsecutive
        End Property


        Public Property Get StopsLegalAdjAccrual
            If DaysBeforeStopsLegalAdjAccrual < 0 then
                StopsLegalAdjAccrual = False
            Else
                StopsLegalAdjAccrual = True
            End If
        End Property
                

        Public Property Get HasEntitlement
            If EntitlementAmount <> 0 then
                HasEntitlement = True
            Else
                HasEntitlement = False
            End If
        End Property


        Private Sub DBLoad
            Dim locCmd
            Dim locParam
            Dim locRS       
            Dim locMyID
            Dim locCmdRevoke
        
            '*** If we are already loaded, then exit.
            If m_blnLoaded then
                Exit Sub
            End If

            m_blnLoaded = True          
            
            '*** If we have no ID set up, we can't find our LeaveType - so exit the routine.
            If ID = 0 and Name = "" then
                exit sub
            End if

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection

            locCmd.CommandType = adCmdStoredProc

            If ID <> 0 then
                locCmd.CommandText = "usp_leavetype_by_id"
                Set locParam = locCmd.CreateParameter("lngID", adInteger, adParamInput, , ID)
            else
                locCmd.CommandText = "usp_leavetype_by_name"
                Set locParam = locCmd.CreateParameter("strName", adChar, adParamInput, 30, Name)
            end if
            
            locCmd.Parameters.Append locParam
            
            Set locRS = locCmd.Execute

            If not locRS.EOF then
                m_lngID = locRS("lngID")
                m_strName = locRS("strLeaveTypeName")
                m_blnEERequests = locRS("blnEERequests")
                m_blnAdminEnters = locRS("blnAdminRequests")
                m_blnRequestBeforeAccrued = locRS("blnRequestBeforeAccrued")
                m_dblMinDays = locRS("dblMinimumDays")
                m_dblEntitlementAmount = locRS("dblEntitlement")
                m_dblDaysBeforeStopsLegalAdjAccrual = locRS("dblDaysBeforeStopsLegalAdjAccrual")
                m_blnDaysBeforeStopsLegalAdjAccrualIsConsecutive = locRS("blnDaysBeforeStopsIsConsecutive")
            Else
                SetEmpty
            End If

            locRS.Close
            
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
        End Sub

    End Class
    

    '*** Other Leave Object ***
    Class cObjOtherLeave
        '=================================================================
        'Description:   Contains definition of a specific type of leave.
        'Properties:                Type:               Perm:           Source:     Properties Required:
        'y      EE                  objUser             RW              Virtual
        'y      Year                int                                 Virtual
        '       LeaveGroups         col(cColLeavePeriods)
        '       LeaveDaysInPeriod
        '
        'Methods:
        '       DBLoad
        'y      TestForChangedYear
        '
        '=================================================================
        Private m_objEE
        Private m_lngYear
        Private m_colLeaveGroups
        Private m_blnLoaded

        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjOtherLeave (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE = new cObjUser
            SetEmpty
            SetNotLoaded
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjOtherLeave (" & m_lngtempobjid & "): " & timer & "<br>"
            Set m_objEE  = nothing
            Set m_colLeaveGroups = nothing
        End Sub


        Private Sub SetNotLoaded
            m_blnLoaded = False
        End Sub
        
        
        Private Sub SetEmpty()
            m_lngYear = 0 
            Set m_colLeaveGroups = nothing
        End Sub


        Public Property Set EE(ByRef objUser)
            SetEmpty
            SetNotLoaded
            Set m_objEE = new cObjUser
            m_objEE.WWID = objUser.WWID
            m_objEE.YearToView = objUser.YearToView
        End Property
        
        Public Property Get EE()
            Set EE = m_objEE
        End Property


        Public Property Get Year()
            TestForChangedYear
            DBLoad
            Year = m_lngYear
        End Property


        Public Property Get LeaveGroups()
            
            TestForChangedYear
            DBLoad
            Set LeaveGroups = m_colLeaveGroups
        End Property


        Public Property Get LeaveDaysInPeriod(ByVal datStart, ByVal datEnd, ByVal blnOnlyAffectingLegalAllowance)
            
            Dim loclngCounter
            Dim loclngCount
            Dim loclngLeaveDaysInPeriod
            
            loclngCounter = 0
            loclngCount = LeaveGroups.Count
            loclngLeaveDaysInPeriod = 0
            
            While loclngCounter < loclngCount
                loclngCounter = loclngCounter + 1
                loclngLeaveDaysInPeriod = loclngLeaveDaysInPeriod + LeaveGroups.Item(loclngCounter).LeaveDaysInPeriod(datStart, datEnd, blnOnlyAffectingLegalAllowance, False)
            Wend
            
            LeaveDaysInPeriod = loclngLeaveDaysInPeriod
        End Property


        Private Sub TestForChangedYear()
            
            If EE.YearToView <> m_lngYear then
                SetEmpty
                SetNotLoaded
                m_lngYear = EE.YearToView
                
            End If
        End Sub


        Private Sub DBLoad()
            Dim locColLeaveTypes
            Dim loclngCounter
            Dim locColLeavePeriods
    
            If m_blnLoaded then
                Exit Sub
            End If

            '*** If we have no EE.WWID set up, we can't find our leave requests for this EE - so exit the routine.
            If EE.WWID = "" then
                exit sub
            End if
            
            m_blnLoaded = True

            Set m_colLeaveGroups = nothing
            Set m_colLeaveGroups = new cObjCollection
    
            mDebugPrint "   Creating cColLeavePeriods - 7 <br>"
            Set locColLeaveTypes = new cColLeaveTypes
            Set locColLeaveTypes.EE = EE
            
            locColLeaveTypes.CollectionType = CONST_LEAVE_TYPE_COLLECTION_TYPE_OTHER_LEAVE_TYPES_FOR_EE
    
            loclngCounter = 0
            While loclngCounter < locColLeaveTypes.Count
            
                loclngCounter = loclngCounter + 1
                
                '*** Set up a new collection of leave periods for the EE and for this particular leave type. ***
                mDebugPrint "   Creating cColLeavePeriods - 3 <br>"
                Set locColLeavePeriods = new cColLeavePeriods
                Set locColLeavePeriods.EE = EE
                locColLeavePeriods.CollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_OTHER_LEAVE
                locColLeavePeriods.CollectionLeaveType = locColLeaveTypes.Item(loclngCounter).Name
                
                m_colLeaveGroups.Add locColLeavePeriods
                Set locColLeavePeriods = nothing
            Wend
            
            Set locColLeaveTypes = nothing
            
        End Sub
    End Class


    '*** PublicHoliday Object ***
    Class cObjPublicHoliday
        '=================================================================
        'Description:   Contains definitions of public holidays.
        'Properties:                Type:   Perm:       Source:     Properties Required:
        '   Date                    dat     RW          DB
        '   Description             str     RW          DB
        '
        'Methods:
        '
        '=================================================================
        Private m_strDescription
        Private m_datDate 
        
        Public m_lngtempobjid
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjPublicHoliday (" & m_lngtempobjid & "): " & timer & "<br>"
            SetEmpty
        End Sub

        
        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjPublicHoliday (" & m_lngtempobjid & "): " & timer & "<br>"
            '(nothing needed)
        End Sub

    
        Private Sub SetEmpty()
            m_strDescription = ""
            m_datDate = ""
        End Sub
        
        
        Public Property Let Date(ByVal datValue)
            If IsDate(datValue) then
                If datValue <> m_datDate then
                    m_datDate = datValue
                End If
            Else
                m_datDate = ""
            End If
        End Property
        Public Property Get Date() 
            Date = m_datDate
        End Property
        
        
        Public Property Let Description(ByVal strValue)
            If strValue <> m_strDescription then
                m_strDescription = strValue
            End If          
        End Property
        Public Property Get Description() 
            Description = m_strDescription
        End Property
    End Class
    

    '*** User Object ***
    Class cObjUser
        '=================================================================
        'Description:   Contains definition of a User (Intel Employee).
        'Properties:                Type:   Perm:       Source:     Properties Required:
        'y  WWID                    str     RW          DB (WDS)
        'y  LocalRecordExists       bln     R           Virtual
        'y  YearToView              lng     RW          Virtual
        'y  IDSID                   str     R           DB (WDS)
        'y  FirstNm               str     R           DB (WDS)
        'y  LastNm                str     R           DB (WDS)
        'y  Email                   str     R           DB (WDS)
        'y  DOB                     dat     RW          DB (Local)
        'y  CompanyCd              str     RW          DB (WDS)
        'y  MailStopTxt                str     RW          DB (WDS)
        'y  LDOH                    dat     R           DB (WDS)
        'y  ODOH                    dat     R           DB (WDS)
        '-  AcquisitionDate         dat     RW          DB (Local)
        'y  TerminationDate         dat     RW          DB (WDS)
        'y  IsExempt                bln     R           DB (WDS)
        'y  ExemptStatus        	str     R           Calc
        'y  IsBlueBadge             bln     R           DB (WDS)
        'y  IsPartTimer             bln     R           DB (WDS)
        'y  ActiveStatus            str     RW          DB (WDS)
        'y  IsExemptStatusChanged   bln                 DB (Local)
        'y  IsException             bln                 DB (Local)
        'y  ExceptionComments       str                 DB (Local)
        'y  IsAdmin                 bln                 DB (Local)
        'y  Manager                 objUser             DB (ManagerWWID -> WWID)
        'y  DirectReports           col (Users)         DB (WWID -> ManagerWWID)
        'y  ActiveDelegate          objUser             DB (DelegateWWID -> WWID)
        'y  DelegateForManagers     col (Users)         DB (WWID -> DelegateWWID)
        'y  LeaveRequests           col (LeavePeriods)  DB (WWID)
        'y  Approvals               col (LeavePeriods)  DB (WWID)
        'y  AnnualVacation          objAnnualVacation   DB (WWID)
        'y  Otherleave              objOtherLeave       DB (WWID)
        'y  IsEELeaveTracked        bln                 DB (Local)
        'y  IsValidUser             bln     R           CALC        (IsAdmin, IsEELeaveTracked, IsLeaveApprover, IsDelegate, IsManager)
        'y  FullName                str     R           Calc        (FirstNm, LastNm)
        'y  FullNameReversed        str     R           Calc        (FirstNm, LastNm)
        'y  Age                     lng     R           Calc        (DOB)
        'y  IsDOBRequired           bln     R           Calc        (IsEELeaveTracked,IsBlueBadge,IsExempt)
        'y  SeniorityDate           dat     R           Calc        (ODOH)
        'y  StartDate               dat     R           Calc        (LDOH)
        'y  ActiveStatusName        str     R           Calc        (ActiveStatus)
        'y  ActiveThisYear          bln                 Calc        (YearToView, TerminationDate, StartDate)
        'y  IsManager               bln                 Calc        (DirectReports)
        'y  IsDelegate              bln                 Calc        (DelegateForManagers)
        'y  IsLeaveApprover         bln                 Calc        (Approvals)
        'y  IsInvalidField(strField)    bln             R           Calc        (m_strErrorFieldList)
        'y  ApprovalsPendingApproval    lng             Calc        (IsLeaveApprover,Approvals)
        'y  HasDelegate             bln                 Calc        (ActiveDelegate)
        'y  DelegateFormErrorMessage    str             Calc        
        'y  DelegateFormIsValid     bln                 Calc
        '-  LeaveStatus             str (Constant)      Calc        (AnnualVacation,OtherLeave)
        'y  AdminSelectUserFormErrorMessage     str     Calc        (WWID,IDSID,ActiveThisYear)
        'y  AdminSelectUserFormIsValid          bln     Calc        (AdminSelectUserFormErrorMessage)
        'y  AdminUserProfileFormErrorMessage    str     Calc
        'y  AdminUserProfileFormIsValid bln             Calc
        'y  CarryOversEOY           col(cObjCarryOver)  DB (WWID)
        'y  CarryOversPreArranged   col(cObjCarryOver)  DB (WWID)
        'y  CarryOverEOYForYear(lngYear)                Virtual
        'y  CarryOverPreArrangedForYear(lngYear)        Virtual
        '   DaysWorkedInPeriod(ByVal locdatStartDate, ByVal locdatEndDate)
		'	DaysConfirmed(ByVal locdatStartDate, ByVal locdatEndDate)
		'   DaysToConfirm  (Added by MOF 1/09)
        'Methods:
        '   DBLoad
        '   DBLoadDelegateForManagers
        '   DBLoadAnnualVacation
        '   DBLoadOtherLeave
        '   SetToLoggedOnUser
        '   RefreshApprovals
        '   LoadDelegateFromForm
        '   RemoveDelegate()
        '   Save()
        '   CreateLocalUserRecord()
        '
        '=================================================================
        Private m_strWWID
        Private m_blnLocalRecordExists
        Private m_lngYearToView
        Private m_strIDSID
        Private m_blnIsEELeaveTracked
        Private m_strFirstNm 
        Private m_strLastNm 
        Private m_strEmail 
        Private m_datDOB
        Private m_endDate
        Private m_orgUnit
        Private m_strCompanyCd
        Private m_strMailStopTxt
        Private m_datLDOH
        Private m_datODOH
        Private m_datAcquisitionDate
        Private m_datTerminationDate
        Private m_blnIsExempt
        Private m_blnIsBlueBadge
        Private m_blnIsPartTimer
        Private m_strActiveStatus
        Private m_blnIsExemptStatusChanged
        Private m_blnIsException
        Private m_strExceptionComments 
        Private m_blnIsAdmin
        Private m_strNextLevelWWID
        Private m_objManager
        Private m_colDirectReports
        Private m_strActiveDelegateWWID
        Private m_objActiveDelegate
        Private m_colDelegateForManagers
        Private m_colLeaveRequests
        Private m_colApprovals
        Private m_objAnnualVacation
        Private m_objOtherLeave
        Private m_colCarryOversEOY
        Private m_colCarryOversPreArranged
        Public m_strErrorFieldList

        Private m_blnLoaded
        Private m_blnDelegateForManagersLoaded
        Private m_blnAnnualVacationLoaded
        Private m_blnOtherLeaveLoaded
        Private m_blnCarryOversLoaded
        
        Public m_lngtempobjid
        
        Private Sub Class_Initialize()
            glbObjectCounter = glbObjectCounter + 1
            m_lngtempobjid = glbObjectCounter
            mDebugPrint "Initializing cObjUser (" & m_lngtempobjid & "): " & timer & "<br>"
            m_strWWID = ""
            SetEmpty
            SetNotLoaded
            m_lngYearToView = DatePart("yyyy",now())
        End Sub


        Private Sub Class_Terminate()
            glbObjectTerminateCounter = glbObjectTerminateCounter + 1
            mDebugPrint "Terminating cObjUser (" & m_lngtempobjid & ") (" & WWID & "): " & timer & "<br>"
            Set m_objManager = nothing
            Set m_colDirectReports = nothing
            Set m_objActiveDelegate = nothing
            Set m_colDelegateForManagers = nothing
            Set m_colLeaveRequests = nothing
            Set m_colApprovals = nothing
            Set m_objAnnualVacation = nothing
            Set m_objOtherLeave = nothing
            Set m_colCarryOversEOY = nothing
            Set m_colCarryOversPreArranged = nothing
        End Sub


        Private Sub SetNotLoaded()
            m_blnLoaded = False
            m_blnDelegateForManagersLoaded = False
            m_blnAnnualVacationLoaded = False
            m_blnOtherLeaveLoaded = False
            m_blnCarryOversLoaded = False
        End Sub
        
            
        Private Sub SetEmpty()
            m_blnLocalRecordExists = False
            m_strIDSID = ""
            m_blnIsEELeaveTracked = False
            m_strFirstNm =""
            m_strLastNm =""
            m_strEmail = ""
            m_datDOB = ""
            m_endDate = "" 
            m_strCompanyCd = ""
            m_strMailStopTxt = ""
            m_datLDOH = ""
            m_datODOH = ""
            m_datAcquisitionDate = ""
            m_datTerminationDate = ""
            m_blnIsExempt = False
            m_blnIsBlueBadge = False
            m_blnIsPartTimer = False
            m_strActiveStatus = ""
            m_blnIsExemptStatusChanged = False
            m_blnIsException = False
            m_strExceptionComments = ""
            m_blnIsAdmin = False
            
            m_strNextLevelWWID = ""
            Set m_objManager = nothing
            
            m_strActiveDelegateWWID = ""
            Set m_objActiveDelegate = nothing

            Set m_colApprovals = nothing
            
            Set m_colLeaveRequests = nothing
            
            Set m_objAnnualVacation = nothing
            
            Set m_objOtherLeave = nothing
            
            Set m_colCarryOversEOY = nothing
            Set m_colCarryOversPreArranged = nothing
            
        End Sub


        Public Property Let WWID(ByVal strValue)
            '==============================================================
            'If the WWID has not already been set then this routine sets it to the value passed to it.
            'If the WWID is set already and the value is the same as that passed to this routine, nothing happens.
            'If the WWID is set already and the value is different, the object is emptied using the SetEmpty routine and
            'the WWID value is set to the new value (the previous user object is effectively lost).
            'If the WWID property is changed, the m_blnLoaded flag is set to false, to force other properties to be
            'loaded from the DB as necessary.
            '==============================================================
            If m_strWWID <> "" then
                If m_strWWID <> strValue then
                    SetNotLoaded
                    SetEmpty
                    m_strWWID = strValue
                End If
            Else
                SetNotLoaded
                SetEmpty
                m_strWWID = strValue
				mDebugPrint "Let WWID - value set as '" & m_strWWID & "'<br>"
            End If
        End Property
        Public Property Get WWID()
            WWID = trim(m_strWWID)
        End Property


        Private Property Let LocalRecordExists(ByVal blnValue)
            m_blnLocalRecordExists = blnValue
        End Property
        Public Property Get LocalRecordExists
            DBLoad
            LocalRecordExists = m_blnLocalRecordExists
			
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - LocalRecordExists" ) 
			'Response.Write( "<br>" ) 	
        End Property
        
        
        Public Property Let YearToView(ByVal lngValue)
            If mGetSafeLongInteger(lngValue,0) <> 0 then
                m_lngYearToView = lngValue
            Else
                m_lngYearToView = 0
            End If
            If m_lngYearToView < CONST_FIRST_YEAR_SYSTEM_ACTIVE then
                m_lngYearToView = CONST_FIRST_YEAR_SYSTEM_ACTIVE
            End If
        End Property
        Public Property Get YearToView()
            YearToView = m_lngYearToView
        End Property


        Public Property Let IDSID(ByVal strValue)
            m_blnLoaded = True
            m_strIDSID = strValue
        End Property
        Public Property Get IDSID() 
            DBLoad
            IDSID = m_strIDSID
        End Property


        Public Property Let FirstNm(ByVal strValue)
            m_blnLoaded = True
            m_strFirstNm = strValue
        End Property
        Public Property Get FirstNm() 
            DBLoad
            FirstNm = m_strFirstNm
        End Property


        Public Property Let LastNm(ByVal strValue)
            m_blnLoaded = True
            m_strLastNm = strValue
        End Property
        Public Property Get LastNm() 
            DBLoad
            LastNm = m_strLastNm
        End Property
        
        
        Public Property Let Email(ByVal strValue)
            m_blnLoaded = True
            m_strEmail = strValue
        End Property
        Public Property Get Email() 
            DBLoad
            Email = m_strEmail
        End Property        

        Public Property Let DOB(ByVal datValue)
            m_blnLoaded = True
            m_datDOB = datValue
        End Property
        Public Property Get DOB()
            DBLoad
            DOB = m_datDOB
        End Property

        Public Property Let EndDate(ByVal datValue)
            m_blnLoaded = True
            m_endDate = datValue
        End Property
        Public Property Get EndDate()
            DBLoad
            EndDate = m_endDate
        End Property
        
        Public Property Let OrgUnit(ByVal strValue)
            m_blnLoaded = True
            m_orgUnit = strValue
        End Property
        Public Property Get OrgUnit()
            DBLoad
            OrgUnit = m_orgUnit
        End Property
        
        Public Property Let CompanyCd(ByVal strValue)
            m_blnLoaded = True
            m_strCompanyCd = strValue
        End Property
        Public Property Get CompanyCd()
            DBLoad
            CompanyCd = m_strCompanyCd
        End Property
        
        
        Public Property Let MailStopTxt(ByVal strValue)
            m_blnLoaded = True
            m_strMailStopTxt = strValue
        End Property
        Public Property Get MailStopTxt()
            MailStopTxt = m_strMailStopTxt
        End Property
        

        Public Property Let LDOH(ByVal datValue)
            m_blnLoaded = True
            m_datLDOH = datValue
        End Property
        Public Property Get LDOH()
            DBLoad
            LDOH = m_datLDOH
        End Property
        

        Public Property Let ODOH(ByVal datValue)
            m_blnLoaded = True
            m_datODOH = datValue
        End Property
        Public Property Get ODOH()
            DBLoad
            ODOH = m_datODOH
        End Property
        
        
        Public Property Get AcquisitionDate() 
            DBLoad
            AcquisitionDate = m_datAcquisitionDate
        End Property
        

        Public Property Let TerminationDate(ByVal datValue)
            m_blnLoaded = True
            If IsDate(datValue) then
                m_datTerminationDate = datValue
            Else
                m_datTerminationDate = ""
            End If
        End Property
        Public Property Get TerminationDate() 
            DBLoad
            TerminationDate = m_datTerminationDate
        End Property

        
        Public Property Let IsExempt(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsExempt = blnValue
        End Property
        Public Property Get IsExempt()
            DBLoad
            IsExempt = m_blnIsExempt
        End Property
           
        Public Property Get ExemptStatus()
            If IsExempt then
                ExemptStatus = "Exempt"
            Else
                ExemptStatus = "Non-exempt"
            End If
        End Property
        Public Property Let IsBlueBadge(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsBlueBadge = blnValue
        End Property
        Public Property Get IsBlueBadge()
            DBLoad
            IsBlueBadge = m_blnIsBlueBadge
        End Property
        
        
        Public Property Let IsPartTimer(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsPartTimer = blnValue
        End Property
        Public Property Get IsPartTimer() 
            DBLoad
            IsPartTimer = m_blnIsPartTimer
        End Property
        

        Public Property Let ActiveStatus(ByVal strValue)
            m_blnLoaded = True
            m_strActiveStatus = strValue        
        End Property
        Public Property Get ActiveStatus()
            'Codes: A ACTIVE
            '       H PRIOR-TO-HIRE
            '       L LEAVE
            '       P PAID LEAVE OF ABSENCE
            '       S SUSPEND
            '       T TERMINATED
            DBLoad
            ActiveStatus = m_strActiveStatus
        End Property
    
    
        Public Property Let IsExemptStatusChanged(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsExemptStatusChanged = blnValue
        End Property
        Public Property Get IsExemptStatusChanged() 
            DBLoad
            IsExemptStatusChanged = m_blnIsExemptStatusChanged
        End Property


        Public Property Let IsException(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsException = blnValue
        End Property
        Public Property Get IsException() 
            DBLoad
            IsException = m_blnIsException
        End Property
        
        
        Public Property Let ExceptionComments(ByVal strValue)
            m_blnLoaded = True
            m_strExceptionComments = strValue
        End Property
        Public Property Get ExceptionComments() 
            DBLoad
            ExceptionComments = m_strExceptionComments
        End Property


        Public Property Let IsAdmin(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsAdmin = blnValue
        End Property
        Public Property Get IsAdmin() 
            DBLoad
            IsAdmin = m_blnIsAdmin
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - IsAdmin" ) 
			'Response.Write( "<br>" ) 
        End Property
        
        
        Public Property Get Manager() 
            DBLoad
            If typename(m_objManager) <> "cObjUser" then
                Set m_objManager = new cObjUser
                m_objManager.WWID = m_strNextLevelWWID
            End If
            Set Manager = m_objManager
        End Property
        
        
        Public Property Let ManagerWWID(ByVal strValue)
            If (m_strNextLevelWWID <> strValue) or (strValue = "") then
                Set m_objManager = nothing
                m_strNextLevelWWID = strValue
            End If
        End Property
        Public Property Get ManagerWWID()
            DBLoad
            ManagerWWID = m_strNextLevelWWID
        End Property
        
        
        Public Property Get DirectReports()
            Initialise_DirectReports
            Set DirectReports = m_colDirectReports
        End Property
        
                
        Public Property Get ActiveDelegate() 
            DBLoad
            If typename(m_objActiveDelegate) <> "cObjUser" then
                Set m_objActiveDelegate = new cObjUser
                m_objActiveDelegate.WWID = m_strActiveDelegateWWID
            End If
            Set ActiveDelegate = m_objActiveDelegate
        End Property
        

        Public Property Let ActiveDelegateWWID(ByVal strValue)
            If (m_strActiveDelegateWWID <> strValue) or (strValue = "") then
                Set m_objActiveDelegate = nothing
                m_strActiveDelegateWWID = strValue
            End If
        End Property
        Public Property Get ActiveDelegateWWID()
            DBLoad
            ActiveDelegateWWID = m_strActiveDelegateWWID
        End Property

        
        Public Property Get DelegateForManagers()
            DBLoadDelegateForManagers
            Set DelegateForManagers = m_colDelegateForManagers
        End Property
        
        
        Public Property Get LeaveRequests() 
            InitialiseLeaveRequests
            Set LeaveRequests = m_colLeaveRequests
        End Property
        
        Public Property Get Approvals() 
            InitialiseApprovals
            Set Approvals = m_colApprovals
			
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - Approvals " ) 
			'Response.Write( "<br>" ) 
			
        End Property
        
        
        Public Property Get AnnualVacation() 
            DBLoadAnnualVacation
            Set AnnualVacation = m_objAnnualVacation
        End Property


        Public Property Get OtherLeave() 
            
            DBLoadOtherLeave
            Set OtherLeave = m_objOtherLeave
        End Property


        Public Property Let IsEELeaveTracked(ByVal blnValue)
            m_blnLoaded = True
            m_blnIsEELeaveTracked = blnValue
        End Property
        Public Property Get IsEELeaveTracked()
            DBLoad
            If IsNull(m_blnIsEELeaveTracked) then
                IsEELeaveTracked = mCalculateIsEELeaveTracked(Me)
            Else
                IsEELeaveTracked = m_blnIsEELeaveTracked
            End If
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - IsEELeaveTracked" ) 
			'Response.Write( "<br>" ) 
        End Property


        Public Property Get IsValidUser()
            If IsAdmin OR IsEELeaveTracked OR IsLeaveApprover OR IsDelegate OR IsManager then
                IsValidUser = True
				'dim endtime
				'endtime = timer()
				'dim benchmark
				'benchmark = endtime - starttime
				'Response.Write( benchmark ) 
				'Response.Write( " - IsValidUser" ) 
				'Response.Write( "<br>" ) 
            Else
                IsValidUser = False
            End If
        End Property
        
        
        Public Property Get FullName()
            FullName = FirstNm & " " & LastNm
        End Property


        Public Property Get FullNameReversed()
            FullNameReversed = UCase(LastNm) & ", " & UCase(FirstNm)
        End Property
        

        Public Property Get Age()
            '==============================================================
            'Calculates the age of the user from the DOB property. If the DOB property is not a valid date,
            'it returns 'Null'. The property checks (using datepart() function) to see if the user has had
            'his or her birthday this year and takes one off the year count if not.
            '==============================================================
            Age = mCountWholeYears(DOB, Now)
        End Property
        
        
        Public Property Get IsDOBRequired()
            If IsEELeaveTracked AND IsBlueBadge AND IsExempt then
                IsDOBRequired = True
            Else
                IsDOBRequired = False
            End If
        End Property
        
        
        Public Property Get SeniorityDate()
            SeniorityDate = ODOH
        End Property
        
        
        Public Property Get StartDate() 
            StartDate = LDOH
        End Property

        
        Public Property Get ActiveStatusName()
            Select Case ActiveStatus
                Case "A"
                    ActiveStatusName = "Active"
                Case "H"
                    ActiveStatusName = "Starts " & mFormatDate(LDOH,"medium")
                Case "L"
                    ActiveStatusName = "Unpaid LOA"
                Case "P"
                    ActiveStatusName = "Paid LOA"
                Case "S"
                    ActiveStatusName = "Suspended"
                Case "T"
                    ActiveStatusName = "Terminated " & mFormatDate(TerminationDate,"medium")
                Case Else
                    ActiveStatusName = ActiveStatus & " - unknown"
            End Select
        End Property
        
                    
        Public Property Get ActiveThisYear()
            'Was this employee terminated before the year being viewed?
            If isDate(TerminationDate) then
                if datepart("yyyy",TerminationDate) < YearToView then
                    ActiveThisYear = False
                    Exit Property
                End If
            End If
            'Did this employee start at all?
            If not isdate(StartDate) then
                ActiveThisYear = False
                Exit Property
            End If
            'Did this employee start after the year being viewed?
            If datepart("yyyy",StartDate) > YearToView then
                ActiveThisYear = False
                Exit Property
            End If
            'The employee is (or has been) active during the year being viewed.
            ActiveThisYear = True
        End Property
        
        
        Public Property Get IsManager() 
            If DirectReports.Count > 0 then
                IsManager = True
            Else
                IsManager = False
            End If
			
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - IsManager" ) 
			'Response.Write( "<br>" ) 
        End Property
        
        
        Public Property Get IsDelegate()
            If DelegateForManagers.Count > 0 then
                IsDelegate = True
            Else
                IsDelegate = False
            End If			
			
			'dim endtime
			'endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - IsDelegate" ) 
			'Response.Write( "<br>" ) 
        End Property
        
        
        Public Property Get IsLeaveApprover() 	
			
			dim endtime
			endtime = timer()
			'dim benchmark
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - IsLeaveApprover" ) 
			'Response.Write( "<br>" ) 
			
            If Approvals.Count > 0 then
                IsLeaveApprover = True
            Else
                IsLeaveApprover = False
            End If
			
			'endtime = timer()
			'benchmark = endtime - starttime
			'Response.Write( benchmark ) 
			'Response.Write( " - IsLeaveApprover" ) 
			'Response.Write( "<br>" ) 
        End Property
        

        Public Property Get IsInvalidField(ByVal strFieldName)
            If instr(m_strErrorFieldList,strFieldName & ";") then
                IsInvalidField = True
            Else
                IsInvalidField = False
            End If
        End Property


        Public Property Get ApprovalsPendingApproval()
            Dim loclngCounter
            Dim loclngApprovalsCount
            Dim loclngApprovalsPendingApprovalCount
            If not IsLeaveApprover then
                ApprovalsPendingApproval = 0
            Else
                loclngCounter = 0
                loclngApprovalsCount = Approvals.Count
                loclngApprovalsPendingApprovalCount = 0
                With Approvals
                    While loclngCounter < loclngApprovalsCount
                        loclngCounter = loclngCounter + 1
                        If .Item(loclngCounter).AppointedApprover.WWID = WWID then
                            loclngApprovalsPendingApprovalCount = loclngApprovalsPendingApprovalCount + 1
                        End If
                    Wend
                End With
                ApprovalsPendingApproval = loclngApprovalsPendingApprovalCount
            End If
        End Property
		
		' Calculates the number of leave periods/requests Employee must confirm/cancel
		' [MOF 12/10/2008]
		Public Property Get LeaveRequestsPendingConfirmation()
			Dim loclngCounter
			Dim loclngLeaveRequestsCount
			Dim loclngConfirmCount
			
			loclngCounter = 1
			loclngLeaveRequestsCount = LeaveRequests.Count
			loclngConfirmCount = 0
			
			While loclngCounter <= loclngLeaveRequestsCount
				' Check if we can confirm the leave, if so then add to count
				If LeaveRequests.Item(loclngCounter).CanConfirmLeave and _
				   (LeaveRequests.Item(loclngCounter).Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
				    LeaveRequests.Item(loclngCounter).Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED) Then
					loclngConfirmCount = loclngConfirmCount + 1
				End If
				
				loclngCounter = loclngCounter + 1
			Wend
			
			LeaveRequestsPendingConfirmation = loclngConfirmCount
		End Property	
        
        
        Public Property Get HasDelegate()
            If not mIsEmptyString(m_strActiveDelegateWWID) then
                HasDelegate = True
            Else
                HasDelegate = False
            End If
        End Property
        
        
        Public Property Get LeaveStatus() 
        
            '***** TO BE CALCULATED *****

        End Property


        Public Property Get DelegateFormErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage

            m_strErrorFieldList = ""
            locblnValid = True
            
            '*** Check for valid ActiveDelegate.WWID***
            If not mIsValidWWID(m_strActiveDelegateWWID) then
                locstrErrorMessage = locstrErrorMessage & "  - The Delegate WWID is not a valid Intel WWID.\n"
                m_strErrorFieldList = m_strErrorFieldList & "ActiveDelegate.WWID;"
                locblnValid = False
            '*** User tring to appoint him / her self as delegate?              
            ElseIf m_strActiveDelegateWWID = WWID then
                locstrErrorMessage = locstrErrorMessage & "  - You are trying to appoint yourself as your delegate.\n"
                m_strErrorFieldList = m_strErrorFieldList & "ActiveDelegate.WWID;"
                locblnValid = False
            '*** Can the Active Delegate be found? ****
            ElseIf ActiveDelegate.IDSID = "" then
                locstrErrorMessage = locstrErrorMessage & "  - The WWID you entered could not be found in the Worker Data Services Database (WDS).\n"
                m_strErrorFieldList = m_strErrorFieldList & "ActiveDelegate.WWID;"
                locblnValid = False
            ElseIf ActiveDelegate.ActiveStatus <> "A" then
                locstrErrorMessage = locstrErrorMessage & "  - The status of " & ActiveDelegate.FullName & " is \'" & ActiveDelegate.ActiveStatusName & "\' which is not valid for a delegate.\n"
                m_strErrorFieldList = m_strErrorFieldList & "ActiveDelegate.WWID;"
                locblnValid = False
            ElseIf ActiveDelegate.HasDelegate then
                locstrErrorMessage = locstrErrorMessage & "  - The employee you specified as your delegate (" & ActiveDelegate.FullName & ") has already appointed a delegate of their own.\n"
                m_strErrorFieldList = m_strErrorFieldList & "ActiveDelegate.WWID;"
                locblnValid = False
            End If

            
            If locblnValid = True then
                DelegateFormErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct and re-submit the form."
                DelegateFormErrorMessage = replace(locstrErrorMessage,"'","")
            End If

        End Property
        

        Public Property Get DelegateFormIsValid()
            If DelegateFormErrorMessage = "" then
                DelegateFormIsValid = True
            Else
                DelegateFormIsValid = False
            End If
        End Property


        Public Property Get AdminSelectUserFormErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage

            m_strErrorFieldList = ""
            locblnValid = True
            
            '*** Check for valid WWID***
            If not mIsValidWWID(WWID) then
                locstrErrorMessage = locstrErrorMessage & "  - The WWID is not a valid Intel WWID.\n"
                m_strErrorFieldList = m_strErrorFieldList & "WWID;"
                locblnValid = False
            '*** Admin user trying to select him / her self to administer?              
            ElseIf WWID = objCurrentUser.WWID then
                locstrErrorMessage = locstrErrorMessage & "  - You can not select yourself when performing administration activities.\n"
                m_strErrorFieldList = m_strErrorFieldList & "WWID;"
                locblnValid = False
            '*** Can the select user be found? ****
            ElseIf IDSID = "" then
                locstrErrorMessage = locstrErrorMessage & "  - The WWID you entered could not be found in the Worker Data Services Database (WDS).\n"
                m_strErrorFieldList = m_strErrorFieldList & "WWID;"
                locblnValid = False
            ElseIf not ActiveThisYear then
                locstrErrorMessage = locstrErrorMessage & "  - The status of " & FullName & " is \'" & ActiveStatusName & "\', which is not a valid selection.\n"
                m_strErrorFieldList = m_strErrorFieldList & "WWID;"
                locblnValid = False
            End If

            
            If locblnValid = True then
                AdminSelectUserFormErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct and try again."
                AdminSelectUserFormErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        End Property
        

        Public Property Get AdminSelectUserFormIsValid()
            If AdminSelectUserFormErrorMessage = "" then
                AdminSelectUserFormIsValid = True
            Else
                AdminSelectUserFormIsValid = False
            End If
        End Property


        Public Property Get AdminUserProfileFormErrorMessage()
            Dim locblnValid
            Dim locstrErrorMessage

            m_strErrorFieldList = ""
            locblnValid = True
            
            If DOB <> "" then
                If not IsDate(DOB) then
                    locstrErrorMessage = locstrErrorMessage & "  - The Date of Birth is invalid.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DOB;"
                    locblnValid = False
                ElseIf not mIsFormattedDate(DOB) then
                    locstrErrorMessage = locstrErrorMessage & "  - The Date of Birth should be in (dd mmm yyyy) format - (e.g. 01 Jan 1970).\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DOB;"
                    locblnValid = False
                ElseIf Age < 15 then
                    locstrErrorMessage = locstrErrorMessage & "  - The Date of Birth does not appear to be correct.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "DOB;"
                    locblnValid = False
                End If
            End If
            
            If EndDate <> "" then
                If not IsDate(EndDate) then
                    locstrErrorMessage = locstrErrorMessage & "  - The Date of End of Contract is invalid.\n"
                    m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                    locblnValid = False
                ElseIf not mIsFormattedDate(EndDate) then
                    locstrErrorMessage = locstrErrorMessage & "  - The Date of End of Contract be in (dd mmm yyyy) format - (e.g. 01 Jan 1970).\n"
                    m_strErrorFieldList = m_strErrorFieldList & "EndDate;"
                    locblnValid = False
                End If
            End If
            
            
            If locblnValid = True then
                AdminUserProfileFormErrorMessage = ""
            Else
                locstrErrorMessage = "The following problems were found:\n\n " & _
                    locstrErrorMessage & _
                    "\nPlease correct and try again."
                AdminUserProfileFormErrorMessage = replace(locstrErrorMessage,"'","")
            End If
        End Property
        

        Public Property Get AdminUserProfileFormIsValid()
            If AdminUserProfileFormErrorMessage = "" then
                AdminUserProfileFormIsValid = True
            Else
                AdminUserProfileFormIsValid = False
            End If
        End Property


        Public Property Get CarryOversEOY
            DBLoadCarryOvers
            Initialise_CarryOversEOY
            Set CarryOversEOY = m_colCarryOversEOY
        End Property
        
        
        Public Property Get CarryOversPreArranged
            DBLoadCarryOvers
            Initialise_CarryOversPreArranged
            Set CarryOversPreArranged = m_colCarryOversPreArranged
        End Property
        
        
        Public Property Get CarryOverEOYForYear(ByVal lngYear)
            Dim loclngCounter
            Dim locobjCarryOver
            loclngCounter = 0
            While loclngCounter < CarryOversEOY.Count
                loclngCounter = loclngCounter + 1
                If CarryOversEOY.Item(loclngCounter).Year = lngYear then
                    Set CarryOverEOYForYear = CarryOversEOY.Item(loclngCounter)
                    Exit Property
                End If
            Wend
            
            Set locobjCarryOver = new cObjCarryOver
            Set locobjCarryOver.EE = Me
            locobjCarryOver.Year = lngYear
            locobjCarryOver.Days = 0
            locobjCarryOver.IsPreArranged = False
            CarryOversEOY.Add locobjCarryOver
            Set locobjCarryOver = nothing
            Set CarryOverEOYForYear = CarryOversEOY.Item(CarryOversEOY.Count)
        End Property
                

        Public Property Get CarryOverPreArrangedForYear(ByVal lngYear)
            Dim loclngCounter
            Dim locobjCarryOver
            loclngCounter = 0
            While loclngCounter < CarryOversPreArranged.Count
                loclngCounter = loclngCounter + 1
                If CarryOversPreArranged.Item(loclngCounter).Year = lngYear then
                    Set CarryOverPreArrangedForYear = CarryOversPreArranged.Item(loclngCounter)
                    Exit Property
                End If
            Wend
            
            Set locobjCarryOver = new cObjCarryOver
            Set locobjCarryOver.EE = Me
            locobjCarryOver.Year = lngYear
            locobjCarryOver.Days = 0
            locobjCarryOver.IsPreArranged = True
            CarryOversPreArranged.Add locobjCarryOver
            Set locobjCarryOver = nothing
            Set CarryOverPreArrangedForYear = CarryOversPreArranged.Item(CarryOversPreArranged.Count)
        End Property


        Public Property Get DaysWorkedInPeriod(ByVal locdatStartDate, ByVal locdatEndDate)
            Dim locdblWorkingDays
            Dim locdblWeekendDays
            Dim locdblPublicHolidays
            Dim locdblAnnualLeaveDays
            Dim locdblOtherLeaveDays
            If not isdate(locdatStartDate) or not isdate(locdatEndDate) then
                DaysWorkedInPeriod = 0
                Exit Property
            End If
            locdatStartDate = cDate(locdatStartDate)
            locdatEndDate = cDate(locdatEndDate)
            if locdatStartDate > locdatEndDate then
                DaysWorkedInPeriod = 0
                Exit Property
            End If
                
            locdblWorkingDays = DateDiff("y",locdatStartDate,locdatEndDate) + 1
            locdblWeekendDays = mCountWeekendDays(locdatStartDate,locdatEndDate)
            locdblPublicHolidays = glbPublicHolidays.CountInPeriod(locdatStartDate, locdatEndDate, False)
            locdblAnnualLeaveDays = AnnualVacation.Leave.LeaveDaysInPeriod(locdatStartDate, locdatEndDate, False, False)
            locdblOtherLeaveDays = OtherLeave.LeaveDaysInPeriod(locdatStartDate, locdatEndDate, False) 'CA remove 4th parameter - wrong LeaveDaysInPeriod
            
            DaysWorkedInPeriod = locdblWorkingDays - _
                locdblWeekendDays - _
                locdblPublicHolidays - _
                locdblAnnualLeaveDays - _
                locdblOtherLeaveDays
        End Property

		' [MOF 12/10/2008] Returns the number of days the user has confirmed for the set year
		Public Property Get DaysConfirmed
			Dim loclngLeavePeriodCount
			Dim locdatStart
			Dim locdatEnd
			Dim loclngDayCount
			loclngLeavePeriodCount = 1
			loclngDayCount = 0
						
			' Loop through all the leave periods
			For loclngLeavePeriodCount = 1 to LeaveRequests.Count
				' Check if the leave period is confirmed
				If LeaveRequests.Item(loclngLeavePeriodCount).Status = CONST_LEAVE_PERIOD_STATUS_CONFIRMED Then
					' Get Start/End dates of leave period
					locdatStart = LeaveRequests.Item(loclngLeavePeriodCount).StartDate
					locdatEnd = LeaveRequests.Item(loclngLeavePeriodCount).EndDate
				
					' Add the number of days taken in that period to the count
					loclngDayCount = loclngDayCount + LeaveRequests.Item(loclngLeavePeriodCount).LeaveDaysInPeriod(locdatStart, locdatEnd)
				End If
			Next
			
			' Return day count
			DaysConfirmed = loclngDayCount
		End Property
		
		' [MOF 1/09] Returns the number of leave days which the user has yet to confirm
		Public Property Get DaysToConfirm
		    Dim loclngLeavePeriodCount
		    Dim locobjLeavePeriod
		    Dim locdatStart 
		    Dim locdatEnd
		    Dim loclngDayCount
		    loclngDayCount = 0
		    
		    ' Loop through all the leave periods
		    For loclngLeavePeriodCount = 1 to LeaveRequests.Count
		        Set locobjLeavePeriod = LeaveRequests.Item(loclngLeavePeriodCount)
		        
		        ' Check if the leave period is expired (now > end-date) and 
		        ' status = approved or confirmed
		        If (locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
		            locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED or _
		            locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED) and _
		           locobjLeavePeriod.CanConfirmLeave Then
		            ' Get Start/End dates of leave period
					locdatStart = LeaveRequests.Item(loclngLeavePeriodCount).StartDate
					locdatEnd = LeaveRequests.Item(loclngLeavePeriodCount).EndDate
				
					' Add the number of days taken in that period to the count
					loclngDayCount = loclngDayCount + LeaveRequests.Item(loclngLeavePeriodCount).LeaveDaysInPeriod(locdatStart, locdatEnd)
		        End If
		    Next
		    
		    ' Return day count
		    DaysToConfirm = loclngDayCount
		End Property


        '***************************************
        '************** METHODS ****************
        '***************************************

        Private Sub DBLoad()
            'local connections
            Dim locCmd
            Dim locParam
            Dim locRS
            Dim locCmd2
            
            Dim locParam2
            Dim locRS2

            Dim condition
            
            Dim locCmdRevoke
			
            'wds connection
            Dim wdsRS
            Dim wdsConnection
            Dim locCmdwds
			
            'vars
            Dim insertRequired 
            Dim updateRequired 
					
            On Error Resume Next
           
            '*** If we are already loaded, then exit.
            If m_blnLoaded then
                Exit Sub
            else 
                m_blnLoaded = True
            End If

            '*** If we have no WWID set up, and no IDSID set up we can't find our ee details - so exit the routine.
            If m_strWWID = "" and m_strIDSID = "" then
                exit sub
            End if
            
            mDebugPrint "DBLoad() - m_strWWID = '" & m_strWWID & "'<br>"

            If m_strWWID <> "" then
                condition = """WWID"" = '" & m_strWWID & "' "  
            else
                condition = " ""Idsid"" = '" & m_strIDSID & "' "
            end if
			
            mDebugPrint "m_strWWID:" & m_strWWID & "<br>"
            mDebugPrint "m_strIDSID:" & m_strIDSID & "<br>"
            mDebugPrint "WWID:" & WWID & "<br>"
            mDebugPrint "IDSID:" & IDSID & "<br>"

            if RENEW_ALL_RECORDS_AND_CLEAN_DB then
                RenewDbCleanDb
            end if

            ' ignore ad_idsid accounts
            If Left(IDSID, 3) = "ad_" then
                Response.Write "<center><div style='margin-top:100px;margin-left:auto;margin-right:auto'><h1>Ignoring 'ad_' account.<br>Please use a normal domain account</h1></div></center>"
                mCloseApplication
                Response.End
                exit sub
            End if
			
            'create the connection
            Set locCmd2 = Server.CreateObject("ADODB.Command")
            Set locCmd2.ActiveConnection = glbConnection
            locCmd2.CommandType = adCmdStoredProc
				
            'get last update for this user in the local copy of wds
            'select proc name
            If m_strWWID <> "" then
                locCmd2.CommandText = "usp_getLastUpdateWdsCopy_from_wwid"
                Set locParam2 = locCmd2.CreateParameter("strWWID", adWChar, adParamInput, 8, WWID)
            elseif m_strIDSID <> "" then
                locCmd2.CommandText = "usp_getLastUpdateWdsCopy_from_shortid"
                Set locParam2 = locCmd2.CreateParameter("strIdsid", adWChar, adParamInput, 8, IDSID)
            end if
            'execute
            locCmd2.Parameters.Append locParam2
            Set locRS2 = locCmd2.Execute

            updateRequired = false
            insertRequired = false

            if isnull(locRS2) or locRS2.eof then ' no entry found
                insertRequired = true
            elseif ((locRS2("WdsCopyUpdateDate") = "") or isnull(locRS2("WdsCopyUpdateDate")) or (locRS2("WdsCopyUpdateDate")<Date())) then 'need to update if not updated today
                updateRequired = true
            end if

            locRS2.Close
            Set locRS2 = nothing
            Set locParam2 = nothing

            if insertRequired or updateRequired then
                updateOrInsert updateRequired,insertRequired,condition
            else
                mDebugPrint "No need to update WDS copy<br>"
            end if	
					 
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc

            If m_strWWID <> "" then
                locCmd.CommandText = "usp_eedetails"
                Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, WWID)
            else
                locCmd.CommandText = "usp_userdetails"
                Set locParam = locCmd.CreateParameter("strIdsid", adWChar, adParamInput, 8, idsid)
            end if       
            locCmd.Parameters.Append locParam
            Set locRS = locCmd.Execute
			
            If not locRS.eof then
                mDebugPrint "DBLoad() - Loading details...<br>"
                m_strWWID = locRS("WWID")
                m_strIDSID = locRS("Idsid")
                m_strFirstNm = locRS("FirstNm")
                m_strLastNm = locRS("LastNm")
                m_strNextLevelWWID = locRS("NextLevelWWID")
                m_strEmail = locRS("CorporateEmailTxt")
                m_strCompanyCd = locRS("CompanyCd")
                m_strMailStopTxt = locRS("MailStopTxt")
                m_datLDOH = locRS("StartDt")
                m_datODOH = locRS("OriginalStartDt")
                m_blnIsExempt = mIf(locRS("FLSACd")="E", True, False)
                m_blnIsBlueBadge = mIf(locRS("blnBlueBadge")=True, True, False)
                m_blnIsPartTimer = mIf(locRS("FullTmPartTmCd")="PT", True, False)
                m_datTerminationDate = locRS("EndDt")
                m_strActiveStatus = locRS("EmployeeStatusCd")
                
                if(isNull(locRS("blnIsEELeaveTracked"))) then 
                m_blnIsEELeaveTracked = false		 	  
            else
                m_blnIsEELeaveTracked = locRS("blnIsEELeaveTracked")
            end if 

            m_datDOB = locRS("datDOB")
            m_endDate = locRS("endDate") 
            m_orgUnit = locRS("DepartmentNm") 
            m_blnIsAdmin = mIf(locRS("blnIsAdmin")=True, True, False)
            m_blnIsExemptStatusChanged = mIf(locRS("blnIsExemptStatusChanged")=True, True, False)
            m_blnIsException = mIf(locRS("blnIsException")=True, True, False)
            m_strExceptionComments = locRS("strExceptionComments")
            m_strActiveDelegateWWID = locRS("strDelegateWWID")

            mDebugPrint "DBLoad() - Loaded m_strWWID='" & m_strWWID & "'<br>"
            if isnull(m_blnIsEELeaveTracked) then
                mDebugPrint "DBLoad() - m_blnIsEELeaveTracked = NULL<br>"
            else
                mDebugPrint "DBLoad() - m_blnIsEELeaveTracked = '" & m_blnIsEELeaveTracked & "'<br>"
            end if
				
            if locRS("LocalWWID") = "" or isnull(locRS("LocalWWID")) then
                LocalRecordExists = False
            Else
                LocalRecordExists = True
            End If

            Else
                mDebugPrint "DBLoad() - Details empty...<br>"
                SetEmpty
            End If
            
            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
        End Sub
        
        
        Private Sub DBLoadDelegateForManagers()
            Dim locCmd
            Dim locParam
            Dim locRS
            Dim fldWWID
            Dim fldIDSID
            Dim fldFName
            Dim fldLName
            Dim fldCorporateEmailTxt
            Dim flddatDOB
            Dim locobjUser
            Dim locCmdRevoke
            
            If m_blnDelegateForManagersLoaded then
                Exit Sub
            End If
            

            
            '*** If we have no WWID set up, we can't find our managers - so exit the routine.
            If WWID = "" then
                exit sub
            End if


            m_blnDelegateForManagersLoaded = True
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_eedelegateformanagers"
            locCmd.CommandType = adCmdStoredProc
                        
            Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, WWID)
            locCmd.Parameters.Append locParam
            
            Set locRS = locCmd.Execute

            Set fldWWID = locRS("WWID")
            Set fldIDSID = locRS("Idsid")
            Set fldFName = locRS("FirstNm")
            Set fldLName = locRS("LastNm")
            Set fldCorporateEmailTxt = locRS("CorporateEmailTxt")
            Set flddatDOB = locRS("datDOB")

            If typename(m_colDelegateForManagers) <> "cObjCollection" then
                Set m_colDelegateForManagers = new cObjCollection
            End If
          

            
            While not locRS.eof
                Set locobjUser = new cObjUser
                locobjUser.WWID = fldWWID
                locobjUser.IDSID = fldIDSID
                locobjUser.FirstNm = fldFName
                locobjUser.LastNm = fldLName
                locobjUser.Email = fldCorporateEmailTxt
                locobjUser.DOB = flddatDOB
                m_colDelegateForManagers.Add locobjUser
                Set locobjUser = nothing
                locRS.movenext
            Wend

            Set fldWWID = nothing
            Set fldIDSID = nothing
            Set fldFName = nothing
            Set fldLName = nothing
            Set fldCorporateEmailTxt = nothing
            Set flddatDOB = nothing

            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
        End Sub


        Private Sub DBLoadAnnualVacation()
            If m_blnAnnualVacationLoaded then
                Exit Sub
            End If

            m_blnAnnualVacationLoaded = True

            Set m_objAnnualVacation = new cObjAnnualVacation
            Set m_objAnnualVacation.EE = Me
            m_objAnnualVacation.EE.YearToView = Me.YearToView
            
            
        End Sub


        Private Sub DBLoadOtherLeave()
            If m_blnOtherLeaveLoaded then
                Exit Sub
            End If

            m_blnOtherLeaveLoaded = True

            Set m_objOtherLeave = new cObjOtherLeave
            Set m_objOtherLeave.EE = Me
        End Sub


        Private Sub DBLoadCarryOvers()
            Dim locCmd
            Dim locParam
            Dim locRS
            Dim locCmdRevoke
            Dim fldlngYear
            Dim fldlngID
            Dim fldlngDays
            Dim fldstrEnteredByWWID
            Dim flddatEntered
            Dim fldstrComments
            Dim fldblnIsPreArranged
            Dim locobjCarryOver

            If m_blnCarryOversLoaded then
                Exit Sub
            End If
            
            '*** If we have no WWID set up, we can't find our CarryOvers - so exit the routine.
            If WWID = "" then
                exit sub
            End if

            m_blnCarryOversLoaded = True
            
            Set locCmd = Server.CreateObject("ADODB.Command")
            
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandText = "usp_carryovers_by_WWID"
            locCmd.CommandType = adCmdStoredProc
                        
            Set locParam = locCmd.CreateParameter("strWWID", adWChar, adParamInput, 8, WWID)
            locCmd.Parameters.Append locParam
            
            Set locRS = locCmd.Execute

            Set fldlngYear = locRS("lngYear")
            Set fldlngID = locRS("lngID")
            Set fldlngDays = locRS("lngDays")
            Set fldstrEnteredByWWID = locRS("strEnteredByWWID")
            Set flddatEntered = locRS("datEntered")
            Set fldstrComments = locRS("strComments")
            Set fldblnIsPreArranged = locRS("blnPreArranged")

            While not locRS.eof
                Set locobjCarryOver = new cObjCarryOver
                Set locobjCarryOver.EE = Me
                locobjCarryOver.Year = fldlngYear
                locobjCarryOver.ID = fldlngID
                locobjCarryOver.Days = fldlngDays
                locobjCarryOver.EnteredBy.WWID = fldstrEnteredByWWID
                locobjCarryOver.DateEntered = flddatEntered
                locobjCarryOver.Comments = fldstrComments
                If fldblnIsPreArranged = True then
                    CarryOversPreArranged.Add locobjCarryOver
                Else
                    CarryOversEOY.Add locobjCarryOver
                End If
                Set locobjCarryOver = nothing
                locRS.movenext
            Wend

            Set fldlngYear = nothing
            Set fldlngID = nothing
            Set fldlngDays = nothing
            Set fldstrEnteredByWWID = nothing
            Set flddatEntered = nothing
            Set fldstrComments = nothing
            Set fldblnIsPreArranged = nothing

            locRS.Close
            Set locRS = nothing
            Set locParam = nothing
            Set locCmd = nothing
            
        End Sub

        
        Public Function SetToLoggedOnUser()
            Dim strRequestLogonUser
            Dim lngOffsetDelimiterPos

            SetEmpty
            WWID = ""
                        
			strRequestLogonUser = Trim(Request.ServerVariables("LOGON_USER"))
			'strRequestLogonUser = "dpknowle"

			' Redirect user to "Under Construction" if user is not Developer
			If CONST_UNDER_CONSTRUCTION and CONST_DEVELOPER_ALIAS <> strRequestLogonUser Then
			    Response.Redirect CONST_APPLICATION_PATH & "/underconstruction.asp"
			End If
				
			' Testing: MOF - pretending to be another user
			If CONST_TEST_MODE and CONST_DEVELOPER_ALIAS = strRequestLogonUser then           		
				Dim LoginAs 
				LoginAs = Request.Form("logon_as")
                mDebugPrint "LoginAs: " & LoginAs
				
				Dim SaveUserCmd
				Dim LoadUserRS
				Dim LoadUserRow
				
				' [MOF 11/21/08] - Load the current user from the DB - dev
				if(""<>LoginAs) then 
					' Save this user in the "development" table (temporary) as a user					
					Set SaveUserCmd = Server.CreateObject("ADODB.Command")
					Set SaveUserCmd.ActiveConnection = glbConnection
					
					SaveUserCmd.CommandText = "UPDATE development SET username='" & LoginAs & "'"
					SaveUserCmd.Execute()
				else 
					' Otherwise load the user from the database
					Set LoadUserRS = Server.CreateObject("ADODB.RecordSet")
					LoadUserRS.Open "SELECT * FROM development", glbConnection
					
					' Extract the user
					LoadUserRow = LoadUserRS.GetRows(1, 0)
					LoginAs = LoadUserRow(0, 0)
					
					LoadUserRS.Close()
				end if
			
				' Swap the logged on user with the user stored in the DB		
				strRequestLogonUser = LoginAs 
			End If		
			' End Test Mode						

		
            'Response.Write "SetToLoggedOnUser.strRequestLogonUser " & strRequestLogonUser & "<BR>"   'JG
            '*** CHECK THAT WE HAVE SUCCESSFULLY RETRIEVED THE USER'S NT LOGON ID. ***
            
            if strRequestLogonUser = "" then
                SetToLoggedOnUser = CONST_LOGON_ERROR_USER_BLANK
            else
			
                '*** REMOVE THE DOMAIN NAME ("GER\", etc.) FROM THE NT LOGON ID ***
                lngOffsetDelimiterPos = InStr(1, strRequestLogonUser, "\")
                if lngOffsetDelimiterPos = 0 then
                    lngOffsetDelimiterPos = InStr(1, strRequestLogonUser, "/")
                    'Response.Write "SetToLoggedOnUser.lngOffsetDelimiterPos: " & lngOffsetDelimiterPos & "<BR>"   'JG
                    
                else
                    strRequestLogonUser = lcase(trim(mid(strRequestLogonUser, lngOffsetDelimiterPos + 1)))
                    'Response.Write "SetToLoggedOnUser.strRequestLogonUser: " & strRequestLogonUser & "<BR>"   'JG
                end if

                m_strIDSID = strRequestLogonUser	
				
				'dim endtime
				'endtime = timer()
				'dim benchmark
				'benchmark = endtime - starttime
				'Response.Write( benchmark ) 
				'Response.Write( " - SetToLoggedOnUser 1" ) 
				'Response.Write( "<br>" ) 	
							
                '*** CHECK IDSID TO FORCE DBLoad AND THEN CHECK IF FOUND ***
                
                If len(IDSID) = 0 then 
                    SetToLoggedOnUser = CONST_LOGON_ERROR_USER_NOT_FOUND
                    'Response.Write "SetToLoggedOnUser.SetToLoggedOnUser1: " & SetToLoggedOnUser & "<BR>"   'JG
					'endtime = timer()
					'benchmark = endtime - starttime
					'Response.Write( benchmark ) 
					'Response.Write( " - SetToLoggedOnUser 2" ) 
					'Response.Write( "<br>" ) 	   
                '*** IS THE CURRENT USER A VALID USER? ***               
                
                ElseIf Not IsValidUser then
                    'SetToLoggedOnUser = CONST_LOGON_ERROR_USER_ACCESS_DENIED
                    'Response.Write "SetToLoggedOnUser.SetToLoggedOnUser2: " & SetToLoggedOnUser & "<BR>"   'JG		
					'set endtime = timer()
					'benchmark = endtime - starttime
					'Response.Write( benchmark ) 
					'Response.Write( " - SetToLoggedOnUser 3" ) 
					'Response.Write( "<br>" ) 	
				
                '*** CHECK IF A LOCAL RECORD NEEDS TO BE SET UP
                ElseIf IsEELeaveTracked AND (NOT LocalRecordExists) then
                    SetToLoggedOnUser = CONST_LOGON_ERROR_USER_SET_UP_REQUIRED
                    'Response.Write "SetToLoggedOnUser.SetToLoggedOnUser3: " & SetToLoggedOnUser & "<BR>"   'JG
						
					'endtime = timer()
					'benchmark = endtime - starttime
					'Response.Write( benchmark ) 
					'Response.Write( " - SetToLoggedOnUser 4" ) 
					'Response.Write( "<br>" ) 	
				
                '*** NO PROBLEMS ***
                Else
                    SetToLoggedOnUser = CONST_LOGON_SUCCESSFUL
                    'Response.Write "SetToLoggedOnUser.SetToLoggedOnUser4: " & SetToLoggedOnUser & "<BR>"   'JG	

					'endtime = timer()
					'benchmark = endtime - starttime
					'Response.Write( benchmark ) 
					'Response.Write( " - SetToLoggedOnUser 5" ) 
					'Response.Write( "<br>" ) 	
                End If

            End if

        End Function
    
    
        '**** REFRESH APPROVALS *****
        Public Sub RefreshApprovals
            Set m_colApprovals = nothing
            InitialiseApprovals
        End Sub


        Private Sub InitialiseApprovals
            If typename(m_colApprovals) <> "cColLeavePeriods" then
                mDebugPrint "   Creating cColLeavePeriods - 5 <br>"
                Set m_colApprovals = new cColLeavePeriods
                m_colApprovals.CollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_APPROVALS
                Set m_colApprovals.EE = Me
				
				'dim endtime
				'endtime = timer()
				'dim benchmark
				'benchmark = endtime - starttime
				'Response.Write( benchmark ) 
				'Response.Write( " - InitialiseApprovals" ) 
				'Response.Write( "<br>" ) 
				
            End If
        End Sub
                

        Private Sub InitialiseLeaveRequests
            If typename(m_colLeaveRequests) <> "cColLeavePeriods" then
                mDebugPrint "   Creating cColLeavePeriods - 4 <br>"
                Set m_colLeaveRequests = new cColLeavePeriods
                m_colLeaveRequests.CollectionType = CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_REQUESTS
                Set m_colLeaveRequests.EE = Me
            End If
        End Sub
        

        '**** LOAD DELEGATE FROM FORM ****
        Public Sub LoadDelegateFromForm
            m_strActiveDelegateWWID = request.form("fldstrDelegateWWID")
        End Sub


        '**** REMOVE DELEGATE ****
        Public Function RemoveDelegate()
            Dim loclngSaveResult
            
            m_strActiveDelegateWWID = ""
            Set m_objActiveDelegate = nothing
            
            loclngSaveResult = Save()
            
            RemoveDelegate = loclngSaveResult
        End Function
        

        '**** Save ****
        Public Function Save
			mDebugPrint("In <b>cObjUser.Save()</b><br>")
            Dim loclngResult
            Dim locCmd
            Dim loclngReturnValue
            Dim locCmdRevoke

            Set locCmd = Server.CreateObject("ADODB.Command")
            Set locCmd.ActiveConnection = glbConnection
            locCmd.CommandType = adCmdStoredProc
            loclngReturnValue = 0

            '*** If no local record exists then create one.
            If not LocalRecordExists then
                loclngReturnValue = CreateLocalUserRecord()
                If loclngReturnValue = 0 then
                    Set locCmd = nothing
                    Save = 0
                    Exit Function
                End If
            End If
            
            locCmd.CommandText = "usp_save_user"
            locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
            locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, trim(WWID))
            locCmd.Parameters.Append locCmd.CreateParameter("strDelegateWWID", adWChar, adParamInput, 8, trim(m_strActiveDelegateWWID))
            locCmd.Parameters.Append locCmd.CreateParameter("datDOB", adDBTimeStamp, adParamInput, , mIf(not isdate(DOB),null,DOB))
            locCmd.Parameters.Append locCmd.CreateParameter("endDate", adDBTimeStamp, adParamInput, , mIf(not isdate(EndDate),null,EndDate))   
            locCmd.Parameters.Append locCmd.CreateParameter("blnIsEELeaveTracked", adBoolean, adParamInput, , IsEELeaveTracked)
            locCmd.Parameters.Append locCmd.CreateParameter("blnIsAdmin", adBoolean, adParamInput, , IsAdmin)

            on error resume next
            
            locCmd.Execute
            
            loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
            
            on error goto 0

            Set locCmd = nothing

            Save = loclngReturnValue
                        
        End Function


        '*** CREATE LOCAL USER RECORD ****
        Public Function CreateLocalUserRecord()
			mDebugPrint("In <b>cObjUser.CreateLocalUserRecord()</b><br>")
            Dim locCmd
            Dim loclngReturnValue
            Dim locCmdRevoke

            'Return values:
            '           2 Already Created.
            '           1 Successfully created
            '           0 Error Occurred
            
            If LocalRecordExists then
                loclngReturnValue = 0
				mDebugPrint "CreateLocalUserRecord() - Local Record Exists<br>"
            Else
                
                Set locCmd = Server.CreateObject("ADODB.Command")
                Set locCmd.ActiveConnection = glbConnection
                locCmd.CommandType = adCmdStoredProc
                loclngReturnValue = 0
    
                locCmd.CommandText = "usp_create_user"
                locCmd.Parameters.Append locCmd.CreateParameter("return", adInteger, adParamReturnValue)
                locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, trim(WWID))

				mDebugPrint "CreateLocalUserRecord() - Created Local Record<br> " 
                
                on error resume next

                locCmd.Execute
                
                loclngReturnValue = mGetSafeLongInteger(locCmd("return"),0)
                
                on error goto 0

            End If
            
            CreateLocalUserRecord = loclngReturnValue
        End Function


        '**** LOAD ADMIN USER PROFILE FROM FORM ****
        Public Function LoadAdminUserProfileFromForm()
			mDebugPrint "In LoadAdminUserProfileFromForm()<br>"
			DOB = Request("flddatDOB")
            endDate = Request("fldendDate") 
            
            If Request("fldblnIsAdmin") = "True" then
                IsAdmin = True
            Else
                IsAdmin = False
            End If
            If Request("fldblnIsEELeaveTracked") = "True" then
                IsEELeaveTracked = True
            Else
                IsEELeaveTracked = False
            End If
        End Function


        '*** INITIALISE DIRECT REPORTS ***
        Private Sub Initialise_DirectReports
            If typename(m_colDirectReports) <> "cColEmployees" then
                mDebugPrint "   Creating cColLeavePeriods - 6 <br>"
                Set m_colDirectReports = new cColEmployees
                m_colDirectReports.CollectionType = CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS
                Set m_colDirectReports.EE = Me
            End If
        End Sub


        '*** INITIALISE CARRYOVERS EOY ****
        Private Sub Initialise_CarryOversEOY
            If typename(m_colCarryOversEOY) <> "cObjCollection" then
                Set m_colCarryOversEOY = new cObjCollection
            End If
        End Sub

        
        '*** INITIALISE CARRYOVERS PRE ARRANGED ***
        Private Sub Initialise_CarryOversPreArranged
            If typename(m_colCarryOversPreArranged) <> "cObjCollection" then
                Set m_colCarryOversPreArranged = new cObjCollection
            End If
        End Sub

        
    End Class
    
%>
