Attribute VB_Name = "modSched"
# Option Explicit
# Private # Const LOG_ENABLED # As Boolean = False   ' set True for dev only

# --- logging (disabled for users) ------------------------------------------------
# Private def LogMsg( msg # As String):
    If not LOG_ENABLED : return
    On Error Resume # Next
    # Dim f # As Integer: f = FreeFile
    Open ThisWorkbook.path + "\GECE_debug.log" For Append # As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss"); " | "; msg
    Close #f
# End Sub

# --- EXE-safe Named Range resolver (workbook, then sheet scope) ------------------
# Public def NR( nm # As String, Optional  ws # As Worksheet): # As Range
    # Dim N # As Name
    On Error Resume # Next
    Set N = ThisWorkbook.Names(nm)
    If not N Is Nothing :
        If InStr(1, N.RefersTo, "#REF", vbTextCompare) = 0 : Set NR = N.RefersToRange: return
    # End If
    If not ws Is Nothing :
        Set N = ws.Names(nm)
        If not N Is Nothing :
            If InStr(1, N.RefersTo, "#REF", vbTextCompare) = 0 : Set NR = N.RefersToRange: return
        # End If
    # End If
    On Error GoTo 0
# End Function

# --- (DEV ONLY) wrapper with log+message; do not wire users to this --------------
# Public def Run_Gantt_Debug():
    On Error GoTo ErrH
    LogMsg "=== RUN START ==="
    If Gantt_by_Phases :
        LogMsg "SUCCESS | Gantt created"
        MsgBox "Gantt created.", vbInformation
    else:
        LogMsg "FAILED | Gantt_by_Phases returned False"
        MsgBox "Export did not complete. See GECE_debug.log.", vbExclamation
    # End If
    return
ErrH:
    LogMsg "ERR " + Err.Number + " at line " + Erl + " | " + Err.Description
    MsgBox "Err " + Err.Number + " at line " + Erl + ": " + Err.Description, vbCritical
# End Sub

# --- PRODUCTION entry (wire the user button to THIS function) --------------------
# Public def Gantt_by_Phases(): # As Boolean
10  On Error GoTo ErrH
    Gantt_by_Phases = False
    If CheckForProject() = 0 : return

    # Dim ws # As Worksheet: Set ws = ThisWorkbook.Worksheets("Schedule")
    LogMsg "Start Gantt_by_Phases"

# 0) Data present?
20  # Dim init_hours # As Double
    init_hours = ThisWorkbook.Worksheets(gstrGECECQAOutputSheet).Range("TOTAL_HOURS").Value
    If init_hours = 0 : return   ' silent for users
    LogMsg "TOTAL_HOURS=" + init_hours

# 1) GoalSeek hardened ? fallback solver if native fails (silent)
30  # Dim rFinal # As Range, rReq # As Range, rFactor # As Range
    Set rFinal = NR("SCHED_FINAL_DURATION", ws)
    Set rReq = NR("SCHED_REQUIRED_DURATION", ws)
    Set rFactor = NR("SCHED_FACTOR", ws)

    If not (rFinal Is Nothing or rReq Is Nothing or rFactor Is Nothing) :
        If rFinal.CountLarge = 1 and rReq.CountLarge = 1 and rFactor.CountLarge = 1 and IsNumeric(rReq.Value) :
            If not (rFactor.Locked and ws.ProtectContents) :
                On Error Resume # Next
                rFinal.GoalSeek CDbl(rReq.Value), rFactor     ' positional args (late-binding safe)
                If Err.Number <> 0 :
                    LogMsg "GoalSeek ERR " + Err.Number + " | " + Err.Description + " -> fallback"
                    Err.Clear
                    Call SolveFactorBinary(rFinal, rReq.Value, rFactor)
                # End If
                On Error GoTo ErrH
            else:
                LogMsg "GoalSeek skipped (factor locked, protected sheet)"
            # End If
        else:
            LogMsg "GoalSeek skipped (inputs invalid)"
        # End If
    else:
        LogMsg "GoalSeek skipped (names missing)"
    # End If
    If not rFactor Is Nothing : If val(rFactor.Value) < 0 : return

# 2) Inputs
40  # Dim rTasks # As Range, rDur # As Range, rPred # As Range, rRes # As Range
    Set rTasks = NR("SCHED_TASK", ws)
    Set rDur = NR("SCHED_DURATION", ws)
    Set rPred = NR("SCHED_PREDECESSORS", ws)
    Set rRes = NR("SCHED_RESOURCE_NAMES", ws)
    If rTasks Is Nothing or rDur Is Nothing or rPred Is Nothing : return

60  # Dim taskArr # As Variant, durArr # As Variant, predArr # As Variant, resArr # As Variant
    taskArr = rTasks.Value: durArr = rDur.Value: predArr = rPred.Value
    If not rRes Is Nothing : resArr = rRes.Value
    If not IsArray(taskArr) : return
    # Dim nRows # As Long: nRows = UBound(taskArr, 1)

# 3) Project/per-task starts
70  # Dim projStart # As Variant
    # Dim rStart # As Range: Set rStart = NR("SCHED_START_DATE", ws)
    projStart = IIf(rStart Is Nothing or IsEmpty(rStart.Value), Date, rStart.Value)

80  # Dim taskStarts() # As Variant, i # As Long
    ReDim taskStarts(1 To nRows, 1 To 1)
    taskStarts(1, 1) = projStart
    for i in range(int(2), int(nRows) + 1):
        # Dim nm # As String: nm = "Start_Date_Task" + CStr(i)   ' legacy names on Schedule
        On Error Resume # Next
        taskStarts(i, 1) = ws.Range(nm).Value
        On Error GoTo ErrH
    # Next i

# 4) MS Project
90  # Dim pj # As Object, proj # As Object
    On Error Resume # Next
    Set pj = GetObject(, "MSProject.Application")
    If pj Is Nothing : Set pj = CreateObject("MSProject.Application")
    On Error GoTo ErrH
    If pj Is Nothing : return

    pj.Visible = True
    pj.DisplayAlerts = False
    Set proj = pj.Projects.Add
    If proj Is Nothing : GoTo CleanUp
    If IsDate(projStart) : proj.ProjectStart = CDate(projStart)

# set default to Auto if supported (no named args; ignore failures)
    On Error Resume # Next
    pj.Application.NewTasksAreManual = False
    pj.Application.NewTasksAreManuallyScheduled = False
    On Error GoTo ErrH

# 5) PASS 1  tasks, duration (hours?minutes), resources, starts (force Auto)
100 # Dim minutesPerHour # As Long: minutesPerHour = 60
    # Dim t # As Object, taskName # As String, resStr # As String, durH # As Double
    # Dim uidMap() # As Long: ReDim uidMap(1 To nRows)

    for i in range(int(1), int(nRows) + 1):
        taskName = Trim$(CStr(taskArr(i, 1)))
        If Len(taskName) > 0 :
            Set t = proj.tasks.Add(taskName)
            If not t Is Nothing :
                On Error Resume # Next
                t.TaskMode = 0      ' 0 = Auto (newer versions)
                t.Manual = False    ' older versions
                On Error GoTo ErrH

                uidMap(i) = t.uniqueID
                durH = val(durArr(i, 1))
                If durH > 0 : t.Duration = CLng(durH * minutesPerHour)
                If IsArray(resArr) :
                    resStr = Trim$(CStr(resArr(i, 1)))
                    If Len(resStr) > 0 : t.ResourceNames = resStr
                # End If
                If IsDate(taskStarts(i, 1)) : t.Start = CDate(taskStarts(i, 1))
            # End If
        # End If
    # Next i

# 6) PASS 2  predecessors via UniqueID (drop self/invalid)
110 # Dim predStr # As String, norm # As String
    for i in range(int(1), int(nRows) + 1):
        If uidMap(i) <> 0 :
            predStr = Trim$(CStr(predArr(i, 1)))
            If Len(predStr) > 0 :
                norm = NormalizePredUID(predStr, i, uidMap)
                If Len(norm) > 0 :
                    Set t = TaskByUID(proj, uidMap(i))
                    If not t Is Nothing :
                        On Error Resume # Next
                        t.UniqueIDPredecessors = norm
                        Err.Clear
                        On Error GoTo ErrH
                    # End If
                # End If
            # End If
        # End If
    # Next i

# 7) final sweep: ensure Auto & remove ? (Estimated)
115 # Dim tt # As Object
    On Error Resume # Next
    for tt in proj.tasks:
        If not tt Is Nothing :
            tt.TaskMode = 0
            tt.Manual = False
            tt.Estimated = False
        # End If
    # Next tt
    On Error GoTo ErrH

# 8) basic cosmetics (safe positional args)
120 On Error Resume # Next
    pj.Application.ViewApply "Gantt Chart"
    pj.Application.TableApply "Entry"
    pj.Application.TimescaleEdit 0, 2, 0, 10, True, True, True, True, 2
    On Error GoTo ErrH

# 9) save silently beside workbook
130 # Dim savePath # As String
    savePath = ThisWorkbook.path + "\" + Format(Date, "yyyymmdd") + "_Schedule_ProjectName_Rev01.mpp"
    On Error Resume # Next
    proj.SaveAs savePath
    pj.DisplayAlerts = True
    On Error GoTo ErrH

    Gantt_by_Phases = True
CleanUp:
    return

ErrH:
    LogMsg "ERR " + Err.Number + " @ line " + Erl + " | " + Err.Description
    Resume CleanUp
# End Function

# --- Helpers ----------------------------------------------------------------
# Private def SolveFactorBinary( rFinal # As Range,  Target # As Double,  rFactor # As Range): # As Boolean
    # Dim lo # As Double, hi # As Double, mid # As Double, f # As Double, i # As Long
    # Dim cur # As Double: cur = val(rFactor.Value)
    lo = IIf(cur > 0, 0, cur - 10): hi = IIf(cur > 0, cur * 2 + 10, 10)
    If hi <= lo : hi = lo + 10

    rFactor.Value = lo: Application.Calculate
    # Dim fLo # As Double: fLo = val(rFinal.Value)
    rFactor.Value = hi: Application.Calculate
    # Dim fHi # As Double: fHi = val(rFinal.Value)
    If Abs(fLo - fHi) < 0.0000001 : SolveFactorBinary = False: return

    for i in range(int(1), int(40) + 1):
        mid = (lo + hi) / 2
        rFactor.Value = mid: Application.Calculate
        f = val(rFinal.Value)
        If (f < Target) = (fLo < fHi) :
            lo = mid: fLo = f
        else:
            hi = mid: fHi = f
        # End If
        If Abs(f - Target) <= 0.0001 : Exit For
    # Next i
    SolveFactorBinary = True
# End Function

# Private def NormalizePredUID( s # As String,  curRow # As Long,  uidMap() # As Long): # As String
    # Dim parts() # As String, p # As String, out # As String, num # As Long, suffix # As String, mapped # As Long, i # As Long
    s = Trim$(Replace(s, ";", ",")): If Len(s) = 0 : return
    parts = Split(s, ",")
    for i in range(int(LBound(parts)), int(UBound(parts)) + 1):
        p = Trim$(parts(i)): If Len(p) = 0 : GoTo nxt
        num = val(p): suffix = mid$(p, Len(CStr(num)) + 1)
        If num > 0 and num <> curRow and num <= UBound(uidMap) :
            mapped = uidMap(num)
            If mapped > 0 :
                If Len(out) > 0 : out = out + ","
                out = out + CStr(mapped) + suffix
            # End If
        # End If
nxt: # Next i
    NormalizePredUID = out
# End Function

# Private def TaskByUID( proj # As Object,  uid # As Long): # As Object
    # Dim t # As Object
    for t in proj.tasks:
        If not t Is Nothing : If t.uniqueID = uid : Set TaskByUID = t: return
    # Next t
# End Function

# --- MS Project presence check ----------------------------------------------
# Public def CheckForProject(): # As Integer
    On Error GoTo ErrH
    # Dim p # As Object
    On Error Resume # Next
    Set p = GetObject(, "MSProject.Application")
    If p Is Nothing : Set p = CreateObject("MSProject.Application")
    On Error GoTo ErrH
    CheckForProject = IIf(p Is Nothing, 0, 1)
CleanUp:
    On Error Resume # Next
    Set p = Nothing
    return
ErrH:
    CheckForProject = 0
    Resume CleanUp
# End Function


# ===== (Your existing function, kept; already guarded & DoneEx-safe) =====
# Public def CashInflowOutflow(): # As Boolean
    On Error GoTo ErrHandler
    CashInflowOutflow = False

    # Dim currentDate # As Date, proposalDate # As Date, projectstartDate # As Date, projectfinishDate # As Date
    # Dim proposalDate_tmp # As Date
    # Dim startDate # As Date, finishDate # As Date
    # Dim ii # As Long, jj # As Long, kk # As Long, escalation_increment # As Long
    # Dim WritecellRow # As Long, WritecellColumn # As Long
    # Dim HourcellRow # As Long, HourcellColumn # As Long
    # Dim StartDatecellRow # As Long, StartDatecellColumn # As Long
    # Dim FinishDatecellRow # As Long, FinishDatecellColumn # As Long
    # Dim RessourcecellRow # As Long, RessourcecellColumn # As Long
    # Dim RemotePCTcellRow # As Long, RemotePCTcellColumn # As Long
    # Dim DurationcellRow # As Long, DurationcellColumn # As Long

    # Dim oldCalc # As XlCalculation, oldEv # As Boolean, oldUpd # As Boolean
    oldCalc = Application.Calculation
    oldEv = Application.EnableEvents
    oldUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If gstrGECEWorkBook = "" : getWorkBookName

    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEScheduleSheet)
        WritecellRow = .Range("SCHED_LOCAL_START").row
        WritecellColumn = .Range("SCHED_LOCAL_START").Column
        HourcellRow = .Range("SCHED_HOUR_START").row
        HourcellColumn = .Range("SCHED_HOUR_START").Column
        StartDatecellRow = .Range("SCHED_START_DATE").row
        StartDatecellColumn = .Range("SCHED_START_DATE").Column
        FinishDatecellRow = .Range("SCHED_FINISH_DATE_START").row
        FinishDatecellColumn = .Range("SCHED_FINISH_DATE_START").Column
        RessourcecellRow = .Range("SCHED_RESSOURCES_START").row
        RessourcecellColumn = .Range("SCHED_RESSOURCES_START").Column
        RemotePCTcellRow = .Range("SCHED_REM_PCT_START").row
        RemotePCTcellColumn = .Range("SCHED_REM_PCT_START").Column
        DurationcellRow = .Range("SCHED_DURATION_START").row
        DurationcellColumn = .Range("SCHED_DURATION_START").Column

        .Range("I320:CZ320").ClearContents
        .Range("I323:CZ329").ClearContents
        .Range("I333:CZ339").ClearContents

        projectstartDate = .Cells(StartDatecellRow + 2, StartDatecellColumn).Value
        projectfinishDate = .Cells(FinishDatecellRow + 2, FinishDatecellColumn).Value

        proposalDate_tmp = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("PROPSAL_DATE").Value
        If proposalDate_tmp = "00:00:00" :
            proposalDate = Now()
        else:
            proposalDate = proposalDate_tmp
        # End If

        If Month(proposalDate) - 4 > 0 and Month(projectstartDate) - 4 < 0 :
            escalation_increment = escalation_increment - 1
        ElseIf Month(proposalDate) - 4 < 0 and Month(projectstartDate) - 4 > 0 :
            escalation_increment = escalation_increment + 1
        # End If
        escalation_increment = escalation_increment + Year(projectstartDate) - Year(proposalDate)

        for kk in range(int(0), int(95) + 1):
            currentDate = .Cells(.Range("SCHED_CALENDAR_START").row, .Range("SCHED_CALENDAR_START").Column + kk).Value

            If Month(currentDate) = 4 :
                escalation_increment = escalation_increment + 1
            # End If
            .Cells(WritecellRow - 3, WritecellColumn + kk).Value = escalation_increment

            for ii in range(int(0), int(50) + 1):
                If .Cells(HourcellRow + ii, HourcellColumn).Value <> 0 :
                    startDate = .Cells(StartDatecellRow + ii, StartDatecellColumn).Value
                    finishDate = .Cells(FinishDatecellRow + ii, FinishDatecellColumn).Value

                    If (Month(currentDate) - Month(startDate) + (Year(currentDate) - Year(startDate)) * 12 >= 0) and _
                       (Month(finishDate) - Month(currentDate) + (Year(finishDate) - Year(currentDate)) * 12 >= 0) :

                        for jj in range(int(0), int(6) + 1):
                            If .Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value <> 0 :
                                .Cells(WritecellRow + jj, WritecellColumn + kk).Value = _
                                    .Cells(WritecellRow + jj, WritecellColumn + kk).Value + _
                                    (.Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value * (1 - .Cells(RemotePCTcellRow + ii, RemotePCTcellColumn).Value))
                            # End If
                        # Next jj

                        for jj in range(int(0), int(6) + 1):
                            If .Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value <> 0 :
                                .Cells(WritecellRow + 10 + jj, WritecellColumn + kk).Value = _
                                    .Cells(WritecellRow + 10 + jj, WritecellColumn + kk).Value + _
                                    (.Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value * .Cells(RemotePCTcellRow + ii, RemotePCTcellColumn).Value)
                            # End If
                        # Next jj
                    # End If
                # End If
            # Next ii
        # Next kk
    End With

    CashInflowOutflow = True
CleanUp:
    With Application
        .Calculation = oldCalc
        .EnableEvents = oldEv
        .ScreenUpdating = oldUpd
    End With
    return

ErrHandler:
    MsgBox Err.Number + "; " + Err.Description + "; " + Err.Source, vbCritical
    Resume CleanUp
# End Function

# ===== Stubs (unchanged) =====
# Public def ShowRessourcePlanningForm():
# End Function

# Public def ShowForm():
    ReplaceCellName.Show vbModal
# End Function
