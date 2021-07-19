Attribute VB_Name = "modMain"
' Transym OCR Demonstration program
'
' THE SOFTWARE IS PROVIDED "AS-IS" AND WITHOUT WARRANTY OF ANY KIND,
' EXPRESS, IMPLIED OR OTHERWISE, INCLUDING WITHOUT LIMITATION, ANY
' WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'
' This program demonstrates calling TOCR version 5.1 from VB6
'
' Copyright (C) 2021 Transym Computer Services Ltd.
'
'
' TOCR5.1DemoVB6 Issue1

Option Explicit
 
 Const SAMPLE_TIFF_FILE = "Sample.tif"
 Const SAMPLE_BMP_FILE = "Sample.bmp"
 Const SAMPLE_PDF_FILE = "Sample.pdf"




'---------------------------------------------------------------------------
' Application start up module
'
Sub Main()


Call Example1   ' demonstrates how to OCR a single file
Call Example2   ' demonstrates how to OCR multiple files
Call Example3   ' demonstrates how to OCR a bitmap using a memory mapped file created by TOCR
Call Example4   ' demonstrates how to OCR a bitmap using a memory mapped file created here
Call Example5   ' retrieves information on job slot usage
Call Example6   ' retrieves information on job slots
Call Example7   ' gets images from a TWAIN compatible device
Call Example8   ' demonstrates TOCRSetConfig and TOCRGetConfig

End

End Sub ' Main

'---------------------------------------------------------------------------
' Demonstrates how to OCR a single file
'
Sub Example1()

Dim Status              As Long
Dim JobInfo_EG          As TOCRJOBINFO_EG
Dim JobNo               As Long
Dim Results             As TOCRRESULTS_EG
Dim Answer              As String

TOCRSetConfig TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX

JobInfo_EG.InputFile = SAMPLE_TIFF_FILE
JobInfo_EG.JobType = TOCRJOBTYPE_TIFFFILE

'or

'JobInfo_EG.InputFile = SAMPLE_BMP_FILE
'JobInfo_EG.JobType = TOCRJOBTYPE_DIBFILE

'or

'JobInfo_EG.InputFile = SAMPLE_PDF_FILE
'JobInfo_EG.JobType = TOCRJOBTYPE_PDFFILE

Status = TOCRInitialise(JobNo)
If Status = TOCR_OK Then
    If OCRWait(JobNo, JobInfo_EG) Then
        If GetResults(JobNo, Results) Then
            If FormatResults(Results, Answer) Then
                MessageBox 0, StrPtr(Answer), StrPtr("Example 1"), vbOKOnly
            End If
        End If
    End If
    TOCRShutdown JobNo
End If

End Sub ' Example1

'---------------------------------------------------------------------------
' Demonstrates how to OCR multiple files
'
Sub Example2()

Dim Status              As Long
Dim JobNo               As Long
Dim JobInfo_EG          As TOCRJOBINFO_EG
Dim Results             As TOCRRESULTS_EG
Dim CountDone           As Long

TOCRSetConfig TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX

Status = TOCRInitialise(JobNo)
If Status = TOCR_OK Then

    ' 1st file
    JobInfo_EG.InputFile = SAMPLE_TIFF_FILE
    JobInfo_EG.JobType = TOCRJOBTYPE_TIFFFILE
    If OCRWait(JobNo, JobInfo_EG) Then
        If GetResults(JobNo, Results) Then
            CountDone = CountDone + 1
        End If
    End If

    ' 2nd file
    JobInfo_EG.InputFile = SAMPLE_BMP_FILE
    JobInfo_EG.JobType = TOCRJOBTYPE_DIBFILE
    If OCRWait(JobNo, JobInfo_EG) Then
        If GetResults(JobNo, Results) Then
            CountDone = CountDone + 1
        End If
    End If

    ' 3rd file
    JobInfo_EG.InputFile = SAMPLE_PDF_FILE
    JobInfo_EG.JobType = TOCRJOBTYPE_PDFFILE
    If OCRWait(JobNo, JobInfo_EG) Then
        If GetResults(JobNo, Results) Then
            CountDone = CountDone + 1
        End If
    End If
    TOCRShutdown (JobNo)
End If

MsgBox CStr(CountDone) & " of 3 jobs done", vbInformation, "Example 2"

End Sub ' Example2

'---------------------------------------------------------------------------
' Demonstrate how to OCR a bitmap using a memory mapped file created by TOCR.
'
Sub Example3()

Dim Status              As Long
Dim JobInfo_EG          As TOCRJOBINFO_EG
Dim JobNo               As Long
Dim Results             As TOCRRESULTS_EG
Dim Answer              As String
Dim hFile               As Long

TOCRSetConfig TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX

Status = TOCRInitialise(JobNo)
If Status = TOCR_OK Then
    If TOCRConvertFormat(JobNo, ByVal SAMPLE_TIFF_FILE, TOCRCONVERTFORMAT_TIFFFILE, hFile, TOCRCONVERTFORMAT_MMFILEHANDLE, 0) = TOCR_OK Then
        
        JobInfo_EG.JobType = TOCRJOBTYPE_MMFILEHANDLE
        
        JobInfo_EG.hMMF = hFile
                
        If OCRWait(JobNo, JobInfo_EG) Then
           If GetResults(JobNo, Results) Then
               If FormatResults(Results, Answer) Then
                   MessageBox 0, StrPtr(Answer), StrPtr("Example 3"), vbOKOnly
               End If
           End If
        End If
        
        CloseHandle (hFile)
    End If
    
    TOCRShutdown JobNo
End If

End Sub ' Example3

'---------------------------------------------------------------------------
' Demonstrate how to OCR a bitmap using a memory mapped file created here.
' In this example the bitmap is simply read from a file but it could easily be
' one in memory that you have processed in some way.
'
Sub Example4()

Dim Status              As Long
Dim JobInfo_EG          As TOCRJOBINFO_EG
Dim JobNo               As Long
Dim Results             As TOCRRESULTS_EG
Dim Answer              As String
Dim BI                  As BMPINFO
Dim hFile               As Long

TOCRSetConfig TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX

Status = TOCRInitialise(JobNo)
If Status = TOCR_OK Then

    ' Get a bitmap into memory
    
    If GetBitmap(BI, SAMPLE_BMP_FILE) = BITMAP_OK Then
        If SaveMonoBitmapToMMFile(BI, hFile) Then
        
            JobInfo_EG.JobType = TOCRJOBTYPE_MMFILEHANDLE
            
            JobInfo_EG.hMMF = hFile
            
            If OCRWait(JobNo, JobInfo_EG) Then
                If GetResults(JobNo, Results) Then
                    If FormatResults(Results, Answer) Then
                        MessageBox 0, StrPtr(Answer), StrPtr("Example 4"), vbOKOnly
                    End If
                End If
            End If
            
            CloseHandle hFile
        End If
        If BI.hBmp Then DeleteObject BI.hBmp
    End If
    
    TOCRShutdown JobNo
End If

End Sub ' Example4

'---------------------------------------------------------------------------
' Retrieve information on job slot usage.
'
Sub Example5()

Dim NumSlots            As Long
Dim SlotUse()           As Long
Dim Msg                 As String
Dim SlotNo              As Long

TOCRSetConfig TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX

' uncomment to see effect on job slot use
'Dim JobNo               As Long
'If Not TOCRInitialise(JobNo) = TOCR_OK Then End

NumSlots = TOCRGetJobDBInfo(ByVal 0&)
If NumSlots > 0 Then
    ReDim SlotUse(0 To NumSlots - 1) As Long
    If TOCRGetJobDBInfo(SlotUse(0)) = TOCR_OK Then
        Msg = "Slot usage is:" & vbCrLf
        For SlotNo = 0 To NumSlots - 1
            Msg = Msg & vbCrLf & "Slot" & Str$(SlotNo) & " is "
            Select Case SlotUse(SlotNo)
                Case TOCRJOBSLOT_FREE
                    Msg = Msg & "free"
                Case TOCRJOBSLOT_OWNEDBYYOU
                    Msg = Msg & "owned by you"
                Case TOCRJOBSLOT_BLOCKEDBYYOU
                    Msg = Msg & "blocked by you"
                Case TOCRJOBSLOT_OWNEDBYOTHER
                    Msg = Msg & "owned by another process"
                Case TOCRJOBSLOT_BLOCKEDBYOTHER
                    Msg = Msg & "blocked by another process"
            End Select
        Next SlotNo
        MsgBox Msg, vbInformation, "Example 5"
    Else
        MsgBox "Failed to get job slot information", vbExclamation, "Example 5"
    End If
Else
    MsgBox "Failed to get number of job slots", vbExclamation, "Example 5"
End If

'TOCRShutdown JobNo

End Sub ' Example5

'---------------------------------------------------------------------------
' Retrieve information on job slots.
'
Sub Example6()

Dim NumSlots            As Long
Dim SlotUse()           As Long
Dim Msg                 As String
Dim SlotNo              As Long
Dim Volume              As Long
Dim Time                As Long
Dim Remaining           As Long
Dim Features            As Long
Dim Licence             As String

NumSlots = TOCRGetJobDBInfo(ByVal 0&)
If NumSlots > 0 Then
    Msg = "Slot usage is" & vbCrLf
    For SlotNo = 0 To NumSlots - 1
        Msg = Msg & vbCrLf & "Slot" & Str$(SlotNo)
        Licence = Space$(19)
        If TOCRGetLicenceInfoEx(SlotNo, Licence, Volume, Time, Remaining, Features) = TOCR_OK Then
            Msg = Msg & " " & Licence
            Select Case Features
                Case TOCRLICENCE_STANDARD
                    Msg = Msg & " STANDARD licence"
                Case TOCRLICENCE_EURO
                    If Licence = "5AD4-1D96-F632-8912" Then
                        Msg = Msg & " EURO TRIAL licence"
                    Else
                        Msg = Msg & " EURO licence"
                    End If
                Case TOCRLICENCE_EUROUPGRADE
                    Msg = Msg & " EURO UPGRADE licence"
                Case TOCRLICENCE_V3SE
                    If Licence = "2E72-2B35-643A-0851" Then
                        Msg = Msg & " V3 TRIAL licence"
                    Else
                        Msg = Msg & " V3 licence"
                    End If
                Case TOCRLICENCE_V3SEUPGRADE
                    Msg = Msg & " V1/2 UPGRADE to V3 SE licence"
                Case TOCRLICENCE_V3PRO
                    Msg = Msg & " V3 Pro/V4 licence"
                Case TOCRLICENCE_V3PROUPGRADE
                    Msg = Msg & " V1/2 UPGRADE to V3 Pro/V4 licence"
                Case TOCRLICENCE_V3SEPROUPGRADE
                    Msg = Msg & " V3 SE UPGRADE to V3 Pro/V4 licence"
            End Select
            If Volume <> 0 Or Time <> 0 Then
                Msg = Msg & Str$(Remaining)
                If Time <> 0 Then
                    Msg = Msg & " days"
                Else
                    Msg = Msg & " A4 pages"
                End If
                Msg = Msg & " remaining on licence"
            End If
        End If
    Next SlotNo
    MsgBox Msg, vbInformation, "Example 6"
Else
    MsgBox "Failed to get number of job slots", vbExclamation, "Example 6"
End If ' NumSlots > 0

End Sub ' Example6

'---------------------------------------------------------------------------
' Get images from a TWAIN compatible device
' NOTE: Untested, Should still work as no changes to TWAIN in V5
'
Sub Example7()

Dim NumImages           As Long         ' number of images acquired
Dim BI                  As BMPINFO
Dim hDIB()              As Long         ' handles to memory blocks holding images
Dim ImgCnt              As Long         ' counter
Dim DIBNo               As Long         ' loop counter


On Error Resume Next
TOCRTWAINAcquire NumImages
If Err = ERRCANTFINDDLLENTRYPOINT Then
    MsgBox "This version of TOCR DLL does not support TWAIN", vbExclamation
    Exit Sub
End If
On Error GoTo 0

If NumImages Then
    ReDim hDIB(0 To NumImages - 1)
    ImgCnt = 0
    TOCRTWAINGetImages hDIB(0)
    ' Convert the memory pointers to bitmap handles
    BI.hBmp = 0
    For DIBNo = 0 To NumImages - 1
        If GetMonoBitmapFromDIB(BI, hDIB(DIBNo)) Then
            ImgCnt = ImgCnt + 1
        End If
        ' Free memory as you go along
        If hDIB(DIBNo) Then GlobalFree hDIB(DIBNo)
        hDIB(DIBNo) = 0
        If BI.hBmp Then DeleteObject BI.hBmp
        BI.hBmp = 0
    Next DIBNo
    MsgBox CStr(ImgCnt) & " images acquired", vbInformation, "Example 7"
Else
    MsgBox "No images acquired", vbInformation, "Example 7"
End If


End Sub ' Example7

'---------------------------------------------------------------------------
' Demonstrate TOCRSetConfig and TOCRGetConfig
'
Sub Example8()

Dim JobNo               As Long
Dim Answer              As String
Dim Value               As Long

Answer = Space(250)

' Override the INI file settings for all new jobs
TOCRSetConfig TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX
TOCRSetConfig TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_SRV_ERRORMODE, TOCRERRORMODE_MSGBOX

TOCRGetConfigStr TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_LOGFILE, Answer
MsgBox "Default Log file name = " & Answer, vbInformation, "Example 8"

TOCRSetConfigStr TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_LOGFILE, "Loggederrs.lis"
TOCRGetConfigStr TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_LOGFILE, Answer
MsgBox "New default Log file name = " & Answer, vbInformation, "Example 8"

TOCRInitialise (JobNo)
TOCRSetConfig JobNo, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_NONE

TOCRGetConfig JobNo, TOCRCONFIG_DLL_ERRORMODE, Value
MsgBox "Job DLL error mode = " & CStr(Value), vbInformation, "Example 8"

TOCRGetConfig JobNo, TOCRCONFIG_SRV_ERRORMODE, Value
MsgBox "Job Service error mode = " & CStr(Value), vbInformation, "Example 8"

TOCRGetConfigStr JobNo, TOCRCONFIG_LOGFILE, Answer
MsgBox "Job Log file name = " & Answer, vbInformation, "Example 8"

' Cause an error - then check Loggederrs.lis
TOCRSetConfig JobNo, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_LOG
TOCRSetConfig JobNo, 1000, TOCRERRORMODE_LOG

End Sub ' Example8


' Wait for the engine to complete
Private Function OCRWait(ByVal JobNo As Integer, JobInfo_EG As TOCRJOBINFO_EG) As Boolean

Dim Status          As Long
Dim JobStatus       As Long
Dim Msg             As String
Dim ErrorMode       As Long


Status = TOCRDoJob_EG(JobNo, JobInfo_EG)
If Status = TOCR_OK Then
    Status = TOCRWaitForJob(JobNo, JobStatus)
End If

If Status = TOCR_OK And JobStatus = TOCRJOBSTATUS_DONE Then
    OCRWait = True
Else
    OCRWait = False

    ' If something hass gone wrong display a message
    ' (Check that the OCR engine hasn't already displayed a message)
    TOCRGetConfig JobNo, TOCRCONFIG_DLL_ERRORMODE, ErrorMode
    If ErrorMode = TOCRERRORMODE_NONE Then
        Msg = Space$(TOCRJOBMSGLENGTH)
        TOCRGetJobStatusMsg JobNo, Msg
        MsgBox Msg, vbCritical, "OCRWait"
    End If
End If

End Function

' Wait for the engine to complete by polling
Private Function OCRPoll(ByVal JobNo As Integer, JobInfo_EG As TOCRJOBINFO_EG) As Boolean

Dim Status                  As Long
Dim JobStatus               As Long
Dim Msg                     As String
Dim ErrorMode               As Long
Dim Progress                As Single
Dim AutoOrientation         As Long

Status = TOCRDoJob_EG(JobNo, JobInfo_EG)
If Status = TOCR_OK Then
    Do
        'Status = TOCRGetJobStatus(JobNo, JobStatus)
        Status = TOCRGetJobStatusEx(JobNo, JobStatus, Progress, AutoOrientation)

        ' Do something whilst the OCR engine runs
        DoEvents: Sleep (100): DoEvents
        Debug.Print ("Progress" & Str$(Int(Progress * 100)) & "%")

    Loop While Status = TOCR_OK And JobStatus = TOCRJOBSTATUS_BUSY
End If

If Status = TOCR_OK And JobStatus = TOCRJOBSTATUS_DONE Then
    OCRPoll = True
Else
    OCRPoll = False

    ' If something hass gone wrong display a message
    ' (Check that the OCR engine hasn't already displayed a message)
    TOCRGetConfig JobNo, TOCRCONFIG_DLL_ERRORMODE, ErrorMode
    If ErrorMode = TOCRERRORMODE_NONE Then
        Msg = Space$(TOCRJOBMSGLENGTH)
        TOCRGetJobStatusMsg JobNo, Msg
        MsgBox Msg, vbCritical, "OCRPoll"
    End If
End If

End Function

'---------------------------------------------------------------------------
' Retrieve the results from the service process and load into 'Results'
' Remember the character numbers returned refer to the Windows character set.
'
Private Function GetResults(ByVal JobNo As Long, Results As TOCRRESULTS_EG) As Boolean

Dim ResultsInf          As Long         ' number of bytes needed for results
Dim Bytes()             As Byte         ' work array

GetResults = False
Results.Hdr.NumItems = 0

If TOCRGetJobResultsEx_EG(JobNo, TOCRGETRESULTS_NORMAL_EG, ResultsInf, ByVal 0&) = TOCR_OK Then
    If ResultsInf > 0 Then
        ReDim Bytes(0 To ResultsInf - 1)
        If TOCRGetJobResultsEx_EG(JobNo, TOCRGETRESULTS_NORMAL_EG, ResultsInf, Bytes(0)) = TOCR_OK Then
            UnpackResults Bytes(), Results
            With Results
                If .Hdr.StructId = 0 Then
                    If .Hdr.NumItems > 0 Then
                        If .Item(0).StructId <> 0 Then
                            MsgBox "Wrong results item structure Id" & _
                                Str$(.Item(0).StructId), vbCritical
                            .Hdr.NumItems = 0
                        End If
                    End If
                Else
                    MsgBox "Wrong results header structure Id" & _
                        Str$(.Hdr.StructId), vbCritical
                End If
            End With ' results
        End If
        GetResults = True
    End If
End If

End Function ' GetResults



'---------------------------------------------------------------------------
' This routine provides one solution to the classic VB problem of how do you
' read 'variable structures' (structures which contain a variable number of
' another structure) in VB.  There are many examples of this in the Windows
' API (BITMAPINFO being one). This typically occurs when chatting to DLLs
' written in/for C because C, having no array bound checking, has no
' difficulty with the problem.
'
' This routine assumes the required data has been read into the Byte array
' 'Bytes' and then it re-dimensions 'Results' appropriately and copies the
' data.
Private Sub UnpackResults(Bytes() As Byte, Results As TOCRRESULTS_EG)

Dim HeaderLen           As Long         ' length of TOCRRESULTSHEADER
Dim ItemLen             As Long         ' length of TOCRRESULTSITEM
Dim ResultsLen          As Long         ' real length of Results
Dim NumItems            As Long         ' number of items in Results

' Notice the use of LenB here (see VB help for difference between Len and LenB)

HeaderLen = LenB(Results.Hdr)
ItemLen = LenB(Results.Item(0))

' Find the number of items in the array

ResultsLen = UBound(Bytes) - LBound(Bytes) + 1
NumItems = (ResultsLen - HeaderLen) / ItemLen

' Copy the header

CopyMemory Results, Bytes(LBound(Bytes)), HeaderLen


' Copy the array of items

If NumItems > 0 Then
    ReDim Results.Item(0 To NumItems - 1)
    CopyMemory Results.Item(0), Bytes(LBound(Bytes) + HeaderLen), ItemLen * NumItems
End If

' Note, the reason you need two CopyMemorys is because 'Item()' in 'Results'
' is in fact just a pointer to the array of items.  You can verify this by
' find LenB(Results), re-dimension Results.Items(), refind LenB(Results) - it
' will be unchanged.
'
' Had 'Items()' been a fixed array in Results (dimensioned to some value in the
' Type declaration) then this routine will still work but you wouldn't need it
' because you could have just sent 'Results' to the API call.

End Sub ' UnpackResults

Private Function FormatResults(Results As TOCRRESULTS_EG, Answer As String) As Boolean

Dim ItemNo As Integer

FormatResults = False
Answer = ""

With Results
    If .Hdr.NumItems > 0 Then
        For ItemNo = 0 To .Hdr.NumItems - 1
            If ChrW(.Item(ItemNo).OCRCharWUnicode) = ChrW$(13) Then
                Answer = Answer & vbCrLf
            Else
                Answer = Answer & ChrW(.Item(ItemNo).OCRCharWUnicode)
            End If
        Next ItemNo
        FormatResults = True
    Else
        MsgBox "No results returned", vbInformation, "FormatResults"
    End If
End With

End Function ' Format Results


