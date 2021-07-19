Attribute VB_Name = "modTOCRUser"
'***************************************************************************
' Module:     modTOCRUser
'
Option Explicit

' User constants, Version 5.1.0.0

Public Const TOCRJOBMSGLENGTH = 512        ' max length of a job status message
Public Const TOCRFONTNAMELENGTH = 65       ' max length of a job status message

Public Const TOCRMAXPPM = 78741            ' max pixels per metre
Public Const TOCRMINPPM = 984              ' min pixels per metre

' Setting for JobNo for TOCRSetErrorMode and TOCRGetErrorMode
Public Const TOCRDEFERRORMODE = -1         ' set/get the default API error mode (applies
                                    ' when there are no jobs and is applied to new jobs)

' Settings for ErrorMode for TOCRSetErrorMode and TOCRGetErrorMode
Public Const TOCRERRORMODE_NONE = 0&        ' errors unseen (use return status of API calls)
Public Const TOCRERRORMODE_MSGBOX = 1&      ' errors will bring up a message box
Public Const TOCRERRORMODE_LOG = 2&         ' errors are sent to a log file

' Setting for TOCRShutdown
Public Const TOCRSHUTDOWNALL = -1          ' stop and shutdown processing for all jobs

' Values returned by TOCRGetJobStatus JobStatus
Public Const TOCRJOBSTATUS_ERROR = -1      ' an error ocurred processing the last job
Public Const TOCRJOBSTATUS_BUSY = 0        ' the job is still processing
Public Const TOCRJOBSTATUS_DONE = 1        ' the job completed successfully
Public Const TOCRJOBSTATUS_IDLE = 2        ' no job has been specified yet

' Settings for TOCRJOBINFO.JobType
Public Const TOCRJOBTYPE_TIFFFILE = 0      ' TOCRJOBINFO.InputFile specifies a tiff file
Public Const TOCRJOBTYPE_DIBFILE = 1       ' TOCRJOBINFO.InputFile specifies a dib (bmp) file
Public Const TOCRJOBTYPE_DIBCLIPBOARD = 2  ' clipboard contains a dib (clipboard format CF_DIB)
Public Const TOCRJOBTYPE_MMFILEHANDLE = 3  ' TOCRJOBINFO.PageNo specifies a handle to a memory mapped DIB file
Public Const TOCRJOBTYPE_PDFFILE = 4       ' TOCRJOBINFO.InputFile specifies a PDF file

' Settings for TOCRJOBINFO.Orientation
Public Const TOCRJOBORIENT_AUTO = 0        ' detect orientation and rotate automatically
Public Const TOCRJOBORIENT_OFF = 255       ' don't rotate
Public Const TOCRJOBORIENT_90 = 1          ' 90 degrees clockwise rotation
Public Const TOCRJOBORIENT_180 = 2         ' 180 degrees clockwise rotation
Public Const TOCRJOBORIENT_270 = 3         ' 270 degrees clockwise rotation

' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.ResultsReference
Public Const TOCRRESULTSREFERENCE_SELFREL = 0   ' relative to the first top left character recognised
Public Const TOCRRESULTSREFERENCE_BEFORE = 1    ' page position before rotation and deskewing
Public Const TOCRRESULTSREFERENCE_BETWEEN = 2   ' page position after rotation but before deskewing
Public Const TOCRRESULTSREFERENCE_AFTER = 3         ' page position after rotation and deskewing

' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.LexMode
Public Const TOCRJOBLEXMODE_AUTO = 0        ' decide whether to apply lex
Public Const TOCRJOBLEXMODE_ON = 1                  ' lex always on
Public Const TOCRJOBLEXMODE_OFF = 2                 ' lex always off

' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.Speed
Public Const TOCRJOBSPEED_SLOW = 0
Public Const TOCRJOBSPEED_MEDIUM = 1
Public Const TOCRJOBSPEED_FAST = 2
Public Const TOCRJOBSPEED_EXPRESS = 3

' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.CCAlgorithm (thresholded Conversions from Colour)
Public Const TOCRJOBCC_AVERAGE = 0              ' (R+G+B)/3
Public Const TOCRJOBCC_LUMA_BT601 = 1           ' 0.299*R + 0.587*G + 0.114*B
Public Const TOCRJOBCC_LUMA_BT709 = 2           ' 0.2126*R + 0.7152*G + 0.0722*B
Public Const TOCRJOBCC_DESATURATION = 3         ' (max(R,G,B) + min(R,G,B))/2
Public Const TOCRJOBCC_DECOMPOSITION_MAX = 4    ' max(R,G,B)
Public Const TOCRJOBCC_DECOMPOSITION_MIN = 5    ' min(R,G,B)
Public Const TOCRJOBCC_RED = 6                  ' R
Public Const TOCRJOBCC_GREEN = 7                ' G
Public Const TOCRJOBCC_BLUE = 8                         ' B

' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.CGAlgorithm (Conversions from Grey)
Public Const TOCRJOBCG_HISTOGRAM = 9
Public Const TOCRJOBCG_REGIONS = 10

' Flags for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.ExtraInfFlags
Public Const TOCREXTRAINF_RETURNBITMAP1 = 1
Public Const TOCREXTRAINF_RETURNBITMAPONLY = 2

' Values returned by TOCRGetJobDBInfo
Public Const TOCRJOBSLOT_FREE = 0          ' job slot is free for use
Public Const TOCRJOBSLOT_OWNEDBYYOU = 1    ' job slot is in use by your process
Public Const TOCRJOBSLOT_BLOCKEDBYYOU = 2  ' blocked by own process (re-initialise)
Public Const TOCRJOBSLOT_OWNEDBYOTHER = -1 ' job slot is in use by another process (can't use)
Public Const TOCRJOBSLOT_BLOCKEDBYOTHER = -2 ' blocked by another process (can't use)

' Values returned in WaitAnyStatus by TOCRWaitForAnyJob
Public Const TOCRWAIT_OK = 0               ' JobNo is the job that finished (get and check it's JobStatus)
Public Const TOCRWAIT_SERVICEABORT = 1     ' JobNo is the job that failed (re-initialise)
Public Const TOCRWAIT_CONNECTIONBROKEN = 2 ' JobNo is the job that failed (re-initialise)
Public Const TOCRWAIT_FAILED = -1          ' JobNo not set - check manually
Public Const TOCRWAIT_NOJOBSFOUND = -2     ' JobNo not set - no running jobs found

' Settings for Mode for TOCRGetJobResultsEx
Public Const TOCRGETRESULTS_NORMAL = 0     ' return results for TOCRRESULTS
Public Const TOCRGETRESULTS_EXTENDED = 1   ' return results for TOCRRESULTSEX

' Settings for Mode for TOCRGetJobResultsEx_EG
Public Const TOCRGETRESULTS_NORMAL_EG = 2     ' return results for TOCRRESULTS_EG
Public Const TOCRGETRESULTS_EXTENDED_EG = 3   ' return results for TOCRRESULTSEX_EG

' Values returned in ResultsInf by TOCRGetJobResults and TOCRGetJobResultsEx
Public Const TOCRGETRESULTS_NORESULTS = -1 ' no results are available

' Flags returned by TOCRResults_EG.Item().FontStyleInfo
' Flags returned by TOCRResultsEx_EG.Item().FontStyleInfo
Public Const TOCRRESULTSFONT_NOTSET = 0               ' character style is not specified
Public Const TOCRRESULTSFONT_NORMAL = 1               ' character is Normal
Public Const TOCRRESULTSFONT_ITALIC = 2               ' character is Italic
Public Const TOCRRESULTSFONT_UNDERLINE = 4            ' character is Underlined
'Public Const TOCRRESULTSFONT_BOLD = 8                 ' character is Bold - removed not yet done

' Values for TOCRConvertFormat InputFormat
Public Const TOCRCONVERTFORMAT_TIFFFILE = TOCRJOBTYPE_TIFFFILE
Public Const TOCRCONVERTFORMAT_PDFFILE = TOCRJOBTYPE_PDFFILE

' Values for TOCRConvertFormat OutputFormat
Public Const TOCRCONVERTFORMAT_DIBFILE = TOCRJOBTYPE_DIBFILE
Public Const TOCRCONVERTFORMAT_MMFILEHANDLE = TOCRJOBTYPE_MMFILEHANDLE

' Values for licence features (returned by TOCRGetLicenceInfoEx)
Public Const TOCRLICENCE_STANDARD = 1      ' V1 standard licence (no higher characters)
Public Const TOCRLICENCE_EURO = 2          ' V2 (higher characters)
Public Const TOCRLICENCE_EUROUPGRADE = 3   ' standard licence upgraded to euro (V1.4->V2)
Public Const TOCRLICENCE_V3SE = 4          ' V3SE version 3 standard edition licence (no API)
Public Const TOCRLICENCE_V3SEUPGRADE = 5   ' versions 1/2 upgraded to V3 standard edition (no API)
' Note V4 licences are the same as V3 Pro licences
Public Const TOCRLICENCE_V3PRO = 6         ' V3PRO version 3 pro licence
Public Const TOCRLICENCE_V3PROUPGRADE = 7  ' versions 1/2 upgraded to version 3 pro
Public Const TOCRLICENCE_V3SEPROUPGRADE = 8 ' version 3 standard edition upgraded to version 3 pro
Public Const TOCRLICENCE_V5 = 9 ' version 5
Public Const TOCRLICENCE_V5UPGRADE3 = 10 ' version 5 upgraded from version 3
Public Const TOCRLICENCE_V5UPGRADE12 = 11 ' version 5 upgraded from version 1/2

' Values for TOCRSetConfig and TOCRGetConfig
Public Const TOCRCONFIG_DEFAULTJOB = -1    ' default job number (all new jobs)
Public Const TOCRCONFIG_DLL_ERRORMODE = 0  ' set the dll ErrorMode
Public Const TOCRCONFIG_SRV_ERRORMODE = 1  ' set the service ErrorMode
Public Const TOCRCONFIG_SRV_THREADPRIORITY = 2 ' set the service thread priority
Public Const TOCRCONFIG_DLL_MUTEXWAIT = 3  ' set the dll mutex wait timeout (ms)
Public Const TOCRCONFIG_DLL_EVENTWAIT = 4  ' set the dll event wait timeout (ms)
Public Const TOCRCONFIG_SRV_MUTEXWAIT = 5  ' set the service mutex wait timeout (ms)
Public Const TOCRCONFIG_LOGFILE = 6        ' set the log file name
