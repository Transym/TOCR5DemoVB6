Attribute VB_Name = "modTOCRDeclares"
'***************************************************************************
' Module:     modTOCRDeclares
'
Option Explicit


' TOCR declares Version 5.1.0

Type TOCRPROCESSOPTIONS_EG
    StructId                    As Long
    InvertWholePage             As Boolean
    DeskewOff                   As Boolean
    Orientation                 As Byte
    NoiseRemoveOff              As Boolean
    ReturnNoiseON               As Boolean
    LineRemoveOff               As Boolean
    DeshadeOff                  As Boolean
    InvertOff                   As Boolean
    SectioningOn                As Boolean
    MergeBreakOff               As Boolean
    LineRejectOff               As Boolean
    CharacterRejectOff          As Boolean
    ResultsReference            As Byte
    LexMode                     As Byte
    OCRBOnly                    As Boolean
    Speed                       As Byte
    FontStyleInfoOff            As Boolean
    CCAlgorithm                 As Byte
    CCThreshold                 As Byte
    CGAlgorithm                 As Byte
    ExtraInfFlags               As Byte
    DisableLangs(0 To 45)       As Boolean
    DisableCharacter(0 To 607)  As Boolean
End Type ' TOCRPROCESSOPTIONS_EG

' Superseded by TOCRPROCESSOPTIONS_EG
Type TOCRPROCESSOPTIONS
    StructId                    As Long
    InvertWholePage             As Boolean
    DeskewOff                   As Boolean
    Orientation                 As Byte
    NoiseRemoveOff              As Boolean
    LineRemoveOff               As Boolean
    DeshadeOff                  As Boolean
    InvertOff                   As Boolean
    SectioningOn                As Boolean
    MergeBreakOff               As Boolean
    LineRejectOff               As Boolean
    CharacterRejectOff          As Boolean
    LexOff                      As Boolean
    DisableCharacter(0 To 255)  As Boolean
End Type ' TOCRPROCESSOPTIONS

Type TOCRJOBINFO_EG
    hMMF                        As Long
    InputFile                   As String
    StructId                    As Long
    JobType                     As Long
    PageNo                      As Long
    ProcessOptions              As TOCRPROCESSOPTIONS_EG
    
End Type ' TOCRJOBINFO_EG

' Superseded by TOCRJOBINFO_EG
Type TOCRJOBINFO2
    StructId                    As Long
    JobType                     As Long
    InputFile                   As String
    hMMF                        As Long
    PageNo                      As Long
    ProcessOptions              As TOCRPROCESSOPTIONS
End Type ' TOCRJOBINFO

' Superseded by TOCRJOBINFO2
Type TOCRJOBINFO
    StructId                    As Long
    JobType                     As Long
    InputFile                   As String
    PageNo                      As Long
    ProcessOptions              As TOCRPROCESSOPTIONS
End Type ' TOCRJOBINFO

Type TOCRRESULTSHEADER_EG
    StructId                    As Long
    XPixelsPerInch              As Long
    YPixelsPerInch              As Long
    NumItems                    As Long
    MeanConfidence              As Single
    DominantLanguage            As Long
End Type ' TOCRRESULTSHEADER_EG

' Superseded by TOCRRESULTSHEADER_EG
Type TOCRRESULTSHEADER
    StructId                    As Long
    XPixelsPerInch              As Long
    YPixelsPerInch              As Long
    NumItems                    As Long
    MeanConfidence              As Single
End Type ' TOCRRESULTSHEADER

Type TOCRRESULTSITEM_EG
    Confidence                  As Single
    StructId                    As Integer
    OCRCharWUnicode             As Integer
    OCRCharWInternal            As Integer
    FontID                      As Integer
    FontStyleInfo               As Integer
    XPos                        As Integer
    YPos                        As Integer
    XDim                        As Integer
    YDim                        As Integer
    YDimRef                     As Integer
    Noise                       As Integer
End Type ' TOCRRESULTSITEM_EG

' Superseded by TOCRRESULTSITEM_EG
Type TOCRRESULTSITEM
    StructId                    As Integer
    OCRCha                      As Integer
    Confidence                  As Single
    XPos                        As Integer
    YPos                        As Integer
    XDim                        As Integer
    YDim                        As Integer
End Type ' TOCRRESULTSITEM

Type TOCRRESULTS_EG
    Hdr                         As TOCRRESULTSHEADER_EG
    Item()                      As TOCRRESULTSITEM_EG
End Type ' TOCRRESULTS_EG

' Superseded by TOCRRESULTS_EG
Type TOCRRESULTS
    Hdr                         As TOCRRESULTSHEADER
    Item()                      As TOCRRESULTSITEM
End Type ' TOCRRESULTS

Type TOCRRESULTSITEMEXALT_EG
    Factor                      As Single
    Valid                       As Integer
    OCRCharWUnicode             As Integer
    OCRCharWInternal            As Integer
    
End Type ' TOCRRESULTSITEMEXALT_EG

' Superseded by TOCRRESULTSITEMEXALT_EG
Type TOCRRESULTSITEMEXALT
    Valid                       As Integer
    OCRCha                      As Integer
    Factor                      As Single
End Type ' TOCRRESULTSITEMEXALT

Type TOCRRESULTSITEMEX_EG
    Confidence                  As Single
    StructId                    As Integer
    OCRCharWUnicode             As Integer
    OCRCharWInternal            As Integer
    XPos                        As Integer
    YPos                        As Integer
    XDim                        As Integer
    YDim                        As Integer
    YDimRef                     As Integer
    Noise                       As Integer
    Alt(0 To 4)                 As TOCRRESULTSITEMEXALT
End Type ' TOCRRESULTSITEMEX_EG

' Superseded by TOCRRESULTSITEMEX_EG
Type TOCRRESULTSITEMEX
    StructId                    As Integer
    OCRCha                      As Integer
    Confidence                  As Single
    XPos                        As Integer
    YPos                        As Integer
    XDim                        As Integer
    YDim                        As Integer
    Alt(0 To 4)                 As TOCRRESULTSITEMEXALT
End Type ' TOCRRESULTSITEMEX

Type TOCRRESULTSEX_EG
    Hdr                         As TOCRRESULTSHEADER
    Item()                      As TOCRRESULTSITEMEX
End Type ' TOCRRESULTSEX_EG

' Superseded by TOCRRESULTSEX_EG
Type TOCRRESULTSEX
    Hdr                         As TOCRRESULTSHEADER
    Item()                      As TOCRRESULTSITEMEX
End Type ' TOCRRESULTSEX

Public Declare Function TOCRInitialise Lib "TOCRDll" (JobNo As Long) As Long
Public Declare Function TOCRShutdown Lib "TOCRDll" (ByVal JobNo As Long) As Long
Public Declare Function TOCRDoJob_EG Lib "TOCRDll" (ByVal JobNo As Long, JobInfo_EG As TOCRJOBINFO_EG) As Long
Public Declare Function TOCRWaitForJob Lib "TOCRDll" (ByVal JobNo As Long, JobStatus As Long) As Long
Public Declare Function TOCRWaitForAnyJob Lib "TOCRDll" (WaitAnyStatus As Long, JobNo As Long) As Long
Public Declare Function TOCRGetJobDBInfo Lib "TOCRDll" (JobSlotInf As Long) As Long
Public Declare Function TOCRGetJobStatus Lib "TOCRDll" (ByVal JobNo As Long, JobStatus As Long) As Long
Public Declare Function TOCRGetJobStatusEx Lib "TOCRDll" (ByVal JobNo As Long, JobStatus As Long, Progress As Single, AutoOrientation As Long) As Long
Public Declare Function TOCRGetJobStatusMsg Lib "TOCRDll" (ByVal JobNo As Long, ByVal Msg As String) As Long
Public Declare Function TOCRGetNumPages Lib "TOCRDll" (ByVal JobNo As Long, ByVal Filename As String, ByVal JobType As Long, NumPages As Long) As Long
Public Declare Function TOCRGetJobResults Lib "TOCRDll" (ByVal JobNo As Long, ResultsInf As Long, Results As Any) As Long
Public Declare Function TOCRGetJobResultsEx Lib "TOCRDll" (ByVal JobNo As Long, ByVal Mode As Long, ResultsInf As Long, ResultsEx As Any) As Long
Public Declare Function TOCRGetJobResultsEx_EG Lib "TOCRDll" (ByVal JobNo As Long, ByVal Mode As Long, ResultsInf As Long, ResultsEx_EG As Any) As Long
Public Declare Function TOCRGetLicenceInfoEx Lib "TOCRDll" (ByVal JobNo As Long, ByVal Licence As String, Volume As Long, Time As Long, Remaining As Long, Features As Long) As Long
Public Declare Function TOCRConvertFormat Lib "TOCRDll" (ByVal JobNo As Long, InputAddr As Any, ByVal InputFormat As Long, OutputAddr As Any, ByVal OutputFormat As Long, ByVal PageNo As Long) As Long
Public Declare Function TOCRSetConfig Lib "TOCRDll" (ByVal JobNo As Long, ByVal Parameter As Long, ByVal Value As Long) As Long
Public Declare Function TOCRGetConfig Lib "TOCRDll" (ByVal JobNo As Long, ByVal Parameter As Long, Value As Long) As Long
Public Declare Function TOCRSetConfigStr Lib "TOCRDll" Alias "TOCRSetConfig" (ByVal JobNo As Long, ByVal Parameter As Long, ByVal Value As String) As Long
Public Declare Function TOCRGetConfigStr Lib "TOCRDll" Alias "TOCRGetConfig" (ByVal JobNo As Long, ByVal Parameter As Long, ByVal Value As String) As Long
Public Declare Function TOCRTWAINAcquire Lib "TOCRDll" (NumberOfImages As Long) As Long
Public Declare Function TOCRTWAINGetImages Lib "TOCRDll" (GlobalMemoryDIBs As Long) As Long
Public Declare Function TOCRTWAINSelectDS Lib "TOCRDll" () As Long
Public Declare Function TOCRTWAINShowUI Lib "TOCRDll" (ByVal Show As Boolean) As Long
' Superseded by TOCRGetConfig
Public Declare Function TOCRGetErrorMode Lib "TOCRDll" (ByVal JobNo As Long, ErrorMode As Long) As Long
' Superseded by TOCRSetConfig
Public Declare Function TOCRSetErrorMode Lib "TOCRDll" (ByVal JobNo As Long, ByVal ErrorMode As Long) As Long
' Superseded by TOCRDOJOB_EG
Public Declare Function TOCRDoJob2 Lib "TOCRDll" (ByVal JobNo As Long, JobInfo As TOCRJOBINFO2) As Long
'Superseded by TOCRDoJob2
Public Declare Function TOCRDoJob Lib "TOCRDll" (ByVal JobNo As Long, JobInfo As TOCRJOBINFO) As Long
Public Declare Function TOCRRotateMonoBitmap Lib "TOCRDll" (hBmp As Long, ByVal Width As Long, ByVal Height As Long, ByVal Orientation As Long) As Long
' Superseded by TOCRConvertFormat
Public Declare Function TOCRConvertTIFFtoDIB Lib "TOCRDll" (ByVal JobNo As Long, ByVal InputFilename As String, ByVal OutputFilename As String, ByVal PageNo As Long) As Long
' UNTESTED superseded by TOCRGetLicenceInfoEx
'Public Declare Function TOCRGetLicenceInfo Lib "TOCRDll" (NumberOfJobSlots As Long, Volume As Long, Time As Long, Remaining As Long) As Long
