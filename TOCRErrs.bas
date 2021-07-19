Attribute VB_Name = "modTOCRErrs"
'***************************************************************************
' Module:     modTOCRErrs
'
Option Explicit

' List of all error codes, Version 5.1.0.0

Public Const TOCR_OK = 0

' Error codes returned by an API function
Public Const TOCRERR_ILLEGALJOBNO = 1
Public Const TOCRERR_FAILLOCKDB = 2
Public Const TOCRERR_NOFREEJOBSLOTS = 3
Public Const TOCRERR_FAILSTARTSERVICE = 4
Public Const TOCRERR_FAILINITSERVICE = 5
Public Const TOCRERR_JOBSLOTNOTINIT = 6
Public Const TOCRERR_JOBSLOTINUSE = 7
Public Const TOCRERR_SERVICEABORT = 8
Public Const TOCRERR_CONNECTIONBROKEN = 9
Public Const TOCRERR_INVALIDSTRUCTID = 10
Public Const TOCRERR_FAILGETVERSION = 11
Public Const TOCRERR_FAILLICENCEINF = 12
Public Const TOCRERR_LICENCEEXCEEDED = 13
Public Const TOCRERR_INCORRECTLICENCE = 14
Public Const TOCRERR_MISMATCH = 15
Public Const TOCRERR_JOBSLOTNOTYOURS = 16

Public Const TOCRERR_FAILGETJOBSTATUS1 = 20
Public Const TOCRERR_FAILGETJOBSTATUS2 = 21
Public Const TOCRERR_FAILGETJOBSTATUS3 = 22
Public Const TOCRERR_FAILCONVERT = 23
Public Const TOCRERR_FAILSETCONFIG = 24
Public Const TOCRERR_FAILGETCONFIG = 25
Public Const TOCRERR_FAILGETJOBSTATUS4 = 26

Public Const TOCRERR_FAILDOJOB1 = 30
Public Const TOCRERR_FAILDOJOB2 = 31
Public Const TOCRERR_FAILDOJOB3 = 32
Public Const TOCRERR_FAILDOJOB4 = 33
Public Const TOCRERR_FAILDOJOB5 = 34
Public Const TOCRERR_FAILDOJOB6 = 35
Public Const TOCRERR_FAILDOJOB7 = 36
Public Const TOCRERR_FAILDOJOB8 = 37
Public Const TOCRERR_FAILDOJOB9 = 38
Public Const TOCRERR_FAILDOJOB10 = 39
Public Const TOCRERR_UNKNOWNJOBTYPE1 = 40
Public Const TOCRERR_JOBNOTSTARTED1 = 41
Public Const TOCRERR_FAILDUPHANDLE = 42

Public Const TOCRERR_FAILGETJOBSTATUSMSG1 = 45
Public Const TOCRERR_FAILGETJOBSTATUSMSG2 = 46

Public Const TOCRERR_FAILGETNUMPAGES1 = 50
Public Const TOCRERR_FAILGETNUMPAGES2 = 51
Public Const TOCRERR_FAILGETNUMPAGES3 = 52
Public Const TOCRERR_FAILGETNUMPAGES4 = 53
Public Const TOCRERR_FAILGETNUMPAGES5 = 54

Public Const TOCRERR_FAILGETRESULTS1 = 60
Public Const TOCRERR_FAILGETRESULTS2 = 61
Public Const TOCRERR_FAILGETRESULTS3 = 62
Public Const TOCRERR_FAILGETRESULTS4 = 63
Public Const TOCRERR_FAILALLOCMEM100 = 64
Public Const TOCRERR_FAILALLOCMEM101 = 65
Public Const TOCRERR_FILENOTSPECIFIED = 66
Public Const TOCRERR_INPUTNOTSPECIFIED = 67
Public Const TOCRERR_OUTPUTNOTSPECIFIED = 68
Public Const TOCRERR_INVALIDPARAMETER = 69

Public Const TOCRERR_FAILROTATEBITMAP = 70

Public Const TOCERR_TWAINPARTIALACQUIRE = 80
Public Const TOCERR_TWAINFAILEDACQUIRE = 81
Public Const TOCERR_TWAINNOIMAGES = 82
Public Const TOCERR_TWAINSELECTDSFAILED = 83
Public Const TOCERR_MMFNOTALLOWED = 84
Public Const TOCRERR_ILLEGALFONTID = 85


Public Const TOCRERR_FAILGETMMF = 90
Public Const TOCRERR_MMFNOTAVAILABLE = 91

Public Const TOCRERR_PDFEXTRACTOR = 95
Public Const TOCRERR_PDFERROR2 = 96
Public Const TOCRERR_PDFARCHIVER = 97

Public Const TOCRERR_FONTSNOTLOADED = -2


' Error codes which may be seen in a msgbox or console but will not be returned by an API function
'Public Const TOCRERR_INVALIDSERVICESTART         = 1000
'Public Const TOCRERR_FAILSERVICEINIT             = 1001
'Public Const TOCRERR_FAILLICENCE1                = 1002
'Public Const TOCRERR_FAILSERVICESTART            = 1003
'Public Const TOCRERR_UNKNOWNCMD                  = 1004
'Public Const TOCRERR_FAILREADCOMMAND             = 1005
'Public Const TOCRERR_FAILREADOPTIONS             = 1006
'Public Const TOCRERR_FAILWRITEJOBSTATUS1         = 1007
'Public Const TOCRERR_FAILWRITEJOBSTATUS2         = 1008
'Public Const TOCRERR_FAILWRITETHREADH            = 1009
'Public Const TOCRERR_FAILREADJOBINFO1            = 1010
'Public Const TOCRERR_FAILREADJOBINFO2            = 1011
'Public Const TOCRERR_FAILREADJOBINFO3            = 1012
'Public Const TOCRERR_FAILWRITEPROGRESS           = 1013
'Public Const TOCRERR_FAILWRITEJOBSTATUSMSG       = 1014
'Public Const TOCRERR_FAILWRITERESULTSSIZE        = 1015
'Public Const TOCRERR_FAILWRITERESULTS            = 1016
'Public Const TOCRERR_FAILWRITEAUTOORIENT         = 1017
'Public Const TOCRERR_FAILLICENCE2                = 1018
'Public Const TOCRERR_FAILLICENCE3                = 1019
'Public Const TOCRERR_TOOMANYCOLUMNS              = 1020
'Public Const TOCRERR_TOOMANYROWS                 = 1021
'Public Const TOCRERR_EXCEEDEDMAXZONE             = 1022
'Public Const TOCRERR_NSTACKTOOSMALL              = 1023
'Public Const TOCRERR_ALGOERR1                    = 1024
'Public Const TOCRERR_ALGOERR2                    = 1025
'Public Const TOCRERR_EXCEEDEDMAXCP               = 1026
'Public Const TOCRERR_CANTFINDPAGE                = 1027
'Public Const TOCRERR_UNSUPPORTEDIMAGETYPE        = 1028
'Public Const TOCRERR_IMAGETOOWIDE                = 1029
'Public Const TOCRERR_IMAGETOOLONG                = 1030
'Public Const TOCRERR_UNKNOWNJOBTYPE2             = 1031
'Public Const TOCRERR_TOOWIDETOROT                = 1032
'Public Const TOCRERR_TOOLONGTOROT                = 1033
'Public Const TOCRERR_INVALIDPAGENO               = 1034
'Public Const TOCRERR_FAILREADJOBTYPENUMBYTES     = 1035
'Public Const TOCRERR_FAILREADFILENAME            = 1036
'Public Const TOCRERR_FAILSENDNUMPAGES            = 1037
'Public Const TOCRERR_FAILOPENCLIP                = 1038
'Public Const TOCRERR_NODIBONCLIP                 = 1039
'Public Const TOCRERR_FAILREADDIBCLIP             = 1040
'Public Const TOCRERR_FAILLOCKDIBCLIP             = 1041
'Public Const TOCRERR_UNKOWNDIBFORMAT             = 1042
'Public Const TOCRERR_FAILREADDIB                 = 1043
'Public Const TOCRERR_NOXYPPM                     = 1044
'Public Const TOCRERR_FAILCREATEDIB               = 1045
'Public Const TOCRERR_FAILWRITEDIBCLIP            = 1046
'Public Const TOCRERR_FAILALLOCMEMDIB             = 1047
'Public Const TOCRERR_FAILLOCKMEMDIB              = 1048
'Public Const TOCRERR_FAILCREATEFILE              = 1049
'Public Const TOCRERR_FAILOPENFILE1               = 1050
'Public Const TOCRERR_FAILOPENFILE2               = 1051
'Public Const TOCRERR_FAILOPENFILE3               = 1052
'Public Const TOCRERR_FAILOPENFILE4               = 1053
'Public Const TOCRERR_FAILREADFILE1               = 1054
'Public Const TOCRERR_FAILREADFILE2               = 1055
'Public Const TOCRERR_FAILFINDDATA1               = 1056
'Public Const TOCRERR_TIFFERROR1                  = 1057
'Public Const TOCRERR_TIFFERROR2                  = 1058
'Public Const TOCRERR_TIFFERROR3                  = 1059
'Public Const TOCRERR_TIFFERROR4                  = 1060
'Public Const TOCRERR_FAILREADDIBHANDLE           = 1061
'Public Const TOCRERR_PAGETOOBIG                  = 1062
'Public Const TOCRERR_FAILSETTHREADPRIORITY       = 1063
'Public Const TOCRERR_FAILSETSRVERRORMODE         = 1064
'Public Const TOCRERR_FAILSENDFONT1               = 1065
'Public Const TOCRERR_FAILSENDFONT2               = 1066
'Public Const TOCRERR_FAILSENDFONT3               = 1067
'Public Const TOCRERR_FAILALLOCFONTMEM            = 1068
'Public Const TOCRERR_FAILWRITEEXTRAINF           = 1069
'
'Public Const TOCRERR_FAILREADFILENAME1           = 1070
'Public Const TOCRERR_FAILREADFILENAME2           = 1071
'Public Const TOCRERR_FAILREADFILENAME3           = 1072
'Public Const TOCRERR_FAILREADFILENAME4           = 1073
'Public Const TOCRERR_FAILREADFILENAME5           = 1074
'
'Public Const TOCRERR_FAILREADFORMAT1             = 1080
'Public Const TOCRERR_FAILREADFORMAT2             = 1081
'
'Public Const TOCRERR_FAILALLOCMEM1               = 1101
'Public Const TOCRERR_FAILALLOCMEM2               = 1102
'Public Const TOCRERR_FAILALLOCMEM3               = 1103
'Public Const TOCRERR_FAILALLOCMEM4               = 1104
'Public Const TOCRERR_FAILALLOCMEM5               = 1105
'Public Const TOCRERR_FAILALLOCMEM6               = 1106
'Public Const TOCRERR_FAILALLOCMEM7               = 1107
'Public Const TOCRERR_FAILALLOCMEM8               = 1108
'Public Const TOCRERR_FAILALLOCMEM9               = 1109
'Public Const TOCRERR_FAILALLOCMEM10              = 1110
'
'Public Const TOCRERR_FAILWRITEMMFH               = 1150
'Public Const TOCRERR_FAILREADACK                 = 1151
'Public Const TOCRERR_FAILFILEMAP                 = 1152
'Public Const TOCRERR_FAILFILEVIEW                = 1153
'Public Const TOCRERR_FAILSENDBMP                 = 1154
'
'Public Const TOCRERR_FAILREADFILE3               = 1155
'Public Const TOCRERR_FAILREADFILE4               = 1156
'
'Public Const TOCRERR_PDFERROR1                   = 1157
'
'Public Const TOCRERR_BUFFEROVERFLOW1             = 2001
'
'Public Const TOCRERR_MAPOVERFLOW                 = 2002
'Public Const TOCRERR_REBREAKNEXTCALL             = 2003
'Public Const TOCRERR_REBREAKNEXTDATA             = 2004
'Public Const TOCRERR_REBREAKEXACTCALL            = 2005
'Public Const TOCRERR_MAXZCANOVERFLOW1            = 2006
'Public Const TOCRERR_MAXZCANOVERFLOW2            = 2007
'Public Const TOCRERR_BUFFEROVERFLOW2             = 2008
'Public Const TOCRERR_NUMKCOVERFLOW               = 2009
'Public Const TOCRERR_BUFFEROVERFLOW3             = 2010
'Public Const TOCRERR_BUFFEROVERFLOW4             = 2011
'Public Const TOCRERR_SEEDERROR                   = 2012
'
'Public Const TOCRERR_FCZYREF                     = 2020
'Public Const TOCRERR_MAXTEXTLINES1               = 2021
'Public Const TOCRERR_LINEINDEX                   = 2022
'Public Const TOCRERR_MAXFCZSONLINE               = 2023
'Public Const TOCRERR_MEMALLOC1                   = 2024
'Public Const TOCRERR_MERGEBREAK                  = 2025
'
'Public Const TOCRERR_DKERNPRANGE1                = 2030
'Public Const TOCRERR_DKERNPRANGE2                = 2031
'Public Const TOCRERR_BUFFEROVERFLOW5             = 2032
'Public Const TOCRERR_BUFFEROVERFLOW6             = 2033
'
'Public Const TOCRERR_FILEOPEN1                   = 2040
'Public Const TOCRERR_FILEOPEN2                   = 2041
'Public Const TOCRERR_FILEOPEN3                   = 2042
'Public Const TOCRERR_FILEREAD1                   = 2043
'Public Const TOCRERR_FILEREAD2                   = 2044
'Public Const TOCRERR_SPWIDZERO                   = 2045
'Public Const TOCRERR_FAILALLOCMEMLEX1            = 2046
'Public Const TOCRERR_FAILALLOCMEMLEX2            = 2047
'
'Public Const TOCRERR_BADOBWIDTH                  = 2050
'Public Const TOCRERR_BADROTATION                 = 2051
'
'Public Const TOCRERR_REJHIDMEMALLOC              = 2055
'
'Public Const TOCRERR_UIDA                        = 2070
'Public Const TOCRERR_UIDB                        = 2071
'Public Const TOCRERR_ZEROUID                     = 2072
'Public Const TOCRERR_CERTAINTYDBNOTINIT          = 2073
'Public Const TOCRERR_MEMALLOCINDEX               = 2074
'Public Const TOCRERR_CERTAINTYDB_INIT            = 2075
'Public Const TOCRERR_CERTAINTYDB_DELETE          = 2076
'Public Const TOCRERR_CERTAINTYDB_INSERT1         = 2077
'Public Const TOCRERR_CERTAINTYDB_INSERT2         = 2078
'Public Const TOCRERR_OPENXORNEAREST              = 2079
'Public Const TOCRERR_XORNEAREST                  = 2079
'
'Public Const TOCRERR_OPENSETTINGS                = 2080
'Public Const TOCRERR_READSETTINGS1               = 2081
'Public Const TOCRERR_READSETTINGS2               = 2082
'Public Const TOCRERR_BADSETTINGS                 = 2083
'Public Const TOCRERR_WRITESETTINGS               = 2084
'Public Const TOCRERR_MAXSCOREDIFF                = 2085
'
'Public Const TOCRERR_YDIMREFZERO1                = 2090
'Public Const TOCRERR_YDIMREFZERO2                = 2091
'Public Const TOCRERR_YDIMREFZERO3                = 2092
'Public Const TOCRERR_ASMFILEOPEN                 = 2093
'Public Const TOCRERR_ASMFILEREAD                 = 2094
'Public Const TOCRERR_MEMALLOCASM                 = 2095
'Public Const TOCRERR_MEMREALLOCASM               = 2096
'Public Const TOCRERR_SDBFILEOPEN                 = 2097
'Public Const TOCRERR_SDBFILEREAD                 = 2098
'Public Const TOCRERR_SDBFILEBAD1                 = 2099
'Public Const TOCRERR_SDBFILEBAD2                 = 2100
'Public Const TOCRERR_MEMALLOCSDB                 = 2101
'Public Const TOCRERR_DEVEL1                      = 2102
'Public Const TOCRERR_DEVEL2                      = 2103
'Public Const TOCRERR_DEVEL3                      = 2104
'Public Const TOCRERR_DEVEL4                      = 2105
'Public Const TOCRERR_DEVEL5                      = 2106
'Public Const TOCRERR_DEVEL6                      = 2107
'Public Const TOCRERR_DEVEL7                      = 2108
'Public Const TOCRERR_DEVEL8                      = 2109
'Public Const TOCRERR_DEVEL9                      = 2110
'Public Const TOCRERR_DEVEL10                     = 2111
'Public Const TOCRERR_DEVEL11                     = 2112
'Public Const TOCRERR_DEVEL12                     = 2113
'Public Const TOCRERR_DEVEL13                     = 2114
'Public Const TOCRERR_FILEOPEN4                   = 2115
'Public Const TOCRERR_FILEOPEN5                   = 2116
'Public Const TOCRERR_FILEOPEN6                   = 2117
'Public Const TOCRERR_FILEREAD3                   = 2118
'Public Const TOCRERR_FILEREAD4                   = 2119
'Public Const TOCRERR_ZOOMGTOOBIG                 = 2120
'Public Const TOCRERR_ZOOMGOUTOFRANGE             = 2121
'
'Public Const TOCRERR_MEMALLOCRESULTS             = 2130
'
'Public Const TOCRERR_MEMALLOCHEAP                = 2140
'Public Const TOCRERR_HEAPNOTINITIALISED          = 2141
'Public Const TOCRERR_MEMLIMITHEAP                = 2142
'Public Const TOCRERR_MEMREALLOCHEAP              = 2143
'Public Const TOCRERR_MEMALLOCFCZBM               = 2144
'Public Const TOCRERR_FCZBMOVERLAP                = 2145
'Public Const TOCRERR_FCZBMLOCATION               = 2146
'Public Const TOCRERR_MEMREALLOCFCZBM             = 2147
'Public Const TOCRERR_MEMALLOCFCHBM               = 2148
'Public Const TOCRERR_MEMREALLOCFCHBM             = 2149
