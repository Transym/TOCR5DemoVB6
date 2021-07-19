Attribute VB_Name = "modTOCRUtil"
'***************************************************************************
' Module:     modTOCRUtil
'
Option Explicit

 Public Const BITMAP_OK = 0                ' all OK
 Public Const BITMAP_ERRS = 1              ' VB error encountered
 Public Const BITMAP_FAIL = 2              ' file is a bitmap but it failed to load
 Public Const BITMAP_NOTBITMAP = 3         ' file is not a bitmap

 Public Type BMPINFO ' just the info on bitmaps that I'm interested in
    hBmp            As Long         ' handle to bitmap
    Width           As Long         ' pixel width of bitmap
    Height          As Long         ' pixel height of bitmap
    XPelsPerMeter   As Long         ' X pixels per metre
    YPelsPerMeter   As Long         ' Y pixels per metre
End Type ' BMPINFO

 Public Const LEN_BITMAPCOREHEADER = 12    ' size in bytes of BITMAPCOREHEADER
 Public Const LEN_BITMAPINFOHEADER = 40    ' size in bytes of BITMAPINFOHEADER
 Public Const LEN_RGBQUAD = 4              ' size in bytes of RGBQUAD
 
 Public Const ERRCANTFINDDLLENTRYPOINT = 453 ' Can't find DLL entry point
 
 '--------------------------------------------------------------------------
 'Unicode MessageBox for characters added in V5
 Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long


'---------------------------------------------------------------------------
' Find the number of colours used in a DIB.  This allows a logical palette
' to be created
'
Public Function DIBNumColours(bmih As BITMAPINFOHEADER, ByVal IsCoreHeader As Boolean)

Dim NumColours          As Long         ' number of colours used in the DIB

If IsCoreHeader Then
    NumColours = 2 ^ bmih.biBitCount
Else
    If bmih.biClrUsed <> 0 Then
        NumColours = bmih.biClrUsed
    Else
        NumColours = 0
        If bmih.biBitCount = 1 Then NumColours = 2
        If bmih.biBitCount = 4 Then NumColours = 16
        If bmih.biBitCount = 8 Then NumColours = 256
    End If
End If

DIBNumColours = NumColours

End Function ' DIBNumColours

'---------------------------------------------------------------------------
' Create a logical palette for a DIB
'
Public Function DIBPalette(ByVal NumColours As Long, bmib() As Byte, ByVal IsCoreHeader As Boolean) As Long

Dim Pal                 As LOGPALETTE   ' logical palette
Dim hPal                As Long         ' handle to Pal
Dim PalEntry()          As PALETTEENTRY ' palette entry in Pal
Dim RGB3()              As RGBTRIPLE    ' an RGB triple colour value
Dim RGB4()              As RGBQUAD      ' an RGB quad value
Dim hMem                As Long         ' handle to a memory block
Dim lpMem               As Long         ' pointer to a memory block
Dim Clrno               As Long         ' loop counter for colours

DIBPalette = 0
If NumColours = 0 Then Exit Function

ReDim PalEntry(0 To NumColours - 1) As PALETTEENTRY

' Load PalEntry from the DIB

If IsCoreHeader Then
    ReDim RGB3(0 To NumColours - 1) As RGBTRIPLE
    
    CopyMemory RGB3(0), bmib(LEN_BITMAPCOREHEADER), NumColours * Len(RGB3(0))
    
    For Clrno = 0 To NumColours - 1
        PalEntry(Clrno).peRed = RGB3(Clrno).rgbtRed
        PalEntry(Clrno).peBlue = RGB3(Clrno).rgbtBlue
        PalEntry(Clrno).peGreen = RGB3(Clrno).rgbtGreen
        PalEntry(Clrno).peFlags = 0
    Next Clrno
    
    Erase RGB3
Else
    ReDim RGB4(0 To NumColours - 1) As RGBQUAD
    
    CopyMemory RGB4(0), bmib(LEN_BITMAPINFOHEADER), NumColours * Len(RGB4(0))
    
    For Clrno = 0 To NumColours - 1
        PalEntry(Clrno).peRed = RGB4(Clrno).rgbRed
        PalEntry(Clrno).peBlue = RGB4(Clrno).rgbBlue
        PalEntry(Clrno).peGreen = RGB4(Clrno).rgbGreen
        PalEntry(Clrno).peFlags = 0
    Next Clrno
    
    Erase RGB4
End If

Pal.PALVERSION = PALVERSION
Pal.palNumEntries = NumColours

' Copy Pal to a memory block and return a handle to it

hMem = GlobalAlloc(GHND, Len(Pal) + Len(PalEntry(0)) * NumColours)
If hMem = 0 Then Exit Function

lpMem = GlobalLock(hMem)
If lpMem = 0 Then
    GlobalFree (hMem)
    Exit Function
End If

CopyMemory ByVal lpMem, Pal, Len(Pal)
CopyMemory ByVal lpMem + 4, PalEntry(0), Len(PalEntry(0)) * NumColours

hPal = CreatePaletteMy(lpMem)
GlobalUnlock (hMem)
GlobalFree (hMem)

DIBPalette = hPal

End Function ' DIBPalette

'---------------------------------------------------------------------------
' Convert a DIB to a mono bitmap
'
Public Function DIBtoMonoBitmap(bmi() As Byte, Data() As Byte, BI As BMPINFO) As Long

Dim bmih                As BITMAPINFOHEADER
Dim bmch                As BITMAPCOREHEADER
Dim hMonoBmp            As Long         ' bitmap handle to mono
Dim hClrBmp             As Long         ' bitmap handle to colour
Dim result              As Long         ' API return status
Dim hDCMem              As Long         ' handle to a memory device context
Dim DataAddr            As Long         ' address of DIB section data
Dim IsCoreHeader        As Boolean      ' flag if bitmap header is BITMAPCOREHEADER
Dim NumColours          As Long         ' number of colours in the DIB palette
Dim hPal                As Long         ' handle to a palette
Dim hPalOld             As Long         ' handle to a palette

DIBtoMonoBitmap = BITMAP_NOTBITMAP

hMonoBmp = 0
hClrBmp = 0
hPalOld = 0
hPal = 0

On Error GoTo DMBErrs

' See what type of header it has

CopyMemory bmih, bmi(0), 4
If bmih.biSize = Len(bmch) Then
    IsCoreHeader = True
    CopyMemory bmch, bmi(0), Len(bmch)
    
    bmih.biBitCount = bmch.bcBitCount
    bmih.biHeight = bmch.bcHeight
    bmih.biPlanes = bmch.bcPlanes
    bmih.biWidth = bmch.bcWidth
    bmih.biCompression = BI_RGB
    bmih.biXPelsPerMeter = 3780 ' assumed values
    bmih.biYPelsPerMeter = 3780 ' assumed values
Else
    IsCoreHeader = False
    CopyMemory bmih, bmi(0), Len(bmih)
End If

' Validate BITMAPINFOHEADER

If bmih.biBitCount <> 1 And bmih.biBitCount <> 4 And bmih.biBitCount <> 8 And _
    bmih.biBitCount <> 16 And bmih.biBitCount <> 24 And bmih.biBitCount <> 32 Then Exit Function
If bmih.biWidth = 0 Or bmih.biHeight = 0 Then Exit Function
If bmih.biCompression <> BI_RGB And bmih.biCompression <> BI_RLE4 And _
     bmih.biCompression <> BI_RLE8 And bmih.biCompression <> BI_BITFIELDS Then Exit Function
If (bmih.biCompression = BI_RLE4 And bmih.biBitCount <> 4) Or _
    (bmih.biCompression = BI_RLE8 And bmih.biBitCount <> 8) Then Exit Function
If bmih.biPlanes <> 1 Then Exit Function

' At this point think it's a bitmap - but may fail to load it

DIBtoMonoBitmap = BITMAP_FAIL

' Create a mono bitmap as this is what the Service process will use
'
' If you implement a better way to do this then when you OCR you should
' pass a bitmap and not the filename to the OCR service process because
' the service process mimics the behaviour below (change JobType from
' TOCRJOBTYPE_DIBFILE to TOCRJOBTYPE_DIBCLIPBOARD in frmViewer.GetFile).

result = 0
hDCMem = GetDC(0&)
If hDCMem Then

    ' Create a logical palette from the DIB if required and load
    
    If GetDeviceCaps(hDCMem, RASTERCAPS) And RC_PALETTE Then
        NumColours = DIBNumColours(bmih, IsCoreHeader)
        If NumColours Then
            hPal = DIBPalette(NumColours, bmi(), IsCoreHeader)
            If hPal Then
                hPalOld = SelectPalette(hDCMem, hPal, False)
                RealizePalette hDCMem
            End If
        End If
    End If
    
    ' Create a colour bitmap for the data
    ' CreateDIBSection fails for run length encoded files
    If bmih.biCompression = BI_RLE4 Or bmih.biCompression = BI_RLE8 Then
        hClrBmp = CreateDIBitmapMy(hDCMem, bmi(0), 0, ByVal 0&, ByVal 0&, DIB_PAL_COLORS)
    Else
        hClrBmp = CreateDIBSectionMy(hDCMem, bmi(0), DIB_RGB_COLORS, DataAddr, 0&, 0&)
    End If
    If hClrBmp Then
        If SetDIBitsMy(hDCMem, hClrBmp, 0, Abs(bmih.biHeight), Data(0), bmi(0), DIB_RGB_COLORS) Then
            GdiFlush

            ' Create mono bitmap info

            bmih.biSize = Len(bmih)
            bmih.biBitCount = 1
            bmih.biClrImportant = 0
            bmih.biClrUsed = 2
            bmih.biCompression = BI_RGB
            bmih.biSizeImage = Int((bmih.biWidth + 31&) / 32) * 4& * Abs(bmih.biHeight)

            ReDim Data(0 To bmih.biSizeImage - 1)
            
            ' Ensure bmi is long enough for the header and colour table

            ReDim bmi(0 To Len(bmih) + LEN_RGBQUAD * 2 - 1)

            ' Load the new header into bmi

            CopyMemory bmi(0), bmih, Len(bmih)

            If GetDIBitsMy(hDCMem, hClrBmp, 0, Abs(bmih.biHeight), Data(0), bmi(0), DIB_RGB_COLORS) Then

                ' Save some more space

                DeleteObject hClrBmp
                hClrBmp = 0

                hMonoBmp = CreateBitmap(bmih.biWidth, Abs(bmih.biHeight), 1, 1, ByVal 0&)
                If hMonoBmp Then
                    result = SetDIBitsMy(hDCMem, hMonoBmp, 0, Abs(bmih.biHeight), Data(0), bmi(0), DIB_RGB_COLORS)
                    GdiFlush
                End If
            End If
            
        End If ' SetDIBitsMy
    End If ' hClrBmp
    
    ' Clean up
    
    If result = 0 And hMonoBmp Then
        DeleteObject hMonoBmp
        hMonoBmp = 0
    End If
    
    If hClrBmp Then
        DeleteObject hClrBmp
        hClrBmp = 0
    End If
    
    If hPalOld Then
        hPal = SelectPalette(hDCMem, hPalOld, False)
        hPalOld = 0
        RealizePalette hDCMem
    End If
    
    If hPal Then
        DeleteObject hPal
        hPal = 0
    End If
    
    ReleaseDC 0&, hDCMem
End If ' hDCMem

If result Then
    If BI.hBmp Then
        DeleteObject BI.hBmp
        BI.hBmp = 0
    End If
    
    ' Load the BitmapInfo structure
    
    With BI
        .hBmp = hMonoBmp
        .Width = bmih.biWidth
        .Height = Abs(bmih.biHeight)
        .XPelsPerMeter = bmih.biXPelsPerMeter
        .YPelsPerMeter = bmih.biYPelsPerMeter
    End With
    
    DIBtoMonoBitmap = BITMAP_OK
End If ' result

Exit Function

' - - - - - - - - Error Handler and Exit  - - - - - -
DMBErrs:

DIBtoMonoBitmap = BITMAP_ERRS

DMBExit:

If hMonoBmp Then
    DeleteObject hMonoBmp
    hMonoBmp = 0
End If
If hClrBmp Then
    DeleteObject hClrBmp
    hClrBmp = 0
End If
If hPalOld Then
    hPal = SelectPalette(hDCMem, hPalOld, False)
    hPalOld = 0
    RealizePalette hDCMem
End If
If hPal Then
    DeleteObject hPal
    hPal = 0
End If
If hDCMem Then ReleaseDC 0, hDCMem

Exit Function

End Function ' DIBtoMonoBitmap

'---------------------------------------------------------------------------
' Opens a bitmap file and creates a memory mono bitmap (relevant info is
' stored in the BitmapInfo structure).
' Assumes the input file exists.
'
Public Function GetBitmap(BI As BMPINFO, ByVal File As String) As Long

Dim bmfh                As BITMAPFILEHEADER
Dim FH                  As Long         ' file handle
Dim bmib()              As Byte         ' BITMAPINFO as a byte array
Dim Data()              As Byte         ' bitmap data

On Error GoTo GBErrs

GetBitmap = BITMAP_NOTBITMAP

' Warning "Open for binary" will try to create a file if it doesn't exist

FH = FreeFile
Open File For Binary Access Read Shared As #FH
Get #FH, , bmfh

' Validate BITMAPFILEHEADER

If bmfh.bfType <> BFT_BITMAP Then GoTo GBExit
If bmfh.bfSize <> LOF(FH) Then GoTo GBExit

' Extract the header and colour palette

ReDim bmib(0 To bmfh.bfOffBits - Len(bmfh) - 1)

Get #FH, Len(bmfh) + 1, bmib()

' Extract the data

ReDim Data(0 To bmfh.bfSize - bmfh.bfOffBits - 1)
Get #FH, bmfh.bfOffBits + 1, Data()
Close #FH
FH = 0

GetBitmap = DIBtoMonoBitmap(bmib(), Data(), BI)

GoTo GBExit

Exit Function

' - - - - - - - - Error Handler and Exit  - - - - - -
GBErrs:

GetBitmap = BITMAP_ERRS

GBExit:

If GetBitmap = BITMAP_ERRS Then MsgBox Err.Description, vbExclamation
If GetBitmap = BITMAP_FAIL Then MsgBox "Failed to load bitmap", vbExclamation
If GetBitmap = BITMAP_NOTBITMAP Then MsgBox "File is not a bitmap", vbExclamation

If FH Then Close #FH

Exit Function

End Function ' GetBitmap

'---------------------------------------------------------------------------
' Get a bitmap held in a global memory block DIB
'
Public Function GetMonoBitmapFromDIB(BI As BMPINFO, ByVal hDIB As Long, Optional ByVal AllowWarning As Boolean = True) As Boolean

Dim bmi                 As BITMAPINFO
Dim NumBytes            As Long         ' number of bytes required to hold the DIB
Dim lpDIB               As Long         ' pointer to memory
Dim bmib()              As Byte         ' BITMAPINFO as a byte array
Dim Data()              As Byte         ' bitmap data

GetMonoBitmapFromDIB = False

On Error GoTo GMBFDErrs

lpDIB = GlobalLock(hDIB)
If lpDIB Then

    ' Get the bitmap header
    
    CopyMemory bmi.bmiHeader, ByVal lpDIB, Len(bmi.bmiHeader)

    ' Calculate the size required for byte arrays
    
    ReDim bmi.bmiColors(0 To 0)
    NumBytes = bmi.bmiHeader.biSize + Len(bmi.bmiColors(0)) * DIBNumColours(bmi.bmiHeader, False)
    
    ReDim bmib(0 To NumBytes - 1)
    
    CopyMemory bmib(0), ByVal lpDIB, NumBytes
    lpDIB = lpDIB + NumBytes
    
    With bmi.bmiHeader
        NumBytes = Int((.biWidth * .biBitCount + 31) / 32) * 4 * Abs(.biHeight)
    End With
    
    ReDim Data(0 To NumBytes - 1)
    CopyMemory Data(0), ByVal lpDIB, NumBytes

    GlobalUnlock hDIB
    hDIB = 0

    If DIBtoMonoBitmap(bmib(), Data(), BI) = BITMAP_OK Then
        GetMonoBitmapFromDIB = True
    End If

   
End If ' lpDIB

Exit Function

' - - - - - - - - Error Handler - - - - - - - - - - -
GMBFDErrs:

If hDIB Then GlobalUnlock hDIB

If AllowWarning Then
    MsgBox "Failed to retrieve image" & vbCrLf & vbCrLf & Err.Description, vbCritical
End If

Exit Function

End Function ' GetMonoBitmapFromDIB

'---------------------------------------------------------------------------
' Save a bitmap held in memory to a memory (only) mapped file.
' The handle to the MM file is returned in hFile.
'
Public Function SaveMonoBitmapToMMFile(BI As BMPINFO, hFile As Long) As Boolean

Const PAGE_READWRITE = 4
Const FILE_MAP_WRITE = 2

Dim bmi                 As BITMAPINFO
Dim ScanWidth           As Long         ' scan line width in bytes
Dim NumBytes            As Long         ' number of bytes required to hold the DIB
Dim lpMap               As Long         ' pointer to mapped file
Dim lp                  As Long         ' pointer
Dim hDCMem              As Long         ' handle to a memory device context
Dim hbmpOld             As Long         ' handle to a bitmap

SaveMonoBitmapToMMFile = False
hFile = 0

' Initialise the header

With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
    .biWidth = BI.Width
    .biHeight = BI.Height
    .biXPelsPerMeter = BI.XPelsPerMeter
    .biYPelsPerMeter = BI.YPelsPerMeter
    .biPlanes = 1
    .biBitCount = 1
    .biCompression = BI_RGB
    .biClrUsed = 2
    .biClrImportant = 0
    ScanWidth = Int((BI.Width + 31) / 32) * 4
    .biSizeImage = ScanWidth * BI.Height
End With

' Initialise the palette

ReDim bmi.bmiColors(0 To 1)

With bmi.bmiColors(0)
    .rgbRed = 0
    .rgbGreen = 0
    .rgbBlue = 0
    .rgbReserved = 0
End With
With bmi.bmiColors(1)
    .rgbRed = 255
    .rgbGreen = 255
    .rgbBlue = 255
    .rgbReserved = 0
End With

' Calculate the size of the memory mapped file memory

NumBytes = Len(bmi.bmiHeader) + Len(bmi.bmiColors(0)) * 2 + bmi.bmiHeader.biSizeImage

' Create a memory only file

hFile = CreateFileMappingMy(&HFFFFFFFF, ByVal 0&, PAGE_READWRITE, 0, NumBytes, ByVal 0&)
If hFile Then
    lpMap = MapViewOfFile(hFile, FILE_MAP_WRITE, 0, 0, 0)
    If lpMap Then
    
        ' Copy the bitmap header to the MM file
        
        lp = lpMap
        CopyMemory ByVal lp, bmi.bmiHeader, Len(bmi.bmiHeader)
        lp = lp + Len(bmi.bmiHeader)
        CopyMemory ByVal lp, bmi.bmiColors(0), Len(bmi.bmiColors(0)) * 2
        lp = lp + Len(bmi.bmiColors(0)) * 2
        
        ' Retrieve the bitmap bits and copy to the MM file
        
        hDCMem = CreateCompatibleDC(0)
        If hDCMem Then
            If GetDIBitsMy(hDCMem, BI.hBmp, 0, BI.Height, ByVal lp, ByVal lpMap, DIB_RGB_COLORS) Then
                SaveMonoBitmapToMMFile = True
            End If
            DeleteDC hDCMem
        End If ' hDCMem
        
        UnmapViewOfFile ByVal lpMap
        
    End If ' lpMap
End If ' hFile

If Not SaveMonoBitmapToMMFile Then
    If hFile Then
        CloseHandle hFile
        hFile = 0
    End If
End If

End Function ' SaveMonoBitmapToMMFile


