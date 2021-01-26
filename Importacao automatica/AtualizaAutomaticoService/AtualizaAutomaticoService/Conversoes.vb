Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module Conversoes
	Public Cancelar As Boolean
	
	
	Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
	
	Public Const PHYSICALWIDTH As Short = 110
	Public Const PHYSICALHEIGHT As Short = 111
	Public Const PHYSICALOFFSETX As Short = 112
	Public Const PHYSICALOFFSETY As Short = 113
	
	'UPGRADE_WARNING: Structure SECURITY_ATTRIBUTES may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function CreateDirectory Lib "kernel32"  Alias "CreateDirectoryA"(ByVal lpPathName As String, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Integer

    Public Structure SECURITY_ATTRIBUTES
        Dim nLength As Integer
        Dim lpSecurityDescriptor As Integer
        Dim bInheritHandle As Integer
    End Structure
	
	Public Structure SYSTEMTIME
		Dim wYear As Short
		Dim wMonth As Short
		Dim wDayOfWeek As Short
		Dim wDay As Short
		Dim wHour As Short
		Dim wMinute As Short
		Dim wSecond As Short
		Dim wMilliseconds As Short
	End Structure
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetDateFormat Lib "kernel32"  Alias "GetDateFormatA"(ByVal Locale As Integer, ByVal dwFlags As Integer, ByRef lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Integer) As Integer
	
	
	
	Private Structure ctrObj
		Dim Name As String
		Dim Index As Integer
		Dim Parrent As String
		Dim Top As Integer
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Height As Integer
		Dim Width As Integer
		Dim ScaleHeight As Integer
		Dim ScaleWidth As Integer
	End Structure
	
	Private FormRecord() As ctrObj
	Private ControlRecord() As ctrObj
	Private MaxForm As Integer
	Private MaxControl As Integer
	
	Const FW_NORMAL As Short = 400
	Const DEFAULT_CHARSET As Short = 1
	Const OUT_DEFAULT_PRECIS As Short = 0
	Const CLIP_DEFAULT_PRECIS As Short = 0
	Const DEFAULT_QUALITY As Short = 0
	Const DEFAULT_PITCH As Short = 0
	Const FF_ROMAN As Short = 16
	Const CF_PRINTERFONTS As Integer = &H2
	Const CF_SCREENFONTS As Integer = &H1
	Const CF_BOTH As Boolean = (CF_SCREENFONTS Or CF_PRINTERFONTS)
	Const CF_EFFECTS As Integer = &H100
	Const CF_FORCEFONTEXIST As Integer = &H10000
	Const CF_INITTOLOGFONTSTRUCT As Integer = &H40
	Const CF_LIMITSIZE As Integer = &H2000
	Const REGULAR_FONTTYPE As Integer = &H400
	Const LF_FACESIZE As Short = 32
	Const CCHDEVICENAME As Short = 32
	Const CCHFORMNAME As Short = 32
	Const GMEM_MOVEABLE As Integer = &H2
	Const GMEM_ZEROINIT As Integer = &H40
	Const DM_DUPLEX As Integer = &H1000
	Const DM_ORIENTATION As Integer = &H1
	Const PD_PRINTSETUP As Integer = &H40
	Const PD_DISABLEPRINTTOFILE As Integer = &H80000
	
	
	Private Structure POINTAPI
		Dim X As Integer
		Dim y As Integer
	End Structure
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	Private Structure OPENFILENAME
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hInstance As Integer
		Dim lpstrFilter As String
		Dim lpstrCustomFilter As String
		Dim nMaxCustFilter As Integer
		Dim nFilterIndex As Integer
		Dim lpstrFile As String
		Dim nMaxFile As Integer
		Dim lpstrFileTitle As String
		Dim nMaxFileTitle As Integer
		Dim lpstrInitialDir As String
		Dim lpstrTitle As String
		Dim flags As Integer
		Dim nFileOffset As Short
		Dim nFileExtension As Short
		Dim lpstrDefExt As String
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
	End Structure
	Private Structure PageSetupDlg
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hDevMode As Integer
		Dim hDevNames As Integer
		Dim flags As Integer
		Dim ptPaperSize As POINTAPI
		Dim rtMinMargin As RECT
		Dim rtMargin As RECT
		Dim hInstance As Integer
		Dim lCustData As Integer
		Dim lpfnPageSetupHook As Integer
		Dim lpfnPagePaintHook As Integer
		Dim lpPageSetupTemplateName As String
		Dim hPageSetupTemplate As Integer
	End Structure
	Private Structure ChooseColor
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hInstance As Integer
		Dim rgbResult As Integer
		Dim lpCustColors As String
		Dim flags As Integer
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
	End Structure
	Private Structure LOGFONT
		Dim lfHeight As Integer
		Dim lfWidth As Integer
		Dim lfEscapement As Integer
		Dim lfOrientation As Integer
		Dim lfWeight As Integer
		Dim lfItalic As Byte
		Dim lfUnderline As Byte
		Dim lfStrikeOut As Byte
		Dim lfCharSet As Byte
		Dim lfOutPrecision As Byte
		Dim lfClipPrecision As Byte
		Dim lfQuality As Byte
		Dim lfPitchAndFamily As Byte
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(31),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=31)> Public lfFaceName() As Char
	End Structure
	Private Structure ChooseFont
		Dim lStructSize As Integer
		Dim hwndOwner As Integer '  caller's window handle
		Dim hdc As Integer '  printer DC/IC or NULL
		Dim lpLogFont As Integer '  ptr. to a LOGFONT struct
		Dim iPointSize As Integer '  10 * size in points of selected font
		Dim flags As Integer '  enum. type flags
		Dim rgbColors As Integer '  returned text color
		Dim lCustData As Integer '  data passed to hook fn.
		Dim lpfnHook As Integer '  ptr. to hook function
		Dim lpTemplateName As String '  custom template name
		Dim hInstance As Integer '  instance handle of.EXE that
		'    contains cust. dlg. template
		Dim lpszStyle As String '  return the style field here
		'  must be LF_FACESIZE or bigger
		Dim nFontType As Short '  same value reported to the EnumFonts
		'    call back with the extra FONTTYPE_
		'    bits added
		Dim MISSING_ALIGNMENT As Short
		Dim nSizeMin As Integer '  minimum pt size allowed &
		Dim nSizeMax As Integer '  max pt size allowed if
		'    CF_LIMITSIZE is used
	End Structure
	Private Structure PRINTDLG_TYPE
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hDevMode As Integer
		Dim hDevNames As Integer
		Dim hdc As Integer
		Dim flags As Integer
		Dim nFromPage As Short
		Dim nToPage As Short
		Dim nMinPage As Short
		Dim nMaxPage As Short
		Dim nCopies As Short
		Dim hInstance As Integer
		Dim lCustData As Integer
		Dim lpfnPrintHook As Integer
		Dim lpfnSetupHook As Integer
		Dim lpPrintTemplateName As String
		Dim lpSetupTemplateName As String
		Dim hPrintTemplate As Integer
		Dim hSetupTemplate As Integer
	End Structure
	Private Structure DEVNAMES_TYPE
		Dim wDriverOffset As Short
		Dim wDeviceOffset As Short
		Dim wOutputOffset As Short
		Dim wDefault As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(100),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=100)> Public extra() As Char
	End Structure
	Private Structure DEVMODE_TYPE
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHDEVICENAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCHDEVICENAME)> Public dmDeviceName() As Char
		Dim dmSpecVersion As Short
		Dim dmDriverVersion As Short
		Dim dmSize As Short
		Dim dmDriverExtra As Short
		Dim dmFields As Integer
		Dim dmOrientation As Short
		Dim dmPaperSize As Short
		Dim dmPaperLength As Short
		Dim dmPaperWidth As Short
		Dim dmScale As Short
		Dim dmCopies As Short
		Dim dmDefaultSource As Short
		Dim dmPrintQuality As Short
		Dim dmColor As Short
		Dim dmDuplex As Short
		Dim dmYResolution As Short
		Dim dmTTOption As Short
		Dim dmCollate As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHFORMNAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCHFORMNAME)> Public dmFormName() As Char
		Dim dmUnusedPadding As Short
		Dim dmBitsPerPel As Short
		Dim dmPelsWidth As Integer
		Dim dmPelsHeight As Integer
		Dim dmDisplayFlags As Integer
		Dim dmDisplayFrequency As Integer
	End Structure
	
	'Informações do Windows
	Public Structure VolumeInf
		Dim Nome As String
		Dim Sistema As String
		Dim Serie As Integer
	End Structure
	
	Public Declare Function GetWindowsDirectory Lib "kernel32"  Alias "GetWindowsDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	Public Declare Function GetVolumeInformation Lib "kernel32"  Alias "GetVolumeInformationA"(ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, ByRef lpVolumeSerialNumber As Integer, ByRef lpMaximumComponentLength As Integer, ByRef lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Integer) As Integer
	
	
	Const THREAD_BASE_PRIORITY_IDLE As Short = -15
	Const THREAD_BASE_PRIORITY_LOWRT As Short = 15
	Const THREAD_BASE_PRIORITY_MIN As Short = -2
	Const THREAD_BASE_PRIORITY_MAX As Short = 2
	Const THREAD_PRIORITY_LOWEST As Short = THREAD_BASE_PRIORITY_MIN
	Const THREAD_PRIORITY_HIGHEST As Short = THREAD_BASE_PRIORITY_MAX
	Const THREAD_PRIORITY_BELOW_NORMAL As Object = (THREAD_PRIORITY_LOWEST + 1)
	Const THREAD_PRIORITY_ABOVE_NORMAL As Object = (THREAD_PRIORITY_HIGHEST - 1)
	Const THREAD_PRIORITY_IDLE As Short = THREAD_BASE_PRIORITY_IDLE
	Const THREAD_PRIORITY_NORMAL As Short = 0
	Const THREAD_PRIORITY_TIME_CRITICAL As Short = THREAD_BASE_PRIORITY_LOWRT
	Const HIGH_PRIORITY_CLASS As Integer = &H80
	Const IDLE_PRIORITY_CLASS As Integer = &H40
	Const NORMAL_PRIORITY_CLASS As Integer = &H20
	Const REALTIME_PRIORITY_CLASS As Integer = &H100
	
	Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Integer, ByVal nPriority As Integer) As Integer
	Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Integer, ByVal dwPriorityClass As Integer) As Integer
	Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Integer) As Integer
	Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Integer) As Integer
	Private Declare Function GetCurrentThread Lib "kernel32" () As Integer
	Private Declare Function GetCurrentProcess Lib "kernel32" () As Integer
	
	
	
	'UPGRADE_WARNING: Structure ChooseColor may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ChooseColor_Renamed Lib "COMDLG32.DLL"  Alias "ChooseColorA"(ByRef pChoosecolor As ChooseColor) As Integer
	'UPGRADE_WARNING: Structure OPENFILENAME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetOpenFileName Lib "COMDLG32.DLL"  Alias "GetOpenFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'UPGRADE_WARNING: Structure OPENFILENAME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetSaveFileName Lib "COMDLG32.DLL"  Alias "GetSaveFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'UPGRADE_WARNING: Structure PRINTDLG_TYPE may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function PrintDialog Lib "COMDLG32.DLL"  Alias "PrintDlgA"(ByRef pPrintdlg As PRINTDLG_TYPE) As Integer
	'UPGRADE_WARNING: Structure PageSetupDlg may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function PageSetupDlg_Renamed Lib "COMDLG32.DLL"  Alias "PageSetupDlgA"(ByRef pPagesetupdlg As PageSetupDlg) As Integer
	'UPGRADE_WARNING: Structure ChooseFont may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ChooseFont_Renamed Lib "COMDLG32.DLL"  Alias "ChooseFontA"(ByRef pChoosefont As ChooseFont) As Integer
	Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
	Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) As Integer
	
	Dim OFName As OPENFILENAME
	Dim CustomColors() As Byte
	
	Public Const AlinhaEsquerda As Short = 0
	Public Const AlinhaDireita As Short = 1
	Public Const AlinhaCentralizado As Short = 2
	Public Const AlinhaJustificado As Short = 3
	
	Public Function DirWindows() As String
		Dim Path, strSave As String
		strSave = New String(Chr(0), 200)
		Path = Left(strSave, GetWindowsDirectory(strSave, Len(strSave)))
		DirWindows = Path
	End Function

    Public Function InformaVolume(ByVal Disco As String) As VolumeInf
        Dim Serial As Integer
        Dim VName, FSName As String
        VName = New String(Chr(0), 255)
        FSName = New String(Chr(0), 255)
        GetVolumeInformation("C:\", VName, 255, Serial, 0, 0, FSName, 255)
        VName = Left(VName, InStr(1, VName, Chr(0)) - 1)
        FSName = Left(FSName, InStr(1, FSName, Chr(0)) - 1)
        If Serial < 0 Then
            Serial = Serial * -1
        End If
        InformaVolume.Nome = VName
        InformaVolume.Serie = Serial
        InformaVolume.Sistema = FSName
    End Function
	
	Public Sub PrioridadeAlta()
		Dim hThread, hProcess As Integer
		'retrieve the current thread and process
		hThread = GetCurrentThread
		hProcess = GetCurrentProcess
		'set the new thread priority to "lowest"
		'UPGRADE_WARNING: Couldn't resolve default property of object THREAD_PRIORITY_TIME_CRITICAL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetThreadPriority(hThread, THREAD_PRIORITY_TIME_CRITICAL)
		'set the new priority class to "idle"
		SetPriorityClass(hProcess, REALTIME_PRIORITY_CLASS)
	End Sub
	
	Public Sub PrioridadeNormal()
		Dim hThread, hProcess As Integer
		'retrieve the current thread and process
		hThread = GetCurrentThread
		hProcess = GetCurrentProcess
		'set the new thread priority to "lowest"
		SetThreadPriority(hThread, THREAD_PRIORITY_NORMAL)
		'set the new priority class to "idle"
		SetPriorityClass(hProcess, NORMAL_PRIORITY_CLASS)
		
	End Sub
	
	' Return the next word from this string. Remove
	' the word from the string.
	Private Function GetWord(ByRef txt As String) As String
		Dim pos As Short
		
		txt = Trim(txt)
		pos = InStr(txt, " ")
		If pos < 1 Then
			GetWord = txt
			txt = ""
		Else
			GetWord = Left(txt, pos - 1)
			txt = Trim(Right(txt, Len(txt) - pos))
		End If
	End Function
	
	Private Sub NonPrintToSpace(ByRef txt As String)
		Dim I As Short
		Dim txtlen As Short
		Dim ch As String
		
		txtlen = Len(txt)
		For I = 1 To txtlen
			ch = Mid(txt, I, 1)
			If ch < " " Or ch > "~" Then Mid(txt, I, 1) = " "
		Next I
	End Sub
	
	Public Function Dec2Sex(ByVal strDecimal As String) As String
		Dim Virgula As Short
		Dim I As Double
		Dim strDecimal2, strInteiro As String
		For I = 1 To Len(strDecimal)
			If Mid(strDecimal, I, 1) = "," Or Mid(strDecimal, I, 1) = "." Then
				Virgula = I
				Exit For
			End If
		Next I
		On Error Resume Next
		If Virgula > 0 Then
			strDecimal2 = Mid(strDecimal, Virgula + 1, 2)
			strDecimal2 = Left(strDecimal2 & "00", 2)
		Else
			strDecimal2 = "00"
		End If
		If Virgula > 0 Then
			strInteiro = Mid(strDecimal, 1, Virgula - 1)
			strInteiro = Right("0000" & Trim(strInteiro), 4)
		Else
			strInteiro = Right("0000" & Trim(strDecimal), 4)
		End If
		strDecimal = strInteiro & "," & strDecimal2
		Virgula = CShort(Mid(strDecimal, 6, 2)) * 0.6
		Dec2Sex = Mid(strDecimal, 1, 4) & ":" & Right("00" & Trim(Str(Virgula)), 2)
	End Function
	
	
	Public Function Sex2Dec(ByVal Sexagesimal As String) As String
		Dim Virgula As Short
		Dim I As Double
		Dim strDecimal, strInteiro As String
		For I = 1 To Len(Sexagesimal)
			If Mid(Sexagesimal, I, 1) = "," Or Mid(Sexagesimal, I, 1) = "." Or Mid(Sexagesimal, I, 1) = ":" Then
				Virgula = I
				Exit For
			End If
		Next I
		On Error Resume Next
		If Virgula > 0 Then
			strDecimal = Mid(Sexagesimal, Virgula + 1, 2)
			strDecimal = Left(strDecimal & "00", 2)
		Else
			strDecimal = "00"
		End If
		If Virgula > 0 Then
			strInteiro = Mid(Sexagesimal, 1, Virgula - 1)
			strInteiro = Right("0000" & strInteiro, 4)
		Else
			strInteiro = Right("0000" & Trim(Sexagesimal), 4)
		End If
		Sexagesimal = strInteiro & ":" & strDecimal
		Virgula = CShort(Mid(Sexagesimal, 6, 2)) / 0.6
		Sex2Dec = Mid(Sexagesimal, 1, 4) & "," & Right("00" & Trim(Str(Virgula)), 2)
	End Function
	
	Public Function MinutoEmHora(ByVal Minutos As Double) As String
		Dim TempHora, tempMinuto As Double
		'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		tempMinuto = Minutos Mod 60
		Minutos = Minutos - tempMinuto
		TempHora = CDbl(Minutos / 60)
		MinutoEmHora = VB6.Format(Trim(Str(TempHora)), "#00") & ":" & VB6.Format(Trim(Str(tempMinuto)), "00")
	End Function
	
	Public Function HoraEmMinutos(ByVal HoraMinuto As String) As Double
		Dim tempMinuto, TempHora As Short
		tempMinuto = CShort(Mid(HoraMinuto, InStr(1, HoraMinuto, ":") + 1, 2))
		TempHora = CShort(Mid(HoraMinuto, 1, InStr(1, HoraMinuto, ":") - 1))
		tempMinuto = tempMinuto + (TempHora * 60)
		HoraEmMinutos = tempMinuto
	End Function
	
	
	Public Function DataInglesa(ByVal Data As String) As String
		On Error Resume Next
		Data = VB6.Format(Month(CDate(Data)), "00") & "/" & VB6.Format(VB.Day(CDate(Data)), "00") & "/" & VB6.Format(Year(CDate(Data)), "0000")
		DataInglesa = Data
	End Function
	
	Public Function DataHoraInglesa(ByVal Data As String) As String
		On Error Resume Next
		Data = VB6.Format(Month(CDate(Data)), "00") & "/" & VB6.Format(VB.Day(CDate(Data)), "00") & "/" & VB6.Format(Year(CDate(Data)), "0000") & " " & VB6.Format(Data, "long time")
		DataHoraInglesa = Data
	End Function
	
	
	Public Function NumeroIngles(ByVal Numero As String) As String
		Numero = Replace(Numero, ".", "")
		Numero = Replace(Numero, ",", ".")
		NumeroIngles = Numero
	End Function
	
	Public Sub Tempo(ByRef PauseTime As Object)
		Dim Start As Object
		'PauseTime = 3   ' ajusta o tempo de duração.
		'UPGRADE_WARNING: Couldn't resolve default property of object Start. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Start = VB.Timer() ' informa quando começa.
		'UPGRADE_WARNING: Couldn't resolve default property of object PauseTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Start. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Do While VB.Timer() < Start + PauseTime
			System.Windows.Forms.Application.DoEvents()
		Loop 
		
	End Sub
	
	Public Function SubtraiHora(ByVal Hora1 As String, ByVal Tempo As String) As String
		Dim Horafinal As String
		Dim Minutos As Double
		
		If Hora1 = "  :  " Or Tempo = "  :  " Then
			SubtraiHora = "  :  "
			Exit Function
		End If
		If Hora1 = "" Then Hora1 = "0000:00"
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		Minutos = DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate("00:00"), CDate(Tempo))
		Horafinal = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, -Minutos, CDate(Hora1)))
		Horafinal = VB6.Format(Horafinal, "short time")
		SubtraiHora = Horafinal
	End Function
	
	Public Function SomaHora(ByVal Hora1 As String, ByVal Tempo As String) As String
		Dim Horafinal As String
		Dim Minutos As Double
		
		If Hora1 = "  :  " Or Tempo = "  :  " Then
			SomaHora = "  :  "
			Exit Function
		End If
		If Hora1 = "" Then Hora1 = "0000:00"
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		Minutos = DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate("00:00"), CDate(Tempo))
		Horafinal = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, Minutos, CDate("01/01/1900 " & Hora1)))
		Horafinal = VB6.Format(Horafinal, "short time")
		SomaHora = Horafinal
	End Function
	
	Public Function DivideHora(ByVal Hora As String, Optional ByRef Divisor As Short = 0) As String
		Dim HoraTemp As Object
		If Hora = "  :  " Then
			DivideHora = "  :  "
			Exit Function
		End If
		If Divisor <= 0 Then Divisor = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object HoraTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HoraTemp = CShort(Mid(Hora, 1, 2)) Mod Divisor
		'UPGRADE_WARNING: Couldn't resolve default property of object HoraTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If HoraTemp = 0 Then
			DivideHora = VB6.Format(CShort(Mid(Hora, 1, 2)) / Divisor, "0#") & ":" & VB6.Format(CShort(Mid(Hora, 4, 2)) / Divisor, "0#")
		Else
			If (CShort(Mid(Hora, 4, 2)) / Divisor) + 30 < 60 Then
				DivideHora = VB6.Format((CShort(Mid(Hora, 1, 2)) - 1) / Divisor, "0#") & ":" & VB6.Format((CShort(Mid(Hora, 4, 2)) / Divisor) + 30, "0#")
			Else
				DivideHora = VB6.Format(((CShort(Mid(Hora, 1, 2)) - 1) / Divisor) + 1, "0#") & ":" & VB6.Format(((CShort(Mid(Hora, 4, 2)) / Divisor) + 30) - 60, "0#")
			End If
		End If
		
	End Function
	
	
	Public Function HoraMaior(ByVal Hora1 As String, ByVal Hora2 As String) As String
		'Esta função retorna uma string com o horário maior
		
		Dim tempHora1, tempHora2 As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object tempHora1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tempHora1 = Hour(CDate(Hora1))
		'UPGRADE_WARNING: Couldn't resolve default property of object tempHora2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tempHora2 = Hour(CDate(Hora2))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object tempHora2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object tempHora1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If tempHora1 > tempHora2 Then
			HoraMaior = Hora1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object tempHora2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tempHora1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If tempHora1 < tempHora2 Then
				HoraMaior = Hora2
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object tempHora1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				tempHora1 = Minute(CDate(Hora1))
				'UPGRADE_WARNING: Couldn't resolve default property of object tempHora2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				tempHora2 = Minute(CDate(Hora2))
				
				'UPGRADE_WARNING: Couldn't resolve default property of object tempHora2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object tempHora1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If tempHora1 > tempHora2 Then
					HoraMaior = Hora1
				Else
					HoraMaior = Hora2
				End If
			End If
		End If
	End Function
	
	
	Public Function DataValida(ByVal Data As String) As Boolean
		Dim A As Date
		
		On Error GoTo trataErro
		
		A = CDate(Data)
		
		DataValida = True
		Exit Function
		
trataErro: 
		DataValida = False
	End Function
	
	Public Function DataMaior(ByVal Data1 As String, ByVal Data2 As String) As Boolean
		Dim Dt1, Dt2 As Date
		Dt1 = CDate(Data1)
		Dt2 = CDate(Data2)
		
		If Dt1 >= Dt2 Then
			DataMaior = True
		Else
			DataMaior = False
		End If
	End Function
	
	Public Function DataMenor(ByVal Data1 As String, ByVal Data2 As String) As Boolean
		Dim Dt1, Dt2 As Date
		Dt1 = CDate(Data1)
		Dt2 = CDate(Data2)
		
		If Dt1 <= Dt2 Then
			DataMenor = True
		Else
			DataMenor = False
		End If
	End Function
	
	Public Function NesteMes(ByVal Divisor As Short, ByVal dtData As Date, ByVal dtData2 As Date) As Boolean
		Dim Meses As Short
		If dtData <> dtData2 Then
			Meses = DateDiff(Microsoft.VisualBasic.DateInterval.Month, dtData, dtData2)
		ElseIf dtData = dtData2 Then 
			NesteMes = True
			Exit Function
		Else
			NesteMes = False
			Exit Function
		End If
		If Meses < 0 Then Meses = Meses * -1
		If Meses Mod Divisor = 0 Then
			NesteMes = True
		Else
			NesteMes = False
		End If
	End Function
	
	
	Public Function HexParaDec(ByVal Hexadecimal As String) As String
		Dim I, Resultado, A As Double
		Dim StrTemp As String
		Dim tempTotal As Double
		Hexadecimal = UCase(Hexadecimal)
		For I = 1 To Len(Hexadecimal)
			StrTemp = Mid(Hexadecimal, I, 1)
			Select Case StrTemp
				Case "A"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 10
				Case "B"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 11
				Case "C"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 12
				Case "D"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 13
				Case "E"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 14
				Case "F"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 15
				Case Else
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * CDbl(Str(CDbl(StrTemp)))
			End Select
			Resultado = Resultado + tempTotal
		Next I
		
		HexParaDec = Str(Resultado)
	End Function
	
	
	Public Function Encripta(ByVal Texto As String) As String
		Dim StrTemp As String
		Dim I As Double
		Dim strCripto As String
        StrTemp = Texto
        strCripto = ""
		For I = 1 To Len(StrTemp)
			strCripto = strCripto & Hex((CShort(Asc(Mid(StrTemp, I, 1))) * 2 + 4) * 5)
		Next I
		
		Encripta = strCripto
	End Function
	
	Public Function Descripta(ByVal Texto As String) As String
		Dim StrTemp As String
		Dim I As Double
		Dim StrTemp2 As String
		Dim A As Short

        StrTemp2 = ""
		A = 2
		StrTemp = Texto
		For I = 1 To Len(StrTemp) Step 3
			StrTemp2 = StrTemp2 & Chr(((CShort(Hex2dec(Trim(Mid(StrTemp, I, 3)))) / 5) - 4) / 2)
			A = I + 1
		Next I
		Descripta = StrTemp2
	End Function
	
	
	
	Private Function Hex2dec(ByVal Hexadecimal As String) As String
		Dim I, Resultado, A As Double
		Dim StrTemp As String
		Dim tempTotal As Double
		Hexadecimal = UCase(Hexadecimal)
		For I = 1 To Len(Hexadecimal)
			StrTemp = Mid(Hexadecimal, I, 1)
			Select Case StrTemp
				Case "A"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 10
				Case "B"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 11
				Case "C"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 12
				Case "D"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 13
				Case "E"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 14
				Case "F"
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * 15
				Case Else
					tempTotal = 1
					For A = Len(Hexadecimal) To I + 1 Step -1
						tempTotal = tempTotal * 16
					Next A
					tempTotal = tempTotal * CDbl(Str(CDbl(StrTemp)))
			End Select
			Resultado = Resultado + tempTotal
		Next I
		
		Hex2dec = Str(Resultado)
	End Function
	
	Public Function Criptografa(ByVal Texto As String, ByVal Chave As Short) As String
		Dim X As Single
		Dim I As Short
		Dim CharNum, RandomInteger As Short
		Dim SingleChar As New VB6.FixedLengthString(1)
		Dim KeyValue As Short
		Dim strBefore, strAfter As String
		If Texto = "" Then
			Criptografa = ""
			Exit Function
		End If
		strBefore = Texto
		KeyValue = Chave
        X = Rnd(-KeyValue)
        strAfter = ""
		For I = 1 To Len(strBefore)
			SingleChar.Value = Mid(strBefore, I, 1)
			CharNum = Asc(SingleChar.Value)
			RandomInteger = Int(256 * Rnd())
			CharNum = CharNum Xor RandomInteger
			SingleChar.Value = Chr(CharNum)
			strAfter = strAfter & SingleChar.Value
		Next I
		Criptografa = strAfter
		
	End Function
	
	Private Function ActualPos(ByRef plLeft As Integer) As Integer
		If plLeft < 0 Then
			ActualPos = plLeft + 75000
		Else
			ActualPos = plLeft
		End If
	End Function
	
	Private Function FindForm(ByRef pfrmIn As Object) As Integer
		Dim I As Integer
		FindForm = -1
		If MaxForm > 0 Then
			For I = 0 To (MaxForm - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object pfrmIn.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If FormRecord(I).Name = pfrmIn.Name Then
					FindForm = I
					Exit Function
				End If
			Next I
		End If
	End Function
	
	
	Public Function RetornaDiretorio(ByVal Arquivo As String) As String
        Dim I As Double
        RetornaDiretorio = ""
		If Arquivo = "" Then Exit Function
		For I = Len(Arquivo) To 1 Step -1
			If Mid(Arquivo, I, 1) = "\" Then
				RetornaDiretorio = Mid(Arquivo, 1, I)
				Exit Function
			End If
		Next I
	End Function
	
	Public Function RetornaArquivo(ByVal Arquivo As String) As String
        Dim I As Double
        RetornaArquivo = ""
		If Arquivo = "" Then Exit Function
		For I = Len(Arquivo) To 1 Step -1
			If Mid(Arquivo, I, 1) = "\" Then
				RetornaArquivo = Mid(Arquivo, I + 1)
				Exit Function
			End If
		Next I
    End Function

    Public Function RemoveString(ByVal strTemp As String) As String
        Dim StrTemp2 As String, I As Double

        RemoveString = ""
        StrTemp2 = ""
        For I = 1 To Len(strTemp)
            If IsNumeric(Mid(strTemp, I, 1)) = True Then
                StrTemp2 = StrTemp2 & Mid(strTemp, I, 1)
            End If
        Next
        RemoveString = StrTemp2
    End Function

    Public Function UltimoDiaDoMes(ByVal dtData As Date) As Date
        Dim tempData As Date
        UltimoDiaDoMes = dtData
        tempData = CDate("01/" & dtData.Month & "/" & dtData.Year)
        tempData = DateAdd(DateInterval.Month, 1, tempData)
        tempData = DateAdd(DateInterval.Day, -1, tempData)
        UltimoDiaDoMes = tempData
    End Function
End Module