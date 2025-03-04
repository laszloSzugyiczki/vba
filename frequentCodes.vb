
'get last column, row
	lastcol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
	lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row	

'send buttons
	SendKeys ("{DOWN}")

'get calendar week
	cw = CInt(Format(D, "ww", FW))	

'add a new sheet after the last existing sheet if it is not yet exists in active workbook, input: String: shName - name of the sheet that you want to add
	Sub addSheetIfNotExists(shName)
		
		ex = False
		
		For i = 1 To Worksheets.Count
			If Worksheets(i).Name = shName Then
				ex = True
				Exit For
			End If
		Next
		
		If Not ex Then
			Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = shName
		End If
		
	End Sub

'get fisrt available file handle number, plus example
	FileNum = FreeFile()
    
    Open historyPath For Input As #FileNum
    
    Do Until EOF(1)
        Line Input #FileNum, textline
        Text = Text & textline
    Loop

'custom open workbook
	Function openWb(myTitle)
				   
		res = "no workbook opened"
		wbPath = Application.GetOpenFilename(MultiSelect:=False, Title:=myTitle)
		
		If wbPath <> False Then
			Workbooks.Open (wbPath)
			If InStr(wbPath, "\") > 0 Then
				res = Right(wbPath, Len(wbPath) - InStrRev(wbPath, "\"))
			Else
				res = Right(wbPath, Len(wbPath) - InStrRev(wbPath, "/"))
			End If
		End If
	   
		openWb = res
		
	End Function

	'usage:
		myFileName = openWb("Open xxx file!")
			
		If myFileName = "no workbook opened" Then
			MsgBox myFileName
			Exit Sub
		End If

' Week Of Year - corrected version
	Function WOY (MyDate As Date) As Integer   

	  WOY = Format(MyDate, "ww", vbMonday, vbFirstFourDays)
	  
	  If WOY > 52 Then
		
		If Format(MyDate + 7, "ww", vbMonday, vbFirstFourDays) = 2 Then WOY = 1
		
	  End If
	  
	End Function

'monthname conversion
	Function monthNameToNumber(monthName)
		
		Select Case monthName
			Case "January":
				monthNameToNumber = 1
			Case "February":
				monthNameToNumber = 2
			Case "March":
				monthNameToNumber = 3
			Case "April":
				monthNameToNumber = 4
			Case "May":
				monthNameToNumber = 5
			Case "June":
				monthNameToNumber = 6
			Case "July":
				monthNameToNumber = 7
			Case "August":
				monthNameToNumber = 8
			Case "September":
				monthNameToNumber = 9
			Case "October":
				monthNameToNumber = 10
			Case "November":
				monthNameToNumber = 11
			Case "December":
				monthNameToNumber = 12
			
		End Select
		
	End Function

	Function monthNumberToName(monthNumber)
		
		Select Case monthNumber
		
			Case 1:
				monthNumberToName = "January"
			Case 2:
				monthNumberToName = "February"
			Case 3:
				monthNumberToName = "March"
			Case 4:
				monthNumberToName = "April"
			Case 5:
				monthNumberToName = "May"
			Case 6:
				monthNumberToName = "June"
			Case 7:
				monthNumberToName = "July"
			Case 8:
				monthNumberToName = "August"
			Case 9:
				monthNumberToName = "September"
			Case 10:
				monthNumberToName = "October"
			Case 11:
				monthNumberToName = "November"
			Case 12:
				monthNumberToName = "December"
		
		End Select
		
	End Function

'search in array
	Function isinArray(arr, item)

		res = -1
		
		For i = LBound(arr) To UBound(arr)
		
			If arr(i) = item Then
			
				res = i
				Exit For
				
			End If
			
		Next
		
		isinArray = res
		
	End Function

'search in collection
	Function isInCollection(mycoll, item)
		
		Dim res As Integer, i As Integer
		
		res = 0
		
		For i = 1 To mycoll.Count
		
			If mycoll(i) = item Then
			
				res = i
				Exit For
				
			End If
			
		Next
		
		isInCollection = res
		
	End Function

'convert collection to array
	Function colToArray(columnnumber, firstindex, lastindex)
			  
		tmp = ""
		
		For i = firstindex To lastindex
			
			tmp = tmp & Application.WorksheetFunction.Clean(Cells(i, columnnumber).Value) & ":;:"
			
		Next
		
		If Len(tmp) > 0 Then
		
			tmp = Left(tmp, Len(tmp) - 3)
			
		End If
		
		colToArray = Split(tmp, ":;:")
	   
	End Function

'awitch to directory if exists
	Private Sub chdirIfExists(pathToDir)

		If Dir(pathToDir, vbDirectory) <> "" Then
		
			ChDir (pathToDir)
			
		End If
		
	End Sub

'isSummertime
	Function isSummerTime(myDate)
	
		isSummerTime = False
		lastSundayofMarch = DateSerial(Year(myDate), 3, 31)
		
		Do While Weekday(lastSundayofMarch, vbMonday) <> 7
		
			lastSundayofMarch = DateAdd("d", -1, lastSundayofMarch)
			
		Loop
		
		lastSundayofOctober = DateSerial(Year(myDate), 10, 31)
		
		Do While Weekday(lastSundayofOctober, vbMonday) <> 7
		
			lastSundayofOctober = DateAdd("d", -1, lastSundayofOctober)
			
		Loop

		If CDbl(myDate) >= CDbl(lastSundayofMarch) And CDbl(myDate) <= CDbl(lastSundayofOctober) Then
		
			isSummerTime = True
			
		End If
		
	End Function

'remove filter if activated
	If ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode Then

		ActiveSheet.ShowAllData
		
	End If
			   
'schedule shutdown
	Function shutdown(inMin)
		
		Shell "shutdown -s -t " & inMin *60
			
	End Function

'sort array
	Function SortArrayAtoZ(myArray As Variant)

		Dim i As Long
		Dim j As Long
		Dim Temp

		'Sort the Array A-Z
		For i = LBound(myArray) To UBound(myArray) - 1
			For j = i + 1 To UBound(myArray)
				If UCase(myArray(i)) > UCase(myArray(j)) Then
					Temp = myArray(j)
					myArray(j) = myArray(i)
					myArray(i) = Temp
				End If
			Next j
		Next i

		SortArrayAtoZ = myArray

	End Function

'Regex
	using REGEX:

	a-z : lower
	0-5 : 0/1/2/4/5
	[a] : only "a"
	[abc] : only one from a/b/c
	[a]{2} : aa
	[a]{1,3}: a, aa, aaa
	a+ : a, aa, aaa, ...
	[a-z]? : only one lower or empty
	[a-z]* : only lower even multiple or empty
	. : anything but \n
	a|b : a/b 
	red|orange : red/orange
	^ : not e.g.: [^0-9] nut number
	a$ : last char is a

	abr    same as       meaning
	\d     [0-9]         Any single digit
	\D     [^0-9]        Any single character that's not a digit
	\w     [a-zA-Z0-9_]  Any word character
	\W     [^a-zA-Z0-9_] Any non-word character
	\s     [ \r\t\n\f]   Any space character
	\S     [^ \r\t\n\f]  Any non-space character
	\n     [\n]          New line

	Dim regEx As New RegExp
	Dim strPattern As String

	strPattern = "^[0-9]{1,3}"
	With regEx
		.Global = True
		.MultiLine = True
		.IgnoreCase = False
		.Pattern = strPattern
	End With
	 
	If regEx.test(strInput) Then
		simpleCellRegex = regEx.Replace(strInput, strReplace)
	Else
		simpleCellRegex = "Not matched"
	End If
	 
	pattern to replace: $1 :  strReplace = "$1" - ()

	String   Regex Pattern                  Explanation
	a1aaa    [a-zA-Z][0-9][a-zA-Z]{3}       Single alpha, single digit, three alpha characters
	a1aaa    [a-zA-Z]?[0-9][a-zA-Z]{3}      May or may not have preceeding alpha character
	a1aaa    [a-zA-Z][0-9][a-zA-Z]{0,3}     Single alpha, single digit, 0 to 3 alpha characters
	a1aaa    [a-zA-Z][0-9][a-zA-Z]*         Single alpha, single digit, followed by any number of alpha characters

	</i8>    \<\/[a-zA-Z][0-9]\>            Exact non-word character except any single alpha followed by any single digit

	"[0-9]{2}[.][0-9]{2}[.][0-9]{4}[ ][0-9]{2}[:][0-9]{2}[ ][-][ ][0-9]{2}[.][0-9]{2}[.][0-9]{4}[ ][0-9]{2}[:][0-9]{2}"


'copy code automatically 
	Sub injectVBACode(srcModulName, trgSheet)

		Dim CodeCopy As VBIDE.CodeModule
		Dim CodePaste As VBIDE.CodeModule
		Dim numLines As Integer
		trg = trgSheet.CodeName
		Set CodeCopy = ActiveWorkbook.VBProject.VBComponents(srcModulName).CodeModule
		Set CodePaste = ActiveWorkbook.VBProject.VBComponents(trg).CodeModule

		numLines = CodeCopy.CountOfLines
		'Use this line to erase all code that might already be in sheet3:
		If CodePaste.CountOfLines > 1 Then CodePaste.DeleteLines 1, CodePaste.CountOfLines

		CodePaste.AddFromString CodeCopy.Lines(1, numLines)
		
	End Sub

'check if a macro exists
	Function MacroExists(MyModuleName, MySub)
		
		MacroExists = "not found"
		
		Dim MyModule As Object

		Dim MyLine As Long

		On Error Resume Next

		Set MyModule = ActiveWorkbook.VBProject.VBComponents(MyModuleName).CodeModule
		
		If Err.Number <> 0 Then
			'MsgBox ("Module : " & MyModuleName & vbCr & "does not exist.")
			Exit Function
		End If

		MacroExists = MyModule.ProcStartLine(MySub, vbext_pk_Proc) & ";" & MyModule.ProcCountLines(MySub, vbext_pk_Proc)
		
	End Function

	Sub MacroExists2()
		Dim MyModule As Object
		Dim MyModuleName As String
		Dim MySub As String
		Dim MyLine As Long
		'---------------------------------------------------------------------------
		'- test data
		MyModuleName = "Sheet2"
		MySub = "Worksheet_BeforeRightClick"
		'----------------------------------------------------------------------------
		On Error Resume Next
		'- MODULE
		Set MyModule = ActiveWorkbook.VBProject.VBComponents(MyModuleName).CodeModule
		If Err.Number <> 0 Then
			MsgBox ("Module : " & MyModuleName & vbCr & "does not exist.")
			Exit Sub
		End If
		'-----------------------------------------------------------------------------
		'- SUBROUTINE
		'- find first line of subroutine (or error)
		MyLine = MyModule.ProcStartLine(MySub, vbext_pk_Proc)
		If Err.Number <> 0 Then
			MsgBox ("Module exists      : " & MyModuleName & vbCr _
				   & "Sub " & MySub & "( )  : does not exist.")
		Else
			MsgBox ("Module : " & MyModuleName & vbCr _
				& "Subroutine   : " & MySub & vbCr _
				& "Line Number : " & MyLine)
		End If
	End Sub

'quick sort
	Call QuickSort(myArray, 0, UBound(myArray))

	Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
	
		Dim pivot   As Variant
		Dim tmpSwap As Variant
		Dim tmpLow  As Long
		Dim tmpHi   As Long

		tmpLow = inLow
		tmpHi = inHi

		pivot = vArray((inLow + inHi) \ 2)

		While (tmpLow <= tmpHi)

		 While (vArray(tmpLow) < pivot And tmpLow < inHi)
		 
			tmpLow = tmpLow + 1
			
		 Wend

		 While (pivot < vArray(tmpHi) And tmpHi > inLow)
		 
			tmpHi = tmpHi - 1
			
		 Wend

		 If (tmpLow <= tmpHi) Then
		 
			tmpSwap = vArray(tmpLow)
			vArray(tmpLow) = vArray(tmpHi)
			vArray(tmpHi) = tmpSwap
			tmpLow = tmpLow + 1
			tmpHi = tmpHi - 1
			
		 End If
		 
		Wend

		If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
		If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
	  
	End Sub
	

'clipboard

	#If VBA7 Then

	Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
	Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
	Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
	Private Declare PtrSafe Function CloseClipboard Lib "User32" () As LongPtr
	Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
	Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
	Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
	Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
	Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As LongPtr
	Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

	#Else

	Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
	Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
	Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
	Private Declare Function CloseClipboard Lib "User32" () As Long
	Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
	Private Declare Function EmptyClipboard Lib "User32" () As Long
	Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
	Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
	Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long

	#End If

	Public Sub SetText(Text As String)

	#If VBA7 Then

	Dim hGlobalMemory As LongPtr
	Dim lpGlobalMemory As LongPtr
	Dim hClipMemory As LongPtr

	#Else

	Dim hGlobalMemory As Long
	Dim lpGlobalMemory As Long
	Dim hClipMemory As Long

	#End If

	Const GHND = &H42
	Const CF_TEXT = 1

	   ' Allocate moveable global memory.
	   '-------------------------------------------
	   hGlobalMemory = GlobalAlloc(GHND, Len(Text) + 1)

	   ' Lock the block to get a far pointer
	   ' to this memory.
	   lpGlobalMemory = GlobalLock(hGlobalMemory)

	   ' Copy the string to this global memory.
	   lpGlobalMemory = lstrcpy(lpGlobalMemory, Text)

	   ' Unlock the memory.
	   If GlobalUnlock(hGlobalMemory) <> 0 Then
		  MsgBox "Could not unlock memory location. Copy aborted."
		  GoTo CloseClipboard
	   End If

	   ' Open the Clipboard to copy data to.
	   If OpenClipboard(0&) = 0 Then
		  MsgBox "Could not open the Clipboard. Copy aborted."
		  Exit Sub
	   End If

	   ' Clear the Clipboard.
	   Call EmptyClipboard

	   ' Copy the data to the Clipboard.
	   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

	CloseClipboard:

	   If CloseClipboard() = 0 Then
		  MsgBox "Could not close Clipboard."
	   End If

	End Sub

	Public Property Get GetText()

	#If VBA7 Then

	Dim hClipMemory As LongPtr
	Dim lpClipMemory As LongPtr

	#Else

	Dim hClipMemory As Long
	Dim lpClipMemory As Long

	#End If

	Dim MaximumSize As Long
	Dim ClipText As String

	Const CF_TEXT = 1

	   If OpenClipboard(0&) = 0 Then
		  MsgBox "Cannot open Clipboard. Another app. may have it open"
		  Exit Property
	   End If

	   ' Obtain the handle to the global memory block that is referencing the text.
	   hClipMemory = GetClipboardData(CF_TEXT)
	   If IsNull(hClipMemory) Then
		  MsgBox "Could not allocate memory"
		  GoTo CloseClipboard
	   End If

	   ' Lock Clipboard memory so we can reference the actual data string.
	   lpClipMemory = GlobalLock(hClipMemory)

	   If Not IsNull(lpClipMemory) Then
		  MaximumSize = 64

		  Do
			MaximumSize = MaximumSize * 2

			ClipText = Space$(MaximumSize)
			Call lstrcpy(ClipText, lpClipMemory)
			Call GlobalUnlock(hClipMemory)

		  Loop Until ClipText Like "*" & vbNullChar & "*"

		  ' Peel off the null terminating character.
		  ClipText = Left$(ClipText, InStrRev(ClipText, vbNullChar) - 1)

	   Else
		  MsgBox "Could not lock memory to copy string from."
	   End If

	CloseClipboard:

	   Call CloseClipboard
	   GetText = ClipText

	End Property
	