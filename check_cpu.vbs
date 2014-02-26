Option Explicit
Dim intPasses, intPause, objWMIService, CPUInfo, i, Item, sum, w, c
Dim outCode, outText, outData, cpu

'Parámetros generales
intPasses = 3	 'Num Checks
intPause  = 3000 'Milliseconds
sum       = 0    'Sumador
outCode   = 0
outText   = " - espacio libre:"
outData   = "| "

'Comprobamos parámetros de entrada (warning y critical)
If Wscript.Arguments.Count <> 2 Then
	'Wscript.Echo "Uso: " & Wscript.ScriptFullName & " <warning>% <critical>%"
	Wscript.Echo "Uso: " & Wscript.ScriptName & " <warning>% <critical>%"
	Wscript.Quit 3
Else
	w = Int(Wscript.Arguments(0))
	c = Int(Wscript.Arguments(1))
	If Not (isNumeric(w) And isNumeric(c) And (0 < w) And (w < c) And (c < 100)) Then
		Wscript.Echo "Los valores son incorrectos. 0 < w (" & w & ") < c (" & c & ") < 100"
		Wscript.Quit 3
	End If
End If



Set objWMIService = GetObject("winmgmts:\\localhost\root\CIMV2")

'Loop Num Checks
For i = 1 to intPasses
	Set CPUInfo = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = '_Total'",,48) 
	'Wscript.Echo vbCrLf & "Pass " & i
	For Each Item in CPUInfo 
		'Wscript.Echo "Caption: " & Item.Caption
		'Wscript.Echo "Description: " & Item.Description
		'Wscript.Echo "Name: " & Item.Name
		'Wscript.Echo "PercentUserTime: " & Item.PercentUserTime
		'Wscript.Echo "PercentProcessorTime: " & Item.PercentProcessorTime
		Sum = Sum + Item.PercentUserTime
	Next
	If i < intPasses Then
		Wscript.Sleep intPause
	End If
Next

cpu = Round(Sum/intPasses)

outText = " - CPU Usada: " & cpu & "% "
outData=outData & "CPU=" & cpu & "%;" & w & ";" & c & ";0;100;"

'Comprobamos el outCode
If cpu > c Then
	outCode = 2	'Critical
ElseIf cpu > w Then
	outCode = 1	'Warning
End If

Select Case outCode
	Case 0
		outText="CPU OK" & outText			'OK
	Case 1
		outText="CPU WARNING" & outText		'WARNING
	Case 2
		outText="CPU CRITICAL" & outText	'CRITICAL
End Select

WScript.Echo outText & outData
WScript.Quit outCode