'
' Mihai Craiu, 20061002
' mihai.craiu@vodafone.ro
'

Set args = WScript.Arguments.Named

If (not args.Exists("w")) or (not args.Exists("c")) or args.Exists("h") Then
	WScript.Echo
	WScript.Echo "Usage: check_disk.vbs /w:INTEGER /c:INTEGER [/p] [/d:DRIVE_LIST | /x:DRIVE_LIST] [/u:UNITS] [/h]"
	WScript.Echo "	/w: warning limit"
	WScript.Echo "	/c: critical limit"
	WScript.Echo "	/p: limits in %, otherwise in UNITS"
	WScript.Echo "	/d: included drives list"
	WScript.Echo "	/x: excluded drives list"
	WScript.Echo "	/u: B | kB | MB | GB, default MB"
	WScript.Echo "	/h: this help"
	WScript.Echo
	WScript.Echo "	check_disk.vbs /w:15 /c:5 /p /d:CDE /u:kB - result will be displayed in kB, limits are in percents"
	WScript.Echo "	check_disk.vbs /w:500 /c:250 - result will be displayed in MB, limits are in MB, all fixed drives"
	WScript.Echo
	WScript.Quit 3
End If

If args.Exists("u") Then
	u=args.Item("u")
	If u<>"B" and u<>"kB" and u<>"MB" and u<>"GB" Then
		WScript.Echo
		WScript.Echo "Units must be one of B, kB, MB, GB"
		WScript.Echo
		WScript.Quit 3	
	End If
Else
	u="MB"
End If

Select Case u
	Case "B"
		uLabel=""
		uVal=1
	Case "kB"
		uLabel="kB"
		uVal=1024
	Case "MB"
		uLabel="MB"
		uVal=1024*1024
	Case "GB"
		uLabel="GB"
		uVal=1024*1024*1024
End Select

w=1*args.Item("w")
c=1*args.Item("c")
p=args.Exists("p")

If w<c Then
	WScript.Echo
	WScript.Echo "Warning limit must be greater than critical limit"
	WScript.Echo
	WScript.Quit 3
End If

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set colDrives=objFSO.Drives

outCode=0
outText=" - espacio libre:"
outData="| "
'sizeLimit es el tamaño en bytes para el cual la comprobación se hace por porcentaje y no por cantidad absoluta.
sizeLimitGB = 6
sizeLimit = sizeLimitGB * 1024 * 1024 * 1024
For Each objDrive in colDrives
	If DriveSelected(objDrive) Then
		If objDrive.IsReady Then
			disk=objDrive.DriveLetter
			freeSpace=objDrive.FreeSpace
			size=objDrive.TotalSize
			busySpace=size-freeSpace
	'		WScript.Echo "Espacio libre: " & freeSpace
	'		WScript.Echo "Espacio total: " & size
	
			'Trabajamos con wtmp y ctmp como variables temporales de w y c
			wtmp=w
			ctmp=c
			
			'Comprobamos que el tamaño de la unidad sea < sizeLimit
			If size < sizeLimit Then
				wtmp=10	'Warning será 10% de la unidad
				ctmp=5	'Critical será 5% de la unidad
				p3=1	'p3 se usa para forzar comprobación en porcentaje cuando tamaño de unidad es menor que sizeLimit
			End If
			
			If outCode=0 Then
					If p Or p3 Then
						If 100*freeSpace/size<wtmp Then
							outCode=1
						End If 
					Else
						If freeSpace/uVal<wtmp Then
							outCode=1
						End If
					End If
			End If
			
			If outCode=1 Then
					If p Or p3 Then
						If 100*freeSpace/size<ctmp Then
							outCode=2
						End If 
					Else
						If freeSpace/uVal<ctmp Then
							outCode=2
						End If
					End If
			End If
			
			'Restablecemos p3=0
			p3=0
			
			If freeSpace>2147483648 Then
				espacioLibre = Round(freeSpace/uVal) & "GB"
			Else
				espacioLibre = Round(1024*freeSpace/uVal) & "MB"
			End If
			outText=outText & " " & disk & ":" & espacioLibre & "/" & Round(size/uVal) & uLabel & " (" & Round(100*freeSpace/size) & "%);"
			outData=outData & "disco" & disk & "=" & Round(busySpace/1048576) & "MB;" 'Se divide entre 1048576 para pasar B a MB
				outData=outData & Round(size*90/104857600) & ";" 'Warning:  90% del tamaño total 'size'
				outData=outData & Round(size*95/104857600) & ";" 'Critical: 95% del tamaño total 'size'
				outData=outData & "0;" & Round(size/1048576) & ";"
		End If
	End If
Next

Select Case outCode
	Case 0
		outText="DISK OK" & outText
	Case 1
		outText="DISK WARNING" & outText
	Case 2
		outText="DISK CRITICAL" & outText
End Select

WScript.Echo outText & outData
WScript.Quit outCode

'Función auxiliar que devuelve si la unidad está seleccionada
'Unidades seleccionadas serían 
Function DriveSelected(d)
	'Inicializamos a false
	DriveSelected=false
	'Si hemos pasado argumentos "d" (lista de unidades a incluir)
	If args.Exists("d") Then
		'Comprueba si la letra de la unidad pasada como argumento
		'está contenida en la lista de unidades incluídas
		If InStr(UCase(args.Item("d")),d.DriveLetter)>0 Then
			DriveSelected=true	'Si está contenida
		End If
	'O si hemos pasado argumentos "x" (lista de unidades a excluir)
	ElseIf args.Exists("x") Then
		'Comprueba si la letra de la unidad pasada como argumento
		'está contenida en la lista de unidades excluídas
		If InStr(UCase(args.Item("x")),d.DriveLetter)>0 Then
			DriveSelected=false	'Si está contenida
		ElseIf d.DriveType=2 Then
			DriveSelected=true	'Si no está contenida y es tipo 2
		End If
	'O si la unidad es de tipo 2 (Fixed drive (hard disk))
	ElseIf d.DriveType=2 Then
		DriveSelected=true	'Si es tipo 2, devolvemos true
	End If
End Function
