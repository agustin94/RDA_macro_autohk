
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1

Run, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Aviso de deuda.xlsx
Sleep 3000
xl :=	ComObjActive("Excel.Application")
NumFila := 1
PathFileWord := "C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\"
Loop{
		;MsgBox %NumFila%
		NumFilaVerificacion = "A" NumFila
		;NumFilaVerificacion := "A" NumFila
		:MsgBox %NumFilaverificacion%
		;xl := ComObjCreate("Excel.Application")
			;NumFilaVerificacion := "A" NumFila
		NClienteVerificacion := xl.Range("A"NumFila).Value
		;MsgBox %NClienteVerificacion%

		If NClienteVerificacion =
		{ Sleep, 2000
			WinActivate, A
			Sleep, 333
			Loop
			{
				CoordMode, Pixel, Window
				ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\archivo excel.png
				If ErrorLevel = 0
					Click, %FoundX%, %FoundY% Left, 1
				
					Sleep, 1000
			SendRaw, {Down 9}
			SendRaw, {Enter}
			break
			}
			Until ErrorLevel = 0

			
		}




		If NClienteVerificacion = NroCliente  ; comparacion 
		{
			NumFila += 1
			MsgBox if %NumFila%
			MsgBox Esta mierda pasa por aca
			StringNroCliente := "A" NumFila
			StringNumero := "B" NumFila 
			StringEmail := "C" NumFila 
			StringMonto := "D" NumFila 
			StringVencimiento:="E" NumFila
			;String := "A" NumFila, "B" NumFila " ,C" NumFila " ,D" NumFila ", E" NumFila
			;MsgBox  %StringNroCliente% %StringNumero% %StringEmail% %StringMonto% %StringVencimiento%
			;MsgBox String 
			
		}

			StringNroCliente := "A" NumFila
			StringNombre := "B" NumFila 
			StringEmail := "C" NumFila 
			StringMonto := "D" NumFila 
			StringVencimiento:="E" NumFila
			;String := "A" NumFila, "B" NumFila " ,C" NumFila " ,D" NumFila ", E" NumFila
			;MsgBox  %StringNroCliente% %StringNumero% %StringEmail% %StringMonto% %StringVencimiento%
			;MsgBox String 


		;String := "A" NumFila, 
		;StrinB := "B" NumFila ;" ,C" NumFila " ,D" NumFila ", E" NumFila
		;MyArray := StrSplit(String, " ")
		;MyArray1 := StrSplit(StringB," ")
		;NroCliente := MyArray[1]
		;MsgBox %NroCliente%
		;Nombre := MyArray1[1]
		;MsgBox %Nombre%
		;Email := MyArray[3]
		;Monto := MyArray[4]
		;Vencimiento := MyArray[5]


		NroCliente := xl.Range(StringNroCliente).Value ;xl.Range(StringNroCliente).Value
		;MsgBox %NroCliente%
		SplitNCliente := StrSplit(NroCliente,".")
		ArrayCliente := SplitNCliente[1]
		;MsgBox %ArrayCliente%
		Nombre := xl.Range(StringNombre).Value ;xl.Range(StringNroCliente).Value
		;MsgBox %Nombre%
		Email := xl.Range(StringEmail).Value
		;MsgBox %Email%
		Monto := xl.Range(StringMonto).Copy
		;MsgBox %Monto%
		Vencimiento := xl.Range(StringVencimiento).Value
		;MsgBox %Vencimiento%
		Monto := xl.Range(StringMonto).Copy
		;MsgBox %Monto%
		;StringReplace, Clipboard, Clipboard, `r`n, `,%A_Space%, All
		;MsgBox %dato%

		Run, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Aviso de deuda.docx
		Sleep 3000
		Send, {Backspace}{Control Down}{l}{Control Up}

		SendRaw, {Nro de cliente}
		Sleep 1000
		Send, {Tab}
		Sleep 1000
		SendRaw, %ArrayCliente%

		Send, {Tab 2}
		Send, {Enter}
		Send, {Tab 6}
		Send, {Enter 2}
		;Send, {Escape}
		Sleep 1000

		SendRaw, {nombre}
		Send, {Tab}
		SendRaw, %Nombre%
		Send, {Tab 2}
		Send, {Enter}
		Send, {Tab 6}
		Send, {Enter 2}
		Sleep 1000
		SendRaw, {Fecha de vencimiento} 
		Send, {Tab}

		SendRaw, %Vencimiento%
		Send, {Tab 2}
		Send, {Enter}
		Send, {Tab 6}
		Send, {Enter 2}
		Sleep 1000

		SendRaw,{día} de {mes} de {año}
		send, {Tab}
		SendRaw, %A_DD% de %A_MMMM% del %A_YYYY%
		Send, {Tab 2}
		Send, {Enter}
		Send, {Tab 6}
		Send, {Enter 2}

		Monto := xl.Range(StringMonto).Copy
		;MsgBox %Monto%
		SendRaw, ${monto}
		Send, {Tab}
		Sleep 1000
		SendRaw, %Clipboard%
		Sleep 1000
		Send, {Enter}
		Send, {Tab 3}
		Send, {Enter}
		Send, {Tab 6}
		Send, {Enter 2}
		Send, {Escape}

		Send, {F12}
		Sleep 1000

		SendRaw, Carta documento_Cliente_%ArrayCliente%
		;Send, {Enter}
		Loop
		{
			CoordMode, Pixel, Window
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\flechapath.png
			If ErrorLevel = 0
				Click, %FoundX%, %FoundY% Left, 1
			If ErrorLevel
				Break
		}
		Until ErrorLevel = 0

		SendRaw, %PathFileWord%
		Send, {Enter}
		Sleep 1000
		Send, {Tab 14}
		Send, {Enter}
		Loop
		{
			CoordMode, Pixel, Window
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\archivo button.png
			If ErrorLevel = 0
				Click, %FoundX%, %FoundY% Left, 1
		}
		Until ErrorLevel = 0
		
		Send, {Down 8}
		Send,{Enter}

		;EMAIL

		Run, C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE
		Sleep 5000
		Send, {Control Down}{u}{Control Up}
		Sleep 2000
		SendRaw, %Email%
		Sleep 1000
		Send, {tab 3}
		Sleep 1000
		SendRaw, Prueba automatizada

		Send, {tab 2}
		Sleep 1000
		SendRaw, Email generado via AutoHotkey
		Sleep 1000
		Loop
		{
			CoordMode, Pixel, Window
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\adjuntar archivo.png
			If ErrorLevel = 0
				Click, %FoundX%, %FoundY% Left, 1
		}
		Until ErrorLevel = 0


		Send, {Down 14}
		Send, {Enter}

		Loop
		{
			CoordMode, Pixel, Window
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\flechapath.png
			If ErrorLevel = 0
				Click, %FoundX%, %FoundY% Left, 1
			If ErrorLevel
				Break
		}
		Until ErrorLevel = 0

		SendRaw, %PathFileWord%

		Send, {Enter}


		Loop
		{
			CoordMode, Pixel, Window
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\flechapath.png
			If ErrorLevel = 0
				Click, %FoundX%, %FoundY% Left, 1
			If ErrorLevel
				Break
		}
		Until ErrorLevel = 0

		SendRaw, Carta documento_Cliente_%ArrayCliente%.docx

		Send, {Enter}

		Sleep 1000
		Loop
		{
			CoordMode, Pixel, Window
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\enviar.png
			If ErrorLevel = 0
				Click, %FoundX%, %FoundY% Left, 1
			If ErrorLevel
				Break
		}
		Until ErrorLevel = 0
		Sleep 10000 
		Loop
			{
				CoordMode, Pixel, Window
				ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Agustin Moreno\Desktop\Proyectos The Eye\RDA-Pulover\Screenshot\archivo outlook.png
				If ErrorLevel = 0
					Click, %FoundX%, %FoundY% Left, 1
				If ErrorLevel
					Break
			}
			Until ErrorLevel = 0
			Sleep, 1000

		Send, {Down 7}
		Send {Enter}
		WinClose, Word
		
		NumFila += 1
		
  }