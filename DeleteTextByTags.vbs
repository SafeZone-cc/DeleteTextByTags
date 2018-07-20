Option Explicit		'ver 1.4
On Error resume next

' Указать список расширений для обработки (через знак ;)
Dim Exts: Exts = "htm;html;txt"

' Файл со списком шаблонов в формате 
' Текст между этими словами будет удален, включая сами слова
Dim PattSrc: PattSrc = "Tags.txt"
' Сравнение без учета регистра букв? [true / false]
Dim IgnoreCase: IgnoreCase = false

Dim aExts: aExts = Split(Exts, ";")

Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oRegEx: Set oRegEx = CreateObject("VBScript.Regexp")
oRegEx.IgnoreCase = IgnoreCase
oRegEx.Global = true
oRegEx.Multiline = false
oRegEx.Pattern = """.*?"""

Dim cur: cur = oFSO.GetParentFolderName(WScript.ScriptFullName)

' Папка, в которой будет производится поиск = папке скрипта
Dim Folder: Folder = cur
PattSrc = oFSO.BuildPath(cur, PattSrc)

' Считываю c внешнего файла слова и создаю из них регулярки
Dim Patterns(): redim Patterns(0)
Dim s
Dim pos, i: i = 0
Dim direction: direction = false
Dim word1, word2, Encode
with oFSO.OpenTextFile(PattSrc, 1)
	Do Until .AtEndOfStream
		s = .ReadLine
		if len(s) <> 0 then ' не пустая строка
			direction = not direction	' true - прочитана 1-я строка (1-е слово)
			pos = instr(s, "=")
			if pos <> 0 then s = mid(s, pos + 1) ' урезаем знаки перед "="
			if direction then word1 = s else word2 = s
			if not direction then ' 2 слова прочитано
				For each Encode in Array("","utf-8")
					Patterns(i) = Reg2Escape(Recode(word1, Encode)) & "[\S\s]*?" & Reg2Escape(Recode(word2, Encode))
					i = i + 1
					redim preserve Patterns(i)
				next
			end if
		end if
	Loop
	.Close
end with
if i > 0 then redim preserve Patterns(i-1)
if len(Patterns(0)) = 0 then msgbox "Во внешнем файле нет информации для составления регулярного выражения!": WScript.Quit
if not oFSO.FolderExists(Folder) then msgbox "Папка " & Folder & " не существует!": WScript.Quit

Dim oRoot: Set oRoot = oFSO.GetFolder(Folder)
scanFolder oRoot
msgbox "Завершено."

Function Reg2Escape(byval str)
	Dim Char, NewLine, n
	' экран для спецсимволов регулярки
	For n = 1 to len(Str)
		Char = mid(Str, n, 1)
		if instr("\^$*+?{}.()|[]<>", Char) <> 0 then
			NewLine = NewLine & "\" & Char
		else
			NewLine = NewLine & "\u" & right("000" & hex(ascW(Char)), 4)
		end if
	Next
	Reg2Escape = NewLine
End Function

Sub scanFolder(oFolder)
    On Error Resume Next    
    Dim oFile, oSubfolder, fPath, content, contentNew, lLast, lNew

    If oFolder.Attributes AND &H600 Then Exit Sub 'проходим мимо симлинков
    
    For Each oFile In oFolder.Files
	  fPath = oFile.Path
	  '	если не этот скрипт и не файл-лог и совпадает с одним из списка заданных расширений
	  if StrComp(fPath, WScript.ScriptFullName, 1) <> 0 AND StrComp(fPath, PattSrc, 1) <> 0 AND IsValidExtension(oFSO.GetExtensionName(fPath)) then
		with oFile.OpenAsTextStream(1)
			content = .ReadAll()
			.Close
		end with
		contentNew = content
		For i = 0 to Ubound(Patterns)
			llast = len(contentNew)
			oRegEx.Pattern = "\r\n" & Patterns(i) & "\r\n"		' если останется пустая строка
			contentNew = oRegEx.Replace(contentNew, vbNewLine)
			lnew = len(contentNew)			
			if llast = lnew then
				oRegEx.Pattern = Patterns(i)
				contentNew = oRegEx.Replace(contentNew, vbnullstring)
			end if
		Next
		if len(contentNew) <> len(content) then	'если были изменения (проверка только по разнице объема содержимого)
			with oFile.OpenAsTextStream(2)
				.Write contentNew
				.Close
			end with
		end if
	  end if
    Next

    For Each oSubfolder In oFolder.Subfolders
        scanFolder oSubfolder 'рекурсия
    Next
End Sub

Function Recode(text, Codepage) ' перекодировка текста из ANSI -> в UTF-8
    If Codepage = "" Then Recode = text: Exit Function
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2     'text
        .Position = 0
        .Charset = "utf-8"
        .WriteText text
        .Flush
        .Position = 0
        .Type = 1     'binary
        .Read (3)     'skip BOM
        Recode = ByteArrayToString(.Read)
        .Close
    End With
End Function

Function ByteArrayToString(varByteArray)
    Dim rs: Set rs = CreateObject("ADODB.Recordset")
    rs.Fields.Append "temp", 201, LenB(varByteArray) 'adLongVarChar
    rs.Open: rs.AddNew: rs("temp").AppendChunk varByteArray: rs.Update
    ByteArrayToString = rs("temp"): rs.Close: Set rs = Nothing
End Function

Function IsValidExtension(Extension) ' проверка на совпадение найденного расширения со списком заданных
	Dim myExt
	IsValidExtension = false
	For each myExt in aExts
		if StrComp(Extension, myExt, 1) = 0 then IsValidExtension = true: Exit For
	Next
End Function