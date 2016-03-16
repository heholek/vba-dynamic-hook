' Code related to ESET's VBA Dynamic Hook research
' For feedback or questions contact us at: github@eset.com
' https://github.com/eset/vba-dynamic-hook/
'
' This code is provided to the community under the two-clause BSD license as
' follows:
'
' Copyright (C) 2016 ESET
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice, this
' list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation
' and/or other materials provided with the distribution.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' Kacper Szurek <kacper.szurek@eset.com>
'
' Main script which runs `unprotect.py`, `parser.py` and `starter.py`, add `class.vba` content to file as another macro

Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetAbsolutePathName(".")  & "\"

input_file = fso.GetFile(WScript.Arguments.Item(0))
input_path = fso.GetAbsolutePathName(input_file)
extension = "." & fso.GetExtensionName(fullpath)

without_password_path = Replace(input_path, extension, "_without" & extension)
output_path = Replace(input_path, extension, "_output" & extension)

Dim shell_object : Set shell_object = CreateObject( "WScript.Shell" )
Dim python_object: Set python_object = shell_object.exec("python " & CurrentDirectory & "unprotect.py """ & input_path & """ " & """" & without_password_path & """" )


Wscript.Echo python_object.StdOut.ReadAll()

Set word_object = CreateObject("Word.Application")
' Dont display window
word_object.Visible = True
word_object.DisplayAlerts = False
' Disable macros
word_object.WordBasic.DisableAutoMacros 1

On Error Resume Next

Set word_document = word_object.Documents.Open(without_password_path)

If Err.Number <> 0 Then
	shell_object.exec("taskkill /f /im winword.exe")
	WScript.Echo "[-] Error: " & Err.Description
	Err.Clear
	WScript.Quit 1
End If

For Each VBComponentVar In word_document.VBProject.VBComponents
	with VBComponentVar.CodeModule
		Wscript.Echo "[+] Parsing " & .Name
		lines_count = .CountOfLines
		' Pass datas to python

		Set python_object = shell_object.exec("python " & CurrentDirectory & "parser.py")
		If lines_count > 1 Then
			python_object.StdIn.Write .Lines(1, lines_count)

			python_object.StdIn.Close()

			parsed_data = python_object.StdOut.ReadAll()
			parsed_data_array = Split(parsed_data, "|*&|VHOOK_SPLITTER|&*|")
			Wscript.Echo parsed_data_array(0)

			.DeleteLines 1, lines_count
			.InsertLines 1, parsed_data_array(1)
		Else
			Wscript.Echo "[-] Empty procedure"
		End if
	end with
Next

' Add reference to Microsoft Scripting Runtime for file write
word_document.VBProject.References.AddFromGUID "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0

' Add new module to existing file
Dim class_file, class_file_content
Set class_file = fso.OpenTextFile(CurrentDirectory & "class.vba")
class_file_content = class_file.ReadAll
class_file.Close

Set class_module = word_document.VBProject.VBComponents.Add(1)
class_module.Name = "vhook"
class_module.CodeModule.AddFromString class_file_content

word_document.SaveAs output_path
word_object.Quit 0

Wscript.Echo "[+] Open " & output_path

shell_object.Run "python " & CurrentDirectory & "starter.py """ & output_path & """", 0, True

Wscript.Echo "[+] END"
