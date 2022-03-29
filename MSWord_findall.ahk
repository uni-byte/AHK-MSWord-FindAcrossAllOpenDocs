FindInAllOpenDocuments() {
global
Try
	oWord := ComObjActive("Word.Application")
Catch err
	MsgBox, Exception thrown:`n%err%

Search := oWord.Selection.Text
Clipboard := Search
Array := []
loop % oWord.Documents.Count {
	Doc 	:= oWord.Documents(A_Index)
	Name	:= Doc.fullname
	Text 	:= Doc.Range.Text
	If InStr(Text, Search , 1)  {
		Array.Push(Name)
	}	
}

Gui, New, ,%Search%
For key, val in Array
	Gui, Add, Button, Section xm w400 h100 vKno%key% gBtnActive, %val%
Gui, Show
return

BtnActive:
	{
	SetTitleMatchMode, 2
	RegExMatch(A_GuiControl, "\d+$", n) ; Gets button's variable (Kno1, Kno2, etc...), extracts integer and assigns it to 'n'
	WordFD_Path := StrSplit(Array[n], ["/", "\"])
	WordFD_Path := WordFD_Path[WordFD_Path.Count()]
	if WinExist(WordFD_Path)
		WinActivate
	else
		MsgBox, Window not found
	oWord.Documents(WordFD_Path).Activate
	myFindObject := oWord.Selection.Find
	myFindObject.ClearFormatting
	myFindObject.Replacement.ClearFormatting
	myFindObject.Execute(Search, 1, 0, 0, 0, 0, 1, 1, 0, "", 0, 0, 0, 0, 0)
	
	return
	}

return
}
