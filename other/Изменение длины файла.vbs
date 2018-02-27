' 1. узнать длину имени файла
' 2. если длина больше положенного, то
' 3. укорачивать или слева или справа (опция)
' 4. переименовать файл и получить TableName

Const InExt="dbf"


ff="65671366-001.dbf"
'Tablename="65671366-001"
LenFF=Len(FF) 'Длина имени'
TableName=Mid(LCase(FF),1,InStr(LCase(FF),"."&InExt)-1) 'Возвращает (строчные буквы, с 1-го символа, по символ с которого начинается искомая подстрока(строчные буквы, маска "." & dbf) и отнять 1)     то есть 65671366-001.dbf=65671366-001'
LenTB=Len(TableName) 'Длина номера накладной'
CheckLenTableName()
'msgbox(TableName)
'If LenTB > 8 Then
'	TableName1=Left(TableName, 8)
'	TableName2=Right(TableName, 8)
'	msgbox(TableName & chr(13) & TableName1 & chr(13) & Tablename2)
'End If



'TableName=Mid(TableName, 1, Instr(1,TableName,"-")-1) & Mid(TableName, Instr(1,TableName,"-")+1, LenTB-1) ' вернуть номер символа - и отнять 1  вернуть '
    
	
Function DelBigName(LR)
	If LR = "L" Then
		TableName=Left(TableName, 8)
	End If
	If LR = "R" Then
		TableName=Right(TableName, 8)
	End If
	FFNew=FDir & TableName & ".dbf"
  'msgbox(FF & "--" & FFNew)
  'FF.Move (FFNew) 
  'msgbox ("Необходимо уменьшить имя накладной до 8-ми символов перед '.dbg'")
  msgbox(TableName)
End Function

Function CheckLenTableName()
  If Len(TableName)>8 Then
    DelBigName("L")
  End If
  
  If Len(TableName)>8 Then
    msgbox ("Необходимо уменьшить имя накладной до 8-ми символов перед '.dbf'. Количество символов=" & Len(TableName) & Chr(13)&Chr(10) & "Будет крах, звонить 911 с корпоративного телефона")
  End If
End Function