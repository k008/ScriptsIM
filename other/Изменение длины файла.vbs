' 1. ������ ����� ����� �����
' 2. ���� ����� ������ �����������, ��
' 3. ����������� ��� ����� ��� ������ (�����)
' 4. ������������� ���� � �������� TableName

Const InExt="dbf"


ff="65671366-001.dbf"
'Tablename="65671366-001"
LenFF=Len(FF) '����� �����'
TableName=Mid(LCase(FF),1,InStr(LCase(FF),"."&InExt)-1) '���������� (�������� �����, � 1-�� �������, �� ������ � �������� ���������� ������� ���������(�������� �����, ����� "." & dbf) � ������ 1)     �� ���� 65671366-001.dbf=65671366-001'
LenTB=Len(TableName) '����� ������ ���������'
CheckLenTableName()
'msgbox(TableName)
'If LenTB > 8 Then
'	TableName1=Left(TableName, 8)
'	TableName2=Right(TableName, 8)
'	msgbox(TableName & chr(13) & TableName1 & chr(13) & Tablename2)
'End If



'TableName=Mid(TableName, 1, Instr(1,TableName,"-")-1) & Mid(TableName, Instr(1,TableName,"-")+1, LenTB-1) ' ������� ����� ������� - � ������ 1  ������� '
    
	
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
  'msgbox ("���������� ��������� ��� ��������� �� 8-�� �������� ����� '.dbg'")
  msgbox(TableName)
End Function

Function CheckLenTableName()
  If Len(TableName)>8 Then
    DelBigName("L")
  End If
  
  If Len(TableName)>8 Then
    msgbox ("���������� ��������� ��� ��������� �� 8-�� �������� ����� '.dbf'. ���������� ��������=" & Len(TableName) & Chr(13)&Chr(10) & "����� ����, ������� 911 � �������������� ��������")
  End If
End Function