'Option Explicit
Dim FSO,C,FDir,FLD,FL,FF,Sh,FDirOut,dbfConn,dbfRS,LetLab,Letdate,LetNum,xlsFiles,xlsStrs,x1,x2,MonthYear,ar1
Dim xlglob,Desktop,Document,sheets,xlWbk
Dim Mass()
Dim aNoArgs()
Dim oMyStyle
Dim OApplication, ODesktop, ODocument, srcFile, srcProps() ' ******************
Dim args(0)

FDir="\\129.186.1.24\holdingswap\03 ЗАВЕДУЮЩИЕ\Отчёты Росздрав\"      ' Путь, где смотреть Документы с сайта в формате Excel
FDirOut="C:\braki\"   ' Путь куда выкладывать файл с отчетом
TemplateFile="C:\braki\ReportIMNA1.ots" ' Файл с шаблоном отчета

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

C=Chr(34) ' Двойные кавычки для строк

MonthYear=InputBox("Введите месяц. Пример: март","Введите месяц")

'********************************************************
'создаем новый ServiceManager
Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
Set oCalcDoc = oServiceManager.createInstance("com.sun.star.frame.Desktop")
' создаем новую книгу OpenOffice.org Calc
Set args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
args(0).Name = "AsTemplate"
args(0).Value = True
Set oBook = oCalcDoc.loadComponentFromURL("file:///"&TemplateFile, "_blank", 0, args)
'получаем ссылку на второй!!!!!!!!!!!!!!!! лист новой книги
Set oSheet = oBook.getSheets().getByIndex(1)
' т.е. чтобы получить ячейку в первом столбце первой строки пишем oSheet.getCellByPosition(0,0)
'кроме того в getCellByPosition первый аргумент столбец, второй строка (в Excel наоборот)
'***************************************************************************************************************************************

xlsFiles=0
xlsStrs=0
n=0
startcol=10
For Each FF in FL
    'msgbox FF.Name
    if (InStr(LCase(FF.Name),LCase(MonthYear) & " имн.xls")) then
        Set xlglob = CreateObject("com.sun.star.ServiceManager") 
        Set Desktop = xlglob.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.LoadComponentFromURL("file:///"&FDir&FF.Name, "_blank", 0, mass )
        Set sheets = Document.getSheets()
        Set xlWbk = sheets.getByIndex(0)
        k=0

        while xlWbk.getCellByPosition(1,k).String<>"Дата письма"
            k=k+1
        wend
        k=k+1
        while Len(Trim(xlWbk.getCellByPosition(3,k).String))>0       ' Пока содержимое первой ячейки текущей строки непустое, берем данные
'            if InStr(xlWbk.getCellByPosition(7,k).String,":")>0 then
'                LetLab = Mid(Replace(xlWbk.getCellByPosition(7,k).String,C,"'"),1,InStr(xlWbk.getCellByPosition(8,k).String,":")-1)
'            else
'                LetLab = ""
'            end if

            Call oSheet.getCellByPosition(0, startcol+n).SetFormula(1+n)
            Call oSheet.getCellByPosition(1, startcol+n).SetFormula(Mid(Replace("Письмо ФСН №" & xlWbk.getCellByPosition(1,k).String & " от " & xlWbk.getCellByPosition(2,k).String,C,"'"),1,200))
            Call oSheet.getCellByPosition(2, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(3,k).String,C,"'"),1,200))
            'серия Call oSheet.getCellByPosition(3, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(3,k).String,C,"'"),1,200))
            Call oSheet.getCellByPosition(4, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(6,k).String,C,"'"),1,200))
            Call oSheet.getCellByPosition(5, startcol+n).SetFormula("ИП Коростеленко М.Е.")
            Call oSheet.getCellByPosition(7, startcol+n).SetFormula("0")
            Call oSheet.getCellByPosition(8, startcol+n).SetFormula("Не выявлено")
            Call oSheet.Rows.insertByIndex(startcol+n+1, 1)

            xlsStrs = xlsStrs+1
            k=k+1
            n=n+1
        wend
        Document.Dispose()
        SET xlWbk = Nothing
        SET sheets = Nothing
        SET Document = Nothing
        SET Desktop = Nothing
        SET xlglob = Nothing
        '  FSO.DeleteFile FDir&FF.Name		' Удаляем обработанный файл
        xlsFiles = xlsFiles+1
    end if
Next

    ' Заполняем месяц
    'ar1=Split(MonthYear,".")
    'select case CInt(ar1(0))
    ' case 1 MonthYear=" январь"
    ' case 2 MonthYear=" февраль"
    ' case 3 MonthYear=" март"
    ' case 4 MonthYear=" апрель"
    ' case 5 MonthYear=" май"
    ' case 6 MonthYear=" июнь"
    ' case 7 MonthYear=" июль"
    ' case 8 MonthYear=" август"
    ' case 9 MonthYear=" сентябрь"
    ' case 10 MonthYear=" октябрь"
    ' case 11 MonthYear=" ноябрь"
    ' case 12 MonthYear=" декабрь"
    'end select
    
    'Call oSheet.getCellByPosition(6, 6).SetFormula(MonthYear)
    MonthYear=" "&LCase(MonthYear)&" месяц "
    'if Len(ar1(1))=2 then
    '  MonthYear=MonthYear&"20"
    'end if
    MonthYear=MonthYear&Year(Now)&" г."
    Call oSheet.getCellByPosition(6, 6).SetFormula(MonthYear)
    '-------------------------------------------------------------
'применение созданного выше стиля "osmorStyle" для форматирования
'диапазона ячеек "K1:L10"
'ссылку на диапазона получаем по имени методом getCellRangeByName
'-------------------------------------------------------------   
                Set oCells = oSheet.getCellRangeByName("A1:L111")
                'Set oMyStyle = oBook.createInstance("com.sun.star.style.CellStyle")
                'Call oBook.getStyleFamilies().getByName("CellStyles").insertByName("osmorStyle", oMyStyle)
            ' oMyStyle.CellBackColor = RGB(255, 220, 220) ' цвет фона
                'oMyStyle.IsCellBackgroundTransparent = False
            ' oMyStyle.CharColor = RGB(0, 0, 200) ' цвет  шрифта
                'oMyStyle.CharWeight = 150 ' толщина шрифта
                'Set oCells = oSheet.getCellRangeByName("A1:L111")
                'oCells.CellStyle = "osmorStyle" ' применяем стиль к выбранному диапазону
                oCells.IsTextWrapped = True ' Переносить по словам
                'Set oMyStyle = Nothing

if xlsStrs=0 then
        MsgBox "Обработано " & xlsFiles & " файлов!"
else
        MsgBox "Обработано " & xlsFiles & " файлов! Получено " & xlsStrs & " позиций!" & Chr(13) & "Отправьте файл в РосЗдрав!"    ' Вывод сообщения о завершении
end if  
