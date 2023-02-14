'常量定义
Public Const LOOPLIMIT = 1000000000'循环上限
'注：
'   当前工作簿路径为 ThisWorkbook.Path
'   当前工作簿名称为 ThisWorkbook.Name
'   Windows系统下的路径分割符      "\"

'判断字符串是否是路径
Public Function IsPath(Byval filepath As String) As Boolean
    '如果接收的参数是字符串且字符串包含路径字符，返回真
    IsPath = (VarType(filepath) = vbString) And InStr(filepath,"\")
End Function

'判断字符串是否是含扩展名文件名
Public Function IsFileName(Byval filename As String) As Boolean
    '如果接收的参数是字符串且字符串包含"."字符，返回真
    IsFileName = (VarType(filename) = vbString) And InStr(filename,".")
End Function

'从全路径名中将文件名取出
Public Function GetFilenameFromPath(ByVal file_path As String)As String
    '参数       包含文件名的全路径
    '返回值     剔除路径后的文件名
    If Not IsPath(file_path) Then'路径验证
        GetFilenameFromPath = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim varSplitArr As Variant                      '定义路径分割数组
    'Split函数将字符串以指定字符分割，返回一个字符串数组
    varSplitArr = Split(file_path, "\")              '分割路径
    'UBOUND函数，返回数字的最后一个元素的下标
    GetFilenameFromPath=varSplitArr(UBOUND(varSplitArr))'返回文件名
End Function

'从文件名中获取扩展名
Public Function GetExtensionFromFilename(ByVal file_name As String)As String
    '参数       文件名
    '返回值     扩展名（含“.”）
    If  Not IsFileName(file_name) Then'文件名验证
        GetExtensionFromFilename = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim varSplitArr As Variant                      '定义路径分割数组
    'Split函数将字符串以指定字符分割，返回一个字符串数字
    varSplitArr = Split(file_name, ".")              '分割路径
    'UBOUND函数，返回数字的最后一个元素的下标
    GetExtensionFromFilename="." & varSplitArr(UBOUND(varSplitArr))'返回扩展名
ENd Function

'从全路径名中将文件名剔除，仅返回路径
Public Function ExceptFilenameFromPath(ByVal file_path As String)As String
    '参数       包含文件名的全路径
    '返回值     路径字符串（“含“\”）
    If Not IsPath(file_path) Then'路径验证
        ExceptFilenameFromPath = ""'返回空字符串
        Exit Function '退出函数
    End If

    '获取移除文件名后的字符串的左侧部分
    'Left函数从字符串左边起，截取指定长度的字符串
    Dim strFilename As String
    strFilename=GetFilenameFromPath(file_path)'获取文件名
    ExceptFilenameFromPath=Left(file_path,Len(file_path)-Len(strFilename))
End Function

'从全文件名中将扩展名剔除，仅返回文件名
Public Function ExceptExtensionFromFilename(ByVal filename As String)As String
    '参数       全文件名
    '返回值     无扩展名的文件名
    If Not IsFileName(filename) Then'路径验证
        ExceptExtensionFromFilename = ""'返回空字符串
        Exit Function '退出函数
    End If

    '获取移除文件名后的字符串的左侧部分
    'Left函数从字符串左边起，截取指定长度的字符串
    Dim strExtension As String
    strExtension=GetExtensionFromFilename(Filename)
    ExceptExtensionFromFilename=Left(Filename,Len(Filename)-Len(strExtension))
End Function

'如果路径尾不是'\',则添加上'\'
Public Function SetPathDelimit(ByVal FloderPath As string)As String
    '参数       不包含文件名的路径字符串
    '返回值     结尾是“\”的路径字符串
    If VarType(FloderPath) = vbString Then'如果参数是字符串
        'RIgh函数从字符串右边起，截取指定长度的字符串
        If Right(FloderPath,1) <> "\" Then
            FloderPath=FloderPath & "\"
        End If
    Else
        SetPathDelimit = None'参数错误
        Exit Function'退出函数
    End If
    SetPathDelimit = FloderPath
End Function

'更改全路径中的文件名
Public Function ReplaceFilename(ByVal file_path As String, ByVal filename As String)As String
    '参数       参数1：全路径字符串      参数2：新文件名
    '返回值     替换文件名后的全路径
    If Not IsPath(file_path) Then'路径验证
        ReplaceFilename = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim strPath As String
    strPath=ExceptFilenameFromPath(file_path)        '获取纯路径
    strPath=SetPathDelimit(strPath)             '保证路径结尾是"\"
    ReplaceFilename=strPath & filename
End Function

'替换文件名的扩展名（后缀）
Public Function ReplaceExtension(ByVal filename As String, ByVal extension As String)As String
    '参数       参数1：带扩展名的文件名      参数2：扩展名
    '返回值     替换扩展名后的文件名
    If Not (IsFileName(filename) And (VarType(extension) = vbString)) Then'路径验证
        If Len(extension) = 0 Then'后缀名验证
            ReplaceExtension = ""'返回空字符串
            Exit Function '退出函数
        End If
    End If

    Dim strFilename As String
    strFilename=ExceptExtensionFromFilename(Filename)   '去除原有扩展名
    If CutLastElem(strFilename) = "." Then      '如果后缀包含了"."
        ReplaceExtension=strFilename & extension
    Else                                            '如果后缀没有包含"."
        ReplaceExtension=strFileName &"."& extension
    End If
End Function

'在接收的文件目录下创建一个文件副本
Public Function CreateFileCopy(ByVal file_path As String)As String
    '参数       包含文件名的全路径
    '返回值     包含文件副本名的全路径
    If Not IsPath(file_path) Then'路径验证
        CreateFileCopy = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim strFilename As String
    strFilename = GetFilenameFromPath(file_path)             '获取写入文件名
    Dim strExPath As String
    strExPath = ExceptFilenameFromPath(file_path)            '获取纯路径
    strExPath=SetPathDelimit(strExPath)                     '保证路径正确
    Dim strExFilename As String
    strExFilename = ExceptExtensionFromFilename(strFilename)'获取无后缀文件名
    Dim strExtension As String
    strExtension = GetExtensionFromFilename(strFilename)    '获取文件后缀

    CreateFileCopy=strExPath & strExFilename & "-Copy" & strExtension'构造副本名
End Function

'接收一个字符串返回字符串的第一个字符
Public Function CutFirstElem(ByVal TargetString As String) As String
    If Not VarType(TargetString) = vbString Then'
        CutFirstElem = ""'返回空字符串
        Exit Function '退出函数
    End If
    Dim strOneElem As String*1  '定义一个仅接受一个字符的字符串
    strOneElem=TargetString     '获取第一个字符
    CutFirstElem=strOneElem     '返回获取的字符
End Function

'接收一个字符串返回字符串的最后一个字符
Public Function CutLastElem(ByVal TargetString As String)As String
    If Not VarType(TargetString) = vbString Then'
        CutLastElem = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim i As Integer                            '声明遍历索引值
    i=Len(TargetString)                         '获取字符串长度
    CutLastElem = Mid(TargetString,i,1)         '获取字符串数组
End Function

'接收一个字符串，剔除字符串的最后一个字符，返回剩余字符
Public Function RemoveLastElem(ByVal TargetString As String)As String
    If Not VarType(TargetString) = vbString Then'
        RemoveLastElem = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim i As Integer                            '声明遍历索引值
    i=Len(TargetString)                         '获取字符串长度
    RemoveLastElem = Left(TargetString,i - 1)       '获取字符串数组
End Function

'接收一个字符串返回字符串的任意位置的单个字符
Public Function CutAnyElem(ByVal TargetString As String,ByVal Pos As String )As String
    If Not VarType(TargetString) = vbString Then'
        CutAnyElem = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim strOneElem As String*1              '定义一个仅接受一个字符的字符串
    strOneElem=Mid(TargetString,Pos,1)      '获取Pos位置处的字符
    CutAnyElem=strOneElem                 '返回获取的字符
End Function

'查看有无指定参数名称指定的文件
Public Function IsFileExist(strFname As String) As Boolean
    '#
    '#  参数：      strFname    想要打开的文件的全路径名。
    '#
    '#  返回值：    True：  文件存在。
    '#             False:  文件不存在。
    '#
    If Not (IsPath(strFname) And IsFileName(strFname)) Then'如果字符串包含"\"和"."
        IsFileExist = False'返回假
        Exit Function '退出函数
    End If

    Dim bFileExist As Boolean       '文件有无
    Dim strFileName As String       '文件名
    Dim varFilePos As Variant       '文件名所在字符串位置
    
    
    bFileExist = False       '初始默认不存在文件
    strFileName = Dir(strFname, vbNormal)

                'Dir函数返回一个文件夹下的一个文件的名字（包含后缀）
                '参数：vbNormal       值：0   指定无属性的文件
                '参数：vbReadOnly     值：1   指定无属性的只读文件
                '参数：vbHidden       值：2   指定无属性的隐藏文件
                '参数：vbSystem       值：4   指定无属性的系统文件
                '参数：vbvolume       值：8   指定卷标文件
                '参数：vbDirectory    值：16  指定无属性文件及其路径
                '参数：vbAlisa        值：64  指定的文件名是别名
    If strFileName <> "" Then'如果文件名不为空
        '获取文件名所在字符串中的位置
        varFilePos = InStr(1, strFname, strFileName, vbTextCompare)
                    'InStr函数获取被替换字符串首字符在源字符串中的位置
                    'InStr(int,string,string,int)
                    '参数：vbTextCompare    值:1    执行文本比较
                    '参数：vbBinaryCompare  值:0    执行二进制比较
                    '参数：起始查找位置，源字符串，被查找字符串，第几次出现的被查找字符串
                    '返回值:未找到：  0
                    '       找到：    被查找字符串首字符所在位置
        If Not IsNull(varFilePos) Then'如果存在位置
            bFileExist = True'文件存在设置为真
        End If
    End If
    
    IsFileExist = bFileExist'返回文件是否存在
    
End Function

'接收一个路径和单元格位置，提取单元格内容并返回
Public Function GetDetailFromCell(ByVal FullPath As String, ByVal SheetIndex As String, ByVal CellPos As String) As String
    If Not (IsPath(FullPath) And IsFileName(FullPath)) Then'如果字符串包含"\"和"."
        GetDetailFromCell = False'返回假
        Exit Function '退出函数
    End If

    If IsFileExist(FullPath) Then
        Set WriteWorkBook = GetObject(FullPath)
        GetDetailFromCell=WriteWorkbook.Sheets(SheetIndex).Range(CellPos).Value
    Else
        Msgbox "There is nothing!"
    End If
End Function

'在字符串末尾加点（如果末尾已经有点了就不添加点）
Public Function AddDotAtLast(Byval strData As String) As String
    If Not VarType(strData) = vbString Then'
        AddDotAtLast = ""'返回空字符串
        Exit Function '退出函数
    End If
    
    '判断末尾是否已经存在"."
    If CutLastElem(strData) <> "." Then '如果不存在
        strData = strData & "."
    End If
    AddDotAtLast = strData
End Function

'打开一个文件路径并返回
Public Function SelectPath() As String
    ' 这个过程用于不同文件路径的选择并输出到不同的单元格内
    ' 打开excel表并将路径输入到单元格中
    Set fileDlg = Application.FileDialog(msoFileDialogFilePicker)
    ' 
    Dim path As String
    With fileDlg
        If .Show = -1 Then
            For Each fld In .SelectedItems
                path = fld
            Next fld
        End If
    End With
    '将路径写入单元格
    SelectPath = path
End Function

'接收一个路径在路径下的指定txt文件内写入数据
'参数 1.写入路径  2.写入数据
Public Function WriteTXT(ByVal strpath As String,ByVal strData As String) As Boolean
    If Not (IsPath(strpath) And VarType(strData) = vbString) Then'
        AddDotAtLast = ""'返回空字符串
        Exit Function '退出函数
    End If

    Open strpath For Append As #1'打开txt文件，如果不存在则创建它
    Print #1, strData'写入数据
    Close #1'关闭TXT
End Function

'打开一个文件路径并存储到指定TXT文件中
Public Function SelectFilePathAndStoreIt(ByVal cstPath As String,Byval strMark As String) As Boolean
    '参数 1.txt文件路径，2.存储路径的前缀
    If Not (IsPath(cstPath) And IsFileName(cstPath) And VarType(strMark) = vbString) Then'
        AddDotAtLast = ""'返回空字符串
        Exit Function '退出函数
    End If

    Dim strpath$'存储路径字符串
    strpath = SelectPath()'打开文件路径
    SelectFilePathAndStoreIt = WriteTXT(cstPath,strMark & strpath)'存储打开的文件路径
End Function


'加密函数：后期追加，暂且只能返回原字符串 At 2021-03-22 By刘加龙
Public Function EncipherLite(ByVal HideString As String) As String
    Dim Enciphered$             '定义字符串变量
    Enciphered = HideString     '获取加密后的字符串
    EncipherLite = Enciphered   '返回加密后的字符串
End Function

'解密函数：后期追加，暂且只能返回原字符串 At 2021-03-22 By刘加龙
Public Function DecipherLite(ByVal HideString As String) As String
    Dim Deciphered$             '定义字符串变量
    Deciphered = HideString     '获取解密后的字符串
    DecipherLite = Deciphered   '返回解密后的字符串
End Function


'密码确认函数 后期还需要修缮，要能自定义密码，要能忘记密码，密钥解密
Public Function CipherVerify() As Boolean

    Dim strPassword$ '定义密码字符串
    strPassword = EncipherLite("141421") '密码加密
    Dim strGetPassword '定义获取密码字符串

    Dim bPass As Boolean '验证密码是否正确
    Dim i% '遍历字符串
    For i = ZEROLITE To 5 '允许5次输入密码
        If i > 5 Then
            '文件毁掉程序
            CipherVerify = False
            Exit For
        Else
            If bPass Then
                CipherVerify=bPass '返回值只允许是真，如果5次是假，毁掉程序
                Exit For '如果密码正确，退出程序
            End If

            strGetPassWord = InputBox("Please Enter The Passward") '请求输入密码
            If EncipherLite(strGetPassWord) <> strPassword Then '如果密码不通过
                MsgBox "Only : " & 5-i & " Times" '警告密码错误，并提醒剩余次数
                bPass=False      '返回失败值
            Else
                bPass=True       '返回真值
            End If
        End If

    Next
    
End Function

' 获取密码，展开数据
Public Sub RowExpand()
    
    '密码确认
    If Not CipherVerify() Then'如果密码不通过退出程序
        ActiveWorkbook.Save '自动保存
        WorkBooks(ThisWorkbook.Name).Close '关闭文件
    End If

    Set MyThisSheet = Thisworkbook.sheets(1) 'sheet(1)

    Dim iRowHead As Byte '起始行
    iRowHead = 1 '起始行=1
    Dim iRowTail% '部品总行
    iRowTail = MyThisSheet.UsedRange.Rows.Count '获取部品选用表总行
    Dim iRowIndex% '行遍历
    iRowIndex = iRowHead '起始行开始遍历

    Rows(CStr(iRowIndex) & ":" & CStr(iRowTail)).RowHeight = 13.8 '展开这一行

End Sub

' 隐藏表格
Public Sub RowHide()

    Set MyThisSheet = Thisworkbook.sheets(1) '此表

    Dim iRowHead As Byte '起始行
    iRowHead = 1 '起始行=1
    Dim iRowTail% '部品总行
    iRowTail = MyThisSheet.UsedRange.Rows.Count '获取部品选用表总行
    Dim iRowIndex% '行遍历
    iRowIndex = iRowHead '起始行开始遍历

    Rows(CStr(iRowIndex) & ":" & CStr(iRowTail)).RowHeight = 0'隐藏数据行

    ActiveWorkbook.Save 'Autosave
End Sub

' 10进制转2进制
Public Function Convert10to2(Value As Long) As String

    Dim lngBit As Long
    Dim strData As String

    Do Until (Value < 2 ^ lngBit)
        If (Value And 2 ^ lngBit) <> 0 Then
            strData = "1" & strData
        Else
            strData = "0" & strData
        End If

        lngBit = lngBit + 1
    Loop

    Convert10to2 = strData

End Function

'将Excel表指定单个单元格中的数据输出为txt文件
Public Function ExportCellToTXTFile(exportpath, filename, workbook, sheet, _ 
    Optional Byval row As Integer = 1, Optional Byval column As Integer = 1, _ 
    Optional Byval writemode As String = "Output",Optional Byval prefix As String = "", _ 
    Optional Byval suffix As String = "", Optional Byval head As String = "", _ 
    Optional Byval tail As String = "")
    ' 参数：导出路径，导出文件名，写入模式(覆盖源文件，追加数据)，读取工作簿，读取数据页，
    ' 读取数据行，读取数据列，数据前缀，数据后缀，文件头行，文件尾行

    blnParaCheck = False'参数验证
    blnExported = False '数据导出成功标志位

    '参数验证
    '路径参数验证 字符串参数验证 sheet参数验证
    If IsPath(exportpath) And (VarType(filename) = vbString) And (VarType(workbook) = vbString) _ 
    And ((VarType(sheet) = vbString) Or (VarType(sheet) = vbInteger)) Then
        blnParaCheck = True'参数验证通过
    End If
    
    If blnParaCheck Then'如果参数验证通过
        If writemode = "Append" Then
            Open SetPathDelimit(exportpath) & filename For Append As #1 '打开或者创建txt文件并追加数据
        ElseIf writemode = "Output" Then
            Open SetPathDelimit(exportpath) & filename For Output As #1 '打开或者创建txt文件并写入数据
        Else:
            blnExported = False '将导出标志位设为假
            ExportCellToTXTFile = blnExported '返回导出失败标志
            Exit Function '退出程序
        End If
        
        If Not prefix = "" Then '如果有前缀
            Print #1, head '输出头数据
        End If

        Print #1, prefix & Workbooks(workbook).Sheets(sheet).Cells(row, column).Value & suffix

        If Not suffix = "" Then '如果有后缀
            Print #1, tail '输出尾数据
        End If
        
        Close #1
        blnExported = True'将导出标志位设置为真
    End If

    ExportCellToTXTFile = blnExported '返回导出成功标志

End Function

'将一列Excel数据同名的数据合并，数量相加，名称数量行数必须相同且对齐
'可能存在缺陷
Public Function MergeNameSumQuan(workbook,sheet,namerow,namecol,numcol) As Boolean
    '首先获取要统计的名称的列
    '然后找到数量的列
    '记录第一个名称，获取所在单元格，获取数量所在单元格
    '向下查找
    '查找到同名，获取数据，加到最初的单元格，隐藏此行
    '遍历一遍后寻找下一个，依次类推

    '参数：工作簿名，sheet,名称行，名称列，数量列

    '合并成功标志位
    blnMerged = False'默认为假

    '参数验证
    If IsFileName(workbook) And (VarType(namerow) = vbInteger) And _ 
    (VarType(namecol) = vbInteger) And (VarType(numcol) = vbInteger) Then'验证通过

        intRowBegin = namerow'获取起始行
        intNameCol = namecol'获取名称列
        intNumCol = numcol'获取数量列
        Set sht = Workbooks(workbook).Sheets(sheet)'获取需要操作的sheet
        
        for s = intRowBegin to LOOPLIMIT'循环上限
            If sht.Cells(s,intNameCol).Value = "" Then'如果出现空单元格
                intRowCount = s - 1'获取上一行行号
                Exit For
            End If
            If Not IsNumeric(sht.Cells(s,intNumCol).Value) Then'如果检测到了数量是非数字项目
                MergeNameSumQuan = blnMerged
                Exit Function'退出程序
            End If
        Next

        for i = intRowBegin To intRowCount'遍历名称

            If sht.Rows(i).EntireRow.Hidden = False Then'如果这一行没有被隐藏

                If sht.Cells(i,intNameCol).Value = "" Then'如果查询到了空数据
                    Exit For'跳出循环
                End If

                for j = i + 1 To intRowCount'遍历名称

                    If sht.Rows(j).EntireRow.Hidden = False Then'如果这一行没有被隐藏

                        If sht.Cells(j,intNameCol).Value = "" Then'如果查询到了空数据
                            Exit For'跳出循环
                        End If

                        If sht.Cells(j,intNameCol).Value = sht.Cells(i,intNameCol).Value Then'如果匹配到了同名单元格
                            '将数量合并填入一开始的名称的一行
                            sht.Cells(i,intNumCol).Value = CInt(sht.Cells(i,intNumCol).Value) + CInt(sht.Cells(j,intNumCol).Value)
                            Rows(j).EntireRow.Hidden = True'隐藏被提取数量的这一行

                        End If

                    End If
                Next
            End If
        Next
        blnMerged = True'设置为真
    End If

    MergeNameSumQuan = blnMerged'返回是否成功合并
End Function

'MergeNameSumQuan的反向操作
'可能存在缺陷
Public Function AntiMergeNameSumQuan(workbook,sheet,namerow,namecol,numcol) As Boolean

    '反合并成功标志位
    blnAntiMerged = False'默认为假

    '参数验证
    If IsFileName(workbook) And (VarType(namerow) = vbInteger) And _ 
    (VarType(namecol) = vbInteger) And (VarType(numcol) = vbInteger) Then'验证通过

        intRowBegin = namerow'获取起始行
        intNameCol = namecol'获取名称列
        intNumCol = numcol'获取数量列
        Set sht = Workbooks(workbook).Sheets(sheet)'获取需要操作的sheet

        for s = intRowBegin to LOOPLIMIT'循环上限
            If sht.Cells(s,intNameCol).Value = "" Then'如果出现空单元格
                intRowCount = s - 1'获取上一行行号
                Exit For
            End If
            If Not IsNumeric(sht.Cells(s,intNumCol).Value) Then'如果检测到了数量是非数字项目
                AntiMergeNameSumQuan = blnMerged
                Exit Function'退出程序
            End If
        Next

        for i = intRowBegin To intRowCount'遍历名称

            If sht.Rows(i).EntireRow.Hidden = False Then'如果这一行没有被隐藏 

                If sht.Cells(i,intNameCol).Value = "" Then'如果查询到了空数据
                    Exit For'跳出循环
                End If

                for j = i + 1 To intRowCount'遍历名称

                    If sht.Rows(j).EntireRow.Hidden = True Then'如果这一行被隐藏了  

                        If sht.Cells(j,intNameCol).Value = "" Then'如果查询到了空数据
                            Exit For'跳出循环
                        End If

                        If sht.Cells(j,intNameCol).Value = sht.Cells(i,intNameCol).Value Then'如果匹配到了同名单元格
                            '将数量合并填入一开始的名称的一行
                            sht.Cells(i,intNumCol).Value = CInt(sht.Cells(i,intNumCol).Value) - CInt(sht.Cells(j,intNumCol).Value)
                            'Rows(j).EntireRow.Hidden = Flase'隐藏被提取数量的这一行

                        End If

                    End If
                Next
            End If
        Next
        sht.Rows(Cstr(intRowBegin) & ":" & CStr(intRowCount)).EntireRow.Hidden = Flase'取消隐藏数据
        blnAntiMerged = True'设置为真
    End If

    AntiMergeNameSumQuan = blnAntiMerged'返回是否成功合并
End Function

'获取电脑硬件信息
Public Function GetHardInfo() As String
    '以逗号做分隔符，返回包含一系列信息的字符串

    On Error Resume Next'遇到错误继续运行
    'On Error Goto 0'遇到错误报错

    Dim strDiv$'分割符变量
    Dim strHardInfo$'定义返回值变量
    strDiv = ","'逗号
    strHardInfo = ""'初始化返回值

    '获取主板序列号
    Dim objs As Object, Obj As Object, WMI As Object', '主板序列号
    Set WMI = GetObject("WinMgmts:")
    Set objs = WMI.InstancesOf("Win32_BaseBoard")
    For Each Obj In objs
        strHardInfo = strHardInfo & Obj.SerialNumber & strDiv'获取主板序列号,并用逗号分割
    Next

    '获取显卡型号，厂商
    Dim tmp1, tmp2
    Set tmp2 = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_VideoController")
    For Each tmp1 In tmp2
        strHardInfo = strHardInfo & tmp1.VideoProcessor & strDiv & tmp1.AdapterCompatibility & strDiv
    Next

    '获取网卡MAC地址
    Dim objNetCard
    Set objNetCard = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
    For Each iAdress In objNetCard
        If iAdress.IPEnabled = True Then
            strHardInfo = strHardInfo & iAdress.MacAddress & strDiv
            Exit For
        End If
    Next

    '获取硬盘型号
    Dim objHardDrive
    Set objHardDrive = GetObject("Winmgmts:").InstancesOf("Win32_DiskDrive")
    For Each mo In objHardDrive
        strHardInfo = strHardInfo & mo.Model & ","
    Next

    ' '获取CPU序列号
    ' '这个不是唯一的，即有可能多个CPU同一一序列号
    ' For Each objSequ In GetObject("Winmgmts:").InstancesOf("Win32_Processor")
    '     strHardInfo = strHardInfo  & CStr(objSequ.ProcessorId) & ","
    ' Next

    GetHardInfo = RemoveLastElem(strHardInfo)'剔除信息字符串的最后一个字符","，并返回
End Function

'检测使用者电脑的硬件信息
Public Sub HardCheck()

    Dim blnCheck As Boolean'验证标志位
    Dim blnChecked As Boolean'验证通过标志位
    Dim varHardInfo As Variant'硬件数据数组
    Dim intHardInfo%'硬件数据数组位数
    Dim intHardInfoRowBegin%'数据写入起始行
    Dim strHardInfo$ '硬件数据字符串

    strHardInfo = GetHardInfo()
    varHardInfo = Split(strHardInfo,",")'获取硬件数据
    intHardInfo = Ubound(varHardInfo)'获取硬件数据数组位数
    intHardInfoRowBegin = 3'赋值数据写入起始行
    blnCheck = False'默认为假
    blnChecked = True'默认为真

    set shtHard = ThisWorkbook.sheets("Hard")'获取数据页

    If shtHard.Cells(7,7).Value = "" Then'判断是否存在个人签名
        shtHard.Cells(7,7).Value = "#4C69754A69614C6F6E67"'个人签名
        blnCheck = False'初始化阶段不需要验证
    Else
        blnCheck = True'存在个人签名，已经完成初始化，需要进行验证
    End If

    If blnCheck Then'进行验证
        For i = intHardInfoRowBegin to intHardInfoRowBegin + intHardInfo
            If shtHard.Cells(i,2).Value <> varHardInfo(i-intHardInfoRowBegin) Then
                blnChecked = False
            End If
        Next
    Else'进行初始化
        For i = intHardInfoRowBegin to intHardInfoRowBegin + intHardInfo
            shtHard.Cells(i,2).Value = varHardInfo(i-intHardInfoRowBegin)
        Next
    End If

    '验证确认
    If Not blnChecked Then '如果密码不通过退出程序
        ActiveWorkbook.Save '自动保存
        Workbooks(ThisWorkbook.Name).Close '关闭文件
    End If
End Sub