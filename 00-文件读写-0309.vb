
Dim strReturn As String, lenReturn As String
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' ******************
' 在窗体初始化时就读取预设内容
' ******************
Private Sub UserForm_Initialize()
    strReturn = vbNullString
    strReturn = Space(&HFE)
    
    lenReturn = GetPrivateProfileString("立柱", "立柱高度", vbNullString, strReturn, &HFF, "D:\swSave\zanshi.ini")
    If lenReturn <> Null Then
        TextBox1.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    End If
    
    lenReturn = GetPrivateProfileString("立柱", "定位高度", vbNullString, strReturn, &HFF, "D:\swSave\zanshi.ini")
    If lenReturn <> Null Then
        TextBox2.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    End If
    
    lenReturn = GetPrivateProfileString("立柱", "横梁间距", vbNullString, strReturn, &HFF, "D:\swSave\zanshi.ini")
    If lenReturn <> Null Then
        TextBox3.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    End If

End Sub

' ******************
' 使用按钮进行存储
' ******************
Private Sub CommandButton1_Click()

    lenReturn = WritePrivateProfileString("立柱", "立柱高度", TextBox1.Text, "D:\swSave\zanshi.ini")
    lenReturn = WritePrivateProfileString("立柱", "定位高度", TextBox2.Text, "D:\swSave\zanshi.ini")
    lenReturn = WritePrivateProfileString("立柱", "横梁间距", TextBox3.Text, "D:\swSave\zanshi.ini")

End Sub

' ******************
' 使用按钮进行读取
' ******************
Private Sub CommandButton2_Click()
    strReturn = vbNullString
    strReturn = Space(&HFE)
    lenReturn = GetPrivateProfileString("立柱", "立柱高度", vbNullString, strReturn, &HFF, "D:\swSave\zanshi.ini")
    ' TextBox1.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    TextBox1.text = strReturn
    ' 上面的函数看不懂，而且会导致文本框读取中出现乱码，直接用【TextBox1.text = strReturn】就可以读出
    lenReturn = GetPrivateProfileString("立柱", "定位高度", vbNullString, strReturn, &HFF, "D:\swSave\zanshi.ini")
    TextBox2.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("立柱", "横梁间距", vbNullString, strReturn, &HFF, "D:\swSave\zanshi.ini")
    TextBox3.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    
End Sub



' *******************************
' 以下是函数说明
' 【写入函数】WritePrivateProfileString
' 类、条目名、字串值、路径与文件名
' 【读取函数】GetPrivateProfileString
' 类、条目名、默认值vbNullString、字串缓冲strReturn、最大长度&HFF、路径与文件名
' 其中strReturn初始被设置成254个空格（Space(&HFE)）
' 
' 请保证写入和读取时，【类、条目、路径】是对应的
' 
' 参考：软件开发技术联盟，《Visual Basic开发实例大全（基础卷）》，p375-377，p726-727
' *******************************
' Public filePath As String
' Public RootPath As String
' Public DocName As String
' 文件完整路径filePath 、根目录RootPath、文件名DocName

' Sub main()

'     Dim fd As FileDialog, vrtSelectedItem As Variant
'     Set fd = Excel.Application.FileDialog(msoFileDialogFilePicker)
'     With fd
'         .AllowMultiSelect = False
'         .ButtonName = "Ok"
'         If .Show = -1 Then
'             For Each vrtSelectedItem In .SelectedItems
'                 filePath = vrtSelectedItem
'             Next
'         Else
'             filePath = "-1"
'         End If
'     End With
'     Set fd = Nothing
    
'     ' 从右往左查找"\"的位置，并返回从左往右的长度：InStrRev(filePath, "\", Len(filePath))
'     DocName = Right(filePath, Len(filePath) - InStrRev(filePath, "\", Len(filePath)))
'     RootPath = Left(filePath, InStrRev(filePath, "\", Len(filePath)))
'     MsgBox "您想打开的文件名为: " & DocName

' End Sub
' *******************************





' ****************
' ModuleOpen模块中添加全局变量
Public RootPath As String
Public filePath As String
Public DocName As String
Public SavePath As String
' 文件完整路径filePath 、根目录RootPath、文件名DocName
' ***************

Option Explicit

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim swSelMgr As SldWorks.SelectionMgr
Dim ExtrudeFeatureData As SldWorks.ExtrudeFeatureData2
Dim Feature As SldWorks.Feature
Dim retval As Boolean
Dim swMoveFaceFeat As SldWorks.MoveFaceFeatureData
Dim swErrors As Long
Dim swWarnings As Long

Dim strReturn As String, lenReturn As String
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



' ******************
' 使用按钮进行存储
' ******************
Private Sub CommandButton1_Click()

' 存储数据文件：选择存储文件和路径，运行写入代码

    Dim fd As FileDialog, vrtSelectedItem As Variant
    Set fd = Excel.Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Ok"
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                filePath = vrtSelectedItem
            Next
        Else
            filePath = "-1"
            Exit Sub
        End If
    End With
    Set fd = Nothing
        
    MsgBox "您想将设置存储到: " & filePath

    lenReturn = WritePrivateProfileString("立柱", "立柱高度", TextBox1.Text, filePath)
    lenReturn = WritePrivateProfileString("立柱", "定位高度", TextBox2.Text, filePath)
    lenReturn = WritePrivateProfileString("立柱", "横梁间距", TextBox3.Text, filePath)

    lenReturn = WritePrivateProfileString("立柱", "立柱高度=" & TextBox1.Text & "，定位高度=" & TextBox2.Text & "，横梁间距=" & TextBox3.Text, 1, "D:\swSave\log.ini")

End Sub

' ******************
' 使用按钮进行读取
' ******************
Private Sub CommandButton2_Click()

' 读取数据文件：选择存储文件和路径，运行读取代码

    Dim fd As FileDialog, vrtSelectedItem As Variant
    Set fd = Excel.Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Ok"
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                filePath = vrtSelectedItem
            Next
        Else
            filePath = "-1"
            Exit Sub
        End If
    End With
    Set fd = Nothing
    
    MsgBox "您想导入的设置为: " & filePath

    strReturn = vbNullString
    strReturn = Space(&HFE)
    lenReturn = GetPrivateProfileString("立柱", "立柱高度", vbNullString, strReturn, &HFF, filePath)
    TextBox1.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("立柱", "定位高度", vbNullString, strReturn, &HFF, filePath)
    TextBox2.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("立柱", "横梁间距", vbNullString, strReturn, &HFF, filePath)
    TextBox3.Text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    
End Sub



Private Sub CommandButton3_Click()

' 新建数据文件：弹出对话框（输入文件名），选择存储路径

    With Excel.Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show = -1 Then
            SavePath = .SelectedItems(1) & IIf(Right(.SelectedItems(1), 1) = "\", "", "\")
        Else
            SavePath = "-1"
            Exit Sub
        End If
    End With
    MsgBox "您选择的文件夹路径为: " & SavePath

    Dim str As String
    str = InputBox("请输入文件名")

    DocName = str & ".ini"
    filePath = SavePath & DocName

    lenReturn = WritePrivateProfileString("立柱", "立柱高度", TextBox1.Text, filePath)
    lenReturn = WritePrivateProfileString("立柱", "定位高度", TextBox2.Text, filePath)
    lenReturn = WritePrivateProfileString("立柱", "横梁间距", TextBox3.Text, filePath)

    
    lenReturn = WritePrivateProfileString("立柱", "立柱高度=" & TextBox1.Text & "，定位高度=" & TextBox2.Text & "，横梁间距=" & TextBox3.Text, 1, "D:\swSave\log.ini")

End Sub




Private Sub CommandButton4_Click()

Dim MyStr
Open "D:\swSave\log.ini" For Input As #1
Do While Not EOF(1)
    MyStr = Input(1, #1)
    TextBox8.Text = TextBox8.Text & MyStr
Loop
Close #1

End Sub



' ********************************************************




' ******************
' 数据文件操作（按钮21-23）
' 修改日期：2023.04.28
' ******************
' 新建数据文件：弹出对话框（输入文件名），选择存储路径
Private Sub CommandButton21_Click()
    With Excel.Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show = -1 Then
            SavePath = .SelectedItems(1) & IIf(Right(.SelectedItems(1), 1) = "\", "", "\")
        Else
            SavePath = "-1"
            Exit Sub
        End If
    End With
    MsgBox "您选择的文件夹路径为: " & SavePath
    Dim str As String
    str = InputBox("请输入文件名")
    DocName = str & ".ini"
    filePath = SavePath & DocName

    lenReturn = WritePrivateProfileString("图框填写h00", "工程名称", TextBox1.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "项目名称", TextBox2.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "项目号", TextBox17.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "图名", TextBox3.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "审核", TextBox4.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "标准化", TextBox5.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "校对", TextBox6.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "设计负责人", TextBox7.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "设计", TextBox8.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "日期", TextBox9.text, filePath)
    
    lenReturn = WritePrivateProfileString("图框填写h00", _
    "工程名称=" & TextBox1.text & _
    "，项目名称=" & TextBox2.text & _
    "，项目号=" & TextBox17.text & _
    "，图名=" & TextBox3.text & _
    "，审核=" & TextBox4.text & _
    "，标准化=" & TextBox5.text & _
    "，校对=" & TextBox6.text & _
    "，设计负责人=" & TextBox7.text & _
    "，设计=" & TextBox8.text & _
    "，日期=" & TextBox9.text, 1, "D:\swlog.ini")

End Sub
' 存储数据文件：选择存储文件和路径，运行写入代码
Private Sub CommandButton22_Click()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    Set fd = Excel.Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Ok"
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                filePath = vrtSelectedItem
            Next
        Else
            filePath = "-1"
            Exit Sub
        End If
    End With
    Set fd = Nothing
    MsgBox "您想将设置存储到: " & filePath

    lenReturn = WritePrivateProfileString("图框填写h00", "工程名称", TextBox1.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "项目名称", TextBox2.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "项目号", TextBox17.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "图名", TextBox3.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "审核", TextBox4.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "标准化", TextBox5.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "校对", TextBox6.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "设计负责人", TextBox7.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "设计", TextBox8.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写h00", "日期", TextBox9.text, filePath)
    
    lenReturn = WritePrivateProfileString("图框填写h00", _
    "工程名称=" & TextBox1.text & _
    "，项目名称=" & TextBox2.text & _
    "，项目号=" & TextBox17.text & _
    "，图名=" & TextBox3.text & _
    "，审核=" & TextBox4.text & _
    "，标准化=" & TextBox5.text & _
    "，校对=" & TextBox6.text & _
    "，设计负责人=" & TextBox7.text & _
    "，设计=" & TextBox8.text & _
    "，日期=" & TextBox9.text, 1, "D:\swlog.ini")

End Sub
' 读取数据文件：选择存储文件和路径，运行读取代码
Private Sub CommandButton23_Click()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    Set fd = Excel.Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Ok"
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                filePath = vrtSelectedItem
            Next
        Else
            filePath = "-1"
            Exit Sub
        End If
    End With
    Set fd = Nothing
    MsgBox "您想导入的设置为: " & filePath

    strReturn = vbNullString
    strReturn = Space(&HFE)
    lenReturn = GetPrivateProfileString("图框填写h00", "工程名称", vbNullString, strReturn, &HFF, filePath)
    TextBox1.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "项目名称", vbNullString, strReturn, &HFF, filePath)
    TextBox2.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "项目号", vbNullString, strReturn, &HFF, filePath)
    TextBox17.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "图名", vbNullString, strReturn, &HFF, filePath)
    TextBox3.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "审核", vbNullString, strReturn, &HFF, filePath)
    TextBox4.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "标准化", vbNullString, strReturn, &HFF, filePath)
    TextBox5.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "校对", vbNullString, strReturn, &HFF, filePath)
    TextBox6.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "设计负责人", vbNullString, strReturn, &HFF, filePath)
    TextBox7.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "设计", vbNullString, strReturn, &HFF, filePath)
    TextBox8.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写h00", "日期", vbNullString, strReturn, &HFF, filePath)
    TextBox9.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))

End Sub















