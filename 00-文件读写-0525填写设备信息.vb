Option Explicit

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim swAssembly As Object
Dim myModelView As Object


Dim strReturn As String, lenReturn As String
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub CommandButton1_Click()

    Unload UserForm00a
      
    If TransNum = 1 Then
    Load UserForm01
    UserForm01.Show vbModeless
    End If


    If TransNum = 2 Then
    Load UserForm02
    UserForm02.Show vbModeless
    End If

String1 = TextBox1.text
' String2 = TextBox2.text
' String3 = TextBox3.text
String4 = TextBox4.text
String5 = TextBox5.text
String6 = TextBox6.text
String7 = TextBox7.text
String8 = TextBox8.text
String9 = TextBox9.text
' String10 = TextBox17.text

End Sub

' ******************
' 数据文件操作（按钮21-23）
' 修改日期：2023.05.25
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

    lenReturn = WritePrivateProfileString("图框填写00a", "工程名称", TextBox1.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "审核", TextBox4.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "标准化", TextBox5.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "校对", TextBox6.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "设计负责人", TextBox7.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "设计", TextBox8.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "日期", TextBox9.text, filePath)
    
    lenReturn = WritePrivateProfileString("图框填写00a", _
    "工程名称=" & TextBox1.text & _
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

    lenReturn = WritePrivateProfileString("图框填写00a", "工程名称", TextBox1.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "审核", TextBox4.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "标准化", TextBox5.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "校对", TextBox6.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "设计负责人", TextBox7.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "设计", TextBox8.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写00a", "日期", TextBox9.text, filePath)
    
    lenReturn = WritePrivateProfileString("图框填写00a", _
    "工程名称=" & TextBox1.text & _
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
    lenReturn = GetPrivateProfileString("图框填写00a", "工程名称", vbNullString, strReturn, &HFF, filePath)
    TextBox1.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写00a", "审核", vbNullString, strReturn, &HFF, filePath)
    TextBox4.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写00a", "标准化", vbNullString, strReturn, &HFF, filePath)
    TextBox5.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写00a", "校对", vbNullString, strReturn, &HFF, filePath)
    TextBox6.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写00a", "设计负责人", vbNullString, strReturn, &HFF, filePath)
    TextBox7.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写00a", "设计", vbNullString, strReturn, &HFF, filePath)
    TextBox8.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写00a", "日期", vbNullString, strReturn, &HFF, filePath)
    TextBox9.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))

End Sub








' ******************
' 数据文件操作（按钮21-23）
' 修改日期：2023.05.25
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

    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, "项目名称", TextBox2.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, "项目号", TextBox17.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, "图名", TextBox3.text, filePath)
    
    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, _
    "，项目名称=" & TextBox2.text & _
    "，项目号=" & TextBox17.text & _
    "，图名=" & TextBox3.text, 1, "D:\swlog.ini")

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

    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, "项目名称", TextBox2.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, "项目号", TextBox17.text, filePath)
    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, "图名", TextBox3.text, filePath)
    
    lenReturn = WritePrivateProfileString("图框填写03" & TransStr, _
    "，项目名称=" & TextBox2.text & _
    "，项目号=" & TextBox17.text & _
    "，图名=" & TextBox3.text, 1, "D:\swlog.ini")

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
    lenReturn = GetPrivateProfileString("图框填写03" & TransStr, "项目名称", vbNullString, strReturn, &HFF, filePath)
    TextBox2.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写03" & TransStr, "项目号", vbNullString, strReturn, &HFF, filePath)
    TextBox17.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))
    lenReturn = GetPrivateProfileString("图框填写03" & TransStr, "图名", vbNullString, strReturn, &HFF, filePath)
    TextBox3.text = Trim(Replace(Left(strReturn, lenReturn), Chr(0), ""))

End Sub






If OptionButton1.Value = True Then
   Unload UserFormb0
   Load UserFormb1
   UserFormb1.Show vbModeless
   Load UserForm03
   UserForm03.Show vbModeless
   TransStr = "转台H=500"
End If




Private Sub CommandButton1_Click()

    With Excel.Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show = -1 Then
            RootPath = .SelectedItems(1) & IIf(Right(.SelectedItems(1), 1) = "\", "", "\")
        Else
            RootPath = "-1"
            Exit Sub
        End If
    End With
    MsgBox "您选择的文件夹路径为: " & RootPath
    
   Unload UserFormh0
    
    If OptionButton1.Value = True Then
       Load UserFormh1
       UserFormh1.Show vbModeless
       DocName = "TJ1550.20 长轴链式滚床A L=5200.SLDASM"
   TransStr = "长轴链式滚床"
    End If
    
    If OptionButton2.Value = True Then
       Load UserFormh2
       UserFormh2.Show vbModeless
       DocName = "TJ1550.29标准皮带滚床L=5000.SLDASM"
   TransStr = "标准皮带滚床"
    End If
    
    If OptionButton3.Value = True Then
       Load UserFormh3
       UserFormh3.Show vbModeless
       DocName = "TJ1550.36 工位滚床A L=6000.SLDASM"
   TransStr = "工位滚床"
    End If
    
    If OptionButton4.Value = True Then
       Load UserFormh4
       UserFormh4.Show vbModeless
       DocName = "TJ1550.43中心旋转滚床L=5200.SLDASM"
   TransStr = "中心旋转滚床"
    End If
    
    If OptionButton5.Value = True Then
       Load UserFormh5
       UserFormh5.Show vbModeless
       DocName = "TJ1550.60偏心滚床C=5200.SLDASM"
   TransStr = "偏心滚床"
    End If
    
    If OptionButton6.Value = True Then
       Load UserFormh6
       UserFormh6.Show vbModeless
       DocName = "1550.30 TJ1500.30.1堆垛滚床.SLDASM"
   TransStr = "堆垛滚床"
    End If
    
    If OptionButton7.Value = True Then
       Load UserFormh7
       UserFormh7.Show vbModeless
       DocName = "TJ1550.4转载滚床A L=5200.SLDASM"
   TransStr = "转载滚床"
    End If
    
    If OptionButton8.Value = True Then
       Load UserFormh8
       UserFormh8.Show vbModeless
       DocName = "TJ1550.9 不锈钢全盖板滚床L=5200.SLDASM"
   TransStr = "不锈钢全盖板滚床"
    End If

    If OptionButton9.Value = True Then
       Load UserFormh9
       UserFormh9.Show vbModeless
       DocName = "TJ1550.10摆杆入口滚床.SLDASM"
   TransStr = "摆杆入口滚床"
    End If
    
    If OptionButton10.Value = True Then
       Load UserFormh10
       UserFormh10.Show vbModeless
       DocName = "11出口滚床.SLDASM"
   TransStr = "出口滚床"
    End If
    
    If OptionButton11.Value = True Then
       Load UserFormh11
       UserFormh11.Show vbModeless
       DocName = "17TJ1550.6 锁紧站滚床17.SLDASM"
   TransStr = "锁紧站滚床"
    End If
    
   Load UserForm03
   UserForm03.Show vbModeless
   
   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc
   Set Part = swApp.OpenDoc6(RootPath & DocName, 2, 192, "", longstatus, longwarnings)
   Set swAssembly = Part
   swApp.ActivateDoc2 DocName, False, longstatus
   Set Part = swApp.ActiveDoc
   Set myModelView = Part.ActiveView
   
End Sub










