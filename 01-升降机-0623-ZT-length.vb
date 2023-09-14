Private Sub cmdGenerate_Click()
  Dim swApp As SldWorks.SldWorks
  Dim swModel As SldWorks.ModelDoc2
  Dim swSelMgr As SldWorks.SelectionMgr
  Dim count As Long
  Dim Feature As SldWorks.Feature
  Dim ExtrudeFeatureData As SldWorks.ExtrudeFeatureData2
  Dim retval As Boolean
  Dim Depth As Double
  Dim Factor As Integer
  
  Factor = CInt(txtDepth.Text)
  Set swApp = Application.SldWorks
  Set swModel = swApp.ActiveDoc
  Set swSelMgr = swModel.SelectionManager
  count = swSelMgr.GetSelectedObjectCount2(-1)
  ' If count <> 1 Then
  '   swApp.SendMsgToUser2 "Please select only Extrude1.", _
  '     swMbWarning, swMbOk
  '   Exit Sub
  ' End If
  Set Feature = swSelMgr.GetSelectedObject6(count, -1)
  ' If Not Feature.GetTypeName2 = "Extrusion" Then
  '   swApp.SendMsgToUser2 "Please select only Extrude1.", _
  '     swMbWarning, swMbOk
  '   Exit Sub
  ' End If
  Set ExtrudeFeatureData = Feature.GetDefinition
  ' 获取选定的特征
  Depth = ExtrudeFeatureData.GetDepth(True)
  ExtrudeFeatureData.SetDepth True, Depth * Factor
  retval = Feature.ModifyDefinition _
    (ExtrudeFeatureData, swModel, Nothing)
End Sub

Private Sub cmdExit_Click()
  End
End Sub





Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()
Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc

' Open
Set Part = swApp.OpenDoc6("D:\Desktop\二次开发\01-SW二次开发修改\5200滚床数模\ZT片架折弯件B（软件开发用）.SLDPRT", 1, 0, "", longstatus, longwarnings)
Set Part = swApp.ActiveDoc

swApp.ActivateDoc2 "ZT片架折弯件B（软件开发用）.SLDPRT", False, longstatus
Set Part = swApp.ActiveDoc

boolstatus = Part.Extension.SelectByID2("拉伸-薄壁1", "SOLIDBODY", 0, 0, 0, False, 0, Nothing, 0)
' value = instance.SelectByID2(Name, Type, X, Y, Z, Append, Mark, Callout, SelectOption) 通过文件名，选定实体，
' Type = BODYFEATURE时，ISelectionMgr::GetSelectedObject6返回feature
' value = instance.GetSelectedObject6(Index, Mark)
Part.ClearSelection2 True
' 清除选择
Part.SelectionManager.EnableContourSelection = False
' 外形选择：非
Set Part = swApp.ActiveDoc

swApp.ActivateDoc2 "TJ1550.29标准皮带滚床L=5000（软件开发用）.SLDASM", False, longstatus
Set Part = swApp.ActiveDoc
End Sub









'-----------------------------------------------
' 前提条件：
' 1. 打开一个钣金零件。
' 2. 选择钣金特征。
' 3. 打开立即窗口。
'
' 后置条件：
' 1. 将默认弯曲半径加倍。
' 2. 检查图形区域和立即窗口。
'-----------------------------------------------
Option Explicit
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swFeat As SldWorks.Feature
    Dim swSheetMetal As SldWorks.SheetMetalFeatureData
    Dim bRet As Boolean
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    Set swFeat = swSelMgr.GetSelectedObject5(1)
    Set swSheetMetal = swFeat.GetDefinition
    Debug.Print "Feature = " & swFeat.Name
    Debug.Print "  Original bend radius = " & swSheetMetal.BendRadius * 1000# & " mm"
    ' 回滚以更改默认弯曲半径
    bRet = swSheetMetal.AccessSelections(swModel, Nothing): Debug.Assert bRet
    ' 将默认弯曲半径值加倍
    swSheetMetal.BendRadius = 2# * swSheetMetal.BendRadius
    ' 应用更改
    bRet = swFeat.ModifyDefinition(swSheetMetal, swModel, Nothing): Debug.Assert bRet
    
    Debug.Print "  Modified bend radius = " & swSheetMetal.BendRadius * 1000# & " mm"
End Sub















Dim swApp As Object
Dim Part As Object
Dim retval As Boolean
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim Feature As SldWorks.Feature
Dim ExtrudeFeatureData As SldWorks.ExtrudeFeatureData2
Dim swSelMgr As SldWorks.SelectionMgr
Dim Factor As Integer
Dim Depth As Double

Sub main()
Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc
' 打开对应的零件，并设置为活动窗口
Set Part = swApp.OpenDoc6("D:\Desktop\二次开发\01-SW二次开发修改\5200滚床数模\ZT片架折弯件B（软件开发用）.SLDPRT", 1, 0, "", longstatus, longwarnings)
Set Part = swApp.ActiveDoc
swApp.ActivateDoc2 "ZT片架折弯件B（软件开发用）.SLDPRT", False, longstatus
Set Part = swApp.ActiveDoc
' 选定拉伸特征
boolstatus = Part.Extension.SelectByID2("拉伸-薄壁1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
Set swSelMgr = Part.SelectionManager
Set Feature = swSelMgr.GetSelectedObject6(1, -1)
Set ExtrudeFeatureData = Feature.GetDefinition
' 获取并修改选定拉伸长度为原来的2倍
Depth = ExtrudeFeatureData.GetDepth(True)
ExtrudeFeatureData.SetDepth True, Depth * 2
retval = Feature.ModifyDefinition _
  (ExtrudeFeatureData, Part, Nothing)
Part.ClearSelection2 True
' 更新零件及装配体
boolstatus = Part.EditRebuild3()
swApp.ActivateDoc2 "TJ1550.29标准皮带滚床L=5000（软件开发用）.SLDASM", False, longstatus
Set Part = swApp.ActiveDoc
boolstatus = Part.EditRebuild3()

End Sub