Attribute VB_Name = "Kernel"
Option Explicit
'===========COLORS===========
'INFO &H00FFFFFF& (SENDPAR=WHITE)
'WARNING &H0000FFFF& (SENDPAR=BLUE)
'ERROR &H00FFFF00& (SENDPAR=RED)

Public Const ShowCntPerMsg As Long = 50
Public Const TwipsPerCM As Long = 567
Public Const Development As Long = 1
Public Const TitleHi As Long = 495
Public Const DefCnt As Long = 1

Public EnableTranslation, PageWidth As Long, PageHeight As Long, TopMargin As Long, BotMargin As Long, LeftMargin As Long, RightMargin As Long, AutoCls As Long

Function ReverseColor(Color As Long) As Long
    If Color = vbRed Then ReverseColor = RGB(15, 255, 255)
    If Color = vbBlue Then ReverseColor = RGB(255, 255, 15)
    If Color = vbBlack Then ReverseColor = vbWhite
End Function

Sub RaiseSysErr(Detail As String, Module As String)
    SystemError.ErrDetail.Caption = translate("An error occured, some operations won't be excuted. Please Report the following contents to us :)") & vbCrLf & "Module: " & Module & vbCrLf & "Details:" & vbCrLf & Detail & vbCrLf & translate("Time: ") & Now & vbCrLf & vbCrLf & translate("Double click to close the window after 10 seconds. Press PrtSc to take a capture (THIS OPERATION WILL COVER YOUR CLIPBOARD). ")
    SystemError.CurrentTime.Caption = Now
    SystemError.Show 1
End Sub

Function InputText(StartText As String, Optional MultiLine As Boolean = True) As String
    InputWin.Text2.Text = StartText
    If MultiLine = False Then
        InputWin.Text1.Visible = True
        InputWin.Text2.Visible = False
        InputWin.Text1.Text = StartText
    End If
    InputWin.Show 1
    If MultiLine = False Then InputWin.Text2.Text = InputWin.Text1.Text
    If InputWin.Caption = translate("UserCancel") Then InputText = StartText Else InputText = InputWin.Text2.Text
    Unload InputWin
End Function

Sub InitPreview()
    If AutoCls = 1 Then Preview.Picture2.Cls
    Preview.Picture2.Height = PageHeight
    Preview.Picture2.Width = PageWidth
'    Preview.HScroll1.Max = PageWidth
 '   Preview.VScroll1.Max = PageHeight
    Preview.Show
    MainFrm.WIP.Left = 0
End Sub

Sub Main()
    AutoCls = GetSetting("FreeExam", "Create", "AutoCls", 1)
    DevWin.Show
End Sub

Function Min(a, b)
    If a < b Then Min = a Else Min = b
End Function

Function Max(a, b)
    If a > b Then Max = a Else Max = b
End Function

Function translate(src As String) As String
    
    
    
    
    translate = src
    If EnableTranslation = 1 Then
        If src = "Unable to get the page size that you've chosen." Then translate = "无法获取您选择的页面尺寸"
        
        If src = "The margin that you've inputed is invaild" Then translate = "您输入的页边距无效"
        If src = "Clear message list only and do not clear message list were both turned on." Then translate = "只清除消息列表和不要清除消息列表都打开了"
        If src = "Create/PageSettings/NewEvent" Then translate = "创建/页面设置/新事件"
        If src = "(Expired)" Then translate = "(过期的)"
        If src = "[Info]" Then translate = "[信息]"
        If src = "[Warning]" Then translate = "[警告]"
        If src = "[Error]" Then translate = "[错误]"
        If src = "Available area to edit is the area in the rectangle." Then translate = "有效编辑区域在矩形内"
        If src = "No new messages." Then translate = "无新消息"
        If src = "An error occured, some operations won't be excuted. Please Report the following contents to us :)" Then translate = "一个错误发生了，一些操作不会被执行，请反馈以下信息给我们:)"
        If src = "Module: " Then translate = "模块："
        If src = "Details:" Then translate = "详情："
        If src = "Time: " Then translate = "时间："
        If src = "Double click to close the window after 10 seconds. Press PrtSc to take a capture (THIS OPERATION WILL COVER YOUR CLIPBOARD). " Then translate = "10秒后双击关闭本窗口。按PrtSc键截图，此操作将覆盖您的剪贴板。"
        
        If src = "FreeExam" Then translate = "自由考"
        If src = "Create" Then translate = "新建"
        If src = "AutoCls" Then translate = "自动清除"
        If src = "The size of the preview is NEAR the actual size." Then translate = "预览尺寸接近真实尺寸"
        If src = "Contents can't be shown" Then translate = "内容不能被显示"
        If src = "Access Denied - You don't have enough privilege to access here. By the way, there is nothing interesting." Then translate = "访问拒绝 - 你没有足够权限访问这里。顺便一提，这里没有好玩的。"
        If src = "DevWin/PrivCheck" Then translate = "开发窗口/权限检查"
        If src = "Authentication Passed." Then translate = "验证通过"
        
        
        
        If src = "Click to edit" Then translate = "点击编辑"
        If src = "Disabled" Then translate = "已停用"
        
        
        If src = "[SysErr]" Then translate = "[系统错误]"
        If src = "Loading Fonts(" Then translate = "加载字体中("
        
        
        If src = "Invaild Format." Then translate = "无效格式"
        If src = "The size you've input is too large" Then translate = "你的输入过大"
        
        
        
        
        
        
        
        If src = "Invaild Format" Then translate = "无效格式"
        
        
        
        
        If src = "The input A is too large that we can't process that." Then translate = "输入A太大以至于我们无法处理它。"
        
        
        If src = "The input B is too large that we can't process that." Then translate = "输入B太大以至于我们无法处理它。"
        
        If src = "The input C is too large that we can't process that." Then translate = "输入C太大以至于我们无法处理它。"
        
        If src = "The input D is too large that we can't process that." Then translate = "输入D太大以至于我们无法处理它。"
        
        
        
        
        
        
        
        
        If src = "Generated and saved in the following path: " Then translate = "已生成并保存在以下路径"
        If src = "FontName: " Then translate = "字体名："
        If src = "FontSize: " Then translate = "字号："
        If src = "Bold: " Then translate = "粗体："
        If src = "italic: " Then translate = "斜体："
        If src = "Alignment: " Then translate = "对齐："
        If src = "* For italic/Bold: 1 is true, 0 is false" Then translate = "对于斜体和粗体：1表示开启，0表示关闭"
        If src = "Text: " Then translate = "文本："
        If src = "DisabledImage: " Then translate = "图像已禁用："
        If src = "False" Then translate = "否"
        If src = "Drive: " Then translate = "磁盘："
        If src = "Path: " Then translate = "路径："
        If src = "File: " Then translate = "文件："
        If src = "Position: " Then translate = "位置："
        If src = "[Image Information Unavailable]" Then translate = "[图像信息不可用]"
        
        
        
        
        
        
        
        If src = "Nothing can be previewed." Then translate = "没有能被预览的东西"
        If src = "You have not selected the position of the image." Then translate = "你还没有选择图像位置"
        If src = "Your input is so large that we can't process it." Then translate = "你的输入过大以至于我们无法处理它"
        If src = "[TYPE=RUNTIME_ERROR][ERRORID=" Then translate = "[类型=运行时错误][错误号="
        If src = "][ERRDESC.=" Then translate = "][错误描述="
        
        If src = "Auto split line is unsupportted for alignment mode 1 or 2." Then translate = "自动换行对于对齐模式1或2不被支持"
        If src = "The input will be split into multi parts" Then translate = "输入将分成多行"
        If src = "select parts that you want to preview in the list." Then translate = "在列表中选定你要预览的部分"
        If src = "The target size is too large, we are unable to process it." Then translate = "目标尺寸太大以至于我们无法处理它"
        
        If src = "from " Then translate = "从"
        If src = " to " Then translate = " 到"
        If src = "Image which tracknumber=" Then translate = "追踪号是以下值的图像："
        If src = " not found" Then translate = " 未找到"
        If src = "CNT=" Then translate = "计数="
        
        
        
        
        If src = "Initiating" Then translate = "加载中"
        If src = "Special Input" Then translate = "特殊输入"
        If src = "Format" Then translate = "格式化"
    End If
End Function
