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
    If Development = 1 Then DevWin.Show
    Lang.Show
End Sub

Function Min(a, b)
    If a < b Then Min = a Else Min = b
End Function

Function Max(a, b)
    If a > b Then Max = a Else Max = b
End Function

Function translatecontrol(src As String)
    If EnableTranslation = 1 Then

        



    




        If src = "InputWin" Then
        InputWin.Caption = "输入"
        InputWin.Label1.AutoSize = True
InputWin.Label28.AutoSize = True
InputWin.Label1.Alignment = 2
InputWin.Label28.Alignment = 2
    InputWin.Label1.Font = "黑体"
InputWin.Label28.Font = "黑体"
InputWin.Label68.Font = "黑体"
        InputWin.Label1.Caption = "取消"
        InputWin.Label28.Caption = "确定"
        InputWin.Label68.Caption = "自动空格"
        End If
        If src = "Integrated" Then
        Integrated.Label1.AutoSize = True
Integrated.Label2.AutoSize = True
Integrated.Label4.AutoSize = True
Integrated.Label8.AutoSize = True
Integrated.Label7.AutoSize = True
Integrated.Label5.AutoSize = True
Integrated.Label1.Alignment = 2
Integrated.Label2.Alignment = 2
Integrated.Label4.Alignment = 2
Integrated.Label8.Alignment = 2
Integrated.Label7.Alignment = 2
Integrated.Label5.Alignment = 2
Integrated.Label1.Font = "黑体"
Integrated.Label2.Font = "黑体"
Integrated.Label4.Font = "黑体"
Integrated.Label8.Font = "黑体"
Integrated.Label7.Font = "黑体"
Integrated.Label5.Font = "黑体"
        Integrated.Label1.Caption = "代码"
        Integrated.Label2.Caption = "结果"
        Integrated.Label4.Caption = "按Enter插入，按Esc退出"
        Integrated.Label8.Caption = "代码"
        Integrated.Label7.Caption = "结果"
        Integrated.Label5.Caption = "按Enter插入，按Esc退出"
        End If
        If src = "PageSettings" Then
        PageSettings.Caption = "页面设置"
        'PageSettings.Frame1.AutoSize = true
'PageSettings.Label1.AutoSize = True
'PageSettings.Frame2.AutoSize = true
'PageSettings.Label4.AutoSize = True
'PageSettings.Label3.AutoSize = True
'PageSettings.Label5.AutoSize = True
PageSettings.Label6.AutoSize = True
PageSettings.PreviewButton.AutoSize = True
PageSettings.Label2.AutoSize = True
'PageSettings.Frame1.Alignment = 2
'PageSettings.Label1.Alignment = 2
'PageSettings.Frame2.Alignment = 2
PageSettings.Label4.Alignment = 1
PageSettings.Label3.Alignment = 1
PageSettings.Label5.Alignment = 1
PageSettings.Label6.Alignment = 1
PageSettings.PreviewButton.Alignment = 2
PageSettings.Label2.Alignment = 2
PageSettings.Message.Font = "黑体"
        PageSettings.Frame1.Font = "黑体"
PageSettings.Label1.Font = "黑体"
PageSettings.Frame2.Font = "黑体"
PageSettings.Label4.Font = "黑体"
PageSettings.Label3.Font = "黑体"
PageSettings.Label5.Font = "黑体"
PageSettings.Label6.Font = "黑体"
PageSettings.Label11.Font = "黑体"
PageSettings.PreviewButton.Font = "黑体"
PageSettings.Label2.Font = "黑体"
        PageSettings.Frame1.Caption = "页面类型"
        PageSettings.Label1.Caption = "注意：8K的宽度是真实宽度的一半，你选择的8K尺寸已减半。显示的尺寸格式为尺寸名称（长×宽），暂不支持自定义尺寸。"
        PageSettings.Frame2.Caption = "边距"
        PageSettings.Label4.Caption = "上"
        PageSettings.Label3.Caption = "下"
        PageSettings.Label11.Caption = "警告：在您点击确定之后你将不能返回这里。"
        PageSettings.Label5.Caption = "左"
        PageSettings.Label6.Caption = "右"
        PageSettings.PreviewButton.Caption = "预览"
        PageSettings.Label2.Caption = "确定"
        End If
        If src = "MainFrm" Then
        MainFrm.Caption = "自由考 创建考卷"
        MainFrm.Label12.AutoSize = True
MainFrm.Label3.AutoSize = True
MainFrm.Label1.AutoSize = True
MainFrm.Label15.AutoSize = True
MainFrm.Label30.AutoSize = True
MainFrm.Label60.AutoSize = True
MainFrm.Label63.AutoSize = True
MainFrm.PreviewButton.AutoSize = True
MainFrm.Label22.AutoSize = True
MainFrm.Label35.AutoSize = True
MainFrm.Label38.AutoSize = True
MainFrm.Label43.AutoSize = True
'MainFrm.Frame19.AutoSize = true
'MainFrm.Frame11.AutoSize = true
'MainFrm.Check22.AutoSize = true
'MainFrm.Check23.AutoSize = true
MainFrm.Label61.AutoSize = True
MainFrm.Label59.AutoSize = True
MainFrm.Label23.AutoSize = True
MainFrm.Label27.AutoSize = True
'MainFrm.Frame9.AutoSize = true
MainFrm.Label24.AutoSize = True
MainFrm.Label25.AutoSize = True
MainFrm.Label33.AutoSize = True
MainFrm.Label39.AutoSize = True
MainFrm.Label42.AutoSize = True
'MainFrm.Frame20.AutoSize = true
MainFrm.Label36.AutoSize = True
MainFrm.Label37.AutoSize = True
MainFrm.Label13.AutoSize = True
MainFrm.Label14.AutoSize = True
'MainFrm.Frame17.AutoSize = true
MainFrm.Label46.AutoSize = True
MainFrm.Label47.AutoSize = True
MainFrm.Label45.AutoSize = True
'MainFrm.Frame18.AutoSize = true
MainFrm.Label52.AutoSize = True
MainFrm.Label53.AutoSize = True
MainFrm.Label54.AutoSize = True
MainFrm.Label55.AutoSize = True
MainFrm.Label56.AutoSize = True
MainFrm.Label57.AutoSize = True
'MainFrm.Frame13.AutoSize = true
'MainFrm.Option1.AutoSize = true
'MainFrm.Option2.AutoSize = true
'MainFrm.Frame14.AutoSize = true
'MainFrm.Frame1.AutoSize = true
MainFrm.Label7.AutoSize = True
MainFrm.Label8.AutoSize = True
MainFrm.Label9.AutoSize = True
MainFrm.Label10.AutoSize = True
'MainFrm.Frame2.AutoSize = true
MainFrm.Label9.AutoSize = True
MainFrm.Label51.AutoSize = True
MainFrm.Label17.AutoSize = True
'MainFrm.Frame6.AutoSize = true
'MainFrm.Frame7.AutoSize = true
MainFrm.Label18.AutoSize = True
MainFrm.Label19.AutoSize = True
'MainFrm.Frame10.AutoSize = true
MainFrm.Label28.AutoSize = True
MainFrm.Label29.AutoSize = True
'MainFrm.Check17.AutoSize = true
MainFrm.Label10.AutoSize = True
MainFrm.Label11.AutoSize = True
MainFrm.Label12.Alignment = 2
MainFrm.Label3.Alignment = 2
MainFrm.Label1.Alignment = 2
MainFrm.Label15.Alignment = 2
MainFrm.Label30.Alignment = 2
MainFrm.Label60.Alignment = 2
MainFrm.Label63.Alignment = 2
MainFrm.PreviewButton.Alignment = 2
MainFrm.Label22.Alignment = 2
MainFrm.Label35.Alignment = 2
MainFrm.Label38.Alignment = 2
MainFrm.Label43.Alignment = 2
'MainFrm.Frame19.Alignment = 2
'MainFrm.Frame11.Alignment = 2
'MainFrm.Check22.Alignment = 2
'MainFrm.Check23.Alignment = 2
MainFrm.Label61.Alignment = 2
MainFrm.Label59.Alignment = 2
MainFrm.Label23.Alignment = 2
MainFrm.Label27.Alignment = 1
'MainFrm.Frame9.Alignment = 2
MainFrm.Label24.Alignment = 2
MainFrm.Label25.Alignment = 2
MainFrm.Label33.Alignment = 2
MainFrm.Label39.Alignment = 2
MainFrm.Label42.Alignment = 2
'MainFrm.Frame20.Alignment = 2
MainFrm.Label36.Alignment = 2
MainFrm.Label37.Alignment = 2
MainFrm.Label13.Alignment = 2
MainFrm.Label14.Alignment = 2
'MainFrm.Frame17.Alignment = 2
MainFrm.Label46.Alignment = 2
MainFrm.Label47.Alignment = 2
MainFrm.Label45.Alignment = 2
'MainFrm.Frame18.Alignment = 2
MainFrm.Label52.Alignment = 2
MainFrm.Label53.Alignment = 2
MainFrm.Label54.Alignment = 2
MainFrm.Label55.Alignment = 2
MainFrm.Label56.Alignment = 2
MainFrm.Label57.Alignment = 2
'MainFrm.Frame13.Alignment = 2
'MainFrm.Option1.Alignment = 2
'MainFrm.Option2.Alignment = 2
'MainFrm.Frame14.Alignment = 2
'MainFrm.Frame1.Alignment = 2
MainFrm.Label7.Alignment = 2
MainFrm.Label8.Alignment = 2
MainFrm.Label9.Alignment = 2
MainFrm.Label10.Alignment = 2
'MainFrm.Frame2.Alignment = 2
MainFrm.Label9.Alignment = 2
MainFrm.Label51.Alignment = 2
MainFrm.Label17.Alignment = 2
'MainFrm.Frame6.Alignment = 2
'MainFrm.Frame7.Alignment = 2
MainFrm.Label18.Alignment = 2
MainFrm.Label19.Alignment = 2
'MainFrm.Frame10.Alignment = 2
MainFrm.Label28.Alignment = 2
MainFrm.Label29.Alignment = 2
'MainFrm.Check17.Alignment = 2
MainFrm.Label10.Alignment = 2
MainFrm.Label11.Alignment = 2

MainFrm.Label5.Alignment = 2
MainFrm.Label6.Alignment = 2
MainFrm.Label12.FontName = "黑体"
MainFrm.Label3.FontName = "黑体"
MainFrm.Label1.FontName = "黑体"
MainFrm.Label15.FontName = "黑体"
MainFrm.Label30.FontName = "黑体"
MainFrm.Label60.FontName = "黑体"
MainFrm.Label63.FontName = "黑体"
MainFrm.PreviewButton.FontName = "黑体"
MainFrm.Label22.FontName = "黑体"
MainFrm.Label35.FontName = "黑体"
MainFrm.Label38.FontName = "黑体"
MainFrm.Label43.FontName = "黑体"



MainFrm.Frame19.Font = "黑体"
MainFrm.Frame11.Font = "黑体"
MainFrm.Check22.Font = "黑体"
MainFrm.Check23.Font = "黑体"
MainFrm.Label61.Font = "黑体"
MainFrm.Label59.Font = "黑体"
MainFrm.Label23.Font = "黑体"
MainFrm.Label27.Font = "黑体"
MainFrm.Frame9.Font = "黑体"
MainFrm.Label24.Font = "黑体"
MainFrm.Label25.Font = "黑体"
MainFrm.Label33.Font = "黑体"
MainFrm.Label39.Font = "黑体"
MainFrm.Label42.Font = "黑体"
MainFrm.Frame20.Font = "黑体"
MainFrm.Label36.Font = "黑体"
MainFrm.Label37.Font = "黑体"
MainFrm.Label13.Font = "黑体"
MainFrm.Label14.Font = "黑体"
MainFrm.Frame17.Font = "黑体"
MainFrm.Label46.Font = "黑体"
MainFrm.Label47.Font = "黑体"
MainFrm.Label45.Font = "黑体"
MainFrm.Frame18.Font = "黑体"
MainFrm.Label52.Font = "黑体"
MainFrm.Label53.Font = "黑体"
MainFrm.Label54.Font = "黑体"
MainFrm.Label55.Font = "黑体"
MainFrm.Label56.Font = "黑体"
MainFrm.Label57.Font = "黑体"
MainFrm.Frame13.Font = "黑体"
MainFrm.Option1.Font = "黑体"
MainFrm.Option2.Font = "黑体"
MainFrm.Frame14.Font = "黑体"
MainFrm.Frame1.Font = "黑体"
MainFrm.Label7.Font = "黑体"
MainFrm.Label8.Font = "黑体"
MainFrm.Label9.Font = "黑体"
MainFrm.Label10.Font = "黑体"
MainFrm.Frame2.Font = "黑体"
MainFrm.Label9.Font = "黑体"
MainFrm.Label51.Font = "黑体"
MainFrm.Label17.Font = "黑体"
MainFrm.Frame6.Font = "黑体"
MainFrm.Frame7.Font = "黑体"
MainFrm.Label18.Font = "黑体"
MainFrm.Label19.Font = "黑体"
MainFrm.Frame10.Font = "黑体"
MainFrm.Label28.Font = "黑体"
MainFrm.Label29.Font = "黑体"
MainFrm.Check17.Font = "黑体"
MainFrm.Label10.Font = "黑体"
MainFrm.Label11.Font = "黑体"
MainFrm.Label5.Font = "黑体"
MainFrm.Label6.Font = "黑体"
MainFrm.Message.Font = "黑体"
        MainFrm.Label12.Caption = "管理/导出"
MainFrm.Label3.Caption = "关于"
MainFrm.Label1.Caption = "通用"
MainFrm.Label15.Caption = "合并预览"
MainFrm.Label30.Caption = "记录"
MainFrm.Label60.Caption = "合并"
MainFrm.Label63.Caption = "列表"
MainFrm.PreviewButton.Caption = "文本"
MainFrm.Label22.Caption = "答题区"
MainFrm.Label35.Caption = "图片"
MainFrm.Label38.Caption = "空白区"
MainFrm.Label43.Caption = "选择题"
MainFrm.Check16.Caption = "英语模式"
        MainFrm.Frame19.Caption = "页眉"
        MainFrm.Frame11.Caption = "页脚"
        MainFrm.Check22.Caption = "分割线"
        MainFrm.Check23.Caption = "分割线"
        MainFrm.Label61.Caption = "预览"
        MainFrm.Label59.Caption = "合并"
        MainFrm.Label23.Caption = "大小"
        MainFrm.Label27.Caption = "数量（-1表示最大）"
        MainFrm.Frame9.Caption = "选项"
        MainFrm.Label24.Caption = "预览"
        MainFrm.Label25.Caption = "保存"
        MainFrm.Label33.Caption = "追踪号"
        MainFrm.Label39.Caption = "尺寸"
        MainFrm.Label42.Caption = "保存"
        MainFrm.Frame20.Caption = "选择一个图像"
        MainFrm.Label36.Caption = "预览"
        MainFrm.Label37.Caption = "保存"
        MainFrm.Label13.Caption = "预览"
        MainFrm.Label14.Caption = "删除"
        MainFrm.Frame17.Caption = "格式"
        MainFrm.Label46.Caption = "字体"
        MainFrm.Label47.Caption = "字号"
        MainFrm.Label45.Caption = "字形"
        MainFrm.Frame18.Caption = "文本 &^"
        MainFrm.Label52.Caption = "点击编辑"
        MainFrm.Label53.Caption = "点击编辑"
        MainFrm.Label54.Caption = "点击编辑"
        MainFrm.Label55.Caption = "点击编辑"
        MainFrm.Label56.Caption = "预览"
        MainFrm.Label57.Caption = "保存"
        
        MainFrm.Label5.Caption = "对齐"
        MainFrm.Label6.Caption = "字形"
        MainFrm.Frame13.Caption = "记录列表"
        MainFrm.Option1.Caption = "页面"
        MainFrm.Option2.Caption = "格式"
        MainFrm.Frame14.Caption = "详情"
        MainFrm.Frame1.Caption = "格式"
        MainFrm.Label7.Caption = "字体"
        MainFrm.Label8.Caption = "字号"
        MainFrm.Label9.Caption = "字形"
        MainFrm.Label10.Caption = "对齐"
        MainFrm.Frame2.Caption = "选项"
        MainFrm.Label9.Caption = "文本 &^"
        MainFrm.Label51.Caption = "点击编辑"
        MainFrm.Label17.Caption = "配图"
        MainFrm.Frame6.Caption = "选择一个图片"
        MainFrm.Frame7.Caption = "选项"
        MainFrm.Label18.Caption = "位置"
        MainFrm.Label19.Caption = "不配图"
        MainFrm.Frame10.Caption = "记录器"
        MainFrm.Label28.Caption = "格式"
        MainFrm.Label29.Caption = "页面"
        MainFrm.Check17.Caption = "自动"
        MainFrm.Label10.Caption = "预览"
        MainFrm.Label11.Caption = "保存"
        End If
        If src = "Preview" Then Preview.Message.FontName = "黑体"
End If
End Function


Function translate(src As String) As String
    
    
    
    
    translate = src
    If EnableTranslation = 1 Then
        If src = "Unable to get the page size that you've chosen." Then translate = "无法获取您选择的页面尺寸"
        If src = "Translating..." Then translate = "正在翻译..."
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
        If src = "Completed on" Then translate = "已完成，完成时间为"
        If src = "Auto split line is unsupportted for alignment mode 1 or 2." Then translate = "自动换行对于对齐模式1或2不被支持"
        If src = "The input will be split into multi parts" Then translate = "输入将分成多行"
        If src = "select parts that you want to preview in the list." Then translate = "在列表中选定你要预览的部分"
        If src = "The target size is too large, we are unable to process it." Then translate = "目标尺寸太大以至于我们无法处理它"
        If src = "Success with tracknumber" Then translate = "成功，追踪号为"
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
