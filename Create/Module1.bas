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
        If src = "Unable to get the page size that you've chosen." Then translate = "�޷���ȡ��ѡ���ҳ��ߴ�"
        
        If src = "The margin that you've inputed is invaild" Then translate = "�������ҳ�߾���Ч"
        If src = "Clear message list only and do not clear message list were both turned on." Then translate = "ֻ�����Ϣ�б�Ͳ�Ҫ�����Ϣ�б�����"
        If src = "Create/PageSettings/NewEvent" Then translate = "����/ҳ������/���¼�"
        If src = "(Expired)" Then translate = "(���ڵ�)"
        If src = "[Info]" Then translate = "[��Ϣ]"
        If src = "[Warning]" Then translate = "[����]"
        If src = "[Error]" Then translate = "[����]"
        If src = "Available area to edit is the area in the rectangle." Then translate = "��Ч�༭�����ھ�����"
        If src = "No new messages." Then translate = "������Ϣ"
        If src = "An error occured, some operations won't be excuted. Please Report the following contents to us :)" Then translate = "һ���������ˣ�һЩ�������ᱻִ�У��뷴��������Ϣ������:)"
        If src = "Module: " Then translate = "ģ�飺"
        If src = "Details:" Then translate = "���飺"
        If src = "Time: " Then translate = "ʱ�䣺"
        If src = "Double click to close the window after 10 seconds. Press PrtSc to take a capture (THIS OPERATION WILL COVER YOUR CLIPBOARD). " Then translate = "10���˫���رձ����ڡ���PrtSc����ͼ���˲������������ļ����塣"
        
        If src = "FreeExam" Then translate = "���ɿ�"
        If src = "Create" Then translate = "�½�"
        If src = "AutoCls" Then translate = "�Զ����"
        If src = "The size of the preview is NEAR the actual size." Then translate = "Ԥ���ߴ�ӽ���ʵ�ߴ�"
        If src = "Contents can't be shown" Then translate = "���ݲ��ܱ���ʾ"
        If src = "Access Denied - You don't have enough privilege to access here. By the way, there is nothing interesting." Then translate = "���ʾܾ� - ��û���㹻Ȩ�޷������˳��һ�ᣬ����û�к���ġ�"
        If src = "DevWin/PrivCheck" Then translate = "��������/Ȩ�޼��"
        If src = "Authentication Passed." Then translate = "��֤ͨ��"
        
        
        
        If src = "Click to edit" Then translate = "����༭"
        If src = "Disabled" Then translate = "��ͣ��"
        
        
        If src = "[SysErr]" Then translate = "[ϵͳ����]"
        If src = "Loading Fonts(" Then translate = "����������("
        
        
        If src = "Invaild Format." Then translate = "��Ч��ʽ"
        If src = "The size you've input is too large" Then translate = "����������"
        
        
        
        
        
        
        
        If src = "Invaild Format" Then translate = "��Ч��ʽ"
        
        
        
        
        If src = "The input A is too large that we can't process that." Then translate = "����A̫�������������޷���������"
        
        
        If src = "The input B is too large that we can't process that." Then translate = "����B̫�������������޷���������"
        
        If src = "The input C is too large that we can't process that." Then translate = "����C̫�������������޷���������"
        
        If src = "The input D is too large that we can't process that." Then translate = "����D̫�������������޷���������"
        
        
        
        
        
        
        
        
        If src = "Generated and saved in the following path: " Then translate = "�����ɲ�����������·��"
        If src = "FontName: " Then translate = "��������"
        If src = "FontSize: " Then translate = "�ֺţ�"
        If src = "Bold: " Then translate = "���壺"
        If src = "italic: " Then translate = "б�壺"
        If src = "Alignment: " Then translate = "���룺"
        If src = "* For italic/Bold: 1 is true, 0 is false" Then translate = "����б��ʹ��壺1��ʾ������0��ʾ�ر�"
        If src = "Text: " Then translate = "�ı���"
        If src = "DisabledImage: " Then translate = "ͼ���ѽ��ã�"
        If src = "False" Then translate = "��"
        If src = "Drive: " Then translate = "���̣�"
        If src = "Path: " Then translate = "·����"
        If src = "File: " Then translate = "�ļ���"
        If src = "Position: " Then translate = "λ�ã�"
        If src = "[Image Information Unavailable]" Then translate = "[ͼ����Ϣ������]"
        
        
        
        
        
        
        
        If src = "Nothing can be previewed." Then translate = "û���ܱ�Ԥ���Ķ���"
        If src = "You have not selected the position of the image." Then translate = "�㻹û��ѡ��ͼ��λ��"
        If src = "Your input is so large that we can't process it." Then translate = "���������������������޷�������"
        If src = "[TYPE=RUNTIME_ERROR][ERRORID=" Then translate = "[����=����ʱ����][�����="
        If src = "][ERRDESC.=" Then translate = "][��������="
        
        If src = "Auto split line is unsupportted for alignment mode 1 or 2." Then translate = "�Զ����ж��ڶ���ģʽ1��2����֧��"
        If src = "The input will be split into multi parts" Then translate = "���뽫�ֳɶ���"
        If src = "select parts that you want to preview in the list." Then translate = "���б���ѡ����ҪԤ���Ĳ���"
        If src = "The target size is too large, we are unable to process it." Then translate = "Ŀ��ߴ�̫�������������޷�������"
        
        If src = "from " Then translate = "��"
        If src = " to " Then translate = " ��"
        If src = "Image which tracknumber=" Then translate = "׷�ٺ�������ֵ��ͼ��"
        If src = " not found" Then translate = " δ�ҵ�"
        If src = "CNT=" Then translate = "����="
        
        
        
        
        If src = "Initiating" Then translate = "������"
        If src = "Special Input" Then translate = "��������"
        If src = "Format" Then translate = "��ʽ��"
    End If
End Function
