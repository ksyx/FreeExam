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
        InputWin.Caption = "����"
        InputWin.Label1.AutoSize = True
InputWin.Label28.AutoSize = True
InputWin.Label1.Alignment = 2
InputWin.Label28.Alignment = 2
    InputWin.Label1.Font = "����"
InputWin.Label28.Font = "����"
InputWin.Label68.Font = "����"
        InputWin.Label1.Caption = "ȡ��"
        InputWin.Label28.Caption = "ȷ��"
        InputWin.Label68.Caption = "�Զ��ո�"
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
Integrated.Label1.Font = "����"
Integrated.Label2.Font = "����"
Integrated.Label4.Font = "����"
Integrated.Label8.Font = "����"
Integrated.Label7.Font = "����"
Integrated.Label5.Font = "����"
        Integrated.Label1.Caption = "����"
        Integrated.Label2.Caption = "���"
        Integrated.Label4.Caption = "��Enter���룬��Esc�˳�"
        Integrated.Label8.Caption = "����"
        Integrated.Label7.Caption = "���"
        Integrated.Label5.Caption = "��Enter���룬��Esc�˳�"
        End If
        If src = "PageSettings" Then
        PageSettings.Caption = "ҳ������"
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
PageSettings.Message.Font = "����"
        PageSettings.Frame1.Font = "����"
PageSettings.Label1.Font = "����"
PageSettings.Frame2.Font = "����"
PageSettings.Label4.Font = "����"
PageSettings.Label3.Font = "����"
PageSettings.Label5.Font = "����"
PageSettings.Label6.Font = "����"
PageSettings.Label11.Font = "����"
PageSettings.PreviewButton.Font = "����"
PageSettings.Label2.Font = "����"
        PageSettings.Frame1.Caption = "ҳ������"
        PageSettings.Label1.Caption = "ע�⣺8K�Ŀ������ʵ��ȵ�һ�룬��ѡ���8K�ߴ��Ѽ��롣��ʾ�ĳߴ��ʽΪ�ߴ����ƣ����������ݲ�֧���Զ���ߴ硣"
        PageSettings.Frame2.Caption = "�߾�"
        PageSettings.Label4.Caption = "��"
        PageSettings.Label3.Caption = "��"
        PageSettings.Label11.Caption = "���棺�������ȷ��֮���㽫���ܷ������"
        PageSettings.Label5.Caption = "��"
        PageSettings.Label6.Caption = "��"
        PageSettings.PreviewButton.Caption = "Ԥ��"
        PageSettings.Label2.Caption = "ȷ��"
        End If
        If src = "MainFrm" Then
        MainFrm.Caption = "���ɿ� ��������"
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
MainFrm.Label12.FontName = "����"
MainFrm.Label3.FontName = "����"
MainFrm.Label1.FontName = "����"
MainFrm.Label15.FontName = "����"
MainFrm.Label30.FontName = "����"
MainFrm.Label60.FontName = "����"
MainFrm.Label63.FontName = "����"
MainFrm.PreviewButton.FontName = "����"
MainFrm.Label22.FontName = "����"
MainFrm.Label35.FontName = "����"
MainFrm.Label38.FontName = "����"
MainFrm.Label43.FontName = "����"



MainFrm.Frame19.Font = "����"
MainFrm.Frame11.Font = "����"
MainFrm.Check22.Font = "����"
MainFrm.Check23.Font = "����"
MainFrm.Label61.Font = "����"
MainFrm.Label59.Font = "����"
MainFrm.Label23.Font = "����"
MainFrm.Label27.Font = "����"
MainFrm.Frame9.Font = "����"
MainFrm.Label24.Font = "����"
MainFrm.Label25.Font = "����"
MainFrm.Label33.Font = "����"
MainFrm.Label39.Font = "����"
MainFrm.Label42.Font = "����"
MainFrm.Frame20.Font = "����"
MainFrm.Label36.Font = "����"
MainFrm.Label37.Font = "����"
MainFrm.Label13.Font = "����"
MainFrm.Label14.Font = "����"
MainFrm.Frame17.Font = "����"
MainFrm.Label46.Font = "����"
MainFrm.Label47.Font = "����"
MainFrm.Label45.Font = "����"
MainFrm.Frame18.Font = "����"
MainFrm.Label52.Font = "����"
MainFrm.Label53.Font = "����"
MainFrm.Label54.Font = "����"
MainFrm.Label55.Font = "����"
MainFrm.Label56.Font = "����"
MainFrm.Label57.Font = "����"
MainFrm.Frame13.Font = "����"
MainFrm.Option1.Font = "����"
MainFrm.Option2.Font = "����"
MainFrm.Frame14.Font = "����"
MainFrm.Frame1.Font = "����"
MainFrm.Label7.Font = "����"
MainFrm.Label8.Font = "����"
MainFrm.Label9.Font = "����"
MainFrm.Label10.Font = "����"
MainFrm.Frame2.Font = "����"
MainFrm.Label9.Font = "����"
MainFrm.Label51.Font = "����"
MainFrm.Label17.Font = "����"
MainFrm.Frame6.Font = "����"
MainFrm.Frame7.Font = "����"
MainFrm.Label18.Font = "����"
MainFrm.Label19.Font = "����"
MainFrm.Frame10.Font = "����"
MainFrm.Label28.Font = "����"
MainFrm.Label29.Font = "����"
MainFrm.Check17.Font = "����"
MainFrm.Label10.Font = "����"
MainFrm.Label11.Font = "����"
MainFrm.Label5.Font = "����"
MainFrm.Label6.Font = "����"
MainFrm.Message.Font = "����"
        MainFrm.Label12.Caption = "����/����"
MainFrm.Label3.Caption = "����"
MainFrm.Label1.Caption = "ͨ��"
MainFrm.Label15.Caption = "�ϲ�Ԥ��"
MainFrm.Label30.Caption = "��¼"
MainFrm.Label60.Caption = "�ϲ�"
MainFrm.Label63.Caption = "�б�"
MainFrm.PreviewButton.Caption = "�ı�"
MainFrm.Label22.Caption = "������"
MainFrm.Label35.Caption = "ͼƬ"
MainFrm.Label38.Caption = "�հ���"
MainFrm.Label43.Caption = "ѡ����"
MainFrm.Check16.Caption = "Ӣ��ģʽ"
        MainFrm.Frame19.Caption = "ҳü"
        MainFrm.Frame11.Caption = "ҳ��"
        MainFrm.Check22.Caption = "�ָ���"
        MainFrm.Check23.Caption = "�ָ���"
        MainFrm.Label61.Caption = "Ԥ��"
        MainFrm.Label59.Caption = "�ϲ�"
        MainFrm.Label23.Caption = "��С"
        MainFrm.Label27.Caption = "������-1��ʾ���"
        MainFrm.Frame9.Caption = "ѡ��"
        MainFrm.Label24.Caption = "Ԥ��"
        MainFrm.Label25.Caption = "����"
        MainFrm.Label33.Caption = "׷�ٺ�"
        MainFrm.Label39.Caption = "�ߴ�"
        MainFrm.Label42.Caption = "����"
        MainFrm.Frame20.Caption = "ѡ��һ��ͼ��"
        MainFrm.Label36.Caption = "Ԥ��"
        MainFrm.Label37.Caption = "����"
        MainFrm.Label13.Caption = "Ԥ��"
        MainFrm.Label14.Caption = "ɾ��"
        MainFrm.Frame17.Caption = "��ʽ"
        MainFrm.Label46.Caption = "����"
        MainFrm.Label47.Caption = "�ֺ�"
        MainFrm.Label45.Caption = "����"
        MainFrm.Frame18.Caption = "�ı� &^"
        MainFrm.Label52.Caption = "����༭"
        MainFrm.Label53.Caption = "����༭"
        MainFrm.Label54.Caption = "����༭"
        MainFrm.Label55.Caption = "����༭"
        MainFrm.Label56.Caption = "Ԥ��"
        MainFrm.Label57.Caption = "����"
        
        MainFrm.Label5.Caption = "����"
        MainFrm.Label6.Caption = "����"
        MainFrm.Frame13.Caption = "��¼�б�"
        MainFrm.Option1.Caption = "ҳ��"
        MainFrm.Option2.Caption = "��ʽ"
        MainFrm.Frame14.Caption = "����"
        MainFrm.Frame1.Caption = "��ʽ"
        MainFrm.Label7.Caption = "����"
        MainFrm.Label8.Caption = "�ֺ�"
        MainFrm.Label9.Caption = "����"
        MainFrm.Label10.Caption = "����"
        MainFrm.Frame2.Caption = "ѡ��"
        MainFrm.Label9.Caption = "�ı� &^"
        MainFrm.Label51.Caption = "����༭"
        MainFrm.Label17.Caption = "��ͼ"
        MainFrm.Frame6.Caption = "ѡ��һ��ͼƬ"
        MainFrm.Frame7.Caption = "ѡ��"
        MainFrm.Label18.Caption = "λ��"
        MainFrm.Label19.Caption = "����ͼ"
        MainFrm.Frame10.Caption = "��¼��"
        MainFrm.Label28.Caption = "��ʽ"
        MainFrm.Label29.Caption = "ҳ��"
        MainFrm.Check17.Caption = "�Զ�"
        MainFrm.Label10.Caption = "Ԥ��"
        MainFrm.Label11.Caption = "����"
        End If
        If src = "Preview" Then Preview.Message.FontName = "����"
End If
End Function


Function translate(src As String) As String
    
    
    
    
    translate = src
    If EnableTranslation = 1 Then
        If src = "Unable to get the page size that you've chosen." Then translate = "�޷���ȡ��ѡ���ҳ��ߴ�"
        If src = "Translating..." Then translate = "���ڷ���..."
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
        If src = "Completed on" Then translate = "����ɣ����ʱ��Ϊ"
        If src = "Auto split line is unsupportted for alignment mode 1 or 2." Then translate = "�Զ����ж��ڶ���ģʽ1��2����֧��"
        If src = "The input will be split into multi parts" Then translate = "���뽫�ֳɶ���"
        If src = "select parts that you want to preview in the list." Then translate = "���б���ѡ����ҪԤ���Ĳ���"
        If src = "The target size is too large, we are unable to process it." Then translate = "Ŀ��ߴ�̫�������������޷�������"
        If src = "Success with tracknumber" Then translate = "�ɹ���׷�ٺ�Ϊ"
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
