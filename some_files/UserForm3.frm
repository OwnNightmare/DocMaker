VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "�������� �����"
   ClientHeight    =   9636.001
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   16128
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    
    KrimskPath = "Z:\�������\ActMakerProject"
    KarskPath = "Z:\�������\ActMakerProject"
    PsekPath = "Z:\�������\ActMakerProject"
    
    
    ' @�������� �� ������ ����
    With Me
    
        ComboBoxSikn.AddItem ("��������")
        ComboBoxSikn.AddItem ("����������")
        ComboBoxSikn.AddItem ("�������")
        
        '@�������� �� ������ ����
        ComboBoxSikn = "��������"
        'ComboBoxSikn = "����������"
        'ComboBoxSikn = "�������"
        
        ComboBoxTopic.AddItem ("�������� �����")
        ComboBoxTopic.AddItem ("���")
        ComboBoxTopic.AddItem ("���")
        ComboBoxTopic.AddItem ("���")
        ComboBoxTopic.AddItem ("�������������")
        ComboBoxTopic.AddItem ("������� �� �������")
        ComboBoxTopic.AddItem ("�������������")
        ComboBoxTopic.AddItem ("������� ���")
        ComboBoxTopic.AddItem ("�� ������ ��� ���")
        ComboBoxTopic.AddItem ("������������� �������")
        ComboBoxTopic.AddItem ("���������(������)")
        ComboBoxTopic.AddItem ("����� ����")
        ComboBoxTopic.AddItem ("������")
        
        FillDateBoxes '���������� ����� � ������ ����� ���������� �������
        ComboBoxDay = day(Now)
        ComboBoxMonth = month(Now)
        ComboBoxYear = year(Now)
        
        ComboBoxNefteauto.AddItem ("�������� �.�.")
        ComboBoxNefteauto.AddItem ("������� �.�.")
        ComboBoxNefteauto.AddItem ("������� �.�.")
        ComboBoxNefteauto.AddItem ("������� �.�.")
        
        ComboBoxPover.AddItem ("�������� �.�.")
        ComboBoxPover.AddItem ("�������� �.�.")
        
        CheckBoxPover = False
        
    End With
    
    
End Sub

Private Sub ComboBoxSikn_Change()
    
    ComboBoxRosneft.Clear
    ComboBoxTransneft.Clear
    CheckBoxSrm1 = False
    CheckBoxSrm2 = False
    CheckBoxSrm3 = False
    CheckBoxSrm4 = False
    
    If ComboBoxSikn = "��������" Then

        LabelSiknNum = "834"
        
        ComboBoxRosneft.AddItem ("����� �.�.")
        ComboBoxRosneft.AddItem ("��������� �.�.")
        ComboBoxRosneft.AddItem ("������ �.�.")
        ComboBoxRosneft.AddItem ("������� �.�.")
        ComboBoxRosneft.AddItem ("������� �.�.")
        
        ComboBoxTransneft.AddItem ("������� �.�.")
        ComboBoxTransneft.AddItem ("�������� �.�.")
        ComboBoxTransneft.AddItem ("�������� �.�.")
        ComboBoxTransneft.AddItem ("��������� �.�.")
        ComboBoxTransneft.AddItem ("������� �.�.")
        
        CheckBoxSrm3.Enabled = True
        CheckBoxSrm4.Enabled = True
        
    ElseIf ComboBoxSikn = "����������" Then
        
        LabelSiknNum = "837"
        
        ComboBoxRosneft.AddItem ("��������� �.�.")
        ComboBoxRosneft.AddItem ("�������� �.�.")
        ComboBoxRosneft.AddItem ("������� �.�.")
        
        ComboBoxTransneft.AddItem ("��������� �.�.")
        ComboBoxTransneft.AddItem ("�������� �.�.")
        ComboBoxTransneft.AddItem ("������� �.�.")
        ComboBoxTransneft.AddItem ("�������� �.�.")
        ComboBoxTransneft.AddItem ("�������� �.�.")
        
        CheckBoxSrm3.Enabled = False
        CheckBoxSrm4.Enabled = False
    
    ElseIf ComboBoxSikn = "�������" Then
    
        LabelSiknNum = "835"
        
        ComboBoxRosneft.AddItem ("�������� �.�.")
        ComboBoxRosneft.AddItem ("�������� �.�.")
        ComboBoxRosneft.AddItem ("�������� �.�.")
        ComboBoxRosneft.AddItem ("�������� �.�.")
        
        ComboBoxTransneft.AddItem ("������� �.�.")
        
        CheckBoxSrm3.Enabled = False
        CheckBoxSrm4.Enabled = False
    End If
    
    End Sub

Private Sub ComboBoxTopic_AfterUpdate()
        
        PoverkaCheck.Visible = False
        PoverkaCheck = False
        CheckBoxPover = False
        LabelMassomers.Visible = False
        CheckBoxSrm1 = False
        CheckBoxSrm1.Visible = False
        CheckBoxSrm2 = False
        CheckBoxSrm2.Visible = False
        CheckBoxSrm3 = False
        CheckBoxSrm3.Visible = False
        CheckBoxSrm4 = False
        CheckBoxSrm4.Visible = False
        LabelCurrentKf.Visible = False
        LabelNewKf.Visible = False

    If ComboBoxTopic = "������� ���" Then
        CheckBoxPover = True
        LabelMassomers.Visible = True
        CheckBoxSrm1.Visible = True
        CheckBoxSrm2.Visible = True
        CheckBoxSrm3.Visible = True
        CheckBoxSrm4.Visible = True
        LabelCurrentKf.Visible = True
        LabelNewKf.Visible = True
        
    ElseIf ComboBoxTopic = "�������������" Then
        PoverkaCheck.Visible = True
        LabelMassomers.Visible = True
        CheckBoxSrm1.Visible = True
        CheckBoxSrm2.Visible = True
        CheckBoxSrm3.Visible = True
        CheckBoxSrm4.Visible = True
    ElseIf ComboBoxTopic = "�� ������ ��� ���" Then
        LabelMassomers.Visible = True
        CheckBoxSrm1.Visible = True
        CheckBoxSrm2.Visible = True
        CheckBoxSrm3.Visible = True
        CheckBoxSrm4.Visible = True
    End If
    
    If ComboBoxTopic = "������� �� �������" Then
        CheckBoxTransneft = False
    Else
        CheckBoxTransneft = True
    End If
    
End Sub
Private Sub CheckBoxTransneft_Change()
    If CheckBoxTransneft = True Then
        ComboBoxTransneft.Enabled = True
    Else
        ComboBoxTransneft.Enabled = False
    End If
End Sub

Private Sub CheckBoxPover_Change()

    If CheckBoxPover = True Then
        ComboBoxPover.Enabled = True
    Else
        ComboBoxPover.Enabled = False
    End If
    
End Sub

Private Sub ComboBoxNefteauto_AfterUpdate()
    Dim man As String
    man = ComboBoxNefteauto
    
    If man <> "" Then
        TextBoxNefteautoPost.Visible = True
        If man = "������� �.�." Then
            TextBoxNefteautoPost = "������� �����"
        ElseIf man = "������� �.�." Then
            TextBoxNefteautoPost = "������� �� ����������"
        ElseIf man = "�������� �.�." Then
            TextBoxNefteautoPost = "��������� �������"
        ElseIf man = "������� �.�." Then
            TextBoxNefteautoPost = "��������� ���������"
        Else
            TextBoxNefteautoPost = "�������"
        End If
    End If
    

End Sub

Private Sub ComboBoxPover_AfterUpdate()
    If ComboBoxPover <> "" Then
        TextBoxPoverPost.Visible = True
        If ComboBoxPover = "�������� �.�." Then
            TextBoxPoverPost = "������� ������� �� ����������"
        ElseIf ComboBoxPover = "�������� �.�." Then
            TextBoxPoverPost = "������� �� ����������"
        Else
            TextBoxPoverPost = "������� �� ����������"
        End If
    End If
    
End Sub

Private Sub ComboBoxRosneft_AfterUpdate()
    If ComboBoxRosneft <> "" Then
        TextBoxRosneftPost.Visible = True
        TextBoxRosneftPost = "��������"
    End If
    
End Sub
Private Sub ComboBoxTransneft_AfterUpdate()
    If ComboBoxTransneft <> "" Then
        TextBoxTransneftPost.Visible = True
        If ComboBoxTransneft = "�������� �.�." Then
           TextBoxTransneftPost = "��������� �������"
        ElseIf ComboBoxTransneft = "������� �.�." Then
            TextBoxTransneftPost = "������� ����� �������"
        Else
            If ComboBoxSikn = "��������" Then
                TextBoxTransneftPost = "�������� ��� �������� ����"
            ElseIf ComboBoxSikn = "����������" Then
                TextBoxTransneftPost = "�������� ��� " & QuoteSmh("����������")
            ElseIf ComboBoxSikn = "�������" Then
                TextBoxTransneftPost = "�������� ��� " & QuoteSmh("�������") '@�������� �� �������
            End If
        End If
    End If
End Sub

' ���������� ���������� � ������ ������ ������ ���� ���� "������� ���"
Private Sub CheckBoxSrm1_Change()
    If CheckBoxSrm1 And Me.ComboBoxTopic = "������� ���" Then
        Me.LabelSrm1.Visible = True
        Me.TextBoxSrm1OldKf.Visible = True
        Me.TextBoxSrm1NewKf.Visible = True
    Else
        Me.LabelSrm1.Visible = False
        Me.TextBoxSrm1OldKf.Visible = False
        Me.TextBoxSrm1NewKf.Visible = False
    End If

End Sub
Private Sub CheckBoxSrm2_Change()
    If CheckBoxSrm2 And Me.ComboBoxTopic = "������� ���" Then
        Me.TextBoxSrm2OldKf.Visible = True
        Me.TextBoxSrm2NewKf.Visible = True
        Me.LabelSrm2.Visible = True
    Else
        Me.TextBoxSrm2OldKf.Visible = False
        Me.TextBoxSrm2NewKf.Visible = False
        Me.LabelSrm2.Visible = False
    End If

End Sub
Private Sub CheckBoxSrm3_Change()
    If CheckBoxSrm3 And Me.ComboBoxTopic = "������� ���" Then
        Me.TextBoxSrm3OldKf.Visible = True
        Me.TextBoxSrm3NewKf.Visible = True
        Me.LabelSrm3.Visible = True
    Else
        Me.TextBoxSrm3OldKf.Visible = False
        Me.TextBoxSrm3NewKf.Visible = False
        Me.LabelSrm3.Visible = False
    End If

End Sub

Private Sub CheckBoxSrm4_Change()
    If CheckBoxSrm4 And Me.ComboBoxTopic = "������� ���" Then
        Me.TextBoxSrm4OldKf.Visible = True
        Me.TextBoxSrm4NewKf.Visible = True
        Me.LabelSrm4.Visible = True
    Else
        Me.TextBoxSrm4OldKf.Visible = False
        Me.TextBoxSrm4NewKf.Visible = False
        Me.LabelSrm4.Visible = False
    End If

End Sub

' ��������� ����� � ������.
Private Sub ButtonOpenFolder_Click()
    OpenFolderActs (UserForm1.ComboBoxSikn)
End Sub

Private Sub ButtonCreateDoc_Click()
    
    Dim sikn_name, autocmb, roscmb, transcmb, povcmb As String
    Dim autochc, roschc, transchc, povchc As Boolean
    Dim theme As String
    Dim response As VbMsgBoxResult
    
    sikn_name = Me.ComboBoxSikn
    autocmb = Me.ComboBoxNefteauto
    roscmb = Me.ComboBoxRosneft
    transcmb = Me.ComboBoxTransneft
    transchc = Me.CheckBoxTransneft
    povcmb = Me.ComboBoxPover
    povchc = Me.CheckBoxPover
    theme = Me.ComboBoxTopic
    
    If ComboBoxTopic = "" Then
        response = MsgBox("�� ������� ���� ����!", vbExclamation)
        Exit Sub
    End If
    
    If (autocmb = "") Or (roscmb = "") Or (transcmb = "" And transchc) Or (povcmb = "" And povchc) Then
         response = MsgBox("������������� ����� �� ������ �� ������. ���������� ������������ ����?", vbYesNo)
         If response = vbNo Then
            Exit Sub
        End If
    End If
            
    If theme = "������� ���" Then
        CreateDocPoverkaSrm
    ElseIf theme = "�������� �����" Then
        CreateDocProtections (sikn_name)
    ElseIf theme = "���" Then
        CreateDocArm (sikn_name)
    ElseIf theme = "���" Then
        CreateDocIbp (sikn_name)
    ElseIf theme = "�������������" Then
        CreateDocProb (sikn_name)
    ElseIf theme = "���" Then
        CreateDocIfs (sikn_name)
    ElseIf theme = "�� ������ ��� ���" Then
        CreateDocSrmFail (sikn_name)
    ElseIf theme = "������� �� �������" Then
        CreateDocPribori (sikn_name)
    ElseIf theme = "������������� �������" Then
        CreateDocTimeSync (sikn_name)
    ElseIf theme = "���������(������)" Then
        CreateDocPlotChistka (sikn_name)
    ElseIf theme = "�������������" And sikn_name = "��������" Then
        CreateDocGermet
    ElseIf theme = "������" Then
        CreateDocFree (sikn_name)
    ElseIf theme = "����� ����" Then
        CreateDocZamerShara (sikn_name)
    Else
        MsgBox "���� ��� ���� ��� �� ����������("
    End If

End Sub
