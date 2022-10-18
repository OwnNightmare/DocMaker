VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Создание актов"
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
    
    
    KrimskPath = "Z:\Проекты\ActMakerProject"
    KarskPath = "Z:\Проекты\ActMakerProject"
    PsekPath = "Z:\Проекты\ActMakerProject"
    
    
    ' @Заменить на нужные пути
    With Me
    
        ComboBoxSikn.AddItem ("Крымский")
        ComboBoxSikn.AddItem ("Псекупский")
        ComboBoxSikn.AddItem ("Карский")
        
        '@Заменить на нужный узел
        ComboBoxSikn = "Крымский"
        'ComboBoxSikn = "Псекупский"
        'ComboBoxSikn = "Карский"
        
        ComboBoxTopic.AddItem ("Проверка защит")
        ComboBoxTopic.AddItem ("АРМ")
        ComboBoxTopic.AddItem ("ИБП")
        ComboBoxTopic.AddItem ("ИФС")
        ComboBoxTopic.AddItem ("Пробоотборник")
        ComboBoxTopic.AddItem ("Приборы на поверку")
        ComboBoxTopic.AddItem ("Герметичность")
        ComboBoxTopic.AddItem ("Поверка СРМ")
        ComboBoxTopic.AddItem ("Не прошел КМХ СРМ")
        ComboBoxTopic.AddItem ("Синхронизация времени")
        ComboBoxTopic.AddItem ("Плотномер(чистка)")
        ComboBoxTopic.AddItem ("Замер шара")
        ComboBoxTopic.AddItem ("Другое")
        
        FillDateBoxes 'Напоолняем боксы с датами всеми возможными числами
        ComboBoxDay = day(Now)
        ComboBoxMonth = month(Now)
        ComboBoxYear = year(Now)
        
        ComboBoxNefteauto.AddItem ("Харченко И.И.")
        ComboBoxNefteauto.AddItem ("Олейник И.В.")
        ComboBoxNefteauto.AddItem ("Борисов Ю.Г.")
        ComboBoxNefteauto.AddItem ("Мотыжев В.В.")
        
        ComboBoxPover.AddItem ("Ефремова М.А.")
        ComboBoxPover.AddItem ("Запашний В.В.")
        
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
    
    If ComboBoxSikn = "Крымский" Then

        LabelSiknNum = "834"
        
        ComboBoxRosneft.AddItem ("Ляпка Т.Н.")
        ComboBoxRosneft.AddItem ("Богданова С.Ю.")
        ComboBoxRosneft.AddItem ("Калита С.Ю.")
        ComboBoxRosneft.AddItem ("Тихолоз И.С.")
        ComboBoxRosneft.AddItem ("Попкова Т.А.")
        
        ComboBoxTransneft.AddItem ("Абалова А.А.")
        ComboBoxTransneft.AddItem ("Чуйченко Ю.А.")
        ComboBoxTransneft.AddItem ("Комарова Е.Д.")
        ComboBoxTransneft.AddItem ("Алексеева Ю.А.")
        ComboBoxTransneft.AddItem ("Синявин Ф.Г.")
        
        CheckBoxSrm3.Enabled = True
        CheckBoxSrm4.Enabled = True
        
    ElseIf ComboBoxSikn = "Псекупский" Then
        
        LabelSiknNum = "837"
        
        ComboBoxRosneft.AddItem ("Корнилова Е.Ю.")
        ComboBoxRosneft.AddItem ("Костенко М.А.")
        ComboBoxRosneft.AddItem ("Ширнина Ю.В.")
        
        ComboBoxTransneft.AddItem ("Черникова Н.Г.")
        ComboBoxTransneft.AddItem ("Федоткин В.В.")
        ComboBoxTransneft.AddItem ("Кадяева Е.С.")
        ComboBoxTransneft.AddItem ("Сорокина Е.Ю.")
        ComboBoxTransneft.AddItem ("Шандыбин Д.Г.")
        
        CheckBoxSrm3.Enabled = False
        CheckBoxSrm4.Enabled = False
    
    ElseIf ComboBoxSikn = "Карский" Then
    
        LabelSiknNum = "835"
        
        ComboBoxRosneft.AddItem ("Гусарова М.Ф.")
        ComboBoxRosneft.AddItem ("Лопатина Н.Н.")
        ComboBoxRosneft.AddItem ("Лопатина Т.Н.")
        ComboBoxRosneft.AddItem ("Угрюмова В.А.")
        
        ComboBoxTransneft.AddItem ("Сметана М.В.")
        
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

    If ComboBoxTopic = "Поверка СРМ" Then
        CheckBoxPover = True
        LabelMassomers.Visible = True
        CheckBoxSrm1.Visible = True
        CheckBoxSrm2.Visible = True
        CheckBoxSrm3.Visible = True
        CheckBoxSrm4.Visible = True
        LabelCurrentKf.Visible = True
        LabelNewKf.Visible = True
        
    ElseIf ComboBoxTopic = "Герметичность" Then
        PoverkaCheck.Visible = True
        LabelMassomers.Visible = True
        CheckBoxSrm1.Visible = True
        CheckBoxSrm2.Visible = True
        CheckBoxSrm3.Visible = True
        CheckBoxSrm4.Visible = True
    ElseIf ComboBoxTopic = "Не прошел КМХ СРМ" Then
        LabelMassomers.Visible = True
        CheckBoxSrm1.Visible = True
        CheckBoxSrm2.Visible = True
        CheckBoxSrm3.Visible = True
        CheckBoxSrm4.Visible = True
    End If
    
    If ComboBoxTopic = "Приборы на поверку" Then
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
        If man = "Борисов Ю.Г." Then
            TextBoxNefteautoPost = "инженер КИПиА"
        ElseIf man = "Олейник И.В." Then
            TextBoxNefteautoPost = "инженер по метрологии"
        ElseIf man = "Харченко И.И." Then
            TextBoxNefteautoPost = "начальник участка"
        ElseIf man = "Мотыжев В.В." Then
            TextBoxNefteautoPost = "начальник отделения"
        Else
            TextBoxNefteautoPost = "инженер"
        End If
    End If
    

End Sub

Private Sub ComboBoxPover_AfterUpdate()
    If ComboBoxPover <> "" Then
        TextBoxPoverPost.Visible = True
        If ComboBoxPover = "Ефремова М.А." Then
            TextBoxPoverPost = "ведущий инженер по метрологии"
        ElseIf ComboBoxPover = "Запашний В.В." Then
            TextBoxPoverPost = "инженер по метрологии"
        Else
            TextBoxPoverPost = "инженер по метрологии"
        End If
    End If
    
End Sub

Private Sub ComboBoxRosneft_AfterUpdate()
    If ComboBoxRosneft <> "" Then
        TextBoxRosneftPost.Visible = True
        TextBoxRosneftPost = "оператор"
    End If
    
End Sub
Private Sub ComboBoxTransneft_AfterUpdate()
    If ComboBoxTransneft <> "" Then
        TextBoxTransneftPost.Visible = True
        If ComboBoxTransneft = "Шандыбин Д.Г." Then
           TextBoxTransneftPost = "начальник УЭСАиТМ"
        ElseIf ComboBoxTransneft = "Синявин Ф.Г." Then
            TextBoxTransneftPost = "инженер КИПиА УЭСАиТМ"
        Else
            If ComboBoxSikn = "Крымский" Then
                TextBoxTransneftPost = "оператор ПСП Крымской ЛПДС"
            ElseIf ComboBoxSikn = "Псекупский" Then
                TextBoxTransneftPost = "оператор НПС " & QuoteSmh("Псекупская")
            ElseIf ComboBoxSikn = "Карский" Then
                TextBoxTransneftPost = "оператор НПС " & QuoteSmh("Карская") '@изменить на карском
            End If
        End If
    End If
End Sub

' Показывать текстбоксы с вводом коэфов только если тема "Поверка СРМ"
Private Sub CheckBoxSrm1_Change()
    If CheckBoxSrm1 And Me.ComboBoxTopic = "Поверка СРМ" Then
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
    If CheckBoxSrm2 And Me.ComboBoxTopic = "Поверка СРМ" Then
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
    If CheckBoxSrm3 And Me.ComboBoxTopic = "Поверка СРМ" Then
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
    If CheckBoxSrm4 And Me.ComboBoxTopic = "Поверка СРМ" Then
        Me.TextBoxSrm4OldKf.Visible = True
        Me.TextBoxSrm4NewKf.Visible = True
        Me.LabelSrm4.Visible = True
    Else
        Me.TextBoxSrm4OldKf.Visible = False
        Me.TextBoxSrm4NewKf.Visible = False
        Me.LabelSrm4.Visible = False
    End If

End Sub

' Открывает папку с актами.
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
        response = MsgBox("Не выбрана тема акта!", vbExclamation)
        Exit Sub
    End If
    
    If (autocmb = "") Or (roscmb = "") Or (transcmb = "" And transchc) Or (povcmb = "" And povchc) Then
         response = MsgBox("Представитель одной из сторон не выбран. Продолжить формирование акта?", vbYesNo)
         If response = vbNo Then
            Exit Sub
        End If
    End If
            
    If theme = "Поверка СРМ" Then
        CreateDocPoverkaSrm
    ElseIf theme = "Проверка защит" Then
        CreateDocProtections (sikn_name)
    ElseIf theme = "АРМ" Then
        CreateDocArm (sikn_name)
    ElseIf theme = "ИБП" Then
        CreateDocIbp (sikn_name)
    ElseIf theme = "Пробоотборник" Then
        CreateDocProb (sikn_name)
    ElseIf theme = "ИФС" Then
        CreateDocIfs (sikn_name)
    ElseIf theme = "Не прошел КМХ СРМ" Then
        CreateDocSrmFail (sikn_name)
    ElseIf theme = "Приборы на поверку" Then
        CreateDocPribori (sikn_name)
    ElseIf theme = "Синхронизация времени" Then
        CreateDocTimeSync (sikn_name)
    ElseIf theme = "Плотномер(чистка)" Then
        CreateDocPlotChistka (sikn_name)
    ElseIf theme = "Герметичность" And sikn_name = "Крымский" Then
        CreateDocGermet
    ElseIf theme = "Другое" Then
        CreateDocFree (sikn_name)
    ElseIf theme = "Замер шара" Then
        CreateDocZamerShara (sikn_name)
    Else
        MsgBox "Пока что этот акт не реализован("
    End If

End Sub
