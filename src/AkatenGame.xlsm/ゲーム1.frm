VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Q�[��1 
   Caption         =   "�ԓ_����V�~�����[�^�["
   ClientHeight    =   11985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19185
   OleObjectBlob   =   "�Q�[��1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�Q�[��1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '�Z���W���[���ϐ�
    '�����������p�����[�^����������
    '�����������l�ۑ�����������
    Dim Hi As Integer '��
    Dim Zikan As Integer '����
    Dim TBen As Boolean '2���ԕ׋�
    Dim Skyoka As Integer '���Ȓl
    '���������������ۑ�����������
    Dim Youbi(7) As String '�j��
    Dim Hzikan(14) As String '�������ԕ\��
    Dim Kzikan(10) As String '�x�����ԕ\��
    Dim Kyoka(5) As String '����

Private Sub UserForm_Initialize()
    '�����������Q�[�������̕\������������
    �E�B���h�E2.Show
    '����������������̕ۑ�����������
    Youbi(1) = "���j��": Youbi(2) = "�Ηj��": Youbi(3) = "���j��": Youbi(4) = "�ؗj��": Youbi(5) = "���j��": Youbi(6) = "�y�j��": Youbi(7) = "���j��"
    Hzikan(1) = "1������": Hzikan(2) = "2������": Hzikan(3) = "3������": Hzikan(4) = "4������": Hzikan(5) = "5������": Hzikan(6) = "6������": _
        Hzikan(7) = "17:00": Hzikan(8) = "18:00": Hzikan(9) = "19:00": Hzikan(10) = "20:00": Hzikan(11) = "21:00": Hzikan(12) = "22:00": Hzikan(13) = "24:00": Hzikan(14) = "02:00"
    Kzikan(1) = "8:00": Kzikan(2) = "10:00": Kzikan(3) = "���H�^�C���I": Kzikan(4) = "15:00": Kzikan(5) = "17:00": _
        Kzikan(6) = "19:00": Kzikan(7) = "21:00": Kzikan(8) = "22:00": Kzikan(9) = "24:00": Kzikan(10) = "02:00"
    Kyoka(1) = "����": Kyoka(2) = "���w": Kyoka(3) = "�p��": Kyoka(4) = "�Ȋw": Kyoka(5) = "���{�j"
    '���������������l�̕ۑ�����������
    If Tuduki = False Or Save = False Then
    '���t�̕ۑ�
    Hi = 1
    �ʒm1.Caption = "������" & Hi & "���ځA" & Youbi(Hi Mod 8) & "�ł�"
    '���Ԃ̕ۑ�
    Zikan = 1
    �ʒm2.Caption = "����" & Hzikan(Zikan) & "�ł�"
    '���Ȃ̕ۑ�
    Call ����
    Call ���ݒl�\��
    '�m�������l
    Neru = 20 '�Q��m���i�����l20�j
    '�R�����g�E���O
    �R�����g.Caption = "�u���������T�ԂŐԓ_��������Ă݂���I�ԓ_���������K�v������͍̂���Ɛ��w�A�����ĉp��ł��Ƃ͎̂ċ��Ȃ��I�v"
    '�w�i
    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1.jpg")
    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1.jpg"
    Else
        Hi = HiSave
        Zikan = ZikanSave
        Skyoka = SkyokaSave
        ���O.Caption = NameSave
        �R�����g.Caption = CommentSave
        �ʒm1.Caption = TutiSave1
        �ʒm2.Caption = TutiSave2
        �ʒm3.Caption = TutiSave3
        Call ���ݒl�\��
        �w�i1.Picture = LoadPicture(URLSave)
    End If
    �g.Caption = ""
    Tuduki = False
End Sub

Private Sub �Z�[�u_Click()
    HiSave = Hi
    ZikanSave = Zikan
    SkyokaSave = Skyoka
    NameSave = ���O.Caption
    CommentSave = �R�����g.Caption
    TutiSave1 = �ʒm1.Caption
    TutiSave2 = �ʒm2.Caption
    TutiSave3 = �ʒm3.Caption
    ���j���[3.Show
End Sub

Private Sub ���[�h_Click()
    ���j���[4.Show
    If Save = True Then
        Hi = HiSave
        Zikan = ZikanSave
        Skyoka = SkyokaSave
        ���O.Caption = NameSave
        �R�����g.Caption = CommentSave
        �ʒm1.Caption = TutiSave1
        �ʒm2.Caption = TutiSave2
        �ʒm3.Caption = TutiSave3
        Call ���ݒl�\��
        �w�i1.Picture = LoadPicture(URLSave)
    End If
    �g.Caption = ""
End Sub

Private Sub ����_Click()
    '�{�v���O����
    If Hi < 6 Then '����
        Zikan = Zikan + 1
        If Zikan < 7 Then '�w�Z��
            TBen = False
            Call �w�Z�w�K
            Call ����
            �ʒm2.Caption = "����" & Hzikan(Zikan) & "�ł�"
        ElseIf Zikan < 11 Then '�i��j�ƕ�
            TBen = False
            Call �ƒ�w�K
            �ʒm3.Caption = ""
            �ʒm2.Caption = "����" & Hzikan(Zikan) & "�ł�"
            If Zikan = 10 Then '�����I������
                �Q�[��3.Show
            End If
        Else '�i���j�ƕ�
            If Zikan = 10 + Nzikan Then '�Q�鎞��
                Call ���t�i�s
            End If
            TBen = True
            Call �ƒ�w�K
            �ʒm2.Caption = "����" & Hzikan(Zikan) & "�ł�"
        End If
    ElseIf Hi < 8 Then '�x��
        Zikan = Zikan + 1
        �ʒm3.Caption = ""
        If Zikan < 7 Then '�i��j�ƕ�
            TBen = True
            Call �ƒ�w�K
            �ʒm2.Caption = "����" & Kzikan(Zikan) & "�ł�"
            If Zikan = 6 Then
                �Q�[��3.Show
            End If
        Else '�i���j�ƕ�
            If Zikan = 6 + Nzikan Then '�Q�鎞��
                Call ���t�i�s
            End If
            TBen = True
            Call �ƒ�w�K
            �ʒm2.Caption = "����" & Kzikan(Zikan) & "�ł�"
        End If
    ElseIf Hi = 8 Then '�e�X�g����
        Call ���t�i�s
    End If
    �g.Caption = ""
End Sub

Private Sub �w�Z�w�K()
    '�ϐ��錾
    Dim Sumaho As Integer '�X�}�z�g�p�ۑ�
    Dim KNeru As Integer '�Q��m������
    Dim Naisyoku As Integer '���E�΂��
    '�{�v���O����
    �Q�[��2.Show
    If Ktai = True Then '�X�}�z���g��
        Randomize '�X�}�z���΂��m���i�����l2����1�j
        Sumaho = Int((2 - 1 + 1) * Rnd + 1)
        If Sumaho = 1 Then '�X�}�z���΂�Ȃ�
            Randomize '�Q��m���i�����l20�j
            KNeru = Int((100 - 1 + 1) * Rnd + 1)
            If KNeru > Neru Then '�Q�Ȃ�
                Yaruki = Yaruki + 10
                If Kamoku = Skyoka Or Kamoku = 6 Then '���Ԋ��ƑI���Ȗڂ����l�̏ꍇ
                    Select Case Kamoku
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "6"
                            Yaruki = Yaruki + 15
                    End Select
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1.jpg"
                    �R�����g.Caption = "�u���̎��Ԃ͌v��ʂ�ɕ׋��ł����I�X�}�z���΂�Ȃ��ĕ׋����Ċy�����ȁI�v"
                Else '���Ԋ��ƑI���Ȗڂ����l�łȂ��ꍇ
                    Randomize '���E�m���i�����l2����1�j
                    Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                    If Naisyoku = 1 Then '���E���΂�Ȃ�
                        Select Case Kamoku
                            Case "1"
                                Kokugo = Kokugo + 10
                            Case "2"
                                Sugaku = Sugaku + 10
                            Case "3"
                                Eigo = Eigo + 10
                            Case "6"
                                Yaruki = Yaruki + 15
                        End Select
                        �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4.jpg")
                        URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4.jpg"
                        �R�����g.Caption = "�u���̎��Ԃ͌v��ʂ�ɕ׋��ł����I�X�}�z�����E���΂�Ȃ������I���̒��q�ŕ׋���i�߂Ă������I�v"
                    Else '���E���΂��
                        Yaruki = Yaruki - 10
                        �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg")
                        URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg"
                        �R�����g.Caption = "�u���܂����A�������̎��Ԃ͓��E���΂�ĕ׋��ǂ��낶��Ȃ������c�v"
                    End If
                End If
            Else '�Q��
                Yaruki = Yaruki - 5
                �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3.jpg")
                URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3.jpg"
                �R�����g.Caption = "��l���͐Q�Ă��܂���"
            End If
        Else  '�X�}�z���΂��
            Yaruki = Yaruki - 15
            �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg")
            URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg"
            �R�����g.Caption = "�u�ň����A�X�}�z���v������Ă��܂����B�V�����̂𔃂�Ȃ��Ɓc�v"
        End If
    Else '�X�}�z���g��Ȃ�
        Randomize '�Q��m���i�����l20�j
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '�Q�Ȃ�
            If Kamoku = Skyoka Or Kamoku = 6 Then '���Ԋ��ƑI���Ȗڂ����l�̏ꍇ
                Select Case Kamoku
                    Case "1"
                        Kokugo = Kokugo + 10
                    Case "2"
                        Sugaku = Sugaku + 10
                    Case "3"
                        Eigo = Eigo + 10
                    Case "6"
                        Yaruki = Yaruki + 15
                End Select
                �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1.jpg")
                URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1.jpg"
                �R�����g.Caption = "�u���̎��Ԃ͌v��ʂ�ɕ׋��ł����B����ς�׋��͂܂��߂ɂ��ׂ����I�v"
            Else '���Ԋ��ƑI���Ȗڂ����l�łȂ��ꍇ
                Randomize '���E�m���i�����l2����1�j
                Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                If Naisyoku = 1 Then '���E���΂�Ȃ�
                    Select Case Kamoku
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "6"
                            Yaruki = Yaruki + 15
                    End Select
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4.jpg"
                    �R�����g.Caption = "�u���̎��Ԃ͌v��ʂ�ɕ׋��ł����B���E���΂�Ȃ������I���̒��q�ŕ׋���i�߂Ă������I�v"
                Else '���E���΂��
                    Yaruki = Yaruki - 10
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg"
                    �R�����g.Caption = "�u���܂����A�������̎��Ԃ͓��E���΂�ĕ׋��ǂ��낶��Ȃ������c�v"
                End If
            End If
        Else '�Q��
            Yaruki = Yaruki - 5
            �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3.jpg")
            URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3.jpg"
            �R�����g.Caption = "��l���͐Q�Ă��܂���"
        End If
    End If
    Call ���ݒl�\��
    �g.Caption = ""
End Sub

Private Sub �ƒ�w�K()
    '�ϐ��錾
    Dim Sumaho As Integer '�X�}�z�g�p�ۑ�
    Dim KNeru As Integer '�Q��m������
    '�{�v���O����
    �Q�[��2.Show
    If Ktai = True Then '�X�}�z���g��
        Randomize '���Ԃ𖳑ʂɂ���i�����l20�j
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '���Ԃ𖳑ʂɂ��Ȃ�
            Yaruki = Yaruki + 10
            Select Case Kamoku
                Case "1"
                    Kokugo = Kokugo + 10
                    If TBen = True Then
                        Kokugo = Kokugo + 10
                    End If
                Case "2"
                    Sugaku = Sugaku + 10
                    If TBen = True Then
                        Sugaku = Sugaku + 10
                    End If
                Case "3"
                    Eigo = Eigo + 10
                    If TBen = True Then
                        Eigo = Eigo + 10
                    End If
                Case "6"
                    Yaruki = Yaruki + 15
                    If TBen = True Then
                        Yaruki = Yaruki + 15
                    End If
            End Select
            Select Case Zikan
                Case Is < 3
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-4.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-4.jpg"
                Case "3"
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4-4.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4-4.jpg"
                Case Is < 6
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-4.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-4.jpg"
                Case "6"
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2-4.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2-4.jpg"
                Case Is > 6
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3-4.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3-4.jpg"
            End Select
            �R�����g.Caption = "�u���̎��Ԃ͌v��ʂ�ɕ׋��ł����I�X�}�z���g���ĕ׋����Ċy�����Ē���ȁI�v"
        Else '���Ԃ𖳑ʂɂ���
            Yaruki = Yaruki - 15
            Select Case Zikan
                Case Is < 3
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-2.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-2.jpg"
                Case "3"
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4-2.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4-2.jpg"
                Case Is < 6
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-2.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-2.jpg"
                Case "6"
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2-2.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2-2.jpg"
                Case Is > 6
                    �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3-2.jpg")
                    URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3-2.jpg"
            End Select
            �R�����g.Caption = "�u�ň����A�X�}�z���g���������Ŏ��Ԃ�Q��Ă��܂����B�v"
        End If
    Else '�X�}�z���g��Ȃ�
        Select Case Kamoku
            Case "1"
                Kokugo = Kokugo + 10
                If TBen = True Then
                    Kokugo = Kokugo + 10
                End If
            Case "2"
                Sugaku = Sugaku + 10
                If TBen = True Then
                    Sugaku = Sugaku + 10
                End If
            Case "3"
                Eigo = Eigo + 10
                If TBen = True Then
                    Eigo = Eigo + 10
                End If
            Case "6"
                Yaruki = Yaruki + 15
                If TBen = True Then
                    Yaruki = Yaruki + 15
                End If
        End Select
        Select Case Zikan
            Case Is < 3
                �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-1.jpg")
                URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-1.jpg"
            Case "3"
                �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4-1.jpg")
                URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4-1.jpg"
            Case Is < 6
                �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-1.jpg")
                URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����1-1.jpg"
            Case "6"
                �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2-1.jpg")
                URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2-1.jpg"
            Case Is > 6
                �w�i1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3-1.jpg")
                URLSave = ThisWorkbook.Path & "\..\gfx\�Z�b�g\����3-1.jpg"
        End Select
        �R�����g.Caption = "�u���̎��Ԃ͌v��ʂ�ɕ׋��ł����B����ς�׋��͂܂��߂ɂ��ׂ����I�v"
    End If
    Call ���ݒl�\��
    �g.Caption = ""
End Sub

Private Sub ���ݒl�\��()
    '�ϐ��錾
    Dim KHozon As String
    Dim SHozon As String
    Dim EHozon As String
    Dim YHozon As String
    Dim X As Integer
    Dim Y As Integer
    '�{�v���O����
    Label1.Caption = "���� "
    X = Int(Kokugo / 25)
    If X <= 0 Then
        X = 1
    End If
    For Y = 1 To X
        Label1.Caption = Label1.Caption & "��"
    Next Y
    If X < 10 Then
        For Y = X + 1 To 10
            Label1.Caption = Label1.Caption & "��"
        Next Y
    End If
    Label2.Caption = "���w "
    X = Int(Sugaku / 25)
    If X <= 0 Then
        X = 1
    End If
    For Y = 1 To X
        Label2.Caption = Label2.Caption & "��"
    Next Y
    If X < 10 Then
        For Y = X + 1 To 10
            Label2.Caption = Label2.Caption & "��"
        Next Y
    End If
    Label3.Caption = "�p�� "
    X = Int(Eigo / 25)
    If X <= 0 Then
        X = 1
    End If
    For Y = 1 To X
        Label3.Caption = Label3.Caption & "��"
    Next Y
    If X < 10 Then
        For Y = X + 1 To 10
            Label3.Caption = Label3.Caption & "��"
        Next Y
    End If
    Label4.Caption = "�ًC "
    X = Int(Yaruki / 30)
    If X <= 0 Then
        X = 1
    End If
    For Y = 1 To X
        Label4.Caption = Label4.Caption & "��"
    Next Y
    If X < 10 Then
        For Y = X + 1 To 10
            Label4.Caption = Label4.Caption & "��"
        Next Y
    End If
    Label5.Caption = Kokugo
    Label6.Caption = Sugaku
    Label7.Caption = Eigo
    Label8.Caption = Yaruki
    �g.Caption = ""
End Sub

Private Sub ���t�i�s()
    Hi = Hi + 1
    If Hi < 8 Then '���j���`���j��
       �ʒm1.Caption = "������" & Hi & "���ځA" & Youbi(Hi Mod 8) & "�ł�"
    ElseIf Hi = 8 Then '�e�X�g����
        �ʒm1.Caption = "������" & "�e�X�g�ł��I"
    ElseIf Hi > 8 Then '���ʔ��\
        �_���v�Z
        Unload Me
        �E�B���h�E3.Show
    End If
    Zikan = 1
    If Hi < 8 Then
        Call ����
    End If
    �g.Caption = ""
End Sub

Private Sub �_���v�Z()
    '�ϐ��錾
    Dim KHozon As Integer
    Dim SHozon As Integer
    Dim EHozon As Integer
    '�{�v���O����
    If Yaruki >= 280 And Neru = 0 Then '���_100�_
        KHozon = Int((Kokugo / 250) * 100)
        If KHozon > 100 Then
            KHozon = 100
        End If
        SHozon = Int((Sugaku / 250) * 100)
        If SHozon > 100 Then
            SHozon = 100
        End If
        EHozon = Int((Eigo / 250) * 100)
        If EHozon > 100 Then
            EHozon = 100
        End If
    ElseIf Yaruki >= 240 And Neru <= 40 Then '���_70�_
        KHozon = Int((Kokugo / 250) * 70)
        SHozon = Int((Sugaku / 250) * 70)
        EHozon = Int((Eigo / 250) * 70)
    ElseIf Yaruki >= 180 And Neru <= 40 Then '���_40�_
        KHozon = Int((Kokugo / 250) * 40)
        SHozon = Int((Sugaku / 250) * 40)
        EHozon = Int((Eigo / 250) * 40)
    ElseIf Neru <= 40 Then '���_35�_
        KHozon = Int((Kokugo / 250) * 35)
        SHozon = Int((Sugaku / 250) * 35)
        EHozon = Int((Eigo / 250) * 35)
    Else '���_30�_
        KHozon = Int((Kokugo / 250) * 30)
        SHozon = Int((Sugaku / 250) * 30)
        EHozon = Int((Eigo / 250) * 30)
    End If
    Kokugo = KHozon
    Sugaku = SHozon
    Eigo = EHozon
    �g.Caption = ""
End Sub

Private Sub ����()
    '�ϐ��錾

    '�v���O����
    Randomize
    Skyoka = Int((5 - 1 + 1) * Rnd + 1)
    �ʒm3.Caption = "����" & Kyoka(Skyoka)
    �g.Caption = ""
End Sub

Private Sub �I��_Click()
    �E�B���h�E1.Show
End Sub
