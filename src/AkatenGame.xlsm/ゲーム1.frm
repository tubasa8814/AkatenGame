VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Q�[��1 
   Caption         =   "�ԓ_����V�~�����[�^�["
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
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
    Dim Yaruki As Integer '���C�p�����[�^
    Dim Kokugo As Integer '����p�����[�^
    Dim Sugaku As Integer '���w�p�����[�^
    Dim Eigo As Integer '�p��p�����[�^
    Dim Neru As Integer '�Q��m��
    '�����������l�ۑ�����������
    Dim Hi As Integer '��
    Dim Zikan As Integer '����
    Dim Skyoka As Integer '���Ȓl
    Dim TBen As Boolean '2���ԕ׋�
    Dim x As Integer
    Dim y As Integer
    '���������������ۑ�����������
    Dim Youbi(7) As String '�j��
    Dim Hzikan(14) As String '�������ԕ\��
    Dim Kzikan(10) As String '�x�����ԕ\��
    Dim Kyoka(5) As String '����
    '�Z�O���[�o���ϐ�
    Public Kamoku As Integer '�Ȗڕۑ�
    Public Ktai As Boolean '�X�}�z�g�p
    
Private Sub UserForm_Initialize()
    '�����������Q�[�������̕\������������
    ����.Show
    '����������������̕ۑ�����������
    Youbi(1) = "���j��": Youbi(2) = "�Ηj��": Youbi(3) = "���j��": Youbi(4) = "�ؗj��": Youbi(5) = "���j��": Youbi(6) = "�y�j��": Youbi(7) = "���j��"
    Hzikan(1) = "1������": Hzikan(2) = "2������": Hzikan(3) = "3������": Hzikan(4) = "4������": Hzikan(5) = "5������": Hzikan(6) = "6������": _
        Hzikan(7) = "17:00": Hzikan(8) = "18:00": Hzikan(9) = "19:00": Hzikan(10) = "20:00": Hzikan(11) = "21:00": Hzikan(12) = "22:00": Hzikan(13) = "24:00": Hzikan(14) = "02:00"
    Kzikan(1) = "8:00": Kzikan(2) = "10:00": Kzikan(3) = "���H�^�C���I": Kzikan(4) = "15:00": Kzikan(5) = "17:00": _
        Kzikan(6) = "19:00": Kzikan(7) = "21:00": Kzikan(8) = "22:00": Kzikan(9) = "24:00": Kzikan(10) = "02:00"
    Kyoka(1) = "����": Kyoka(2) = "���w": Kyoka(3) = "�p��": Kyoka(4) = "�Ȋw": Kyoka(5) = "���{�j"
    '���������������l�̕ۑ�����������
    '���t�̕ۑ�
    Hi = 1
    �ʒm1.Caption = "������" & Hi & "���ځA" & Youbi(Hi Mod 8) & "�ł�"
    '���Ԃ̕ۑ�
    Zikan = 1
    �ʒm2.Caption = "����" & Hzikan(Zikan) & "�ł�"
    '���Ȃ̕ۑ�
    Call ����
    '�p�����[�^�\���̕ۑ�
    ���C.Caption = "���C ��������������������"
    ����.Caption = "���� ��������������������"
    ���w.Caption = "���w ��������������������"
    �p��.Caption = "�p�� ��������������������"
    '�m�������l
    Neru = 20 '�Q��m���i�����l20�j
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
                Call �����I��
            End If
        Else '�i���j�ƕ�
            If Zikan = 10 + y Then '�Q�鎞��
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
                Call �����I��
            End If
        Else '�i���j�ƕ�
            If Zikan = 6 + y Then '�Q�鎞��
                Call ���t�i�s
            End If
            TBen = True
            Call �ƒ�w�K
            �ʒm2.Caption = "����" & Kzikan(Zikan) & "�ł�"
        End If
    ElseIf Hi = 8 Then '�e�X�g����
        Call ���t�i�s
    End If
End Sub

Private Sub �w�Z�w�K()
    '�ϐ��錾
    Dim Sumaho As Integer '�X�}�z�g�p�ۑ�
    Dim KNeru As Integer '�Q��m������
    Dim Naisyoku As Integer '���E�΂��
    Dim Gakusen As Integer '�w�K�I��ۑ�
    '�{�v���O����
    For x = 1 To 4 '�w�K�I��
        If Me.Controls("�׋�" & x).Value = True Then
            Gakusen = x
        End If
    Next x
    If Me.Controls("�X�}�z").Value = True Then '�X�}�z�g�p�m�F
        Ktai = True
    End If
    If Ktai = True Then '�X�}�z���g��
        Randomize '�X�}�z���΂��m���i�����l2����1�j
        Sumaho = Int((2 - 1 + 1) * Rnd + 1)
        If Sumaho = 1 Then '�X�}�z���΂�Ȃ�
            Randomize '�Q��m���i�����l20�j
            KNeru = Int((100 - 1 + 1) * Rnd + 1)
            If KNeru > Neru Then '�Q�Ȃ�
                Yaruki = Yaruki + 10
                If Gakusen = Skyoka Or Gakusen = 4 Then '���Ԋ��ƑI���Ȗڂ����l�̏ꍇ
                    Select Case Gakusen
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "4"
                            Yaruki = Yaruki + 5
                    End Select
                Else '���Ԋ��ƑI���Ȗڂ����l�łȂ��ꍇ
                    Randomize '���E�m���i�����l2����1�j
                    Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                    If Naisyoku = 1 Then '���E���΂�Ȃ�
                        Select Case Gakusen
                            Case "1"
                                Kokugo = Kokugo + 10
                            Case "2"
                                Sugaku = Sugaku + 10
                            Case "3"
                                Eigo = Eigo + 10
                            Case "4"
                                Yaruki = Yaruki + 5
                        End Select
                    Else '���E���΂��
                        Yaruki = Yaruki - 20
                    End If
                End If
            End If
        Else  '�X�}�z���΂��
            Yaruki = Yaruki - 35
        End If
    Else '�X�}�z���g��Ȃ�
        Randomize '�Q��m���i�����l20�j
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '�Q�Ȃ�
            If Gakusen = Skyoka Or Gakusen = 4 Then '���Ԋ��ƑI���Ȗڂ����l�̏ꍇ
                Select Case Gakusen
                    Case "1"
                        Kokugo = Kokugo + 10
                    Case "2"
                        Sugaku = Sugaku + 10
                    Case "3"
                        Eigo = Eigo + 10
                    Case "4"
                        Yaruki = Yaruki + 5
                End Select
            Else '���Ԋ��ƑI���Ȗڂ����l�łȂ��ꍇ
                Randomize '���E�m���i�����l2����1�j
                Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                If Naisyoku = 1 Then '���E���΂�Ȃ�
                    Select Case Gakusen
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "4"
                            Yaruki = Yaruki + 5
                    End Select
                Else '���E���΂��
                    Yaruki = Yaruki - 20
                End If
            End If
        End If
    End If
    Label1.Caption = Yaruki
    Label2.Caption = Kokugo
    Label3.Caption = Sugaku
    Label4.Caption = Eigo
End Sub

Private Sub �ƒ�w�K()
    '�ϐ��錾
    Dim KNeru As Integer '�Q��m������
    Dim Gakusen As Integer '�w�K�I��ۑ�
    '�{�v���O����
    For x = 1 To 4 '�w�K�I��
        If Me.Controls("�׋�" & x).Value = True Then
            Gakusen = x
        End If
    Next x
    If Me.Controls("�X�}�z").Value = True Then '�X�}�z�g�p�m�F
        Ktai = True
    End If
    If Ktai = True Then '�X�}�z���g��
        Randomize '�Q��m���i�����l20�j
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '�Q�Ȃ�
            Yaruki = Yaruki + 10
            Select Case Gakusen
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
                Case "4"
                    Yaruki = Yaruki + 5
            End Select
        Else '�Q��
            Yaruki = Yaruki - 25
        End If
    Else '�X�}�z���g��Ȃ�
        Select Case Gakusen
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
            Case "4"
                Yaruki = Yaruki + 5
        End Select
    End If
    Label1.Caption = Yaruki
    Label2.Caption = Kokugo
    Label3.Caption = Sugaku
    Label4.Caption = Eigo
End Sub

Private Sub �����I��()
    For x = 1 To 4
        If Me.Controls("����" & x).Value = True Then
            y = x
        End If
    Next x
    Select Case y
    Case "1"
        If Neru > 0 Then
            Neru = Neru - 5
        End If
    Case "2"
    Case "3"
        If Neru < 100 Then
            Neru = Neru + 5
        End If
    Case "4"
        If Neru < 95 Then
            Neru = Neru + 10
        End If
    End Select
End Sub

Private Sub ���t�i�s()
    Hi = Hi + 1
    If Hi < 8 Then '���j���`���j��
       �ʒm1.Caption = "������" & Hi & "���ځA" & Youbi(Hi Mod 8) & "�ł�"
    ElseIf Hi = 8 Then '�e�X�g����
        �ʒm1.Caption = "������" & "�e�X�g�ł��I"
    ElseIf Hi > 8 Then '���ʔ��\
        Unload Me
        ����.Show
    End If
    Zikan = 1
    If Hi < 8 Then
        Call ����
    End If
End Sub

Private Sub ����()
    Randomize
    Skyoka = Int((5 - 1 + 1) * Rnd + 1)
    �ʒm3.Caption = Kyoka(Skyoka)
End Sub

Private Sub �I��_Click()
    �m�F.Show
End Sub

Private Sub �Q�鎞��_Click()

End Sub
