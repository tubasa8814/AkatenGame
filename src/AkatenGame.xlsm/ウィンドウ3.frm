VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �E�B���h�E3 
   Caption         =   "�ԓ_����V�~�����[�^�[[���ʔ��\]"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19155
   OleObjectBlob   =   "�E�B���h�E3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�E�B���h�E3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Label1.Caption = "����" & Kokugo & "�_"
    Label2.Caption = "���w" & Sugaku & "�_"
    Label3.Caption = "�p��" & Eigo & "�_"
    If Kokugo >= 30 Then
        If Sugaku >= 30 Then
            If Eigo >= 30 Then
                �w�i.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����4.jpg")
                �R�����g.Caption = "�u�������A����e�X�g�����z�����I" & Chr(13) & Chr(10) & "����ő��Ƃ��邱�Ƃ��ł��邼�I�v"
            Else
                �w�i.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg")
                �R�����g.Caption = "�u�����I��肾�A�e�X�g�Őԓ_���Ƃ��Ă��܂��Ă������Ƃł��˂��B" & Chr(13) & Chr(10) & "������N���̊w�Z��...�v"
            End If
        Else
            �w�i.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg")
            �R�����g.Caption = "�u�����I��肾�A�e�X�g�Őԓ_���Ƃ��Ă��܂��Ă������Ƃł��˂��B" & Chr(13) & Chr(10) & "������N���̊w�Z��...�v"
        End If
    Else
        �w�i.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\�Z�b�g\����2.jpg")
        �R�����g.Caption = "�u�����I��肾�A�e�X�g�Őԓ_���Ƃ��Ă��܂��Ă������Ƃł��˂��B" & Chr(13) & Chr(10) & "������N���̊w�Z��...�v"
    End If
End Sub

Private Sub �I��_Click()
    Unload Me
    ���j���[1.Show
End Sub
