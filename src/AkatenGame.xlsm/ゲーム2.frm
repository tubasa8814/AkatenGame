VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Q�[��2 
   Caption         =   "�ԓ_����V�~�����[�^�[[�ȖڑI��]"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "�Q�[��2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�Q�[��2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub �I������_Click()
    '�ϐ��錾
    Dim X As Integer '�Y����
    '�{�v���O����
    For X = 1 To 4 '�w�K�I��
        If Me.Controls("�I��" & X).Value = True Then
            Kamoku = X
            If Kamoku = 4 Then
                Kamoku = 6
            End If
        End If
    Next X
    If Me.Controls("�X�}�z�g�p").Value = True Then '�X�}�z�g�p�m�F
        Ktai = True
    Else
        Ktai = False
    End If
    Unload Me
End Sub
