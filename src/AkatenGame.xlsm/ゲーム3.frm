VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Q�[��3 
   Caption         =   "�ԓ_����V�~�����[�^�[[�������ԑI��]"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "�Q�[��3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�Q�[��3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub �I������_Click()
    '�ϐ��錾
    Dim X As Integer '�Y����
    '�{�v���O����
    For X = 1 To 4
        If Me.Controls("�I��" & X).Value = True Then
            Nzikan = X
        End If
    Next X
    Select Case Nzikan
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
    Unload Me
End Sub

