VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �E�B���h�E3 
   Caption         =   "�ԓ_����V�~�����[�^�[[���ʔ��\]"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5940
   OleObjectBlob   =   "�E�B���h�E3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�E�B���h�E3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Label1.Caption = Kokugo & "�_"
    Label2.Caption = Sugaku & "�_"
    Label3.Caption = Eigo & "�_"
End Sub

Private Sub �I��_Click()
    Unload Me
    ���j���[1.Show
End Sub
