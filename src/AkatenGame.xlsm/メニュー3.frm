VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���j���[3 
   Caption         =   "�ԓ_����V�~�����[�^[�Z�[�u�I��]"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "���j���[3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���j���[3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X As Integer
    Dim Y As Integer
    Dim Z As String
    Dim Tuti1 As String
    Dim Tuti2 As String
    Dim Tuti3 As String
Private Sub UserForm_Initialize()
    For X = 1 To 3
        If Dir(ThisWorkbook.Path & "\..\save\save" & X & ".txt") <> "" Then
            Open ThisWorkbook.Path & "\..\save\save" & X & ".txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, Z
                Select Case Y
                    Case "13"
                        Tuti1 = Z
                    Case "14"
                        Tuti2 = Z
                    Case "15"
                        Tuti3 = Z
                End Select
                Y = Y + 1
            Loop
            Close #1
            Me.Controls("�I����e" & X).Caption = "�Z�[�u������܂�" & Chr(13) & Chr(10) & _
                Tuti1 & "   " & Tuti2 & "   " & Tuti3
        Else
            Me.Controls("�I����e" & X).Caption = "�Z�[�u�͂���܂���"
        End If
    Next X
End Sub

Private Sub �I��1_Click()
    Open ThisWorkbook.Path & "\..\save\save1.txt" For Output As #1
    Print #1, Kamoku
    Print #1, Ktai
    Print #1, Neru
    Print #1, Nzikan
    Print #1, Yaruki
    Print #1, Kokugo
    Print #1, Sugaku
    Print #1, Eigo
    Print #1, HiSave
    Print #1, ZikanSave
    Print #1, SkyokaSave
    Print #1, NameSave
    Print #1, CommentSave
    Print #1, TutiSave1
    Print #1, TutiSave2
    Print #1, TutiSave3
    Print #1, URLSave
    Close #1
    �I����e1.Caption = "�Z�[�u������܂�" & Chr(13) & Chr(10) & _
        TutiSave1 & "   " & TutiSave2 & "   " & TutiSave3
End Sub

Private Sub �I��2_Click()
    Open ThisWorkbook.Path & "\..\save\save2.txt" For Output As #1
    Print #1, Kamoku
    Print #1, Ktai
    Print #1, Neru
    Print #1, Nzikan
    Print #1, Yaruki
    Print #1, Kokugo
    Print #1, Sugaku
    Print #1, Eigo
    Print #1, HiSave
    Print #1, ZikanSave
    Print #1, SkyokaSave
    Print #1, NameSave
    Print #1, CommentSave
    Print #1, TutiSave1
    Print #1, TutiSave2
    Print #1, TutiSave3
    Print #1, URLSave
    Close #1
    �I����e2.Caption = "�Z�[�u������܂�" & Chr(13) & Chr(10) & _
        TutiSave1 & "   " & TutiSave2 & "   " & TutiSave3
End Sub

Private Sub �I��3_Click()
    Open ThisWorkbook.Path & "\..\save\save3.txt" For Output As #1
    Print #1, Kamoku
    Print #1, Ktai
    Print #1, Neru
    Print #1, Nzikan
    Print #1, Yaruki
    Print #1, Kokugo
    Print #1, Sugaku
    Print #1, Eigo
    Print #1, HiSave
    Print #1, ZikanSave
    Print #1, SkyokaSave
    Print #1, NameSave
    Print #1, CommentSave
    Print #1, TutiSave1
    Print #1, TutiSave2
    Print #1, TutiSave3
    Print #1, URLSave
    Close #1
    �I����e3.Caption = "�Z�[�u������܂�" & Chr(13) & Chr(10) & _
        TutiSave1 & "   " & TutiSave2 & "   " & TutiSave3
End Sub

Private Sub �߂�_Click()
    Unload Me
End Sub
