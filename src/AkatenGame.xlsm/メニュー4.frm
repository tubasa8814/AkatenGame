VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���j���[4 
   Caption         =   "�ԓ_����V�~�����[�^[���[�h�I��]"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "���j���[4.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���j���[4"
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
    If Dir(ThisWorkbook.Path & "\..\save\save1.txt") <> "" Then
        Open ThisWorkbook.Path & "\..\save\save1.txt" For Input As #1
        Line Input #1, Z
            Kamoku = CInt(Z)
        Line Input #1, Z
            Ktai = CBool(Z)
        Line Input #1, Z
            Neru = CInt(Z)
        Line Input #1, Z
            Nzikan = CInt(Z)
        Line Input #1, Z
            Yaruki = CInt(Z)
        Line Input #1, Z
            Kokugo = CInt(Z)
        Line Input #1, Z
            Sugaku = CInt(Z)
        Line Input #1, Z
            Eigo = CInt(Z)
        Line Input #1, Z
            HiSave = CInt(Z)
        Line Input #1, Z
            ZikanSave = CInt(Z)
        Line Input #1, Z
            SkyokaSave = CInt(Z)
        Line Input #1, Z
            NameSave = CStr(Z)
        Line Input #1, Z
            CommentSave = CStr(Z)
        Line Input #1, Z
            TutiSave1 = CStr(Z)
        Line Input #1, Z
            TutiSave2 = CStr(Z)
        Line Input #1, Z
            TutiSave3 = CStr(Z)
        Line Input #1, Z
            URLSave = CStr(Z)
        Close #1
        Save = True
    Else
        Save = False
    End If
End Sub

Private Sub �I��2_Click()
    If Dir(ThisWorkbook.Path & "\..\save\save2.txt") <> "" Then
        Open ThisWorkbook.Path & "\..\save\save2.txt" For Input As #1
        Line Input #1, Z
            Kamoku = CInt(Z)
        Line Input #1, Z
            Ktai = CBool(Z)
        Line Input #1, Z
            Neru = CInt(Z)
        Line Input #1, Z
            Nzikan = CInt(Z)
        Line Input #1, Z
            Yaruki = CInt(Z)
        Line Input #1, Z
            Kokugo = CInt(Z)
        Line Input #1, Z
            Sugaku = CInt(Z)
        Line Input #1, Z
            Eigo = CInt(Z)
        Line Input #1, Z
            HiSave = CInt(Z)
        Line Input #1, Z
            ZikanSave = CInt(Z)
        Line Input #1, Z
            SkyokaSave = CInt(Z)
        Line Input #1, Z
            NameSave = CStr(Z)
        Line Input #1, Z
            CommentSave = CStr(Z)
        Line Input #1, Z
            TutiSave1 = CStr(Z)
        Line Input #1, Z
            TutiSave2 = CStr(Z)
        Line Input #1, Z
            TutiSave3 = CStr(Z)
        Line Input #1, Z
            URLSave = CStr(Z)
        Close #1
        Save = True
    Else
        Save = False
    End If
End Sub

Private Sub �I��3_Click()
    If Dir(ThisWorkbook.Path & "\..\save\save3.txt") <> "" Then
        Open ThisWorkbook.Path & "\..\save\save3.txt" For Input As #1
        Line Input #1, Z
            Kamoku = CInt(Z)
        Line Input #1, Z
            Ktai = CBool(Z)
        Line Input #1, Z
            Neru = CInt(Z)
        Line Input #1, Z
            Nzikan = CInt(Z)
        Line Input #1, Z
            Yaruki = CInt(Z)
        Line Input #1, Z
            Kokugo = CInt(Z)
        Line Input #1, Z
            Sugaku = CInt(Z)
        Line Input #1, Z
            Eigo = CInt(Z)
        Line Input #1, Z
            HiSave = CInt(Z)
        Line Input #1, Z
            ZikanSave = CInt(Z)
        Line Input #1, Z
            SkyokaSave = CInt(Z)
        Line Input #1, Z
            NameSave = CStr(Z)
        Line Input #1, Z
            CommentSave = CStr(Z)
        Line Input #1, Z
            TutiSave1 = CStr(Z)
        Line Input #1, Z
            TutiSave2 = CStr(Z)
        Line Input #1, Z
            TutiSave3 = CStr(Z)
        Line Input #1, Z
            URLSave = CStr(Z)
        Close #1
        Save = True
    Else
        Save = False
    End If
End Sub

Private Sub �߂�_Click()
    Unload Me
    If Tuduki = True Then
    �Q�[��1.Show
    End If
End Sub
