VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} メニュー3 
   Caption         =   "赤点回避シミュレータ[セーブ選択]"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "メニュー3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "メニュー3"
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
            Me.Controls("選択内容" & X).Caption = "セーブがあります" & Chr(13) & Chr(10) & _
                Tuti1 & "   " & Tuti2 & "   " & Tuti3
        Else
            Me.Controls("選択内容" & X).Caption = "セーブはありません"
        End If
    Next X
End Sub

Private Sub 選択1_Click()
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
    選択内容1.Caption = "セーブがあります" & Chr(13) & Chr(10) & _
        TutiSave1 & "   " & TutiSave2 & "   " & TutiSave3
End Sub

Private Sub 選択2_Click()
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
    選択内容2.Caption = "セーブがあります" & Chr(13) & Chr(10) & _
        TutiSave1 & "   " & TutiSave2 & "   " & TutiSave3
End Sub

Private Sub 選択3_Click()
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
    選択内容3.Caption = "セーブがあります" & Chr(13) & Chr(10) & _
        TutiSave1 & "   " & TutiSave2 & "   " & TutiSave3
End Sub

Private Sub 戻る_Click()
    Unload Me
End Sub
