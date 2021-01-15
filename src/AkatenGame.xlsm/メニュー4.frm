VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} メニュー4 
   Caption         =   "赤点回避シミュレータ[ロード選択]"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "メニュー4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "メニュー4"
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

Private Sub 選択2_Click()
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

Private Sub 選択3_Click()
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

Private Sub 戻る_Click()
    Unload Me
    If Tuduki = True Then
    ゲーム1.Show
    End If
End Sub
