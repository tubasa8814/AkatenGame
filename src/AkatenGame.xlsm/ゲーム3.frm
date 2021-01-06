VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ゲーム3 
   Caption         =   "赤点回避シミュレーター[睡眠時間選択]"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "ゲーム3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ゲーム3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 選択完了_Click()
    '変数宣言
    Dim X As Integer '添え字
    '本プログラム
    For X = 1 To 4
        If Me.Controls("選択" & X).Value = True Then
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

