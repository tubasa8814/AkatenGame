VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ゲーム2 
   Caption         =   "赤点回避シミュレーター[科目選択]"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "ゲーム2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ゲーム2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 選択完了_Click()
    '変数宣言
    Dim X As Integer '添え字
    '本プログラム
    For X = 1 To 4 '学習選択
        If Me.Controls("選択" & X).Value = True Then
            Kamoku = X
            If Kamoku = 4 Then
                Kamoku = 6
            End If
        End If
    Next X
    If Me.Controls("スマホ使用").Value = True Then 'スマホ使用確認
        Ktai = True
    Else
        Ktai = False
    End If
    Unload Me
End Sub
