VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ウィンドウ3 
   Caption         =   "赤点回避シミュレーター[結果発表]"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19155
   OleObjectBlob   =   "ウィンドウ3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ウィンドウ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Label1.Caption = "国語" & Kokugo & "点"
    Label2.Caption = "数学" & Sugaku & "点"
    Label3.Caption = "英語" & Eigo & "点"
    If Kokugo >= 30 Then
        If Sugaku >= 30 Then
            If Eigo >= 30 Then
                背景.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室4.jpg")
                コメント.Caption = "「よっしゃ、定期テストを乗り越えた！" & Chr(13) & Chr(10) & "これで卒業することができるぞ！」"
            Else
                背景.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室2.jpg")
                コメント.Caption = "「もう終わりだ、テストで赤点をとってしまってもう卒業できねえ。" & Chr(13) & Chr(10) & "もう一年この学校か...」"
            End If
        Else
            背景.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室2.jpg")
            コメント.Caption = "「もう終わりだ、テストで赤点をとってしまってもう卒業できねえ。" & Chr(13) & Chr(10) & "もう一年この学校か...」"
        End If
    Else
        背景.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室2.jpg")
        コメント.Caption = "「もう終わりだ、テストで赤点をとってしまってもう卒業できねえ。" & Chr(13) & Chr(10) & "もう一年この学校か...」"
    End If
End Sub

Private Sub 終了_Click()
    Unload Me
    メニュー1.Show
End Sub
