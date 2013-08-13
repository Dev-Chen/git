Attribute VB_Name = "Get_INI"
Option Explicit

' INIファイル文字列情報取得関数(API)の定義
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'定数
Const C_INIFLIEFAME       As String = "config.ini"    'INIファイル名
Const C_INI_SECTIONNAME   As String = "DB接続情報"    'セクション名
Const C_INI_KEYNAKME1     As String = "SERVERNAME"    'キー名(サーバー名)
Const C_INI_KEYNAKME2     As String = "DATABASE"      'キー名(データベース名)

'パブリック変数
Public P_SERVERNAME As String
Public P_DATABASE   As String
  
'INIファイル情報の取得
Public Function Fnc_ReadIni() As Boolean
On Error GoTo Err

    Dim str         As String * 1024    '格納バッファ
    Dim FilePath    As String
    Dim ret         As Long             '戻り値 (取得した値の文字数)

    Fnc_ReadIni = False
    
    'ファイルパス取得
    FilePath = App.Path & "\" & C_INIFLIEFAME
    
    'サーバー名取得
    ret = GetPrivateProfileString(C_INI_SECTIONNAME, C_INI_KEYNAKME1, "", str, Len(str), FilePath)
    If ret <> 0 Then
        P_SERVERNAME = Left(str, InStr(str, Chr(0)) - 1)
    Else
        Exit Function
    End If

    'データベース名取得
    ret = GetPrivateProfileString(C_INI_SECTIONNAME, C_INI_KEYNAKME2, "", str, Len(str), FilePath)
    If ret <> 0 Then
        P_DATABASE = Left(str, InStr(str, Chr(0)) - 1)
    Else
        Exit Function
    End If
    
    Fnc_ReadIni = True
    
    Exit Function
Err:
    MsgBox "INIファイル取得エラー(" & Err.Number & ":" & Err.Description & ")", vbCritical
End Function
