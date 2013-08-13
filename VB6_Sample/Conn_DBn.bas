Attribute VB_Name = "Conn_DB"
Option Explicit

'データベース接続処理
Public Function Fnc_DBConect(ByRef cn As ADODB.Connection) As Boolean
On Error GoTo Err

    Dim StrConn As String
    
    Fnc_DBConect = False
    
    If cn.State = 0 Then
        '接続文字列を生成する
        StrConn = ""
        StrConn = StrConn & "Driver={SQL Server};"
        StrConn = StrConn & "Server=" & P_SERVERNAME & ";"  'サーバー名
        StrConn = StrConn & "Database=" & P_DATABASE & ";"  'データーベース名
        'データベース接続
        cn.Open (StrConn)
    End If

    Fnc_DBConect = True
    
    Exit Function
Err:
    MsgBox "データベース接続エラー(" & Err.Number & ":" & Err.Description & ")", vbCritical
End Function


