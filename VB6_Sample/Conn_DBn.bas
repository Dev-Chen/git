Attribute VB_Name = "Conn_DB"
Option Explicit

'�f�[�^�x�[�X�ڑ�����
Public Function Fnc_DBConect(ByRef cn As ADODB.Connection) As Boolean
On Error GoTo Err

    Dim StrConn As String
    
    Fnc_DBConect = False
    
    If cn.State = 0 Then
        '�ڑ�������𐶐�����
        StrConn = ""
        StrConn = StrConn & "Driver={SQL Server};"
        StrConn = StrConn & "Server=" & P_SERVERNAME & ";"  '�T�[�o�[��
        StrConn = StrConn & "Database=" & P_DATABASE & ";"  '�f�[�^�[�x�[�X��
        '�f�[�^�x�[�X�ڑ�
        cn.Open (StrConn)
    End If

    Fnc_DBConect = True
    
    Exit Function
Err:
    MsgBox "�f�[�^�x�[�X�ڑ��G���[(" & Err.Number & ":" & Err.Description & ")", vbCritical
End Function


