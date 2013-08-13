Attribute VB_Name = "Get_INI"
Option Explicit

' INI�t�@�C����������擾�֐�(API)�̒�`
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'�萔
Const C_INIFLIEFAME       As String = "config.ini"    'INI�t�@�C����
Const C_INI_SECTIONNAME   As String = "DB�ڑ����"    '�Z�N�V������
Const C_INI_KEYNAKME1     As String = "SERVERNAME"    '�L�[��(�T�[�o�[��)
Const C_INI_KEYNAKME2     As String = "DATABASE"      '�L�[��(�f�[�^�x�[�X��)

'�p�u���b�N�ϐ�
Public P_SERVERNAME As String
Public P_DATABASE   As String
  
'INI�t�@�C�����̎擾
Public Function Fnc_ReadIni() As Boolean
On Error GoTo Err

    Dim str         As String * 1024    '�i�[�o�b�t�@
    Dim FilePath    As String
    Dim ret         As Long             '�߂�l (�擾�����l�̕�����)

    Fnc_ReadIni = False
    
    '�t�@�C���p�X�擾
    FilePath = App.Path & "\" & C_INIFLIEFAME
    
    '�T�[�o�[���擾
    ret = GetPrivateProfileString(C_INI_SECTIONNAME, C_INI_KEYNAKME1, "", str, Len(str), FilePath)
    If ret <> 0 Then
        P_SERVERNAME = Left(str, InStr(str, Chr(0)) - 1)
    Else
        Exit Function
    End If

    '�f�[�^�x�[�X���擾
    ret = GetPrivateProfileString(C_INI_SECTIONNAME, C_INI_KEYNAKME2, "", str, Len(str), FilePath)
    If ret <> 0 Then
        P_DATABASE = Left(str, InStr(str, Chr(0)) - 1)
    Else
        Exit Function
    End If
    
    Fnc_ReadIni = True
    
    Exit Function
Err:
    MsgBox "INI�t�@�C���擾�G���[(" & Err.Number & ":" & Err.Description & ")", vbCritical
End Function
