VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   7680
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7680
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7680
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7680
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7680
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7680
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7680
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Txt_Birth 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "YYYY/MM/DD"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��"
      Height          =   495
      Left            =   3960
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��"
      Height          =   495
      Left            =   2640
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2760
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   4440
      Width           =   4935
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   720
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      MaxLength       =   3
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   4080
      Width           =   4935
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   720
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   120
      MaxLength       =   3
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   3720
      Width           =   4935
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   720
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      MaxLength       =   3
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   3360
      Width           =   4935
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   720
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      MaxLength       =   3
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   3000
      Width           =   4935
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      MaxLength       =   3
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Delete_BS 
      Caption         =   "�폜"
      Height          =   495
      Left            =   1320
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton RegisterBS 
      Caption         =   "�o�^"
      Height          =   495
      Left            =   0
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      MaxLength       =   3
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   2640
      Width           =   4935
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      MaxLength       =   3
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      MaxLength       =   3
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   1920
      Width           =   4935
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      MaxLength       =   3
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m�m"
      Top             =   1200
      Width           =   4935
   End
   Begin VB.TextBox Txt_Name 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "�m�m�m�m�m�m�m�m"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Txt_ID 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "999"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "���N����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   47
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�Z��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   42
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "���O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   41
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�f�[�^�x�[�X
Private P_cn As New ADODB.Connection
Private P_rs As New ADODB.Recordset
Private P_index As Integer
'�萔
Const C_DCNT As Integer = 10 '���׍s��


'�N����
Private Sub Form_Load()
    On Error GoTo Err
  
    '��ʂ̏�����
    Call Prc_FormInit
    
    'INI���̎擾
    If Fnc_ReadIni = False Then
        Exit Sub
    End If

    '�f�[�^�x�[�X�ڑ�
    If Fnc_DBConect(P_cn) = False Then
        Exit Sub
    End If

    '�f�[�^�\��
    If Disp_Data = False Then
        Exit Sub
    End If
    
    Exit Sub
Err:
    MsgBox Err.Number & ":" & Err.Description
End Sub

'������
Private Sub Prc_FormInit()
On Error Resume Next

    Dim i As Integer
    
    For i = 0 To C_DCNT - 1
        Txt_ID(i).Text = ""
        Txt_ID(i).Locked = False
        Txt_Name(i).Text = ""
        Txt_Address(i).Text = ""
        Txt_Birth(i).Text = ""
    Next i

End Sub

'�f�[�^�\��
Private Function Disp_Data() As Boolean
On Error GoTo Err

    Dim Sql As String
    Dim i   As Integer
    
    Disp_Data = False
    
    i = 0

    'SQL����
    Sql = ""
    Sql = Sql & "SELECT TOP " & C_DCNT & "* "
    Sql = Sql & "FROM ���O�}�X�^ "
    Sql = Sql & "ORDER BY ID"

    'SQL���s
    P_rs.Open Sql, P_cn
    
    Do Until P_rs.EOF
        '�f�[�^�\��
        Txt_ID(i).Text = Format(P_rs!ID, "000")
        Txt_ID(i).Locked = True                   '�o�^�ς�ID���̓��b�N
        Txt_Name(i).Text = P_rs!���O
        Txt_Address(i).Text = P_rs!�Z��
        Txt_Birth(i).Text = Format(P_rs!���N����, "@@@@/@@/@@")
        
        '�m�F�p
        Debug.Print "ID :" & P_rs!ID & " �o�^���t :" & P_rs!�o�^���t
        Debug.Print "ID :" & P_rs!ID & " �X�V���t :" & P_rs!�X�V���t
        
        
        i = i + 1     '�J�E���g
        P_rs.MoveNext '�����R�[�h
        
    Loop
    
    
    '�N���[�Y
    P_rs.Close: Set P_rs = Nothing
    
          
    Disp_Data = True
    
    
    '�����J�[�\���ʒu
    Form1.Show
    If i < 10 Then
        Txt_ID(i).SetFocus
    Else
        Txt_Birth(i - 1).SetFocus
    End If
    
    
    Exit Function
Err:
    MsgBox Err.Number & ":" & Err.Description
End Function



'�o�^�{�^������
Private Sub RegisterBS_Click()

    '���̓G���[�`�F�b�N
    If ERR_CHK = False Then
        Exit Sub
    End If
    
    '�o�^�����J�n
    If Register = False Then
        Exit Sub
    End If
    
    '��ʂ̏�����
    Call Prc_FormInit
    
    '�f�[�^�\��
    If Disp_Data = False Then
        Exit Sub
    End If
    
    Debug.Print "Register OK!"
    
End Sub

'���̓G���[�`�F�b�N
Private Function ERR_CHK() As Boolean
    
    Dim i As Integer
    i = 0
    
    ERR_CHK = False
    
    Do Until i = 10
    
        'ID�����̓`�F�b�N
        If Txt_ID(i).Text = "" Then
            If Txt_Name(i).Text <> "" Or Txt_Address(i).Text <> "" Or Txt_Birth(i).Text <> "" Then
                MsgBox "ID�������͂ł�", vbOKOnly, "�G���["
                Txt_ID(i).SetFocus
                Exit Function
            End If
        End If
        
        '���O�����̓`�F�b�N
        If Txt_ID(i).Text <> "" Then
            If Txt_Name(i).Text = "" Then
                MsgBox "���O�������͂ł�", vbOKOnly, "�G���["
                Txt_Name(i).SetFocus
                Exit Function
            End If
        End If
        
        '���N�������̓`�F�b�N
        If Txt_Birth(i).Text <> "" Then
            If Not IsDate(Txt_Birth(i).Text) Or Len(Txt_Birth(i).Text) < 10 Then
                MsgBox "���N�����̓��͂��s���ł�", vbOKOnly, "�G���["
                Txt_Birth(i).SetFocus
                Exit Function
            End If
        End If
        
        'ID�d���`�F�b�NSQL
        If Txt_ID(i).Text <> "" And Txt_ID(i).Locked = False Then
        
        On Error GoTo Err

        Dim Sql As String
    
        'SQL����
        Sql = "SELECT ID FROM ���O�}�X�^"

        'SQL���s
        P_rs.Open Sql, P_cn
        
        'ID�d���`�F�b�N
        
        Do Until P_rs.EOF
            If Txt_ID(i).Text = Format(P_rs!ID, "000") Then
                MsgBox "ID���d�����Ă��܂�", vbOKOnly, "�G���["
                Txt_ID(i).SetFocus
                P_rs.Close
                Exit Function
            End If
            P_rs.MoveNext '�����R�[�h
        Loop

        '�N���[�Y
        P_rs.Close: Set P_rs = Nothing
        End If
        
        'ID�d���`�F�b�N
        If Txt_ID(i).Locked = False And Txt_ID(i).Text <> "" Then
        Dim j As Integer
            For j = i + 1 To 9
                If Txt_ID(i).Text = Txt_ID(j).Text Then
                    MsgBox "ID���d�����Ă��܂�", vbOKOnly, "�G���["
                    Txt_ID(i).SetFocus
                    Exit Function
                End If
            Next
        End If
        i = i + 1     '�J�E���g
    Loop
    
    ERR_CHK = True
    
    Debug.Print "ERROR CHECK OK!"
    
    Exit Function
Err:
    MsgBox Err.Number & ":" & Err.Description
    
End Function

'�o�^�����J�n
Private Function Register() As Boolean

    Register = False
    
    Dim Sql As String
    Dim i As Integer
    i = 0
    
        'SQL����
        Sql = "SELECT * FROM ���O�}�X�^"

        'SQL���s
        P_rs.Open Sql, P_cn, adOpenDynamic, adLockOptimistic
        
        
        For i = 0 To 9
            If Txt_ID(i).Locked = True Then
                Do Until P_rs.EOF
                    If Txt_ID(i).Text = Format(P_rs!ID, "000") Then
                        P_rs!���O = Txt_Name(i).Text
                        P_rs!�Z�� = Txt_Address(i).Text
                        P_rs!���N���� = Replace(Txt_Birth(i).Text, "/", "")
                        P_rs!�X�V���t = Date
                        P_rs.Update
                        P_rs.MoveFirst
                        Exit Do
                    End If
                    P_rs.MoveNext '�����R�[�h
                Loop
            ElseIf Txt_ID(i).Text <> "" Then
                P_rs.AddNew
                P_rs!ID = Txt_ID(i).Text
                P_rs!���O = Txt_Name(i).Text
                P_rs!�Z�� = Txt_Address(i).Text
                P_rs!���N���� = Replace(Txt_Birth(i).Text, "/", "")
                P_rs!�o�^���t = Date
                P_rs.Update
                P_rs.MoveFirst
            Else
            End If
            
        Next
        
    Register = True
    P_rs.Close: Set P_rs = Nothing

End Function

'�폜�{�^��
Private Sub Delete_BS_Click()
    Dim msg As String
    Dim Del As Integer
    
    If Txt_ID(P_index).Locked = True Then
        msg = "ID:" + Txt_ID(P_index).Text + "���폜���Ă���낵���ł����H"
        Del = MsgBox(msg, vbYesNo, "�m�F")
        
        If Del = vbYes Then
            Call Del_SQL
        Else
        End If
    End If
    
    '��ʂ̏�����
    Call Prc_FormInit
    
    '�f�[�^�\��
    If Disp_Data = False Then
        Exit Sub
    End If
    
End Sub

'�폜���s
Private Sub Del_SQL()

    Dim Sql As String
    
    'SQL����
    Sql = "SELECT * FROM ���O�}�X�^ WHERE ID = '" & Txt_ID(P_index).Text & "'"

    'SQL���s
    P_rs.Open Sql, P_cn, adOpenDynamic, adLockOptimistic
        
        
    If Not P_rs.EOF Then
        P_rs.Delete        '���R�[�h���폜
    End If

    P_rs.Close: Set P_rs = Nothing
        
End Sub


'TAB ENTER�L�[�̃J�[�\���ړ�
Private Sub Txt_ID_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc(vbTab) Then
        Txt_Name(Index).SetFocus
        KeyAscii = 0
    End If

End Sub

'TAB ENTER�L�[�̃J�[�\���ړ�
Private Sub Txt_Name_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc(vbTab) Then
        Txt_Address(Index).SetFocus
        KeyAscii = 0
    End If
    
End Sub

'TAB ENTER�L�[�̃J�[�\���ړ�
Private Sub Txt_Address_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc(vbTab) Then
        Txt_Birth(Index).SetFocus
        KeyAscii = 0
    End If

End Sub

'TAB ENTER�L�[�̃J�[�\���ړ�
Private Sub Txt_Birth_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc(vbTab) Then
    
        If Index = 9 Then                       '�ŏI�s
            If Txt_ID(0).Locked = True Then     '�C���s
                Txt_Name(0).SetFocus
                KeyAscii = 0
            Else                                '�V�K�s
                Txt_ID(0).SetFocus
                KeyAscii = 0
            End If
        Else
            If Txt_ID(Index + 1).Locked = True Then     '�C���s
                Txt_Name(Index + 1).SetFocus
                KeyAscii = 0
            Else                                '�V�K�s
                Txt_ID(Index + 1).SetFocus
                KeyAscii = 0
            End If
        End If
    
    End If

End Sub


'ID�̐������͐���
Private Sub Txt_ID_LostFocus(Index As Integer)
    
    If Not IsNumeric(Txt_ID(Index).Text) Then
        Txt_ID(Index).Text = ""
    End If
        Txt_ID(Index) = Format(Txt_ID(Index).Text, "000")
    P_index = Index
    
End Sub

Private Sub Txt_Name_LostFocus(Index As Integer)
    P_index = Index
End Sub

Private Sub Txt_Address_LostFocus(Index As Integer)
    P_index = Index
End Sub

Private Sub Txt_Birth_LostFocus(Index As Integer)
    P_index = Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_cn.Close: Set P_cn = Nothing
End Sub
