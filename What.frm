VERSION 5.00
Begin VB.Form What 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ʲô��"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7275
   Icon            =   "What.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7275
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ�� (&O) "
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   $"What.frx":3482
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "����һ���� TNT ������������޹�˾��Ʒ�ģ����Ա��� C++ Դ�ļ���ʵ�ó���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "What"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
