VERSION 5.00
Begin VB.Form frm 
   BackColor       =   &H00C0C000&
   Caption         =   "Suprimentos de Informática"
   ClientHeight    =   3840
   ClientLeft      =   2220
   ClientTop       =   3960
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   Picture         =   "frm.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   10290
   Begin VB.TextBox Txt_Total 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Txt_Desc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Txt_Quantidade 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox lstN1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "frm.frx":0442
      Left            =   120
      List            =   "frm.frx":0458
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txt_Preco 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmd_Limpar 
      BackColor       =   &H00C0C000&
      Caption         =   "Limpar"
      DownPicture     =   "frm.frx":048C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7560
      MaskColor       =   &H00000000&
      Picture         =   "frm.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Preço"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Total (R$)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Desconto (%)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Quantidade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Img_prod 
      Height          =   1815
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Camin As String



Private Sub cmd_Limpar_Click()
            txt_Preco = Clear
            Txt_Total = Clear
            Txt_Quantidade = Clear
            Txt_Desc = Clear
            Img_prod.Picture = LoadPicture(Empty)
            lstN1.SetFocus
            
End Sub



Private Sub Form_Load()
            Camin = "C:\Arquivos de programas\Microsoft Visual Studio\Common\Graphics\Icons\Computer\"
End Sub

Private Sub lstN1_Click()
             If lstN1.Text = "CD-ROM" Then
                Img_prod.Picture = LoadPicture(Camin & "CDROM01.ico")
                txt_Preco = Format(1.5, "currency")
             ElseIf lstN1.Text = "CPU" Then
                Img_prod.Picture = LoadPicture(Camin & "MAC02.ico")
                txt_Preco = Format(700#, "currency")
                 ElseIf lstN1.Text = "Disquete" Then
                Img_prod.Picture = LoadPicture(Camin & "DISK06.ico")
                txt_Preco = Format(0.8, "currency")
                     ElseIf lstN1.Text = "Monitor" Then
                Img_prod.Picture = LoadPicture(Camin & "MONITR01.ico")
                txt_Preco = Format(200#, "currency")
                         ElseIf lstN1.Text = "Mouse" Then
                Img_prod.Picture = LoadPicture(Camin & "MOUSE01.ico")
                txt_Preco = Format(20#, "currency")
                             ElseIf lstN1.Text = "Teclado" Then
                Img_prod.Picture = LoadPicture(Camin & "KEYBRD01.ico")
                txt_Preco = Format(29.9, "currency")
            End If
                
                
                
End Sub





Private Sub Txt_Desc_KeyPress(KeyAscii As Integer)
            If KeyAscii < 48 Then
            If KeyAscii <> 44 Then
                KeyAscii = 8
                Else
            If Len(Txt_Desc) >= 3 Then KeyAscii = 0
            End If
            End If
            If KeyAscii > 58 Then
                KeyAscii = 8
                End If
            
End Sub

Private Sub Txt_Desc_LostFocus()
            
            On Error GoTo B
            If Txt_Desc > 100 Then GoTo C
            
            Txt_Total = Format((txt_Preco * Txt_Quantidade) - (txt_Preco * Txt_Quantidade) * (Txt_Desc / 100), "currency")
            Txt_Desc = Txt_Desc + "%"
            Exit Sub
            
            
B:
           MsgBox "Por Favor! Escolha um produto, sua respectiva quantidade e o seu desconto", vbExclamation, "ATENÇÃO"
            Exit Sub
C:
            MsgBox "Por favor!Escolha um desconto igual ou menor que 100", vbExclamation, "Atenção, Houve um equivoco"
            
            


           
End Sub

Private Sub Txt_Quantidade_KeyPress(KeyAscii As Integer)
            If KeyAscii < 48 Then
            KeyAscii = 8
            End If
            If KeyAscii > 58 Then
                KeyAscii = 8
            End If
           
               
End Sub

Private Sub Txt_Quantidade_LostFocus()
                        
            On Error GoTo A
            Txt_Total = Format(txt_Preco * Txt_Quantidade, "currency")
            Exit Sub
A:
            MsgBox "Por Favor! Escolha um produto", vbExclamation, "ATENÇÃO"
            
            

            
End Sub

