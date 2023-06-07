VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconectado"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboChrEspecial 
      Height          =   315
      ItemData        =   "Form1.frx":169B2
      Left            =   6480
      List            =   "Form1.frx":169C2
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtSend 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   4920
      Width           =   6375
   End
   Begin VB.ComboBox cboBaudRate 
      Height          =   315
      ItemData        =   "Form1.frx":169FA
      Left            =   8280
      List            =   "Form1.frx":16A16
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox cboCommPort 
      Height          =   315
      ItemData        =   "Form1.frx":16A4F
      Left            =   9720
      List            =   "Form1.frx":16A51
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtReceive 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mFormat 
         Caption         =   "Format"
         Begin VB.Menu mCorEditor 
            Caption         =   "Cor do Editor"
         End
         Begin VB.Menu mCorTexto 
            Caption         =   "Cor do Texto"
         End
         Begin VB.Menu mFonteTexto 
            Caption         =   "Fonte do Texto"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Variável global
Dim scan As Boolean

Private Sub Form_Load()
   Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " by DALÇOQUIO AUTOMAÇÃO"
   writeMensagem ("Desconectado !!!")
   cboChrEspecial.Text = "None"
   cboBaudRate.Text = "9600"
   
   txtReceive.ToolTipText = "De um Duplo Click para Clear, para limpar todos os dados recebidos."
   txtSend.ToolTipText = "Digite o dado a ser enviado, depois de um Duplo Click para Send, ou pressione [Enter]."
   cboCommPort.ToolTipText = "Selecione a Porta COM para Conectar, Disconnect para Desconectar, Scan para Scanear Porta COM disponível."
   cboBaudRate.ToolTipText = "Selecione o Baud Rate para a Velocidade de Comunicação Serial"
   cboChrEspecial.ToolTipText = "Selecione o caracter especial para ser enviado no final do dado, ou None para nenhum."
   Call scanCommPort
   
End Sub

Private Sub scanCommPort()
   scan = True
   cboCommPort.Enabled = False
   
   cboCommPort.Clear
   Dim i As Integer
   For i = 1 To 16 'Procura portas COM de 1 a 16
      MSComm1.CommPort = i
      On Error Resume Next 'ignora o tratamento de erro
      MSComm1.PortOpen = True 'tenta abrir a porta
      If Err.Number = 0 Then 'a porta está disponível
         cboCommPort.AddItem "COM" & i
         cboCommPort.ListIndex = 1
         MSComm1.PortOpen = False 'fecha a porta
      End If
      On Error GoTo 0 'ativa o tratamento de erro novamente
   Next i
   
   cboCommPort.Enabled = True
   cboCommPort.AddItem "Scan"
   cboCommPort.AddItem "Disconnect"
   cboCommPort.Text = "Disconnect"
   scan = False

End Sub

Private Sub cboBaudRate_Click()
      MSComm1.Settings = cboBaudRate.Text & "n,8,1"
      
      If MSComm1.PortOpen = True Then
         writeMensagem ("Conectado na " & cboCommPort.Text & "," & MSComm1.Settings)
      End If
    
End Sub

Private Sub cboCommPort_Click()
   On Error GoTo Erro
      If scan = True Then Exit Sub
      If MSComm1.PortOpen = True Then
         MSComm1.PortOpen = False
         writeMensagem ("Desconectado !!!")
      End If
      If cboCommPort.Text = "Scan" Then
         writeMensagem ("Scanning...")
         Call scanCommPort
         writeMensagem ("Scan finalizado !!!")
         Exit Sub
      ' Desconecta
      ElseIf cboCommPort.Text = "Disconnect" Then
         Exit Sub
      ' Conecta
      Else
         MSComm1.Settings = cboBaudRate.Text & "n,8,1"
         MSComm1.CommPort = Mid(cboCommPort.Text, 4, 6)
         MSComm1.PortOpen = True
         writeMensagem ("Conectado na " & cboCommPort.Text & "," & MSComm1.Settings)
      End If
   Exit Sub
   
Erro:
MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
   
End Sub

Private Sub MSComm1_OnComm()
   On Error GoTo Erro
      Dim strData As String
      Do While MSComm1.InBufferCount > 0
          If strData = Empty Then Exit Do
          strData = MSComm1.Input(MSComm1.InBufferCount)
      Loop
      txtReceive.Text = txtReceive.Text + MSComm1.Input
      txtReceive.SelStart = Len(txtReceive.Text)
   Exit Sub
   
Erro:
   MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
   
End Sub

Private Sub txtSend_DblClick()
   mSend.Visible = True
   mClear.Visible = False
   mFormat.Visible = False
   PopupMenu mMenu
   mSend.Visible = True
   mClear.Visible = True
   mFormat.Visible = True
   
End Sub
Private Sub mSend_Click()
   If MSComm1.PortOpen = True Then
      Call sendDado
      writeMensagem ("Enviado com sucesso...")
   End If
   
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And MSComm1.PortOpen = True Then
      Call sendDado
   End If

End Sub

Private Sub sendDado()
   If cboChrEspecial = "None" Then
      MSComm1.Output = txtSend.Text
   ElseIf cboChrEspecial = "Nova Linha" Then
      MSComm1.Output = txtSend.Text & vbLf 'vbNewLine
   ElseIf cboChrEspecial = "Retorno de Carro" Then
      MSComm1.Output = txtSend.Text & vbCr
   ElseIf cboChrEspecial = "Ambos, NL e CR" Then
      MSComm1.Output = txtSend.Text & vbCrLf
   End If

End Sub

Private Sub txtReceive_DblClick()
   mSend.Visible = False
   mClear.Visible = True
   mFormat.Visible = True
   PopupMenu mMenu
   mSend.Visible = True
   mClear.Visible = True
   mFormat.Visible = True

End Sub

Private Sub mClear_Click()
   txtReceive.Text = Empty
   
End Sub

Private Sub mCorTexto_Click()
    CommonDialog1.ShowColor
    txtReceive.ForeColor = CommonDialog1.Color
End Sub

Private Sub mCorEditor_Click()
    CommonDialog1.ShowColor
    txtReceive.BackColor = CommonDialog1.Color
End Sub

Private Sub mFonteTexto_Click()
    CommonDialog1.Flags = CommonDialog1CFBoth
    CommonDialog1.ShowFont
    txtReceive.Font = CommonDialog1.FontName
    txtReceive.FontBold = CommonDialog1.FontBold
    txtReceive.FontItalic = CommonDialog1.FontItalic
    txtReceive.FontSize = CommonDialog1.FontSize
End Sub

Private Sub writeMensagem(mensagem As String)
   txtReceive.Text = txtReceive.Text & "> " & mensagem & vbCrLf
   txtReceive.SelStart = Len(txtReceive.Text)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MSComm1.PortOpen = True Then
      MSComm1.PortOpen = False
      writeMensagem ("Desconectado !!!")
   End If
   writeMensagem ("Fechando o sistema...")
   DoEvents
   Sleep (1000)
   End

End Sub

