VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Registro de Vendas"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub button_cancel_Click()

Unload UserForm1

End Sub

Private Sub ListBox1_Click()

caminho = "C:\Users\costa\OneDrive\Área de Trabalho\VBA_impressionador\userform\Cadastro Vendas\Imagens\"
Image1.Picture = LoadPicture(caminho & LCase(ListBox1.Value) & ".jpg", 100, 100)

End Sub

Private Sub ScrollBar1_Change()

Desconto.Value = ScrollBar1.Value & "%"

End Sub

Private Sub UserForm_Initialize()

Dim ult_lin As Integer

Application.DisplayAlerts = False

Sheets("Dados").Activate

'Moto
ult_lin = Range("A2").End(xlDown).Row

ListBox1.RowSource = "Dados!A2:A" & ult_lin

'Modelo

ult_lin = Range("C2").End(xlDown).Row

ComboBox1.RowSource = "Dados!C2:C" & ult_lin

Sheets("Principal").Activate

End Sub
