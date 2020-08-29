VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pa'luego DEMO"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRead 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   6480
   End
   Begin VB.ListBox lstFormats 
      Height          =   1425
      ItemData        =   "frmMain.frx":0000
      Left            =   240
      List            =   "frmMain.frx":0002
      TabIndex        =   6
      Top             =   4920
      Width           =   8655
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Descargar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "https://www.mitele.es/programas-tv/cuarto-milenio/temporada-15/Programa-620-40_1008325075016/player/"
      Top             =   360
      Width           =   6495
   End
   Begin VB.CommandButton btnDown 
      Caption         =   "Leer"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtOut 
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   8535
   End
   Begin VB.Label lblDescripcion 
      Caption         =   "Descripción"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   8535
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Título"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   8535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Datos As Object
Dim ejecutaYT As New cExec
Dim resultado As Boolean
Dim StdOut As String
Dim formato As Object

Private Sub btnDown_Click()

'On Error GoTo ShellError
' Leemos en formato JSON y sin descargar
resultado = ejecutaYT.Run(App.Path & "\youtube-dl.exe", "--print-json --skip-download " & txtURL.Text, False, False)

' Si se ha ejecutado OK
If (resultado) Then
    ' Leemos la salida
    StdOut = ejecutaYT.ReadAllOutput()
    txtOut.Text = StdOut
    ' Parseamos el JSON
    Set Datos = JSON.parse(StdOut)
    lblTitulo.Caption = Datos.Item("title")
    lblDescripcion.Caption = Datos.Item("description")
    ' Formatos
    lstFormats.Clear
    For Each formato In Datos.Item("formats")
        lstFormats.AddItem (formato.Item("format_id"))
    Next
    
    ' Seleccionamos el primero
    If (lstFormats.ListCount > 0) Then
        lstFormats.ListIndex = 0
        btnSave.Enabled = True
    Else
        btnSave.Enabled = False
    End If
    
    'Objeto: JSON.toString(Datos)
    'Coleccion: Datos.Item("items").Item(1).Item("url")
End If

Exit Sub

ShellError:
MsgBox "Error: " & Err.Description, vbCritical, "Error"

End Sub

Private Sub btnSave_Click()

' Ya tenemos el objeto JSON, por lo que podemos descargar
' de momento simplemente descargamos el item 1, de menor calidad
'"-f " + format + " -o \"" + output_filename + "\" " + Url;
Dim seleccionado As String
seleccionado = lstFormats.List(lstFormats.ListIndex)
resultado = ejecutaYT.Run(App.Path & "\youtube-dl.exe", "-f " & seleccionado & " " & txtURL.Text, False, False)

' Leemos la salida
If (resultado) Then
    ' Activamos la lectura cada 500ms
    tmrRead.Enabled = True
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tmrRead_Timer()
Dim exitcode As Long
exitcode = -1 'ejecutaYT.GetExitCode()

If (exitcode = 0) Then
    tmrRead.Enabled = False
    Exit Sub
End If

'StdOut = ejecutaYT.ReadPendingOutput()
StdOut = ejecutaYT.ReadPendingOutput()
txtOut.Text = StdOut

End Sub
