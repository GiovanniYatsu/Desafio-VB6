VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Desafio VB6"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PostgreSQL35W"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PostgreSQL35W"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from ""base"""
      Caption         =   "Dados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar dados"
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar XLSX"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim oExcel As Excel.Application
    Dim oBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    Dim i As Integer
    Dim j As Integer
    
    'Conecta BD
    Set con = OpenConnection()
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM base", con, adOpenKeyset, adLockOptimistic
    
    'Abrir XLSX
    Set oExcel = New Excel.Application
    Set oBook = oExcel.Workbooks.Open("C:\Users\gylel\Desktop\Prova programador\dados.xlsx")
    Set oSheet = oBook.Worksheets(1)
    
    'Inserir informação BD
    For i = 1 To oSheet.UsedRange.Rows.Count
        rs.AddNew
        For j = 1 To oSheet.UsedRange.Columns.Count
            rs.Fields(j - 1) = oSheet.Cells(i, j)
        Next j
        rs.Update
    Next i
    
    'Fechar processos
    rs.Close
    con.Close
    oBook.Close False
    Set rs = Nothing
    Set con = Nothing
    Set oSheet = Nothing
    Set oBook = Nothing
    Set oExcel = Nothing
    
    MsgBox "Importação concluída com sucesso!"
End Sub

End Sub

Private Sub Command2_Click()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSql As String
    Dim fileName As String
    
    'Conexão BD
    Set cn = New ADODB.Connection
    cn.ConnectionString = "DSN=PostgreSQL35W;Server=localhost;Database=Avaliacao;Uid=postgres;Pwd=admin;Port=5432;"
    cn.Open
    
    'Consulta SQL
    strSql = "SELECT * FROM base"
    Set rs = New ADODB.Recordset
    rs.Open strSql, cn, adOpenStatic, adLockReadOnly
    
    'exportar CSV
    fileName = "C:\Users\gylel\Desktop\Prova programador\arquivo.csv"
    Open fileName For Output As #1
    Do Until rs.EOF
        Print #1, rs("login") & "," & rs("nome") & "," & rs("idade")
        rs.MoveNext
    Loop
    Close #1
    
    'Fechar processos
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
    MsgBox "Arquivo extraído com sucesso!"
End Sub
