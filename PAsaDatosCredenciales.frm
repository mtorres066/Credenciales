VERSION 5.00
Begin VB.Form PasaInformacion 
   Caption         =   "Pasa Información de Diferentes BD"
   ClientHeight    =   7065
   ClientLeft      =   1020
   ClientTop       =   1410
   ClientWidth     =   7875
   Icon            =   "PAsaDatosCredenciales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEjecutaCerraduras 
      Caption         =   "Pasa Datos a Cerraduras de Huella"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox TxtCerraduras 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Carga 
      Appearance      =   0  'Flat
      Caption         =   "Carga Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton Ejecuta 
      Caption         =   "Pasar Datos a Credenciales"
      Height          =   1335
      Left            =   360
      MaskColor       =   &H00E0E0E0&
      Picture         =   "PAsaDatosCredenciales.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton EjecutaComidas 
      Caption         =   "Pasar Datos a Comidas"
      Height          =   1335
      Left            =   360
      MaskColor       =   &H00E0E0E0&
      Picture         =   "PAsaDatosCredenciales.frx":5E8C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   975
      Left            =   360
      Picture         =   "PAsaDatosCredenciales.frx":9356
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox TextComidas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox TextCredenciales 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox TextReloj 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Pasa Datos de Reloj Checador --> Cerraduras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   15
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label LabelReloj 
      Caption         =   "# Registros en BD Cerraduras Huella"
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   12
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label LabelReloj 
      Caption         =   "# Registros en BD Comidas"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label LabelReloj 
      Caption         =   "# Registros en BD Credenciales"
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label LabelReloj 
      Caption         =   "# Registros en Reloj Checador"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Pasa Datos de Reloj Checador --> Comidas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Pasa Datos de Reloj Checador --> Credenciales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "PasaInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************************************************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************************************************************************************************************************************
' Módulo   : PasaInformacion
' Fecha    : 07/04/2007 17:56
' Autor    : Miguel
' Propósito:PASAR INFORMACION ENTRE BD DE COMIDAS, RELOJ Y COMIDAS
'**************************************************************************************************************************************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************************************************************************************************************************************
Option Explicit
Dim r As New ADODB.Recordset
Dim r2 As New ADODB.Recordset
Dim r3 As New ADODB.Recordset
Dim r4 As New ADODB.Recordset
Dim RBuscaClaveEmpleado As New ADODB.Recordset

Dim vconteoComidas As String
Dim vconteoCredenciales As String
Dim vconteoReloj As String
Dim vCostoComida As String
Dim vconteoCerraduras As String
Dim vClaveEmp As String
Dim vNombreEmp As String
Dim vDepto As String
Dim vPuesto As String
Dim vNSS As String
Dim vCodigoEmpNuevo As String
Dim vSumaCodigoEmpNuevo As String
Dim vCodigoEmpNuevo2 As String

Dim ContadorComidas As Currency
Dim ContadorCerraduras As Currency
Dim ContadorCredenciales As Currency



Private Sub Form_Load()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load-------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load
' Procedimiento : Form_Load
' Fecha         : 07/04/2007 17:56
' Autor         : Miguel
' Propósito     :CARGA LOS VALORES DE CONEXION CON CADA BD
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load------------------------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo Form_Load_Error

' BD COMIDAS ----------------------------------------------
'Conexion a la bd de comidas ACCESS - Servidor/////////////
'GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=\\respaldosfepsa\comidas\comidasfepsa.mdb; Jet OLEDB:Database Password="

'Conexion a Comidas MySQL en Servidor MySQL ///////////////////// Desde PC Soporte Win XP Virtual
GConectionString = "DRIVER={MySQL ODBC 5.2a Driver};DATABASE=comidas;SERVER=192.168.0.83;UID=miguel;PASSWORD=Matl066;PORT=3306;"

' BD CREDENCIALES -----------------------------------------
'Conexion a la Base de Datos del Diseñador de Credenciales - ACCESS -PC Soporte //////////////////
GConectionString2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\IDCARDDESIGN\DATOS\Fepsa\Fepsa.bid;Persist Security Info=True;Jet OLEDB:Database Password=alx09;"

'Conexion a la Base de Datos del Diseñador de Credenciales - ACCESS -Portatil XP-Virtual /////////////////////
'GConectionString2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\IDCARDDESIGN\DATOS\Fepsa\Fepsa.bid;Persist Security Info=True;Jet OLEDB:Database Password=alx09;"

' BD RELOJ CHECADOR HUELLA  -------------------------------
'Conexion a la Base de Datos del Reloj Checador de Huella ACCESS - Servidor  ////////////
GConectionString3 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=\\Respaldosfepsa\relojchecadorhuella\Fepsa2.mdb; Jet OLEDB:Database Password="

'Conexion a la Base de Datos del Reloj Checador de Huella ACCESS - Portatil XP-Virtual /////////////////////
'GConectionString3 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=c:\relojchecadorhuella\Fepsa2.mdb; Jet OLEDB:Database Password="

'Conexion a la Base de Datos de Cerraduras de Huella Digital "Baños"   ////////////
'GConectionString4 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CerradurasBaños;Data Source=RESPALDOSFEPSA\SQLEXPRESS"


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Conexion para el reloj checador de tarjeta  //////////
'GConectionString3 = "DRIVER=Firebird/InterBase(r) driver;UID=sysdba;PWD=masterkey;DBNAME=192.168.0.128:C:\Reloj_Checador\Database\Reloj_Checador.FDB;"
'GConectionString3 = "DRIVER=Firebird/InterBase(r) driver;UID=sysdba;PWD=masterkey;DBNAME=192.168.0.247:\Reloj_Checador\Database\Reloj_Checador.FDB;"

'Inabilita el boton de SQL Server.
'CmdEjecutaCerraduras.Enabled = False

'INICIALIZA O CREA LA INSTANCIA DE LA CONECCION
'Set Conexion = New ADODB.Connection
'Set Conexion2 = New ADODB.Connection
'Set Conexion3 = New ADODB.Connection

'Conexion.ConnectionString = GConectionString
'Conexion2.ConnectionString = GConectionString2
'Conexion3.ConnectionString = GConectionString3

'Conexion.Open
'Conexion2.Open
'Conexion3.Open

'EjecutaBatch     'Para respaldos previos
Conteo   'Hace conteo de registros en la Bd
Carga.Enabled = False


On Error GoTo 0
    Exit Sub
Form_Load_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario PasaInformacion"

End Sub

Public Sub Conteo()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Conteo-------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Conteo
' Procedimiento : Conteo
' Fecha         : 07/04/2007 17:56
' Autor         : Miguel
' Propósito     :HACE UN CONTEO PARA SABER LA CANT. DE REGISTROS EN CADA BD
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Conteo
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Conteo------------------------------------------------------------------------------------------------------------------------------------------------------------
'Hace los conteos de registros de cada BD

On Error GoTo Conteo_Error
                
                'Abrimos Conexion a bd reloj checador huella - access
                Set Conexion3 = New ADODB.Connection
                Conexion3.ConnectionString = GConectionString3
                Conexion3.Open
                               
                  Set r3 = New ADODB.Recordset  'Conteo de Registros de Reloj Checador Huella
                  Call Abrir_Recordset3(r3, "SELECT U.Badgenumber, U.street, U.TITLE, U.OPHONE, D.Deptname FROM USERINFO U, DEPARTMENTS D WHERE D.DeptID LIKE U.DefaultDeptID")
                  
                  If r3.RecordCount > 0 Then
                        vconteoReloj = r3.RecordCount
                        TextReloj.Text = vconteoReloj
                        TextReloj.Locked = True
                  End If
                
                'Abrimos Conexion a bd de comidas - mysql -
                Set Conexion = New ADODB.Connection
                Conexion.ConnectionString = GConectionString
                Conexion.Open
                
                 Set r = New ADODB.Recordset   'Conteo de Registros de Comidas
                  Call Abrir_Recordset(r, "Select * From EmpleadosComidas")
                  
                  If r.RecordCount > 0 Then
                        vconteoComidas = r.RecordCount
                        TextComidas.Text = vconteoComidas
                        TextComidas.Locked = True
                        If vconteoComidas <> vconteoReloj Then
                              TextComidas.ForeColor = &HFF
                        Else
                              TextComidas.ForeColor = &HC00000
                        End If
                  End If
                  
                  
                'Abrimos Conexion a bd de credenciales - access -
                Set Conexion2 = New ADODB.Connection
                Conexion2.ConnectionString = GConectionString2
                Conexion2.Open
                
                  
                  Set r2 = New ADODB.Recordset  'Conteo de Registros de Credenciales
                  Call Abrir_Recordset2(r2, "Select * From DAALX")
                  
                  If r2.RecordCount > 0 Then
                        vconteoCredenciales = r2.RecordCount
                        TextCredenciales.Text = vconteoCredenciales
                        TextCredenciales.Locked = True
                        If vconteoCredenciales <> vconteoReloj Then
                              TextCredenciales.ForeColor = &HFF
                        Else
                              TextCredenciales.ForeColor = &HC00000
                        End If
                  End If
                  
                  'Set r4 = New ADODB.Recordset  'Conteo de Registros de Cerradura de Huella
                  'Call Abrir_Recordset4(r4, "SELECT * FROM Principal")
                  
                  'If r4.RecordCount > 0 Then
                  '      vconteoCerraduras = r4.RecordCount
                  '      TxtCerraduras.Text = vconteoCerraduras
                  '      TxtCerraduras.Locked = True
                  'End If
                  
                  'Cerramos Conexiones ------
                r2.Close
                Set r2 = Nothing
                r.Close
                Set r = Nothing
                r3.Close
                Set r3 = Nothing
                Conexion2.Close
                Set Conexion2 = Nothing
                Conexion.Close
                Set Conexion = Nothing
                Conexion3.Close
                Set Conexion3 = Nothing

On Error GoTo 0
    Exit Sub
Conteo_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Conteo de Formulario PasaInformacion"
                                    
End Sub


Private Sub ejecuta_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ejecuta_Click-------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ejecuta_Click
' Procedimiento : ejecuta_Click
' Fecha         : 07/04/2007 17:56
' Autor         : Miguel
' Propósito     :PASA LOS REGISTROS DE DE BD DE RELOJ CHECADOR A BD DE CREDENCIALES
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ejecuta_Click
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ejecuta_Click------------------------------------------------------------------------------------------------------------------------------------------------------------

On Error GoTo ejecuta_Click_Error

vClaveEmp = ""
vNombreEmp = ""
vDepto = ""
vPuesto = ""
vNSS = ""

ContadorCredenciales = 0
vCodigoEmpNuevo = ""
vSumaCodigoEmpNuevo = ""
vCodigoEmpNuevo = ""

Select Case MsgBox("Esta seguro de pasar los registros del Reloj Checador a BD de Credenciales?", vbOKCancel Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "Pregunta")

      Case vbOK  'Eligió pasar reloj -->> Credenciales
                
                'Abrimos Conexion a bd de credenciales - access -
                Set Conexion2 = New ADODB.Connection
                Conexion2.ConnectionString = GConectionString2
                Conexion2.Open
                
                  Set r2 = New ADODB.Recordset  'BD de Credenciales  ------------------------------------
                  Call Abrir_Recordset2(r2, "Select * From DAALX")
                  
                  'Abrimos Conexion a bd reloj checador huella - access
                    Set Conexion3 = New ADODB.Connection
                    Conexion3.ConnectionString = GConectionString3
                    Conexion3.Open
                  
                  Set r3 = New ADODB.Recordset  'BD de Reloj Checador ----------------------------------
                  'Call Abrir_Recordset3(r3, "SELECT E.CLAVE, E.NOMBRE, D.NOMBRE NOMDEP, P.NOMBRE NOMP, E.DIR1 FROM EMPLEADOS E, DEPTOS D, PUESTOS P WHERE D.IDDEPTO = E.IDDEPTO AND P.IDPUESTO = E.IDPUESTO")
                  '
                  Call Abrir_Recordset3(r3, "SELECT U.Badgenumber, U.street, U.TITLE, U.OPHONE, D.Deptname FROM USERINFO U, DEPARTMENTS D WHERE D.DeptID LIKE U.DefaultDeptID")
                                            'If r2.RecordCount > 0 Then
                          If r3.RecordCount > 0 Then
                          
                                r2.MoveFirst
                                r3.MoveFirst
                          
                                  Do Until r3.EOF   'Hasta que sea el ultimo registro de Reloj
                                            
                                            If IsNull(r3!Badgenumber) Then
                                                  
                                                  Call MsgBox("Existe una clave en blanco, revise la bd Reloj Checador Huella!", vbExclamation Or vbSystemModal Or vbDefaultButton1, "Dato erroneo")
                                                  Exit Sub
                                                  
                                             Else
                                                                                                                           
                                                        If Format(r2!Campo1, "0##") <> Format(r3!Badgenumber, "0##") Then
                                                        
                                                        Set RBuscaClaveEmpleado = New ADODB.Recordset
                                                        Call Abrir_Recordset2(RBuscaClaveEmpleado, "SELECT * FROM DAALX WHERE CAMPO1 LIKE '" & Format(r3!Badgenumber, "0##") & "'")
                                                        
                                                                  If RBuscaClaveEmpleado.RecordCount = 0 Then
                                                                                                                                                
                                                                            r2.AddNew
                                                                                        
                                                                                      'Clave Empleado  /////////////////
                                                                                      vClaveEmp = Format(r3!Badgenumber, "0##")
                                                                                      r2!Campo1 = UCase(vClaveEmp)
                                                                                      
                                                                                        'Nombre Empleado  //////////////////
                                                                                        If IsNull(r3!street) Then
                                                                                              Call MsgBox("Error .- Falta Nombre del Empleado " & "'" & r3!Badgenumber & "'" & ". Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                                                              Screen.MousePointer = vbDefault
                                                                                              Exit Sub
                                                                                        Else
                                                                                            vNombreEmp = r3!street
                                                                                            r2!Campo2 = UCase(vNombreEmp)
                                                                                        End If
                                                                                      
                                                                                        'IMSS del Empleado  //////////////
                                                                                        If IsNull(r3!OPHONE) Then
                                                                                              vNSS = ""
                                                                                        Else ' Lo agrega
                                                                                              vNSS = r3!OPHONE
                                                                                              r2!Campo14 = UCase(vNSS)
                                                                                        End If
                                                                                                                                        
                                                                                        'Departamento del Empleado  /////////
                                                                                        If IsNull(r3!DEPTNAME) Then
                                                                                            Call MsgBox("Error .- Falta el Depto. del Empleado. Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                                                            Screen.MousePointer = vbDefault
                                                                                            Exit Sub
                                                                                        Else
                                                                                            vDepto = r3!DEPTNAME
                                                                                            r2!Campo16 = UCase(vDepto)
                                                                                        End If
                                                                                                                                         
                                                                                        'Puesto del Empleado  /////////////
                                                                                        If IsNull(r3!Title) Then
                                                                                            Call MsgBox("Error .- Falta el Puesto del Empleado. Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                                                            Screen.MousePointer = vbDefault
                                                                                            Exit Sub
                                                                                        Else
                                                                                            vPuesto = r3!Title
                                                                                            r2!Campo17 = UCase(vPuesto)
                                                                                        End If

                                                                            r2.Update
                                                                                        
                                                                              If Err <> 0 Then
                                                                                  MsgBox Err.Description
                                                                                  Err.Clear
                                                                              End If
                                                                        
                                                                              ContadorCredenciales = ContadorCredenciales + 1
                                                                              vCodigoEmpNuevo = vClaveEmp
                                               
                                                                              vSumaCodigoEmpNuevo = vSumaCodigoEmpNuevo & vCodigoEmpNuevo & ","
                                            
                                                                  End If   'Fin de Pregunta si existe Clave Empleado de Reloj en Credenciales
                                                                    
                                                        End If  'Fin de Pregunta si es igual el codigo de empleado
                                                                                                                                           
                                            End If  'Fin de pregunta si es IsNULL
                                                                                                                    
                                            r3.MoveNext  'Avanza un registro en Reloj
                                            'r2.MoveNext  'Avanza un registro en Credenciales
                                                                                                                            
                                   Loop  'Reloj
                    
                    End If  'Fin de pregunta si Reloj tiene datos

      Case vbCancel

End Select

     If Err <> 0 Then
         MsgBox Err.Description
     End If
    'MsgBox "Se transferieron con exito los registros !"
    
    'CERRAMOS CONEXIONES
    'Cerramos Conexiones ------
    r2.Close
    Set r2 = Nothing
    r3.Close
    Set r3 = Nothing
    
    Conexion2.Close
    Set Conexion2 = Nothing
    Conexion3.Close
    Set Conexion3 = Nothing
        
    Call MsgBox("Se transferieron con exito " & ContadorCredenciales & " registros !! " & vSumaCodigoEmpNuevo & " ", vbExclamation Or vbSystemModal, "# Registros transferidos")
    
    Conteo

On Error GoTo 0
    Exit Sub
ejecuta_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ejecuta_Click de Formulario PasaInformacion"

End Sub


Private Sub EjecutaComidas_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------EjecutaComidas_Click-------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------EjecutaComidas_Click
' Procedimiento : EjecutaComidas_Click
' Fecha         : 07/04/2007 17:56
' Autor         : Miguel
' Propósito     :PASA LOS VALORES DE BD DE RELOJ CHECADOR A LA BD DE COMIDAS
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------EjecutaComidas_Click
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------EjecutaComidas_Click------------------------------------------------------------------------------------------------------------------------------------------------------------

On Error GoTo EjecutaComidas_Click_Error

ContadorComidas = 0
vCodigoEmpNuevo = ""
vSumaCodigoEmpNuevo = ""
vCodigoEmpNuevo2 = ""

Select Case MsgBox("Esta seguro de pasar los registros del Reloj Checador a BD de Comidas?", vbOKCancel Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "Pregunta")

      Case vbOK  'Eligió pasar reloj -->> Comidas
            
            'Abrimos Conexion a bd de comidas - mysql -
            Set Conexion = New ADODB.Connection
            Conexion.ConnectionString = GConectionString
            Conexion.Open
            
            Set r = New ADODB.Recordset  'BD Comidas -------------------------------------
            Call Abrir_Recordset(r, "SELECT ClaveEmp, NombreEmp, Puesto, Depto, ClaveCosto From EmpleadosComidas WHERE ClaveCosto <> 'CASTIGO'")
            'Call Abrir_Recordset(r, "Select * From EmpleadosComidas")
                
            'Abrimos Conexion a bd reloj checador huella - access
            Set Conexion3 = New ADODB.Connection
            Conexion3.ConnectionString = GConectionString3
            Conexion3.Open
                
             Set r3 = New ADODB.Recordset  'BD Reloj Checador  --------------------------
             'Call Abrir_Recordset3(r3, "SELECT E.CLAVE, E.NOMBRE, D.NOMBRE NOMDEP, P.NOMBRE NOMP FROM EMPLEADOS E, DEPTOS D, PUESTOS P WHERE D.IDDEPTO = E.IDDEPTO AND P.IDPUESTO = E.IDPUESTO")
             Call Abrir_Recordset3(r3, "SELECT U.Badgenumber, U.street, U.TITLE, U.OPHONE, D.Deptname FROM USERINFO U, DEPARTMENTS D WHERE D.DeptID LIKE U.DefaultDeptID")
             '
              If r.RecordCount > 0 Then
                              
                    If r3.RecordCount > 0 Then

                                If IsNull(r3!Badgenumber) Then
                                      
                                    Call MsgBox("Existe una clave en blanco, revise la bd del Reloj de Huella!", vbExclamation Or vbSystemModal Or vbDefaultButton1, "Dato erroneo")
                                    Exit Sub
                                      
                                 Else
                                 
                                    r.MoveFirst
                                    r3.MoveFirst
                                       
                                    Do Until r3.EOF
                                                                                                                                                                              
                                            Set RBuscaClaveEmpleado = New ADODB.Recordset
                                            Call Abrir_Recordset(RBuscaClaveEmpleado, "SELECT * FROM EmpleadosComidas WHERE ClaveEmp LIKE '" & Format(r3!Badgenumber, "0##") & "'")
                                            
                                            If RBuscaClaveEmpleado.RecordCount = 0 Then ' No existe la clave en BD Comidas
                                                                                                                                            
                                                  r.AddNew
                                                                                                                
                                                        vClaveEmp = Format(r3!Badgenumber, "0##")   'Clave Empleado.
                                                        r!ClaveEmp = UCase(vClaveEmp)
                                                        
                                                        'Nombre
                                                        If IsNull(r3!street) Then
                                                              Call MsgBox("Error .- Falta Nombre del Empleado " & "'" & r3!Badgenumber & "'" & ". Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                              Screen.MousePointer = vbDefault
                                                              Exit Sub
                                                        Else
                                                            vNombreEmp = r3!street  'Nombre Empleado.
                                                            r!NombreEmp = UCase(vNombreEmp)
                                                        End If
                                                        
                                                        'Departamento.
                                                        If IsNull(r3!DEPTNAME) Then
                                                              Call MsgBox("Error .- Falta Depto. del Empleado " & "'" & r3!Badgenumber & "'" & ". Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                              Screen.MousePointer = vbDefault
                                                              Exit Sub
                                                        Else
                                                            vDepto = r3!DEPTNAME  'Departamento del Empleado.
                                                            r!Depto = UCase(vDepto)
                                                        End If
                                                        
                                                        'Puesto.
                                                        If IsNull(r3!Title) Then
                                                            Call MsgBox("Error .- Falta Puesto del Empleado " & "'" & r3!Badgenumber & "'" & ". Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                            Screen.MousePointer = vbDefault
                                                            Exit Sub
                                                        Else
                                                            vPuesto = r3!Title   'Puesto del Empleado.
                                                            r!Puesto = UCase(vPuesto)
                                                        End If
                                                        
                                                        'Tipo de Cobro.
                                                        If r!ClaveEmp = 111 Or r!ClaveEmp = 112 Then
                                                            vCostoComida = "INVITADO"
                                                        Else
                                                            vCostoComida = "FEPSA"
                                                        End If
                                                        r!ClaveCosto = vCostoComida
                                                                  
                                                r.Update
                                                              
                                                    If Err <> 0 Then
                                                        MsgBox Err.Description
                                                        Err.Clear
                                                    End If
                                                                      
                                                    ContadorComidas = ContadorComidas + 1
                                                    vCodigoEmpNuevo = vClaveEmp
                                                    
                                                    vSumaCodigoEmpNuevo = vSumaCodigoEmpNuevo & vCodigoEmpNuevo & ","
                                                             
                                            End If
                                                                             
                                            r3.MoveNext  'Avanza un registro en Reloj
                                            'r.MoveNext  'Avanza un registro en Credenciales
                                                                        
                                    Loop
 
                                                                                                                                                                                                                
                                End If  'Fin de pregunta si es IsNULL
                    
                    End If  'Fin de pregunta si Reloj tiene datos
                                                                                                            
              End If  'Fin de pregunta si Comidas tiene datos

      Case vbCancel

End Select

    If Err <> 0 Then
        MsgBox Err.Description
    End If
    
    'Cerramos Conexiones ------
    r.Close
    Set r = Nothing
    r3.Close
    Set r3 = Nothing
    
    Conexion.Close
    Set Conexion = Nothing
    Conexion3.Close
    Set Conexion3 = Nothing
    
    Call MsgBox("Se transferieron con exito " & ContadorComidas & " registros !! " & vSumaCodigoEmpNuevo & " ", vbExclamation Or vbSystemModal, "# Registros transferidos")
                                   
    Conteo
                                  
 On Error GoTo 0
    Exit Sub
EjecutaComidas_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento EjecutaComidas_Click de Formulario PasaInformacion"

End Sub

Private Sub CmdSalir_Click()
      Unload Me
      Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
      Unload Me
      Close
End Sub

Public Sub EjecutaBatch()

On Error GoTo EjecutaBatch_Error

Shell "C:\respaldosantes.bat", vbNormalFocus

On Error GoTo 0
Exit Sub
EjecutaBatch_Error:
      MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento EjecutaBatch del Formulario PasaInformacion"
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Carga_Click
' Fecha         : 30/10/08 17:27
' Autor         : Miguel Angel
' Propósito     : PARA CARGAR LOS DATOS DE LA TABLA SABANA
'---------------------------------------------------------------------------------------
'
Private Sub Carga_Click()

On Error GoTo Carga_Click_Error


    Set r2 = New ADODB.Recordset  'TABLA SABANA  ------------------------------------
    Call Abrir_Recordset2(r2, "SELECT * FROM Sabana")
    Dim conteo2 As Double
    conteo2 = r2.RecordCount
    
    Set r3 = New ADODB.Recordset  'BD de Reloj Checador ----------------------------------
    'Call Abrir_Recordset3(r3, "SELECT E.CLAVE, E.NOMBRE, D.NOMBRE NOMDEP, P.NOMBRE NOMP, E.DIR1 FROM EMPLEADOS E, DEPTOS D, PUESTOS P WHERE D.IDDEPTO = E.IDDEPTO AND P.IDPUESTO = E.IDPUESTO")
    '
    Call Abrir_Recordset3(r3, "SELECT U.Badgenumber, U.street, U.TITLE, U.OPHONE, D.Deptname FROM USERINFO U, DEPARTMENTS D WHERE D.DeptID LIKE U.DefaultDeptID")
                              'If r2.RecordCount > 0 Then
            If r3.RecordCount > 0 Then
            
                  r2.MoveFirst
                  r3.MoveFirst
            
                    Do Until r3.EOF   'Hasta que sea el ultimo registro de Reloj
                                                                                                             
                                If r3!Badgenumber <> r2!CodigoEmpleado Then
                                
                                Set RBuscaClaveEmpleado = New ADODB.Recordset
                                Call Abrir_Recordset3(RBuscaClaveEmpleado, "SELECT * FROM Sabana WHERE CodigoEmpleado LIKE '" & r3!Badgenumber & "'")
                                
                                          If RBuscaClaveEmpleado.RecordCount > 0 Then

                                                            vNombreEmp = r2!Nombre
                                                            r3!street = UCase(vNombreEmp)
                                                        
                                                            vNSS = r2!IMSS
                                                            r3!OPHONE = UCase(vNSS)

                                                            vPuesto = r2!Puesto
                                                            r3!Title = UCase(vPuesto)

                                                        r3.Update
                                                                
                                                      If Err <> 0 Then
                                                          MsgBox Err.Description
                                                          Err.Clear
                                                      End If
                                                
                                                      ContadorCredenciales = ContadorCredenciales + 1
                                                                             
                                                      vSumaCodigoEmpNuevo = vSumaCodigoEmpNuevo & vCodigoEmpNuevo & ","
                                                                         
                                          End If   'Fin de Pregunta si existe Clave Empleado de Reloj en Credenciales
                                          
                                Else
                                
                                        vNombreEmp = r2!Nombre
                                        r3!street = UCase(vNombreEmp)
                                    
                                        vNSS = r2!IMSS
                                        r3!OPHONE = UCase(vNSS)

                                        vPuesto = r2!Puesto
                                        r3!Title = UCase(vPuesto)

                                    r3.Update
                                            
                                          If Err <> 0 Then
                                              MsgBox Err.Description
                                              Err.Clear
                                          End If
                                    
                                          ContadorCredenciales = ContadorCredenciales + 1
                                                                 
                                          vSumaCodigoEmpNuevo = vSumaCodigoEmpNuevo & vCodigoEmpNuevo & ","
                                            
                                End If  'Fin de Pregunta si es igual el codigo de empleado
                              r2.MoveNext
                              r3.MoveNext  'Avanza un registro en Credenciales
                                                                                                              
                     Loop  'Reloj
      
      End If  'Fin de pregunta si Reloj tiene datos

    Call MsgBox("Se transferieron con exito " & ContadorCredenciales & " registros !! " & vSumaCodigoEmpNuevo & " ", vbExclamation Or vbSystemModal, "# Registros transferidos")

On Error GoTo 0
Exit Sub
Carga_Click_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Carga_Click of Formulario PasaInformacion"
End Sub


Private Sub CmdEjecutaCerraduras_Click()

On Error GoTo CmdEjecutaCerraduras_Click_Error

ContadorCerraduras = 0
vCodigoEmpNuevo = ""
vSumaCodigoEmpNuevo = ""
vCodigoEmpNuevo2 = ""

Select Case MsgBox("Esta seguro de pasar los registros del Reloj Checador a BD de Cerraduras?", vbOKCancel Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "Pregunta")

      Case vbOK  'Eligió pasar reloj -->> Cerraduras

            Set r4 = New ADODB.Recordset  'Conteo de Registros de Cerradura de Huella
            Call Abrir_Recordset4(r4, "SELECT * FROM Principal")
                
                Set r3 = New ADODB.Recordset  'BD Reloj Checador  --------------------------
                'Call Abrir_Recordset3(r3, "SELECT E.CLAVE, E.NOMBRE, D.NOMBRE NOMDEP, P.NOMBRE NOMP FROM EMPLEADOS E, DEPTOS D, PUESTOS P WHERE D.IDDEPTO = E.IDDEPTO AND P.IDPUESTO = E.IDPUESTO")
                Call Abrir_Recordset3(r3, "SELECT U.Badgenumber, U.street, U.TITLE, U.OPHONE, D.Deptname FROM USERINFO U, DEPARTMENTS D WHERE D.DeptID LIKE U.DefaultDeptID")
                '
                 If r4.RecordCount > 0 Then
                              
                    If r3.RecordCount > 0 Then

                                If IsNull(r3!Badgenumber) Then
                                      
                                    Call MsgBox("Existe una clave en blanco, revise la bd del Reloj de Huella!", vbExclamation Or vbSystemModal Or vbDefaultButton1, "Dato erroneo")
                                    Exit Sub
                                      
                                 Else
                                 
                                    r4.MoveFirst
                                    r3.MoveFirst
                                       
                                    Do Until r3.EOF
                                                                                                                                                                              
                                            Set RBuscaClaveEmpleado = New ADODB.Recordset
                                            Call Abrir_Recordset4(RBuscaClaveEmpleado, "SELECT * FROM Principal WHERE Cod_Nomina = '" & Format(r3!Badgenumber, "###") & "'")
                                            
                                            If RBuscaClaveEmpleado.RecordCount = 0 Then ' No existe la clave en BD Cerraduras
                                                
                                                  r4.AddNew
                                                                                                                
                                                        vClaveEmp = Format(r3!Badgenumber, "###")   'Clave Empleado.
                                                        r4!Cod_Nomina = UCase(vClaveEmp)
                                                        
                                                        'Nombre
                                                        If IsNull(r3!street) Then
                                                              Call MsgBox("Error .- Falta Nombre del Empleado " & "'" & r3!Badgenumber & "'" & ". Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                              Screen.MousePointer = vbDefault
                                                              Exit Sub
                                                        Else
                                                            vNombreEmp = r3!street  'Nombre Empleado.
                                                            r4!Nombre = UCase(vNombreEmp)
                                                        End If
                                                        
                                                        'Departamento.
                                                        If IsNull(r3!DEPTNAME) Then
                                                              Call MsgBox("Error .- Falta Depto. del Empleado " & "'" & r3!Badgenumber & "'" & ". Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                              Screen.MousePointer = vbDefault
                                                              Exit Sub
                                                        Else
                                                            vDepto = r3!DEPTNAME  'Departamento del Empleado.
                                                            r4!Depto = UCase(vDepto)
                                                        End If
                                                        
                                                        'Puesto.
                                                        If IsNull(r3!Title) Then
                                                            Call MsgBox("Error .- Falta Puesto del Empleado " & "'" & r3!Badgenumber & "'" & ". Regrese a la Aplicacion del Reloj Checador de Huella e ingrese el dato faltante.", vbCritical Or vbSystemModal, "Falta Nombre Empleado !!")
                                                            Screen.MousePointer = vbDefault
                                                            Exit Sub
                                                        Else
                                                            vPuesto = r3!Title   'Puesto del Empleado.
                                                            r4!Puesto = UCase(vPuesto)
                                                        End If
                                                        
                                                        'Tipo de Usuario Cerradura.
                                                        r4!Tipo = "U"
                                                        
                                                        'Activo.
                                                        r4!Activo = "SI"
                                                        
                                                        'Time Update.
                                                        r4!TimeUpdate = Date
                                                                  
                                                r4.Update
                                                              
                                                    If Err <> 0 Then
                                                        MsgBox Err.Description
                                                        Err.Clear
                                                    End If
                                                                      
                                                    ContadorCerraduras = ContadorCerraduras + 1
                                                    vCodigoEmpNuevo = vClaveEmp
                                                    
                                                    vSumaCodigoEmpNuevo = vSumaCodigoEmpNuevo & vCodigoEmpNuevo & ","
                                                             
                                            End If
                                                                             
                                            r3.MoveNext  'Avanza un registro en Reloj
                                            'r.MoveNext  'Avanza un registro en Credenciales
                                                                        
                                    Loop
 
                                                                                                                                                                                                                
                                End If  'Fin de pregunta si es IsNULL
                    
                    End If  'Fin de pregunta si Reloj tiene datos
                                                                                                            
              End If  'Fin de pregunta si Comidas tiene datos

      Case vbCancel

End Select

    If Err <> 0 Then
        MsgBox Err.Description
    End If
                                   
    Conteo
    
    Call MsgBox("Se transferieron con exito " & ContadorComidas & " registros !! " & vSumaCodigoEmpNuevo & " ", vbExclamation Or vbSystemModal, "# Registros transferidos")

On Error GoTo 0
Exit Sub
CmdEjecutaCerraduras_Click_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CmdEjecutaCerraduras_Click of Formulario PasaInformacion"

End Sub
