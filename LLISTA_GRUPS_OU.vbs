'--------------------------------------------------------------------
'OBTENIR DISTINGUSHED NAME OU
OU = inputbox("Entra OU (Grups/Usuaris): ")
Set objShell = WScript.CreateObject("WScript.Shell")
Set ObjExec = objShell.Exec("cmd.exe /c dsquery ou -name " & OU & " | findstr GRUPS" )
Do
    strOU = ObjExec.StdOut.ReadLine()
    strOU = truncate_one_L(strOU)
    strOU = truncate_one_R(strOU)
    'wscript.echo strOU
Loop While Not ObjExec.Stdout.atEndOfStream
    '------------------------------------------------
    'FUNCIONS RETALLA COMES PRINCIPI I FINAL DEL TEXT
    Function truncate_one_L(s)
        If Left(s, 1) = """" Then 
        truncate_one_L = Right(s, Len(s) - 1) 
        Else 
        truncate_one_L = s
        End If
    End Function

    Function truncate_one_R(s)
        If Right(s, 1) = """" Then 
        truncate_one_R = Left(s, Len(s) - 1) 
        Else 
        truncate_one_R = s
        End If
    End Function
        '------------------------------------------------
'--------------------------------------------------------------------

'------------------------------------------------------------
'FINESTRA EMERGENT MOSATRADA PER PANTALLA MENTRE DURA L'ACCIÓ
'--------------------------------------
Set objExplorer = CreateObject ("InternetExplorer.Application")

objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Left = 200
objExplorer.Top = 200
objExplorer.Width = 200
objExplorer.Height = 400 
objExplorer.Visible = 1             

objExplorer.Document.Title = "LLISTAT USUARIS OU " & OU

objExplorer.Document.Body.InnerHTML = log_pantalla 

'--------------------------------------
'CONNNEXIO AMB ACTIVE DIRECTORY
'--------------------------------------
Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = "SELECT Name FROM 'LDAP://" & strOU & "' WHERE objectCategory='Group'"
'EXECUTA CONSULTA
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	'Wscript.Echo objRecordSet.Fields("Name").Value
	log_pantalla = log_pantalla & objRecordSet.Fields("Name").Value & "<br>"
	objExplorer.Document.Body.InnerHTML = log_pantalla 
    objRecordSet.MoveNext
Loop

'TANCA FINESTRA IE

