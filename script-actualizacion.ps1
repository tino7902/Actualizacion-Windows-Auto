Import-Module PSWindowsUpdate

function historial-manual{
    Write-Host "historial-manual"

    $Equipos = Read-Host -Prompt "Ingrese el nombre del equipo"

    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    $dias = 15
    if ($r -eq "Si"){
        $nom = Read-Host -Prompt "Ingrese un nombre para el archivo"
        $nom = ".\ $nom"
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias) | 
        Select-Object -Property ComputerName, Result, Date, Title, KB, Description |
        Format-Table -AutoSize -Wrap | Out-File -Encoding utf8 -FilePath $nom -ErrorAction SilentlyContinue
        Write-Host "archivo guardado"
    }
    elseif ($r -eq "No"){
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias) | 
        Select-Object -Property ComputerName, Result, Date, Title, KB |
        Format-Table -AutoSize -Wrap
    }
}
function historial-archivo{
    Write-Host "historial-archivo"
    
    $Equipos = Get-Content -Path (Read-Host -Prompt "Ingrese la ubicación del archivo")

    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    $dias = 15
    if ($r -eq "Si"){
        $nom = Read-Host -Prompt "Ingrese un nombre para el archivo"
        $nom = ".\ $nom"
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias)|
        Select-Object -Property ComputerName, Result, Date, Title, KB, Description|
        Format-Table -AutoSize -Wrap | Out-File -Encoding UTF8 -FilePath $nom -ErrorAction SilentlyContinue
        Write-Host "archivo guardado"
    }
    elseif ($r -eq "No"){
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias) | 
        Select-Object -Property ComputerName, Result, Date, KB, Title| 
        Format-Table -AutoSize -Wrap
    }
}


function revisar-manual{
    Write-Host "revisar-act-manual"

    $Equipos = Read-Host -Prompt "Ingrese el nombre del equipo"

    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    if ($r -eq "Si"){
        $nom = Read-Host -Prompt "Ingrese un nombre para el archivo"
        $nom = ".\ $nom"
        Get-WindowsUpdate -ComputerName $Equipos | Format-Table -AutoSize -Wrap | 
        Out-File -Encoding utf8 -FilePath $nom -ErrorAction SilentlyContinue
        Write-Host "archivo guardado"
    }
    elseif ($r -eq "No"){
        Get-WindowsUpdate -ComputerName $Equipos |
        Format-Table -AutoSize -Wrap
    }
}
function revisar-archivo{
    Write-Host "revisar-act-archivo"
    
    $Equipos = Get-Content -Path (Read-Host -Prompt "Ingrese la ubicación del archivo")

    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    $dias = 15
    if ($r -eq "Si"){
        $nom = Read-Host -Prompt "Ingrese un nombre para el archivo"
        $nom = ".\ $nom"
        Get-WindowsUpdate -ComputerName $Equipos |
        Format-Table -AutoSize -Wrap | Out-File -Encoding utf8 -FilePath $nom -ErrorAction SilentlyContinue
        Write-Host "archivo guardado"
    }
    elseif ($r -eq "No"){
        Get-WindowsUpdate -ComputerName $Equipos |
        Format-Table -AutoSize -Wrap
    }
}

# SMTP Variables
#$username = "helpdesk@coopexsanjo.org.py"
#$password = ConvertTo-SecureString "Paraguay_2022" -AsPlainText -Force
#$smpt_host_server = "mail.coopexsanjo.org.py"
#$smpt_port = 587
#destinatarios: "mirun@coopexsanjo.org.py", "dquinonez@coopexsanjo.org.py"
#$destinatarios = "dquinonez@coopexsanjo.org.py"
#$use_SSL = 1 #tiene ssl
#$SMTP_Sender    = "helpdesk@coopexsanjo.org.py"
#$SMTP_Recipient = "dquinonez@coopexsanjo.org.py"
#$cred = New-Object System.Management.Automation.PSCredential ($username, $password)

function enviar-reporte{
    #SMTP configuration: username, password, SSL and so on

    $email_username = "helpdesk@coopexsanjo.org.py";
    $email_password = "Paraguay_2022";
    $email_smtp_host = "mail.coopexsanjo.org.py";
    $email_smtp_port = 587;
    $email_smtp_SSL = 1;
    $email_from_address = "helpdesk@coopexsanjo.org.py";
    $email_to_addressArray = @("mirun@coopexsanjo.org.py", "dquinonez@coopexsanjo.org.py");

    # E-Mail message configuration: from, to, subject, body

    $message = new-object Net.Mail.MailMessage;
    $message.From = $email_from_address;
    foreach ($to in $email_to_addressArray) {
        $message.To.Add($to);
    }
    $message.Subject = "Informe Automático de Actualización de Windows";
    $log_loc = "C:\Users\pasante\Documents\Script Actualizacion (local)\logs\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log"
    $message.Body = "Se adjunta log de instalación de actualizaciones.";
    $Attachment  = New-Object System.Net.Mail.Attachment($log_loc)
    $message.Attachments.Add($Attachment)
    # ------------------------------------------------------ 
    # Create SmtpClient object and send the e-mail message
    # ------------------------------------------------------ 
    $smtp = new-object Net.Mail.SmtpClient($email_smtp_host, $email_smtp_port);
    $smtp.EnableSSL = $email_smtp_SSL;
    $smtp.Credentials = New-Object System.Net.NetworkCredential($email_username, $email_password);
    $smtp.send($message);
    $message.Dispose();
}

function instalar-manual{
    Write-Host "instalar-act-manual"

    $Equipos = Read-Host -Prompt "Ingrese el nombre del equipo"
    $log_loc = ".\logs\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log"
    $r = Read-Host -Prompt "¿Desea reiniciar automáticamente de ser necesario? (si/no)"
    if ($r -eq "si"){
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -AutoReboot -Verbose |
        Format-Table -AutoSize -Wrap | Out-File $log_loc -force
    }
    elseif ($r -eq "no"){
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -IgnoreReboot -Verbose|
        Format-Table -AutoSize -Wrap | Out-File $log_loc -force
    }
    enviar-reporte
}

function instalar-archivo{
    Write-Host "instalar-act-archivo"
    
    $Equipos = Get-Content -Path (Read-Host -Prompt "Ingrese la ubicación del archivo")
    
    $r = Read-Host -Prompt "¿Desea reiniciar automáticamente de ser necesario? (si/no)"
    if ($r -eq "si"){
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -AutoReboot -Verbose |
        Format-Table -AutoSize -Wrap | Out-File ".\logs\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log" -force
    }
    elseif ($r -eq "no"){
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -IgnoreReboot -Verbose |
        Format-Table -AutoSize -Wrap | Out-File ".\logs\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log" -force
    }
    enviar-reporte
}


$ban = 1
while($ban -ne 0){
    Write-Host "Que acción desea ejecutar:"
    Write-Host "[1] -> obtener historial de actualización (ingreso manual nom. comp.)"
    Write-Host "[2] -> obtener historial de actualización (ingreso por archivo .txt de nom. comp.)"
    Write-Host "[3] -> revisar actualizaciones disponibles  (ingreso manual nom. comp.)"
    Write-Host "[4] -> revisar actualizaciones disponibles (ingreso por archivo .txt de nom. comp.))"
    Write-Host "[5] -> instalar actualizaciones disponibles  (ingreso manual nom. comp.)"
    Write-Host "[6] -> instalar actualizaciones disponibles (ingreso por archivo .txt de nom. comp.))"
    Write-Host "[7] -> SALIR"
    $resp = Read-Host
    switch($resp){
        1 {historial-manual}          #devuelve el historial de actualizaciones de el equipo especificado mediate ingreso de nombre del equipo manualmente
        2 {historial-archivo}         #devuelve el historial de actualizaciones de los equipos especificado mediate ingreso de un archivo con los nombres de estos
        3 {revisar-manual}        #revisa si hay actualizaciones disponibles para el equipo especificado mediate ingreso de nombre del equipo manualmente
        4 {revisar-archivo}       #revisa si hay actualizaciones disponibles para los equipos especificado mediate ingreso de un archivo con los nombres de estos
        5 {instalar-manual}       #instala actualizaciones disponibles para el equipo especificado mediate ingreso de nombre del equipo manualmente
        6 {instalar-archivo}      #instala actualizaciones disponibles para para los equipos especificado mediate ingreso de un archivo con los nombres de estos
        7 {Write-Host "saliendo..."; $ban = 0} 
        8 {enviar-reporte}

        default {Write-Host "opción no válida"}
    }
}