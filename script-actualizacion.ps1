#
#requires -version 4
<#
.SYNOPSIS
    Revisa si hay actualizaciones de windows disponibles y las realiza.
.DESCRIPTION
    El script genera una lista de las actualizaciones disponibles, y si se desea, las aplica
.INPUTS
    [System.String]
.OUTPUTS
    [System.Object]
.NOTES
    Version:        1.0
    Author:         Santino Gianninoto
    Creation Date:  Febrero 2022
    Purpose/Change: Actualizar windows automaticamente de manera remota
    Useful URLs: localhost
.EXAMPLE
    PS C:\>.\Script-Actualizacion.ps1
    "Que acción desea ejecutar:" 2

    Este ejemplo devuelve el historial de actualizaciones de los equipos especificado mediate ingreso de un archivo con los nombres de estos.
.EXAMPLE
    PS C:\>.\Script-Actualizacion.ps1
    "Que acción desea ejecutar:" 4
cd
    Este ejemplo revisa si hay actualizaciones disponibles para los equipos especificado mediate ingreso de un archivo con los nombres de estos
.EXAMPLE
    PS C:\>.\Script-Actualizacion.ps1
    "Que acción desea ejecutar:" 6

    Este ejemplo instala actualizaciones disponibles para para los equipos especificado mediate ingreso de un archivo con los nombres de estos
#>

# ==============================
# Importación de módulo requerido
# ============================== 
Install-Module PSWindowsUpdate
Import-Module PSWindowsUpdate

function historial-manual{
    # ==============================
    # Obtención de datos requeridos
    # ============================== 
    $Equipos = Read-Host -Prompt "Ingrese el nombre del equipo" #Ingresar nombre del equipo

    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    $dias = 15
    if ($r -eq "Si"){
        #si se guarda log
        $log_loc = Read-Host -Prompt "Ingrese la ruta completa a la carpeta donde quiera guardar el log" #ruta completa de la carpeta donde guardar los logs
        $log_loc = "$($log_loc)/$(get-date -Format yyyy-MM-dd)-HistorialAct.log" #se añade el nombre del archivo log a la ruta
        #revision de historial de actualizaciones
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias) | 
        Select-Object -Property ComputerName, Result, Date, Title, KB, Description |
        Format-Table -AutoSize -Wrap | Out-File -Encoding utf8 -FilePath $log_loc -Force
        Write-Host "archivo guardado" #confirmación de log
    }
    elseif ($r -eq "No"){
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias) | 
        Select-Object -Property ComputerName, Result, Date, Title, KB |
        Format-Table -AutoSize -Wrap
    }
}
function historial-archivo{
    # ==============================
    # Obtención de datos requeridos
    # ============================== 
    $Equipos = Get-Content -Path (Read-Host -Prompt "Ingrese la ubicación del archivo") #Ingresar ruta con el archivo con nombres de máquinas o direcciones IP

    # =========================================
    # Revisión de Historial de actualizaciones
    # =========================================
    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    $dias = 15 #cantidad de días atras para tener en cuenta
    if ($r -eq "Si"){
        #si se guarda log
        $log_loc = Read-Host -Prompt "Ingrese la ruta completa a la carpeta donde quiera guardar el log" #ruta completa de la carpeta donde guardar los logs
        $log_loc = "$($log_loc)/$(get-date -Format yyyy-MM-dd)-HistorialAct.log" #se añade el nombre del archivo log a la ruta
        #revision de historial de actualizaciones
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias)|
        Select-Object -Property ComputerName, Result, Date, Title, KB, Description|
        Format-Table -AutoSize -Wrap | Out-File -Encoding UTF8 -FilePath $log_loc -Force
        Write-Host "log guardado" #confirmación de log
    }
    elseif ($r -eq "No"){
        #revision de historial de actualizaciones
        Get-WUHistory -ComputerName $Equipos -MaxDate (Get-Date).AddDays(-$dias) |
        Select-Object -Property ComputerName, Result, Date, KB, Title |
        Format-Table -AutoSize -Wrap
    }
}

function revisar-manual{
    # ==============================
    # Obtención de datos requeridos
    # ============================== 
    $Equipos = Read-Host -Prompt "Ingrese el nombre del equipo" #Ingresar nombre del equipo

    # ===============================
    # Revisión de actualizaciones
    # =============================== 
    #preguntar si desea guardar log
    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    if ($r -eq "Si"){
        #si se guarda log
        $log_loc = Read-Host -Prompt "Ingrese la ruta completa a la carpeta donde quiera guardar el log" #ruta completa de la carpeta donde guardar los logs
        $log_loc = "$($log_loc)/$(get-date -Format yyyy-MM-dd)-ActPendientes.log" #se añade el nombre del archivo log a la ruta
        #revision de actualizaciones pendientes
        Get-WindowsUpdate -ComputerName $Equipos | Format-Table -AutoSize -Wrap | 
        Out-File -Encoding utf8 -FilePath $log_loc -ErrorAction SilentlyContinue
        Write-Host "archivo guardado" #confirmacion de guardado de logs
    }
    elseif ($r -eq "No"){
        #si no se guarda log
        #revisión de actualizaciones pendientes
        Get-WindowsUpdate -ComputerName $Equipos |
        Select-Object -Property ComputerName, Result, Date, KB, Title |
        Format-Table -AutoSize -Wrap
    }
}



function revisar-archivo{
    # ==============================
    # Obtención de datos requeridos
    # ============================== 
    $Equipos = Get-Content -Path (Read-Host -Prompt "Ingrese la ubicación del archivo") #Ingresar ruta con el archivo con nombres de máquinas o direcciones IP

    # ===============================
    # Revisión de actualizaciones
    # =============================== 
    #preguntar si desea guardar log
    $r = Read-Host -Prompt "¿Desea guardar el resultado en un archivo? (Si/No)"
    if ($r -eq "Si"){
        #si se guarda log
        $log_loc = Read-Host -Prompt "Ingrese la ruta completa a la carpeta donde quiera guardar el log" #ruta completa de la carpeta donde guardar los logs
        $log_loc = "$($log_loc)/$(get-date -Format yyyy-MM-dd)-ActPendientes.log" #se añade el nombre del archivo log a la ruta
        #revision de actualizaciones pendientes
        Get-WindowsUpdate -ComputerName $Equipos | Format-Table -AutoSize -Wrap | 
        Out-File -Encoding utf8 -FilePath $log_loc -ErrorAction SilentlyContinue
        Write-Host "log guardado" #confirmacion de guardado de logs
    }
    elseif ($r -eq "No"){
        #si no se guarda log
        #revisión de actualizaciones pendientes
        Get-WindowsUpdate -ComputerName $Equipos |
        Format-Table -AutoSize -Wrap
    }
}



function enviar-reporte($log_loc){
    # ===================================================
    # Configuración SMPT: usuario, contraseña, SSL, etc.
    # ===================================================
    $email_username = "helpdesk@gmail.org.py" #usuario de correo
    $email_password = "contraseña segura" #contraseña del correo
    $email_smtp_host = "mail.mail.org.py" #servidor de correo
    $email_smtp_port = 587 #puerto
    $email_smtp_SSL = 1 #usar ssl -> 1
    $email_from_address = "helpdesk@gmail.org.py" #correo desde el cual enviar
    $email_to_addressArray = @("mail-1@gmail.org.py", "mail-2@gmail.org.py") #correos a los que enviar

    # ======================================================================
    # Configuración del mensaje: de, para quienes, asunto, cuerpo, adjuntar
    # ======================================================================
    $message = new-object Net.Mail.MailMessage #creación de objeto mensaje
    $message.From = $email_from_address #se enstablece el correo desde el que enviar
    #se añade a los destinatarios de la array a la lista de quienes enviar el correo
    foreach ($to in $email_to_addressArray) {
        $message.To.Add($to)
    }
    $message.Subject = "Informe Automático de Actualización de Windows" #asunto del correo
    $message.Body = "Se adjunta log de instalación de actualizaciones." #cuerpo del correo
    #se crea el oobjeto de archivos adjuntos y se adjunta el log
    $Attachment  = New-Object System.Net.Mail.Attachment($log_loc) 
    $message.Attachments.Add($Attachment)

    # ============================================
    # Creación de objeto SMPT y envío del mensaje
    # ============================================ 
    $smtp = new-object Net.Mail.SmtpClient($email_smtp_host, $email_smtp_port) #se crea el onjeto SMPT client y se establece el servidr y el puerto
    $smtp.EnableSSL = $email_smtp_SSL #se configura SSL
    $smtp.Credentials = New-Object System.Net.NetworkCredential($email_username, $email_password) #creacion del objeto credenciales, se añade el usuario y contraseña
    #se envía el mensaje
    $smtp.send($message)
    $message.Dispose()
    Write-Host "reporte enviado" #confirmación de envío
}

function instalar-manual{
    # ==============================
    # Obtención de datos requeridos
    # ============================== 
    $Equipos = Read-Host -Prompt "Ingrese el nombre del equipo" #nombre o IP del equipo a actualizar
    $log_loc = Read-Host -Prompt "Ingrese la ruta completa a la carpeta donde quiera guardar el log" #ruta completa de la carpeta donde guardar los logs
    $log_loc = "$($log_loc)/$(get-date -Format yyyy-MM-dd)-WindowsUpdate.log" #se añade el nombre del archivo log a la ruta

    # ===============================
    # Instalación de actualizaciones
    # =============================== 
    $r = Read-Host -Prompt "¿Desea reiniciar automáticamente de ser necesario? (si/no)" #se pregunta si se desea reiniciar automáticamente
    if ($r -eq "si"){
        #instala actualizaciones y genera un log - reinicio autománico
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -AutoReboot -Verbose |
        Format-Table -AutoSize -Wrap | Out-File -Encoding utf8 -FilePath $log_loc -Force
        Write-Host "log guardado" #confirmación de guardado de log
    }
    elseif ($r -eq "no"){
        #instala actualizaciones y genera un log - sin reinicio autománico
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -IgnoreReboot -Verbose|
        Format-Table -AutoSize -Wrap | Out-File $log_loc -force
        Write-Host "log guardado" #confirmación de guardado de log
    }

    # ==================================
    # Envío de reporte mediante función
    # ==================================
    enviar-reporte $log_loc
}

function instalar-archivo{
    # ==============================
    # Obtención de datos requeridos
    # ============================== 
    $Equipos = Get-Content -Path (Read-Host -Prompt "Ingrese la ubicación del archivo")
    $log_loc = Read-Host -Prompt "Ingrese la ruta completa a la carpeta donde quiera guardar el log"
    $log_loc = "$($log_loc)/$(get-date -Format yyyy-MM-dd)-WindowsUpdate.log"
    
    # ===============================
    # Instalación de actualizaciones
    # =============================== 
    $r = Read-Host -Prompt "¿Desea reiniciar automáticamente de ser necesario? (si/no)" #se pregunta si se desea reiniciar automáticamente
    if ($r -eq "si"){
        #instala actualizaciones y genera un log - reinicio autománico
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -AutoReboot -Verbose |
        Format-Table -AutoSize -Wrap | Out-File $log_loc -force
        Write-Host "log guardado" #confirmación de guardado de log
    }
    elseif ($r -eq "no"){
        #instala actualizaciones y genera un log - sin reinicio autománico
        Get-WindowsUpdate -ComputerName $Equipos -Install -AcceptAll -IgnoreReboot -Verbose |
        Format-Table -AutoSize -Wrap | Out-File $log_loc -force
        Write-Host "log guardado" #confirmación de guardado de log
    }

    # ==================================
    # Envío de reporte mediante función
    # ==================================
    enviar-reporte $log_loc
}

# ============================================================================================================================================================================= #
# ================================================================================  MAIN  ===================================================================================== #
# ============================================================================================================================================================================= #
$ban = 1
while($ban -ne 0){
    # ==============================
    # Menú a imprimir en pantalla
    # ============================== 
    Write-Host "Que acción desea ejecutar:"
    Write-Host "[1] -> obtener historial de actualización (ingreso manual nom. comp.)"
    Write-Host "[2] -> obtener historial de actualización (ingreso por archivo .txt de nom. comp.)"
    Write-Host "[3] -> revisar actualizaciones disponibles  (ingreso manual nom. comp.)"
    Write-Host "[4] -> revisar actualizaciones disponibles (ingreso por archivo .txt de nom. comp.))"
    Write-Host "[5] -> instalar actualizaciones disponibles  (ingreso manual nom. comp.)"
    Write-Host "[6] -> instalar actualizaciones disponibles (ingreso por archivo .txt de nom. comp.))"
    Write-Host "[7] -> SALIR"
    $resp = Read-Host #opción elegida por usuario

    # ======================================================
    # Estructura Switch - Que ejecutar según opción elegida
    # ======================================================
    switch($resp){
        1 {historial-manual}                    #devuelve el historial de actualizaciones de el equipo especificado mediate ingreso de nombre del equipo manualmente
        2 {historial-archivo}                   #devuelve el historial de actualizaciones de los equipos especificado mediate ingreso de un archivo con los nombres de estos
        3 {revisar-manual}                      #revisa si hay actualizaciones disponibles para el equipo especificado mediate ingreso de nombre del equipo manualmente
        4 {revisar-archivo}                     #revisa si hay actualizaciones disponibles para los equipos especificado mediate ingreso de un archivo con los nombres de estos
        5 {instalar-manual}                     #instala actualizaciones disponibles para el equipo especificado mediate ingreso de nombre del equipo manualmente
        6 {instalar-archivo}                    #instala actualizaciones disponibles para para los equipos especificado mediate ingreso de un archivo con los nombres de estos
        7 {Write-Host "saliendo..."; $ban = 0}  #termina la ejecución del programa

        default {Write-Host "opción no válida"} #opción default, pide ingresar una opción válida
    }
}