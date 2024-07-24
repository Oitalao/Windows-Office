# Verifica se o script está sendo executado com permissões de administrador
If (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Cria um novo processo com permissões elevadas
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo
    $newProcess.FileName = "powershell.exe"
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"" + $MyInvocation.MyCommand.Path + "`""
    $newProcess.Verb = "runas"
    [System.Diagnostics.Process]::Start($newProcess)
    Exit
}


Function Menu {
    Clear-Host
    Write-Host "==================================================="
    Write-Host "MENU INTERATIVO DE ATIVACAO E MANUTENCAO DE SISTEMA"
    Write-Host "MADE BY: OITALAO"
    Write-Host "==================================================="
    Write-Host "0. Modo Hacker"
    Write-Host "1. Limpeza e Manutencao"
    Write-Host "2. Ativar Windows"
    Write-Host "3. Ativar Office"
    Write-Host "4. Debloat Windows 11"
    Write-Host "5. Sair"
    Write-Host "======================"
    $op = Read-Host "Escolha uma opcao"

    Switch ($op) {
        0 { ModoHacker }
        1 { LimpezaEManutencao }
        2 { AtivarWindowsMenu }
        3 { AtivarOfficeMenu }
        4 { DebloatWindows11 }
        5 { Exit }
        Default { Menu }
    }
}

Function ModoHacker {
    Write-Host "Ativando Modo Hacker..." -ForegroundColor Green
    # Configurações de cores para o modo hacker
    $host.UI.RawUI.BackgroundColor = "Black"
    $host.UI.RawUI.ForegroundColor = "Green"
    Clear-Host
    Write-Host "Modo Hacker Ativado" -ForegroundColor Green
    Pause
    Menu
}

Function LimpezaEManutencao {
    Write-Host "Deletando arquivos temporários..."
    Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Windows\Prefetch\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Limpeza de arquivos temporários concluída!"

    Write-Host "Executando a Limpeza de Disco..."
    Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait

    Write-Host "Executando CHKDSK..."
    Start-Process -FilePath "chkdsk.exe" -ArgumentList "/f /r" -NoNewWindow -Wait

    Write-Host "Resetando a configuração de IP..."
    Invoke-Expression "netsh int ip reset"

    Write-Host "Limpando DNS..."
    Invoke-Expression "ipconfig /flushdns"

    Write-Host "Executando SFC..."
    sfc /scannow

    Write-Host "Executando DISM - ScanHealth..."
    dism /online /cleanup-image /scanhealth

    Write-Host "Executando DISM - RestoreHealth (se necessário)..."
    dism /online /cleanup-image /restorehealth

    Write-Host "Atualizando todos os pacotes Winget..."
    winget upgrade --all

    Write-Host "Registrando pacotes AppX para todos os usuários..."
    Get-AppxPackage -AllUsers | Where-Object { $_.InstallLocation -like '*SystemApps*' } | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" }

    Write-Host "Todas as tarefas foram executadas!"
    Pause
    Menu
}

Function AtivarWindowsMenu {
    Clear-Host
    Write-Host "======================"
    Write-Host "ATIVAR WINDOWS"
    Write-Host "======================"
    Write-Host "1. Windows 10"
    Write-Host "2. Windows 11"
    Write-Host "3. Voltar ao menu"
    Write-Host "======================"
    $op2 = Read-Host "Escolha uma opcao"

    Switch ($op2) {
        1 { AtivarWindows10 }
        2 { AtivarWindows11 }
        3 { Menu }
        Default { AtivarWindowsMenu }
    }
}

Function AtivarWindows10 {
    Write-Host "Ativando Windows 10..."
    slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
    slmgr /skms kms8.msguides.com
    slmgr /ato
    Write-Host "Windows 10 ativado!"
    Pause
    AtivarWindowsMenu
}

Function AtivarWindows11 {
    Write-Host "Ativando Windows 11..."
    slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
    slmgr /skms kms8.msguides.com
    slmgr /ato
    Write-Host "Windows 11 ativado!"
    Pause
    AtivarWindowsMenu
}

Function AtivarOfficeMenu {
    Clear-Host
    Write-Host "======================"
    Write-Host "ATIVAR OFFICE"
    Write-Host "======================"
    Write-Host "LEIA COM ATENCAO OS AVISOS!!!!!!!!!!!!!!!!!!!!"
    Write-Host "O Pacote Office se trata de uma aplicação externa ao Windows, por isso deve ser instalado pelo link disponibilizado na opcao 3."
    Write-Host "VOCE DEVE FAZER O PROCEDIMENTO DE INSTALACAO ANTES DA ATIVACAO!"
    Write-Host "SE UM NAO FUNCIONAR, TENTE O OUTRO! O Office tende a dar problema para ativar!"
    Write-Host "1. Metodo 1"
    Write-Host "2. Metodo 2"
    Write-Host "3. Nao tenho o Office"
    Write-Host "4. Voltar ao menu"
    Write-Host "======================"
    $op3 = Read-Host "Escolha uma opcao"

    Switch ($op3) {
        1 { AtivarOfficeMetodo1 }
        2 { AtivarOfficeMetodo2 }
        3 { NaoTenhoOffice }
        4 { Menu }
        Default { AtivarOfficeMenu }
    }
}

Function AtivarOfficeMetodo1 {
    Write-Host "Navegando para o diretório do Microsoft Office..."
    If (Test-Path "$env:ProgramFiles\Microsoft Office\Office16\ospp.vbs") {
        Set-Location -Path "$env:ProgramFiles\Microsoft Office\Office16"
    } ElseIf (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\Office16\ospp.vbs") {
        Set-Location -Path "$env:ProgramFiles(x86)\Microsoft Office\Office16"
    }

    ForEach ($lic in (Get-ChildItem ..\root\Licenses16\ProPlus2019VL*.xrm-ms)) {
        cscript ospp.vbs /inslic:"$($lic.FullName)" > $null
    }

    Write-Host "Ativando o seu Office..."
    cscript //nologo slmgr.vbs /ckms > $null
    cscript //nologo ospp.vbs /setprt:1688 > $null
    cscript //nologo ospp.vbs /unpkey:6MWKP > $null
    cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP > $null

    $i = 1
    Do {
        Switch ($i) {
            1 { $KMS_Srv = "kms7.MSGuides.com" }
            2 { $KMS_Srv = "kms8.MSGuides.com" }
            3 { $KMS_Srv = "kms9.MSGuides.com" }
            Default { Write-Host "Desculpa! Nao ha suporte a sua versao do Office."; Exit }
        }
        cscript //nologo ospp.vbs /sethst:$KMS_Srv > $null
        If ((cscript //nologo ospp.vbs /act) -match "successful") {
            Write-Host "Office ativado com sucesso!"
            Start-Process "http://MSGuides.com"
            Break
        } Else {
            Write-Host "A conexão com o servidor KMS falhou! Tentando conectar a outro..."
            $i++
        }
    } While ($i -le 3)

    Pause
    Menu
}

Function AtivarOfficeMetodo2 {
    Write-Host "Navegando para o diretório do Microsoft Office (Program Files x86)..."
    Set-Location -Path "$env:ProgramFiles(x86)\Microsoft Office\Office16"
    
    Write-Host "Navegando para o diretório do Microsoft Office (Program Files)..."
    Set-Location -Path "$env:ProgramFiles\Microsoft Office\Office16"

    Write-Host "Instalando as licenças KMS..."
    ForEach ($lic in (Get-ChildItem ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms)) {
        cscript ospp.vbs /inslic:"$($lic.FullName)"
    }

    Write-Host "Configurando a porta KMS..."
    cscript ospp.vbs /setprt:1688

    Write-Host "Removendo a chave do produto..."
    cscript ospp.vbs /unpkey:6F7TH > $null

    Write-Host "Inserindo a nova chave do produto..."
    cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

    Write-Host "Configurando o host KMS..."
    cscript ospp.vbs /sethst:e8.us.to

    Write-Host "Ativando o Office..."
    cscript ospp.vbs /act

    Write-Host "Processo concluído."
    Pause
    Menu
}

Function NaoTenhoOffice {
    Write-Host "Abrindo o site do Office..."
    Start-Process "https://msguides.com/download-microsoft-office-windows-os"
    Menu
}

Function DebloatWindows11 {
    Set-ExecutionPolicy Unrestricted -Force
    Invoke-Expression (Invoke-RestMethod 'https://raw.githubusercontent.com/Raphire/Win11Debloat/master/Get.ps1')
    Pause
    Menu
}

Menu
