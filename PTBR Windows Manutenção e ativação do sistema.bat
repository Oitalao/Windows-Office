@echo off
:: Verifica se o script está sendo executado com permissões de administrador
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    echo Por favor, execute este script como administrador.
    echo Clique com o botao direito do mouse no arquivo e "executar como administrador"
    pause
    exit /b
)

:menu
cls
echo ===================================================
echo MENU INTERATIVO DE ATIVACAO E MANUTENCAO DE SISTEMA
echo MADE BY: OITALAO
echo ===================================================
echo 0. Modo Hacker
echo 1. Limpeza e Manutencao
echo 2. Ativar Windows
echo 3. Ativar Office
echo 4. Debloat Windows 11 
echo 5. Sair
echo =======================
set /p op="Escolha uma opcao: "

if %op%==0 goto modo_hacker
if %op%==1 goto limpeza_e_manutencao
if %op%==2 goto ativar_windows
if %op%==3 goto ativar_office
if %op%==4 goto debloat_windows11
if %op%==5 goto sair

:modo_hacker
color 0a
goto menu

:limpeza_e_manutencao
echo Deletando arquivos temporarios...
del /q /f /s "%TEMP%\*.*"
del /q /f /s "C:\Windows\Temp\*.*"
del /q /f /s "C:\Windows\Prefetch\*.*"
echo Limpeza de arquivos temporarios concluida!

echo Executando a Limpeza de Disco...
cleanmgr /sageset:1
cleanmgr /sagerun:1

echo Executando CHKDSK...
chkdsk /f /r

echo Resetando a configuracao de IP...
netsh int ip reset

echo Limpando DNS...
ipconfig /flushdns

echo Executando SFC...
sfc /scannow

echo Executando DISM - ScanHealth...
dism /online /cleanup-image /scanhealth

echo Executando DISM - RestoreHealth (se necessario)...
dism /online /cleanup-image /restorehealth

echo Atualizando todos os pacotes Winget...
winget upgrade --all

echo Registrando pacotes AppX para todos os usuarios...
powershell -Command "Get-AppXPackage -AllUsers | Where-Object { $_.InstallLocation -like '*SystemApps*' } | Foreach { Add-AppxPackage -DisableDevelopmentMode -Register '$($_.InstallLocation)\AppXManifest.xml' }"

echo Todas as tarefas foram executadas!
pause
goto menu

:ativar_windows
cls
echo =======================
echo ATIVAR WINDOWS
echo =======================
echo 1. Windows 10
echo 2. Windows 11
echo 3. Voltar ao menu
echo =======================
set /p op2="Escolha uma opcao: "

if %op2%==1 goto ativar_windows10
if %op2%==2 goto ativar_windows11
if %op2%==3 goto menu

:ativar_windows10
echo Ativando Windows 10...
slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
slmgr /skms kms8.msguides.com
slmgr /ato
echo Windows 10 ativado!
pause
goto ativar_windows

:ativar_windows11
echo Ativando Windows 11...
slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
slmgr /skms kms8.msguides.com
slmgr /ato
echo Windows 11 ativado!
pause
goto ativar_windows

:ativar_office
cls
echo =======================
echo ATIVAR OFFICE
echo =======================
echo LEIA COM ATENCAO OS AVISOS!!!!!!!!!!!!!!!!!!!!
echo O Pacote Office se trate de uma aplicacao externa ao Windows, por isso deve ser instalado pelo link disponibilizado na opcao 3.
echo VOCE DEVE FAZER O PROCEDIMENTO DE INSTALACAO ANTES DA ATIVACAO!
echo SE UM NAO FUNCIONAR, TENTE O OUTRO! O Office tende a dar problema para ativar!
echo 1. Metodo 1
echo 2. Metodo 2
echo 3. Nao tenho o Office
echo 4. Voltar ao menu
echo =======================
set /p op3="Escolha uma opcao: "

if %op3%==1 goto ativar_office_metodo1
if %op3%==2 goto ativar_office_metodo2
if %op3%==3 goto nao_tenho_office
if %op3%==4 goto menu

:ativar_office_metodo1
echo Navegando para o diretório do Microsoft Office...
& (if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16") ^
&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16") ^
&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul) ^
&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul) ^
& echo. ^
& echo ============================================================================ ^
& echo Ativando o seu Office... ^
& cscript //nologo slmgr.vbs /ckms >nul ^
& cscript //nologo ospp.vbs /setprt:1688 >nul ^
& cscript //nologo ospp.vbs /unpkey:6MWKP >nul ^
& cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul ^
& set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul ^
& echo ============================================================================ ^
& echo. ^
& echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo. ^& echo ============================================================================ ^& if errorlevel 2 exit) || (echo A conexao com o meu servidor KMS Falhou! Tentando conectar a outro... & echo Por favor aguarde... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"
goto halt
:notsupported
echo. ^
& echo ============================================================================ ^
& echo Desculpa! Nao ha suporte a sua versao do Office. ^
& echo Por favor tente instalar a ultima disponivel: bit.ly/aiomsp
:halt
pause >nul
goto menu

:ativar_office_metodo2
echo Navegando para o diretório do Microsoft Office (Program Files x86)...
cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

echo Navegando para o diretório do Microsoft Office (Program Files)...
cd /d "%ProgramFiles%\Microsoft Office\Office16"

echo Instalando as licencas KMS...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"

echo Configurando a porta KMS...
cscript ospp.vbs /setprt:1688

echo Removendo a chave do produto...
cscript ospp.vbs /unpkey:6F7TH >nul

echo Inserindo a nova chave do produto...
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

echo Configurando o host KMS...
cscript ospp.vbs /sethst:e8.us.to

echo Ativando o Office...
cscript ospp.vbs /act

echo Processo concluido.
pause
goto menu

:nao_tenho_office
echo Abrindo o site do Office...
start https://msguides.com/download-microsoft-office-windows-os
goto menu

:debloat_windows11
@echo off
:: Definir a política de execução como Unrestricted
powershell -Command "Set-ExecutionPolicy Unrestricted -Force"

:: Executar o script de debloat do Windows 11
powershell -Command "& {Invoke-Expression (Invoke-RestMethod 'https://raw.githubusercontent.com/Raphire/Win11Debloat/master/Get.ps1')}"

pause
goto menu

:sair
echo Saindo em 1 segundo...
timeout /t 1 /nobreak >nul
exit
