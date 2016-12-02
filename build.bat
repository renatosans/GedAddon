
@REM REQUISITOS:
@REM � necess�rio adicionar o caminho do ANT na vari�vel de ambiente PATH
@REM o ANT por sua vez procura pela vari�vel de ambiente JAVA_HOME que deve apontar para o JDK


@REM CARACTER�STICAS:
@REM Os caminhos utilizados no build s�o relativos, definidos em rela��o a CURRENT_DIR
@REM Nos projetos abaixo o build � feito com a task <exec/> do ANT chamando o compilador csc.exe do .NET Framework
@REM O caminho do .NET Framework 3.5 foi adicionado na vari�vel de ambiente TARGET_FRAMEWORK que � onde o csc.exe se encontra
SET CURRENT_DIR=%CD%
SET TARGET_FRAMEWORK=C:\Windows\Microsoft.NET\Framework\v3.5


CD ..\..\subversion_jobAccounting\Source\DotNet\ClassLibraries\DocMageFramework
CALL ANT
CD /d %CURRENT_DIR%

CD ..\..\subversion_jobAccounting\Source\DotNet\ClassLibraries\SharpZipLib
CALL ANT
CD /d %CURRENT_DIR%

CD GedAddon
CALL ANT
CD /d %CURRENT_DIR%

CD GedAddonSetup
CALL ANT
CD /d %CURRENT_DIR%

