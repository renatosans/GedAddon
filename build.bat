
@REM REQUISITOS:
@REM É necessário adicionar o caminho do ANT na variável de ambiente PATH
@REM o ANT por sua vez procura pela variável de ambiente JAVA_HOME que deve apontar para o JDK


@REM CARACTERÍSTICAS:
@REM Os caminhos utilizados no build são relativos, definidos em relação a CURRENT_DIR
@REM Nos projetos abaixo o build é feito com a task <exec/> do ANT chamando o compilador csc.exe do .NET Framework
@REM O caminho do .NET Framework 3.5 foi adicionado na variável de ambiente TARGET_FRAMEWORK que é onde o csc.exe se encontra
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

