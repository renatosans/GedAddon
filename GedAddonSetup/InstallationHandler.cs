using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using DocMageFramework.FileUtils;
using ICSharpCode.SharpZipLib.Zip;


namespace GedAddonSetup
{
    public class InstallationHandler
    {
        private String installationFilesFolder;


        // EndInstall - Signals SBO that the installation is complete.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern Int32 EndInstall();

        // EndUninstall - Signals SBO that the addon removal is is complete.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern Int32 EndUninstall(String path, Boolean succeed);

        // SetAddOnFolder - Use it if you want to change the installation folder.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern Int32 SetAddOnFolder(String srrPath);

        // RestartNeeded - Use it if your installation requires a restart, it will cause
        // the SBO application to close itself after the installation is complete.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern Int32 RestartNeeded();


        public InstallationHandler()
        {
            this.installationFilesFolder = null;
        }

        public Boolean ExtractInstallationFiles()
        {
            // Para gerar o arquivo de instalação execute o build do sistema (build.bat), que vai montar
            // a pasta DebugData com os arquivos necessários.
            Assembly thisExe = Assembly.GetExecutingAssembly();
            Stream zipStream = thisExe.GetManifestResourceStream("GedAddonFiles.zip");
            if (!File.Exists("GedAddonFiles.zip") && (zipStream == null))
            {
                ShowError("Não foi possivel encontrar o arquivo de instalação.");
                return false;
            }

            installationFilesFolder = Path.Combine(Path.GetTempPath(), "GedAddonFiles");
            if (Directory.Exists(installationFilesFolder))
                Directory.Delete(installationFilesFolder, true);

            // Verifica se o arquivo de instalação está embarcado(dentro do executável), escolhendo entre
            // descompactar a partir do arquivo em disco ou o arquivo embarcado
            FastZip zipManager = new FastZip();
            if (zipStream == null)
                zipManager.ExtractZip("GedAddonFiles.zip", installationFilesFolder, null);
            else
                zipManager.ExtractZip(zipStream, installationFilesFolder, FastZip.Overwrite.Always, null, null, null, false, true);

            return true;
        }

        public Boolean CopyInstallationFiles(String destinationFolder)
        {
            EventLog.WriteEntry("Ged Addon Setup", "Copiando arquivos para " + destinationFolder);
            // Procura pelos arquivos de instalação 
            FileInfo[] sourceFiles = null;
            try
            {
                DirectoryInfo sourceDirectory = new DirectoryInfo(installationFilesFolder);
                sourceFiles = sourceDirectory.GetFiles("*.*", SearchOption.AllDirectories);
            }
            catch (Exception exc)
            {
                ShowError("Falha ao copiar arquivos! " + Environment.NewLine + exc.Message);
                return false;
            }

            // Remove a instalação pré existente
            try
            {
                // Derruba as instancias que estão rodando do addon ao abrir o SAP
                Process[] addonProcesses = Process.GetProcessesByName("GedAddon");
                foreach ( Process runningAddon in addonProcesses)
                {
                    runningAddon.Kill();
                    runningAddon.WaitForExit(9000);
                }
                // Remove arquivos instalados anteriormente
                if (Directory.Exists(destinationFolder))
                    Directory.Delete(destinationFolder, true);
            }
            catch (Exception exc)
            {
                ShowError("Falha ao copiar arquivos! " + Environment.NewLine + exc.Message);
                return false;
            }

            TargetDirectory targetDirectory = new TargetDirectory(destinationFolder);
            if (!targetDirectory.Mount())
            {
                ShowError("Falha ao copiar arquivos! " + Environment.NewLine + targetDirectory.GetLastError());
                return false;
            }
            if (!targetDirectory.CopyFilesFrom(sourceFiles))
            {
                ShowError("Falha ao copiar arquivos! " + Environment.NewLine + targetDirectory.GetLastError());
                return false;
            }

            try
            {
                // Também copia o instalador para o diretório de destino
                String destFilename = Path.Combine(destinationFolder, "GedAddonSetup.exe"); // uninstname do .ARD
                File.Copy(Application.ExecutablePath, destFilename);
            }
            catch (Exception exc)
            {
                ShowError("Falha ao copiar arquivos! " + Environment.NewLine + exc.Message);
                return false;
            }

            return true;
        }

        public Boolean AddToWindowsRegistry(String installationFolder, String addonInstallDllFolder)
        {
            RegistryKey parentKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", true);
            if (parentKey == null)
            {
                ShowError("Erro ao registrar addon. Não foi encontrada a chave SOFTWARE.");
                return false;
            }

            RegistryKey addonKey = parentKey.OpenSubKey("Ged Addon", true);
            try
            {
                if (addonKey == null) addonKey = parentKey.CreateSubKey("Ged Addon");
                addonKey.SetValue("InstallationFolder", installationFolder);
                addonKey.SetValue("AddonInstallDllFolder", addonInstallDllFolder);
            }
            catch (Exception exc)
            {
                ShowError("Erro ao registrar addon." + exc.Message);
                return false;
            }

            addonKey.Close();
            parentKey.Close();

            return true;
        }

        public Boolean Uninstall()
        {
            EventLog.WriteEntry("Ged Addon Setup", "Removendo Addon...");
            RegistryKey parentKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", true);
            if (parentKey == null) return false;

            RegistryKey addonKey = parentKey.OpenSubKey("Ged Addon", true);
            if (addonKey == null) return false;

            String installationFolder = (String)addonKey.GetValue("InstallationFolder");
            String addonInstallDllFolder = (String)addonKey.GetValue("AddonInstallDllFolder");
            addonKey.Close();
            if (String.IsNullOrEmpty(installationFolder)) return false;
            if (String.IsNullOrEmpty(addonInstallDllFolder)) return false;

            try
            {
                // Remove o diretório de instalação
                if (Directory.Exists(installationFolder))
                    Directory.Delete(installationFolder, true);
                // Remove o registro do Addon
                parentKey.DeleteSubKey("Ged Addon");
                parentKey.Close();
            }
            catch (Exception exc)
            {
                EventLog.WriteEntry("Ged Addon Setup", "Falha na remoção: " + exc.Message);
                return false;
            }

            // Avisa o SAP sobre o termino da operação, trata exceções
            AddonUninstallFinished(addonInstallDllFolder);
            // Avisa o usuário sobre o termino da operação
            ShowInfo("Ged Addon removido com sucesso");
            return true;
        }

        public void AddonInstallFinished(String addonInstallDllFolder)
        {
            // Muda o diretório corrente para o lugar onde está a DLL de instalação de Addons do SAP para
            // que seja possível utilizar as funções EndInstall, EndUninstall, SetAddOnFolder e RestartNeeded
            Environment.CurrentDirectory = addonInstallDllFolder;
            try
            {
                // Informa o SAP Business One que a instalação foi concluída
                EndInstall();
            }
            catch (Exception exc)
            {
                EventLog.WriteEntry("Ged Addon Setup", "Falha chamando EndInstall(). " + exc.Message);
            }
        }

        public void AddonUninstallFinished(String addonInstallDllFolder)
        {
            // Muda o diretório corrente para o lugar onde está a DLL de instalação de Addons do SAP para
            // que seja possível utilizar as funções EndInstall, EndUninstall, SetAddOnFolder e RestartNeeded
            Environment.CurrentDirectory = addonInstallDllFolder;
            try
            {
                // Informa o SAP Business One que a remoção do addon foi concluída
                EndUninstall(null, true);
            }
            catch (Exception exc)
            {
                EventLog.WriteEntry("Ged Addon Setup", "Falha chamando EndUninstall(). " + exc.Message);
            }
        }

        private void ShowError(String errorMessage)
        {
            MessageBox.Show(errorMessage, "Falha na instalação do Addon", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowInfo(String message)
        {
            MessageBox.Show(message, "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

}
