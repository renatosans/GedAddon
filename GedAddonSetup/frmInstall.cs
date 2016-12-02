//
// 1) Addon installer program should be able to accept a command line parameter from SBO.
//    This parameter is a string built from 2 strings devided by "|".
//    The first string is the path recommended by SBO for installation folder.
//    The second string is the location of "AddOnInstallAPI.dll".
//    For example, a command line parameter that looks like this:
//    "C:\MyAddon|C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
//    Means that the recommended installation folder for this addon is "C:\MyAddon" and the
//    location of "AddOnInstallAPI.dll" is "C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
//
// 2) When the installation is complete the installer must call the function 
//    "EndInstall" from "AddOnInstallAPI.dll" to inform SBO the installation is complete.
//    This dll contains 3 functions that can be used during the installation.
//    The functions are: 
//         1) EndInstall - Signals SBO that the installation is complete.
//         2) SetAddOnFolder - Use it if you want to change the installation folder.
//         3) RestartNeeded - Use it if your installation requires a restart, it will cause
//            the SBO application to close itself after the installation is complete.
//    All 3 functions return a 32 bit integer. There are 2 possible values for this integer.
//    0 - Success, 1 - Failure.
//
// 3) After your installer is ready you need to create an add-on registration file.
//    In order to create it you have a utility - "Add-On Registration Data Creator"
//    "..\SAP Manage\SAP Business One SDK\Tools\AddOnRegDataGen\AddOnRegDataGen.exe"
//    This utility creates a file with the extention 'ard', you will be asked to 
//    point to this file when you register your addon.

using System;
using System.IO;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using DocMageFramework.AppUtils;


namespace GedAddonSetup
{
    public partial class frmInstall : Form
    {
        private String destinationFolder;

        private String addonInstallDllFolder;


        public frmInstall()
        {
            InitializeComponent();
            this.Icon = new Icon(IOHandler.GetEmbeddedResource("Setup.ico"));
            this.imgLogo.Image = new Bitmap(IOHandler.GetEmbeddedResource("Logo.png"));
            this.txtInfo.ForeColor = Color.Navy;
            this.txtInfo.Font = new Font("Arial", 11.25F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            this.txtInfo.Text = "Bem vindo ao instalador do Addon Datacopy GED. Este instalador irá configurar a estação de " +
                                "trabalho para o Gerenciamento de Documentos através da interface do SAP Business One.";
        }

        private void frmInstall_Shown(object sender, EventArgs e)
        {
            EventLog.WriteEntry("Ged Addon Setup", "Setup Commandline:" + Environment.NewLine + Environment.CommandLine);
            String commandLine = Environment.GetCommandLineArgs()[1];
            String[] commandLineElements = commandLine.Split(char.Parse("|"));
            if (commandLineElements.Length != 2)
            {
                btnInstall.Enabled = false;
                return;
            }
            destinationFolder = commandLineElements[0];
            addonInstallDllFolder = Path.GetDirectoryName(commandLineElements[1]);
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            btnInstall.Enabled = false;

            InstallationHandler handler = new InstallationHandler();
            // Extrai os arquivos de instalação do pacote (zip)
            if (!handler.ExtractInstallationFiles()) return;
            // Copia os arquivos extraídos para o diretório de destino
            if (!handler.CopyInstallationFiles(destinationFolder)) return;
            // Adiciona dados do Addon ao registro do windows
            if (!handler.AddToWindowsRegistry(destinationFolder, addonInstallDllFolder)) return;

            // Informa o SAP Business One que a instalação foi concluída e encerra o programa
            handler.AddonInstallFinished(addonInstallDllFolder);
            this.Close();
        }
    }

}
