using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Security.Principal;


namespace GedAddonSetup
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(String[] args)
        {
            if (args.Length != 1)
            {
                MessageBox.Show("This installer must be run from Sap Business One", "GedAddon setup", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // Verifica se o instalador está sendo executado com permissões administrativas
            WindowsIdentity windowsIdentity = WindowsIdentity.GetCurrent();
            WindowsPrincipal windowsPrincipal = new WindowsPrincipal(windowsIdentity);
            Boolean executingAsAdmin = windowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator);

            // Para debugar este programa execute o Visual Studio como administrador, caso contrário
            // o programa não vai parar nos "breakpoints" (isso se deve ao código de controle do UAC)
            Process process = Process.GetCurrentProcess();
            if ((process.ProcessName.ToUpper().Contains("VSHOST")) && (!executingAsAdmin))
            {
                String errorMessage = "Execute o Visual Studio com permissões administrativas para debugar!";
                MessageBox.Show(errorMessage, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Verifica se a caixa de dialogo do UAC (User Account Control) é necessária
            if (!executingAsAdmin)
            {
                // Pede elevação de privilégios (executa como administrador se o usuário concordar), o programa
                // atual é encerrado e uma nova instancia é executada com os privilégios concedidos
                ProcessStartInfo processInfo = new ProcessStartInfo();
                processInfo.Verb = "runas";
                processInfo.FileName = Application.ExecutablePath;
                processInfo.Arguments = '"' + Environment.GetCommandLineArgs()[1] + '"';
                try { Process.Start(processInfo); } catch { }
                return;
            }

            if (!EventLog.SourceExists("Ged Addon Setup"))
                EventLog.CreateEventSource("Ged Addon Setup", null);
            // EventLog.WriteEntry("Ged Addon Setup", "Iniciando instalador...");
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveHandler);

            if (args[0] == "/U")
            {
                InstallationHandler handler = new InstallationHandler();
                handler.Uninstall();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmInstall());
        }

        // Tenta carregar a dll embarcada(dentro do executável) caso não encontre ela no diretório da aplicação
        private static Assembly AssemblyResolveHandler(Object sender, ResolveEventArgs args)
        {
            String[] assemblyDetails = args.Name.Split(new Char[] { ',' });
            String assemblyFilename = assemblyDetails[0] + ".dll";

            Assembly thisExe = Assembly.GetExecutingAssembly();
            Stream dllStream = thisExe.GetManifestResourceStream(assemblyFilename);
            if (dllStream != null)
            {
                Byte[] rawAssembly = new Byte[dllStream.Length];
                dllStream.Read(rawAssembly, 0, (int)dllStream.Length);
                return Assembly.Load(rawAssembly);
            }

            return null; // falha no carregamento
        }
    }

}
