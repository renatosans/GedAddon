using System;
using System.IO;
using System.Threading;
using System.Reflection;
using System.Windows.Forms;


namespace GedAddon
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(String[] args)
        {
            // Cria o handler para carregamento de dlls embarcadas no executável
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveHandler);
            // Cria o handler para exceções não tratadas
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(NotifyUnhandledException);
            Application.ThreadException += new ThreadExceptionEventHandler(NotifyThreadException);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            // Cria o objeto que controla os recursos do addon
            AddonController controller = new AddonController();

            // Inicia o loop de mensagens(fica aguardando os eventos)
            if (controller.IsAttached) Application.Run();
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

        private static void NotifyUnhandledException(Object sender, UnhandledExceptionEventArgs e)
        {
            Exception unhandledException = (Exception)e.ExceptionObject;
            MessageBox.Show(unhandledException.Message);
        }

        private static void NotifyThreadException(Object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message);
        }
    }

}
