using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office;
using Microsoft.Office.Interop.Outlook;

namespace MailFileFormatter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        private int existemFalhas;

        public MainWindow()
        {
            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            //Stream iconStream = System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/AssemblyInfo;component/rename.ico")).Stream;
            //MyNotifyIcon.Icon = new System.Drawing.Icon(iconStream);
            //MyNotifyIcon.Icon = new System.Drawing.Icon("rename.ico");

            var iconHandle = MailFileFormatter.Properties.Resources.rename.Handle;
            MyNotifyIcon.Icon = System.Drawing.Icon.FromHandle(iconHandle);
            MyNotifyIcon.MouseDoubleClick += MyNotifyIcon_MouseDoubleClick;
            
        }

        private string tratarPreExistencia(string [] arquivos, string nomeAntigo, string nomeNovo, int numeroVersao)
        {
            string numeroVersaoFormatado = " (" + numeroVersao + ")";
            string nomeArquivo = numeroVersao > 0 ? nomeNovo.Insert(nomeNovo.Length - 4, numeroVersaoFormatado) : nomeNovo;
            
            foreach (string a in arquivos)
            {
                if (a.Contains(nomeAntigo))
                {
                    return tratarPreExistencia(arquivos, nomeAntigo, nomeNovo, numeroVersao+1);
                }
            }

            return nomeArquivo;
        }

        private void btnBuscar_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            tbResultado.Text = string.Empty;
            OpenFileDialog ofd = new OpenFileDialog();
            Nullable<bool> result = ofd.ShowDialog();
            string pathPastaSelecionada = string.Empty;
            existemFalhas = 0;

            if (result == true)
            {
                pathPastaSelecionada = System.IO.Path.GetDirectoryName(ofd.FileName);
                tbFolder.Text = pathPastaSelecionada;

                string[] arquivos = Directory.GetFiles(pathPastaSelecionada);
                try
                {
                    foreach (string arq in arquivos)
                    {
                        Stream messageStream = File.Open(arq, FileMode.Open, FileAccess.Read);
                        OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
                        messageStream.Close();
                        string data = string.Empty;

                        string nome = arq.Substring(arq.LastIndexOf("\\") + 1);
                        data = (message.ReceivedDate == DateTime.MinValue ? string.Empty : string.Concat(message.ReceivedDate.Year, "_", message.ReceivedDate.Month, "_", message.ReceivedDate.Day));
                        if(string.IsNullOrEmpty(data))
                        {
                            throw new System.Exception();
                        }
                        string remetente = message.From;
                        string complementoNome = string.Concat(data, " - ", remetente, " - ");
                        string nome1 = string.Concat(complementoNome, nome);

                        if (!string.IsNullOrEmpty(data) && !arq.Contains(data))
                        {
                            if(arq.Length <= 200)
                            {
                                File.Move(arq, arq.Replace(nome, tratarPreExistencia(arquivos, nome, nome1, 0)));
                            }
                            else
                            {
                                existemFalhas++;
                            }
                        }
                    }

                    if (existemFalhas == 0)
                    {
                        tbResultado.Text = "OK";
                    }
                    else
                    {
                        tbResultado.Text = "OK, porém "+existemFalhas+ " arquivos não puderam ser renomeados.";
                    }
                    tbResultado.Foreground = Brushes.Green;
                }
                catch (System.Exception ex)
                {
                    tbResultado.Text = "NOK - Um ou mais arquivos não puderam ser renomeados.";
                    tbResultado.Foreground = Brushes.Red;
                }
            }
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            var desktopWorkingArea = System.Windows.SystemParameters.WorkArea;
            this.Left = desktopWorkingArea.Right - this.Width - 5;
            this.Top = desktopWorkingArea.Bottom - this.Height - 5;
        }

        private void MyNotifyIcon_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            this.WindowState = WindowState.Normal;
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                //MyNotifyIcon.BalloonTipTitle = "Minimize Sucessful";
                //MyNotifyIcon.BalloonTipText = "Minimized the app ";
                //MyNotifyIcon.ShowBalloonTip(400);
                MyNotifyIcon.Visible = true;
            }
            else if (this.WindowState == WindowState.Normal)
            {
                MyNotifyIcon.Visible = false;
                this.ShowInTaskbar = true;
            }
        }

        private void btnLimpar_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            tbFolder.Text = string.Empty;
            tbResultado.Text = string.Empty;
            existemFalhas = 0;
        }
        
    }
}
