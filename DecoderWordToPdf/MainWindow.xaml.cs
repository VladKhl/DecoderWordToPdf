using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace DecoderWordToPdf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        //Document readyDOC;
        string directory;
        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            
            
        }

        private void UpLoadWordBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofdDoc = new OpenFileDialog();
            ofdDoc.Filter = "Word Documents (.docx)|*.docx";
            ofdDoc.FilterIndex = 1;
            if (ofdDoc.ShowDialog() == true)
            {
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                wordDocument = appWord.Documents.Open(ofdDoc.FileName);
                directory = ofdDoc.FileName;
                uploadpb.Visibility = Visibility.Visible;
                uploadnametb.Visibility = Visibility.Visible;
                var maximum = uploadpb.Maximum;
                Action action = () => { uploadpb.Value++; uploadnametb.Text = "Loading..."; };
                Action actionEnd = () => { uploadnametb.Text = ofdDoc.SafeFileName; };
                var task = new System.Threading.Tasks.Task(() =>
                {
                    for (var i = 0; i < maximum; i++)
                    {
                        uploadpb.Dispatcher.Invoke(action);
                        Thread.Sleep(100);
                        if (i == maximum - 1)
                        {
                            uploadpb.Dispatcher.Invoke(actionEnd);
                        }
                    }
                });
                task.Start();
            }
        }

        private void DownloadPDFBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "Pdf Files|*.pdf";

            if (saveFileDialog1.ShowDialog() == true)
            {
                wordDocument.ExportAsFixedFormat(saveFileDialog1.FileName, WdExportFormat.wdExportFormatPDF);
                Process.Start("taskkill", "/F /IM WINWORD.EXE* /T");
                MessageBox.Show("Successfull");
            }
        }
    }
}
