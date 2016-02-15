using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.IO;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;

namespace WordGeneration
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string DEFAULT_FOLDER = "docs";

        private Stopwatch _sw;

        public MainWindow()
        {
            InitializeComponent();

            // Set default values
            txtGenerationFolder.Text = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), DEFAULT_FOLDER);
            pbProgress.Value = 0;
        }

        async private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            bool parsed = true;
            long nbDocs;
            int nbFilesPerFolder, nbFoldersPerFolder;
            parsed &= long.TryParse(txtNbDocs.Text, out nbDocs);
            parsed &= int.TryParse(txtFilesPerFolder.Text, out nbFilesPerFolder);
            parsed &= int.TryParse(txtFoldersPerFolder.Text, out nbFoldersPerFolder);

            if (!parsed)
            {
                System.Windows.MessageBox.Show(this, "Error while parsing values : check your string formats", "Word Generation", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                var gen = new DocGeneration();
                gen.Progress += Gen_Progress;
                btnGo.IsEnabled = false;

                _sw = Stopwatch.StartNew();
                await gen.Generate(nbDocs, txtGenerationFolder.Text, nbFilesPerFolder, nbFoldersPerFolder, (int)sldParaPerDoc.Value, (int)sldWordsPerPara.Value);
                pbProgress.Value = 100;
                btnGo.IsEnabled = true;
                Process.Start(txtGenerationFolder.Text);
                System.Windows.MessageBox.Show(this, $"{nbDocs} documents generated in {_sw.ElapsedMilliseconds} ms", "Word Generation", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }



        private void Gen_Progress(object sender, DocGenerationEventArgs e)
        {
            // Update progress bar
            Dispatcher.BeginInvoke(new Action(() => pbProgress.Value = (int)(e.DocsGenerated * 100 / e.NumberOfDocs)));
        }

        private void btnGenerationFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Choose the folder where the documents will be generated";
                if (!string.IsNullOrWhiteSpace(txtGenerationFolder.Text)) dlg.SelectedPath = txtGenerationFolder.Text;
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtGenerationFolder.Text = dlg.SelectedPath;
                }
            }
        }
    }
}
