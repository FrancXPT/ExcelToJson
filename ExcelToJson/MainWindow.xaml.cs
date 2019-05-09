using Excel2Json;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
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

namespace ExcelToJson
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Excel2JsonConverter Conver = new Excel2JsonConverter();
        SaveFileDialog sv = new SaveFileDialog();
        OpenFileDialog op = new OpenFileDialog();
        public MainWindow()
        {
            InitializeComponent();
            Expande.Header = "Changer Location de Sauvegarde";
            SaveLocationtxt.IsEnabled = false;

        }
        private void ChosseBtn_Click(object sender, RoutedEventArgs e)
        {
            op.Filter = "Excel File (*.xslx)|*.xlsx;*.xsl";
            Nullable<bool> result = op.ShowDialog();
            if (result == true)
            {
                chossedFiletxt.IsEnabled = false;
                chossedFiletxt.Text = op.FileName;
                Expande.IsExpanded = true;
            }
        }

        private void Close_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.Close();
        }

        private void SaveLocationBtn_Click(object sender, RoutedEventArgs e)
        {
            sv.Filter = "Json File (*.Json)|*.json";
            Nullable<bool> result = sv.ShowDialog();
            if (result == true)
            {
                SaveLocationtxt.Text = sv.FileName;
                
            }
        }

        private void ConvertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (chossedFiletxt.Text != "" && SaveLocationtxt.Text != "")
            {
                try
                {
                    Conver.JsonPath = SaveLocationtxt.Text;
                    progressPanel.Visibility = Visibility.Visible;
                    Conver.ExcelFileToJson(chossedFiletxt.Text);
                    progressPanel.Visibility = Visibility.Hidden;
                    MessageBox.Show("File Converted \n"+SaveLocationtxt.Text);
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Remlpir Tout Les Champts. !!");
            }
        }
    }
}
