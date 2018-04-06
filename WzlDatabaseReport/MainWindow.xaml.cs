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
using WzlDatabaseReport.Report;

namespace WzlDatabaseReport
{
    /// <summary>
    /// Logika interakcji dla klasy MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreateReportBtn_Click(object sender, RoutedEventArgs e)
        {
            // Tworzenie raportu PDF
            var pdfReport = new PdfReport
            {
                Title = "Raport z bazy AzureDB, tabela Customer",
                Author = "Ja",
                CoverImagePath = "image.png",
                SearchName = GetFirstLetterTb.Text
            };
            // Generowanie raportu
            var report = pdfReport.CreateReport();
            // Zapis do pliku
            report.Save($"report_{Guid.NewGuid()}.pdf");
            // Komunikat o udanym zapisie
            MessageBox.Show("Zapisano plik", "Sukces", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CreateExcelReportBtn_Click(object sender, RoutedEventArgs e)
        {
            // Tworzenie raportu PDF
            var exelReport = new ExcellReport
            {
                Title = "Raport z bazy AzureDB, tabela Customer",
                Author = "Ja",                
                SearchName = GetFirstLetterTb.Text
            };
            // Generowanie raportu
            var report = exelReport.CreateReport();
            // Zapis do pliku
            report.SaveAs(new System.IO.FileInfo( $"report_{Guid.NewGuid()}.xlsx"));
            // Komunikat o udanym zapisie
            MessageBox.Show("Zapisano plik", "Sukces", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
