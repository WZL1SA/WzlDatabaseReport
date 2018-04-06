using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WzlDatabaseReport.Report
{
    internal class ExcellReport : Report<ExcelPackage>
    {
        private ExcelWorkbook _workbook;
        private ExcelPackage _package;

        public override ExcelPackage CreateReport()
        {
         
            _package = new ExcelPackage();
            _workbook = _package.Workbook;
            _package.Workbook.Properties.Author = Author;
            _package.Workbook.Properties.Title = Title;
            _package.Workbook.Properties.Keywords = "training Exel C#";


            CreateTableSheet();

            CreateChart();

                return _package ;
            
               
        }

        private void CreateChart()
        {
            var sheet = _workbook.Worksheets.Add("Wykres");

            using (var context = new wzlEntities())
            {
                var chartData = context.Customer.GroupBy(customer => customer.LastName)
                    .Select(group => new { group.Key, Counter = group.Count() })
                        .GroupBy(entry => entry.Counter)
                        .Select(value => new { value.Key, Counter = (double)value.Count() });

                sheet.Cells[1, 1].Value = "Klucz";
                sheet.Cells[1, 2].Value = "Wartość";

                var k = 2;

                foreach (var item in chartData)
                {
                    sheet.Cells[k, 1].Value = item.Key;
                    sheet.Cells[k, 2].Value = item.Counter;
                    k++;
                }

                var chart = (sheet.Drawings.AddChart("PieChart", eChartType.Pie3D) as ExcelPieChart);

                chart.Title.Text = "Total";
                chart.SetPosition(0, 0, 5, 5);
                chart.SetSize(600, 300);
                var valueAddress = new ExcelAddress(2, 2, k - 1, 2); // zakres dla warości
                var legendAddress = new ExcelAddress(2, 1, k - 1, 1); // zakres danych dla legendy
                chart.Series.Add(valueAddress.Address, legendAddress.Address);

                chart.DataLabel.ShowCategory = true;
                chart.DataLabel.ShowPercent = true;

                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Border.Fill.Color = Color.DarkBlue;


                

                sheet.Cells[k, 2].Formula = $"SUM({valueAddress.Address})";
            }
            sheet.Calculate();
           
            }

        private void CreateTableSheet()
        {
            var sheet = _workbook.Worksheets.Add("Dane w tabeli");
            //dodać nagłówek
            sheet.Cells[1, 1].Value = "ID";
            sheet.Cells[1, 2].Value = "Imię";
            sheet.Cells[1, 3].Value = "Drugie imię";
            sheet.Cells[1, 4].Value = "Nazwisko";
            sheet.Cells[1, 5].Value = "e-mail";
            sheet.Cells[1, 6].Value = "Hasło";



            //przygotować dane
            //Ok now format the values;
            using (var range = sheet.Cells[1, 1, 1, 6])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                range.Style.Font.Color.SetColor(Color.White);
            }


            using (var context = new wzlEntities())
            {
                var r = 2;

                var items = context.Customer
                    .Where(customer => customer.FirstName.StartsWith(SearchName));
                //var items = context.Customer;

                foreach (var item in items)    // gdy chcemy wszystkie to używamy context.Customer)
                {

                    sheet.Cells[r, 1].Value = item.CustomerID.ToString();
                    sheet.Cells[r, 2].Value = item.FirstName;
                    sheet.Cells[r, 3].Value = item.MiddleName;
                    sheet.Cells[r, 4].Value = item.LastName;
                    sheet.Cells[r, 5].Value = item.EmailAddress;
                    sheet.Cells[r, 6].Value = item.PasswordHash;
                    sheet.Cells[r, 1, r, 6].Style.Fill.PatternType = ExcelFillStyle.LightVertical;
                    sheet.Cells[r, 1, r, 6].Style.Fill.BackgroundColor
                        .SetColor(r % 2 == 0 ? Color.LightBlue : Color.LightSalmon);
                    r++;
                }
            }

            sheet.Cells.AutoFitColumns(); // dopasowanie szerokości kolumn do tekstu


            

            //iterować dane do arkusza

        }
    }
}
