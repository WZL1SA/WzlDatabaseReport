using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WzlDatabaseReport.Report
{
    class PdfReport
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Author { get; set; }
        public DateTime Time { get; set;}
        public string CoverImagePath { get; set; }

        private Document document;

        public PdfDocument CreateReport()
        {
            var pdfDocument = new PdfDocument();

            document = new Document();
            DefineStyles(document);

            CreateTitlePage();

            CreateTableOfContents();

            CreateIntroduction();

            CreateDataTable();

            var renderer = new DocumentRenderer(document);
            renderer.PrepareDocument();
            for (int i = 0; i < renderer.FormattedDocument.PageCount; i++)
            {
                var gfx = XGraphics.FromPdfPage(pdfDocument.AddPage());
                renderer.RenderPage(gfx, i+1);
            }
            return pdfDocument;
        }

        private void CreateDataTable()
        {
            var section = document.AddSection();
            var paragraph = section.AddParagraph("Tabela danych");
            paragraph.Style = "Title";
            paragraph.AddBookmark("Tabela");

            var table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = getColor(System.Drawing.Color.Black);
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            // Before you can add a row, you must define the columns
            var column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn("4.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;

            // Create the header of the table
            var row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = getColor(System.Drawing.Color.LightBlue);

            string[] headers = new string[] { "ID", "Imię","Drugie imię", "Nazwisko", "e-mail", "Hasło"};

            for (int i = 0; i < headers.Length; i++)
            {
                row.Cells[i].AddParagraph(headers[i]);
                row.Cells[i].Format.Font.Bold = true;
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
                row.Cells[i].MergeDown = 1;
            }
            row = table.AddRow();
            row.Shading.Color = getColor(System.Drawing.Color.LightBlue);

            table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, getColor(System.Drawing.Color.Black));


            using (wzlEntities context = new wzlEntities())
            {
                int k = 0;
                foreach(var item in context.Customer)
                {
                    row = table.AddRow();                    
                    row.Cells[0].AddParagraph(item.CustomerID.ToString());
                    row.Cells[1].AddParagraph(item.FirstName);
                    row.Cells[2].AddParagraph(item.MiddleName ?? "");
                    row.Cells[3].AddParagraph(item.LastName);
                    row.Cells[4].AddParagraph(item.EmailAddress ?? "");
                    row.Cells[5].AddParagraph(item.PasswordHash);
                    if (k % 2 == 0)
                    {
                        row.Format.Shading.Color = getColor(System.Drawing.Color.LightGray);
                    } else
                    {
                        row.Format.Shading.Color = getColor(System.Drawing.Color.NavajoWhite);
                    }
                    k++;
                }
            }
        }

        private void CreateIntroduction()
        {
            var section = document.AddSection();
            
            var paragraph = section.AddParagraph("Wprowadzenie");
            paragraph.AddBookmark("Wprowadzenie");
            paragraph.Style = "Title";

            paragraph = section.AddParagraph();            
            paragraph.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin nibh augue, suscipit a, scelerisque sed, lacinia in, mi. Cras vel lorem. Etiam pellentesque aliquet tellus. Phasellus pharetra nulla ac diam. Quisque semper justo at risus. Donec venenatis, turpis vel hendrerit interdum, dui ligula ultricies purus, sed posuere libero dui id orci. Nam congue, pede vitae dapibus aliquet, elit magna vulputate arcu, vel tempus metus leo non est. Etiam sit amet lectus quis est congue mollis. Phasellus congue lacus eget neque. Phasellus ornare, ante vitae consectetuer consequat, purus sapien ultricies dolor, et mollis pede metus eget nisi. Praesent sodales velit quis augue. Cras suscipit, urna at aliquam rhoncus, urna quam viverra nisi, in interdum massa nibh nec erat.");

            section.AddParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin nibh augue, suscipit a, scelerisque sed, lacinia in, mi. Cras vel lorem. Etiam pellentesque aliquet tellus. Phasellus pharetra nulla ac diam. Quisque semper justo at risus. Donec venenatis, turpis vel hendrerit interdum, dui ligula ultricies purus, sed posuere libero dui id orci. Nam congue, pede vitae dapibus aliquet, elit magna vulputate arcu, vel tempus metus leo non est. Etiam sit amet lectus quis est congue mollis. Phasellus congue lacus eget neque. Phasellus ornare, ante vitae consectetuer consequat, purus sapien ultricies dolor, et mollis pede metus eget nisi. Praesent sodales velit quis augue. Cras suscipit, urna at aliquam rhoncus, urna quam viverra nisi, in interdum massa nibh nec erat.");

            paragraph = section.AddParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin nibh augue, suscipit a, scelerisque sed, lacinia in, mi. Cras vel lorem. Etiam pellentesque aliquet tellus. Phasellus pharetra nulla ac diam. Quisque semper justo at risus. Donec venenatis, turpis vel hendrerit interdum, dui ligula ultricies purus, sed posuere libero dui id orci. Nam congue, pede vitae dapibus aliquet, elit magna vulputate arcu, vel tempus metus leo non est. Etiam sit amet lectus quis est congue mollis. Phasellus congue lacus eget neque. Phasellus ornare, ante vitae consectetuer consequat, purus sapien ultricies dolor, et mollis pede metus eget nisi. Praesent sodales velit quis augue. Cras suscipit, urna at aliquam rhoncus, urna quam viverra nisi, in interdum massa nibh nec erat.");
            paragraph.Format.Font.Color = getColor(System.Drawing.Color.DarkGreen);          

        }

        private void CreateTableOfContents()
        {
            if (document == null) return;
            var section = document.AddSection();

            var paragraph = section.AddParagraph("Spis treści");
            paragraph = section.AddParagraph();
            var hyperlink = paragraph.AddHyperlink("Wprowadzenie");
            hyperlink.AddText("Wprowadzenie do raportu");
            paragraph.AddLineBreak();
            hyperlink = paragraph.AddHyperlink("Tabela");
            hyperlink.AddText("Podsumowanie tabelaryczne");
            paragraph.AddLineBreak();
            hyperlink = paragraph.AddHyperlink("Wykres");
            hyperlink.AddText("Grupowanie na wykresie");
            paragraph.AddLineBreak();
            hyperlink = paragraph.AddHyperlink("Changes");
            hyperlink.AddText("Wykaz zmian");
        }

        private void CreateTitlePage()
        {
            if (document == null) return;

            var section = document.AddSection();
            var image = section.AddImage(CoverImagePath);

            image.Width = Unit.FromCentimeter(8);
            image.LockAspectRatio = true;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.RelativeVertical = RelativeVertical.Line;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Left;
            // otaczanie tekstu- "Ramka" w MS Word
            image.WrapFormat.Style = WrapStyle.Through;

            var titleFrame = section.AddTextFrame();
            titleFrame.Height = "3.0cm";
            titleFrame.Width = "7.0cm";
            titleFrame.Left = ShapePosition.Left;
            titleFrame.RelativeHorizontal = RelativeHorizontal.Margin;
            titleFrame.Top = "10.0cm";
            titleFrame.RelativeVertical = RelativeVertical.Page;

            var titleBox = titleFrame.AddParagraph(Title);
            titleBox.Style = "Title";
            titleBox.Format.Font.Color = getColor(System.Drawing.Color.DarkGray);

            var authorBox = titleFrame.AddParagraph(Author);
            authorBox.Style = "Normal";
            
        }


        void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            Style style = document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Verdana";

            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called Table based on style Normal
            style = document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Verdana";
            style.Font.Name = "Times New Roman";
            style.Font.Size = 9;

            // Create a new style called Reference based on style Normal
            style = document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);
            style.ParagraphFormat.Font.Color = getColor(System.Drawing.Color.Black);

            // Define new style "Title", which we will use for cover Title
            style = document.Styles.AddStyle("Title", "Normal");
            style.ParagraphFormat.Font.Size = 20;
            style.ParagraphFormat.Font.Name = "Calibri";
            style.ParagraphFormat.Font.Color = getColor(System.Drawing.Color.DarkGray);
            style.ParagraphFormat.SpaceBefore = "10mm";
            style.ParagraphFormat.SpaceAfter = "10mm";

            

        }

        Color getColor(System.Drawing.Color color)
        {
            return new Color(
                color.R,
                color.G,
                color.B);
        }
    }
}
