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
    /// <summary>
    /// Klasa do tworzenia raportów z bazy danych
    /// </summary>
    internal class PdfReport
    {
        /// <summary>
        /// Tytuł raportu
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// Opis raportu (podtytuł)
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Dane autora/autorów
        /// </summary>
        public string Author { get; set; }
        /// <summary>
        /// Czas raportu
        /// </summary>
        public DateTime Time { get; set;}
        /// <summary>
        /// Ścieżka do pliku z logo
        /// </summary>
        public string CoverImagePath { get; set; }

        private Document _document;

        /// <summary>
        /// Metoda eksportuje raport do pliku PDF
        /// </summary>
        /// <returns>Obiekt reprezentujący dokument PDF</returns>
        public PdfDocument CreateReport()
        {
            // Dokument MigraDoc
            _document = new Document();
            
            // Definicja styli
            DefineStyles(_document);

            // Tworzymy stronę tytułową
            CreateTitlePage();

            // Dodanie spisu treści
            CreateTableOfContents();

            // Sekcja wprowadzenie
            CreateIntroduction();

            // Dodanie tabeli z danymi
            CreateDataTable();

            // Generator dokumentu PDF, zachowuje polskie znaki
            var renderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always) {Document = _document};
            renderer.RenderDocument();
            return renderer.PdfDocument;
        }

        private void CreateDataTable()
        {
            // Utowrzenie sekcji i nagłówka
            var section = _document.AddSection();
            var paragraph = section.AddParagraph("Tabela danych");
            paragraph.Style = "Title";
            // Dodanie bookmarku pozwala klikać w spisie treści (interaktywny PDF)
            paragraph.AddBookmark("Tabela");

            // Dodanie tabeli
            var table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = GetColor(System.Drawing.Color.Black);
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
            row.Shading.Color = GetColor(System.Drawing.Color.LightBlue);

            // Nagłówki
            string[] headers = { "ID", "Imię","Drugie imię", "Nazwisko", "e-mail", "Hasło"};

            // Dodanie nagłówków
            for (int i = 0; i < headers.Length; i++)
            {
                row.Cells[i].AddParagraph(headers[i]);
                row.Cells[i].Format.Font.Bold = true;
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Center;
                row.Cells[i].MergeDown = 1;
            }

            row = table.AddRow();
            row.Shading.Color = GetColor(System.Drawing.Color.LightBlue);

            table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, GetColor(System.Drawing.Color.Black));

            // Dodanie wierszy z danymi- jeden wiersz w tabeli w bazie danych to jeden wiersz w tabeli w raporcie
            using (var context = new wzlEntities())
            {
                var k = 0;
                foreach(var item in context.Customer)
                {
                    row = table.AddRow();                    
                    row.Cells[0].AddParagraph(item.CustomerID.ToString());
                    row.Cells[1].AddParagraph(item.FirstName);
                    row.Cells[2].AddParagraph(item.MiddleName ?? "");
                    row.Cells[3].AddParagraph(item.LastName);
                    row.Cells[4].AddParagraph(item.EmailAddress ?? "");
                    row.Cells[5].AddParagraph(AdjustIfTooWideToFitIn(row.Cells[5], item.PasswordHash));
                    // Wiersze parzyste i nieparzyste podkreślane różnymi kolorami
                    row.Shading.Color = GetColor(k % 2 == 0 ? System.Drawing.Color.LightGray : System.Drawing.Color.NavajoWhite);
                    k++;
                }
            }
        }

        private void CreateIntroduction()
        {
            var section = _document.AddSection();
            
            var paragraph = section.AddParagraph("Wprowadzenie");
            paragraph.AddBookmark("Wprowadzenie");
            paragraph.Style = "SectionTitle";

            paragraph = section.AddParagraph();            
            paragraph.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin nibh augue, suscipit a, scelerisque sed, lacinia in, mi. Cras vel lorem. Etiam pellentesque aliquet tellus. Phasellus pharetra nulla ac diam. Quisque semper justo at risus. Donec venenatis, turpis vel hendrerit interdum, dui ligula ultricies purus, sed posuere libero dui id orci. Nam congue, pede vitae dapibus aliquet, elit magna vulputate arcu, vel tempus metus leo non est. Etiam sit amet lectus quis est congue mollis. Phasellus congue lacus eget neque. Phasellus ornare, ante vitae consectetuer consequat, purus sapien ultricies dolor, et mollis pede metus eget nisi. Praesent sodales velit quis augue. Cras suscipit, urna at aliquam rhoncus, urna quam viverra nisi, in interdum massa nibh nec erat.");

            section.AddParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin nibh augue, suscipit a, scelerisque sed, lacinia in, mi. Cras vel lorem. Etiam pellentesque aliquet tellus. Phasellus pharetra nulla ac diam. Quisque semper justo at risus. Donec venenatis, turpis vel hendrerit interdum, dui ligula ultricies purus, sed posuere libero dui id orci. Nam congue, pede vitae dapibus aliquet, elit magna vulputate arcu, vel tempus metus leo non est. Etiam sit amet lectus quis est congue mollis. Phasellus congue lacus eget neque. Phasellus ornare, ante vitae consectetuer consequat, purus sapien ultricies dolor, et mollis pede metus eget nisi. Praesent sodales velit quis augue. Cras suscipit, urna at aliquam rhoncus, urna quam viverra nisi, in interdum massa nibh nec erat.");

            paragraph = section.AddParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin nibh augue, suscipit a, scelerisque sed, lacinia in, mi. Cras vel lorem. Etiam pellentesque aliquet tellus. Phasellus pharetra nulla ac diam. Quisque semper justo at risus. Donec venenatis, turpis vel hendrerit interdum, dui ligula ultricies purus, sed posuere libero dui id orci. Nam congue, pede vitae dapibus aliquet, elit magna vulputate arcu, vel tempus metus leo non est. Etiam sit amet lectus quis est congue mollis. Phasellus congue lacus eget neque. Phasellus ornare, ante vitae consectetuer consequat, purus sapien ultricies dolor, et mollis pede metus eget nisi. Praesent sodales velit quis augue. Cras suscipit, urna at aliquam rhoncus, urna quam viverra nisi, in interdum massa nibh nec erat.");
            paragraph.Format.Font.Color = GetColor(System.Drawing.Color.DarkGreen);          

        }

        private void CreateTableOfContents()
        {
            if (_document == null) return;
            var section = _document.AddSection();

            var paragraph = section.AddParagraph("Spis treści");
            paragraph.Style = "SectionTitle";


            paragraph = section.AddParagraph();
            // Styl- TOC (tabulacja odsyła w prawo, uzupełnianie kropkami)
            paragraph.Style = "TOC";
            // Dodanie odnośnika do znacznika (bookmark) "Wprowadzenie"
            var hyperlink = paragraph.AddHyperlink("Wprowadzenie");
            // Tekst do wyświetlenia
            hyperlink.AddText("Wprowadzenie do raportu\t");
            // Numer strony
            hyperlink.AddPageRefField("Wprowadzenie");
            // Przejście do nowej linii
            paragraph.AddLineBreak();

            hyperlink = paragraph.AddHyperlink("Tabela");
            paragraph.Style = "TOC";
            hyperlink.AddText("Podsumowanie tabelaryczne\t");
            hyperlink.AddPageRefField("Tabela");
            paragraph.AddLineBreak();

            hyperlink = paragraph.AddHyperlink("Wykres");
            paragraph.Style = "TOC";
            hyperlink.AddText("Grupowanie na wykresie\t");
            hyperlink.AddPageRefField("Wykres");
            paragraph.AddLineBreak();

            hyperlink = paragraph.AddHyperlink("Changes");
            paragraph.Style = "TOC";
            hyperlink.AddText("Wykaz zmian\t");
            hyperlink.AddPageRefField("Changes");
        }

        private void CreateTitlePage()
        {
            if (_document == null) return;

            var section = _document.AddSection();
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
            titleBox.Style = "SectionTitle";
            titleBox.Format.Font.Color = GetColor(System.Drawing.Color.DarkGray);

            var authorBox = titleFrame.AddParagraph(Author);
            authorBox.Style = "Normal";
            
        }

        /// <summary>
        /// Metoda ustawia style dokumentu
        /// </summary>
        /// <param name="document">Dokument do skonfigurowania</param>
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
            style.ParagraphFormat.Font.Color = GetColor(System.Drawing.Color.Black);

            // Define new style "Title", which we will use for cover Title
            style = document.Styles.AddStyle("Title", "Normal");
            style.ParagraphFormat.Font.Size = 20;
            style.ParagraphFormat.Font.Name = "Calibri";
            style.ParagraphFormat.Font.Color = GetColor(System.Drawing.Color.DarkGray);
            style.ParagraphFormat.SpaceBefore = "10mm";
            style.ParagraphFormat.SpaceAfter = "10mm";

            // Define new style "Title", which we will use for cover Title
            style = document.Styles.AddStyle("SectionTitle", "Normal");
            style.ParagraphFormat.Font.Size = 18;
            style.ParagraphFormat.Font.Name = "Calibri";
            style.ParagraphFormat.Font.Color = GetColor(System.Drawing.Color.DarkGray);
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";

            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);

        }

        Color GetColor(System.Drawing.Color color)
        {
            return new Color(
                color.R,
                color.G,
                color.B);
        }

        /// <summary>
        /// Metoda umożliwia złamanie zbyt długich tekstów tak, aby zmieściły się w tabeli
        /// </summary>
        /// <param name="cell">Komórka tabeli</param>
        /// <param name="text">tekst do dostosowania</param>
        /// <returns></returns>
        private string AdjustIfTooWideToFitIn(Cell cell, string text)
        {
            Column column = cell.Column;
            Unit availableWidth = column.Width - column.Table.Borders.Width - cell.Borders.Width;

            var tooWideWords = text.Split(" ".ToCharArray()).Distinct().Where(s => TooWide(s, availableWidth));

            var adjusted = new StringBuilder(text);
            foreach (string word in tooWideWords)
            {
                var replacementWord = MakeFit(word, availableWidth);
                adjusted.Replace(word, replacementWord);
            }

            return adjusted.ToString();
        }

        private bool TooWide(string word, Unit width)
        {
            var tm = new TextMeasurement(_document.Styles["Table"].Font.Clone());
            float f = tm.MeasureString(word, UnitType.Point).Width;
            return f > width.Point;
        }

        /// <summary>
        /// Makes the supplied word fit into the available width
        /// </summary>
        /// <returns>modified version of the word with inserted Returns at appropriate points</returns>
        private string MakeFit(string word, Unit width)
        {
            var adjustedWord = new StringBuilder();
            var current = string.Empty;
            foreach (char c in word)
            {
                if (TooWide(current + c, width))
                {
                    adjustedWord.Append(current);
                    adjustedWord.Append(Chars.CR);
                    current = c.ToString();
                }
                else
                {
                    current += c;
                }
            }
            adjustedWord.Append(current);

            return adjustedWord.ToString();
        }

    }
}
