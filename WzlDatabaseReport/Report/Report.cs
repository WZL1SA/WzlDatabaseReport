using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WzlDatabaseReport.Report
{
    abstract class Report <T> //zapis ten oznacza  że ma metodę która zwraca obiekt zadelkarowany
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
        public DateTime Time { get; set; }
        /// <summary>
        /// Ścieżka do pliku z logo
        /// </summary>
        public string CoverImagePath { get; set; }

        public string SearchName { get; set; }

        public abstract T CreateReport();

    }
}
