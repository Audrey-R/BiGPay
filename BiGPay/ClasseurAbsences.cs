using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    class ClasseurAbsences : ClasseurExcel
    {
        public DateTime DateDepartAbsence {get; set;}
        public DateTime DateRetourAbsence { get; set; }
        public List<DateTime> JoursOuvresAbsence { get; set; } // NetworkDays(DateDepartAbsence, DateRetourAbsence, JoursFeries)
        public int ColonneDepartAbsence = 8;
        public int ColonneRetourAbsence = 10;
        public int ColonneNombreJoursAbsence = 12;


        public ClasseurAbsences(string libelleClasseur)
        {
            ExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            ExcelApp.Application.DisplayAlerts = true;
            ExcelApp.Visible = true;
            Classeur = ExcelApp.Workbooks.Open(libelleClasseur);
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            DerniereLigne = FeuilleActive.Cells[FeuilleActive.Rows.Count, 2].End(XlDirection.xlUp).Row;
            DerniereColonne = FeuilleActive.Cells[2, FeuilleActive.Columns.Count].End(XlDirection.xlToLeft).Column;
            Donnees = FeuilleActive.Range["A2", ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne];
            CelluleA1 = FeuilleActive.get_Range("A1", Type.Missing);
            TrierFeuille(4);
            SupprimerDoublons();
        }
    }
}
