using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    public class ClasseurAstreintes : ClasseurExcel
    {
        public string Code { get; set; }
        public const int _ColonneCollaborateurs = 1;
        public const int _ColonneCode= 9;

        public ClasseurAstreintes() { }
        public ClasseurAstreintes(string libelleClasseur)
        {
            ExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            ExcelApp.Application.DisplayAlerts = true;
            ExcelApp.Visible = true;
            Classeur = ExcelApp.Workbooks.Open(libelleClasseur);
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            DerniereLigne = FeuilleActive.Cells[FeuilleActive.Rows.Count, _ColonneCollaborateurs].End(XlDirection.xlUp).Row;
            DerniereColonne = FeuilleActive.Cells[_PremiereColonne, FeuilleActive.Columns.Count].End(XlDirection.xlToLeft).Column;
            Donnees = FeuilleActive.Range[ConvertirColonneEnLettre(_ColonneCollaborateurs) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne];
            CelluleA1 = FeuilleActive.get_Range("A1", Type.Missing);
            TrierFeuille(FeuilleActive.Range[ConvertirColonneEnLettre(_PremiereColonne) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne], 1);
        }

        public string ObtenirCodeAstreinte(int index)
        {
            Code = FeuilleActive.Cells[index, _ColonneCode].Text;
            return Code;
        }
    }
}
