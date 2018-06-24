using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    public class ClasseurExcel
    {
        public Application ExcelApp { get; set; }
        public Workbook Classeur { get; set; }
        public string Libelle { get; set; }
        public _Worksheet FeuilleActive { get; set; }
        public long DerniereLigne { get; set; }
        public long DerniereColonne { get; set; }
        public Range Donnees { get; set; }
        public Range Collaborateur { get; set; }
        public Range CelluleA1 { get; set; }
        public long LigneAcompleter { get; set; }
        public long LigneACopier { get; set; }

        public void OuvrirClasseur(string excelFile)
        {
            System.Diagnostics.Process.Start(excelFile);
        }

        public void ActiverClasseur()
        {
            Classeur.Activate();
        }

        protected string ConvertirColonneEnLettre(long colonne)
        {
            long dividend = colonne;
            string lettre = String.Empty;
            long modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                lettre = Convert.ToChar(65 + modulo).ToString() + lettre;
                dividend = (long)((dividend - modulo) / 26);
            }

            return lettre;
        }

        public void TrierFeuille(int colonne)
        {
            Donnees.Sort(
                FeuilleActive.Columns[colonne, Missing.Value], XlSortOrder.xlAscending);
        }

        public void SupprimerDoublons()
        {
            Donnees.RemoveDuplicates(XlYesNoGuess.xlYes);
        }

        public void FermerTousLesProcessus()
        {
            foreach (System.Diagnostics.Process process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
        }

        void Application_WorkbookAfterSave(Workbook Classeur, bool Success)
        {
            Classeur.Close();
            ExcelApp.Quit();
        }
    }
}
