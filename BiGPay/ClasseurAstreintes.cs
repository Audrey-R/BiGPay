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
            ExcelApp.Application.DisplayAlerts = false;
            ExcelApp.Visible = false;
            Classeur = ExcelApp.Workbooks.Open(libelleClasseur);
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            FeuilleActive.Columns["A:ZZ"].AutoFit();
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

        public string ObtenirCollaborateursQuiOntEteDAstreinteSurLaPeriode()
        {
            int compteur = 0;
            string chaine = "";
            for(int index = 1; index <= DerniereLigne; index++)
            {
                compteur++;
                string collaborateur = FeuilleActive.Cells[index, _ColonneCollaborateurs].Text;
                if (chaine == "" && collaborateur != "" && collaborateur != "Collaborateur")
                {
                    
                    chaine = compteur + " Collaborateur(s) concerné(s) par les tickets d'astreinte, sur cette période. Le dossier CRA contient-il le(s) CRA pdf de : " + collaborateur ;
                }
                else if (chaine != "" && collaborateur != "" && collaborateur != "Collaborateur")
                {
                    chaine = chaine + ", " + collaborateur;
                }
                else
                {
                    chaine = "";
                }
            }
            return chaine +" ?";
        }
    }
}
