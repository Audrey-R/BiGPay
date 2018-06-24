using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    public class ClasseurCollaborateurs : ClasseurExcel
    {
        public ClasseurCollaborateurs(string libelleClasseur)
        {
            ExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            ExcelApp.Application.DisplayAlerts = true;
            ExcelApp.Visible = true;
            Classeur = ExcelApp.Workbooks.Open(libelleClasseur);
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            DerniereLigne = FeuilleActive.Cells[FeuilleActive.Rows.Count, 2].End(XlDirection.xlUp).Row;
            DerniereColonne = FeuilleActive.Cells[2, FeuilleActive.Columns.Count].End(XlDirection.xlToLeft).Column;
            Donnees = FeuilleActive.Range["B2", ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne];
            TrierFeuille(2);
            SupprimerDoublons();
        }

        public string ObtenirEntreesEtSortiesDuMois(int index)
        {
            ActiverClasseur();
            DateTime? DateEntreeSortie = null;
            string dateEntreeSortieSplit = "";
            string entreeSortie = Donnees.Cells[index, 0].Text;
            string dateEntreeSortie = Donnees.Cells[index, 7].Text;
            Char delimiter = '-';

            if (entreeSortie != "")
            {
                if (entreeSortie == "Le collaborateur a démarré ce mois")
                {
                    //Extraction des caractères concernant la date d'entrée et initialisation de la variable à insérer
                    dateEntreeSortieSplit = dateEntreeSortie.Split(delimiter)[0];
                    entreeSortie = "Entrée le " + dateEntreeSortieSplit;
                    DateEntreeSortie = Convert.ToDateTime(dateEntreeSortieSplit);
                }
                else if (entreeSortie == "Le collaborateur quitte ce mois-ci")
                {
                    //Extraction des caractères concernant la date de sortie et initialisation de la variable à insérer
                    if (dateEntreeSortie.Length > 10)
                    {
                        dateEntreeSortieSplit = dateEntreeSortie.Split(delimiter)[1];
                    }
                    else
                    {
                        dateEntreeSortieSplit = dateEntreeSortie;
                    }
                    entreeSortie = "Sortie le " + dateEntreeSortieSplit;
                    DateEntreeSortie = Convert.ToDateTime(dateEntreeSortieSplit);
                }
            }
            return entreeSortie;
        }
    }
}
