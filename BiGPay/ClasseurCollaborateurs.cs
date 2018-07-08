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
        public const int _ColonneEntreeSortie = 1;
        public const int _ColonneCollaborateurs = 2;
        public const int _ColonneMatricules = 3;
        public const int _ColonneDatesEntreeSortie = 8;

        public ClasseurCollaborateurs() { }
        public ClasseurCollaborateurs(string libelleClasseur)
        {
            ExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            ExcelApp.Application.DisplayAlerts = false;
            ExcelApp.Visible = false;
            Classeur = ExcelApp.Workbooks.Open(libelleClasseur);
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            DerniereLigne = FeuilleActive.Cells[FeuilleActive.Rows.Count, _ColonneCollaborateurs].End(XlDirection.xlUp).Row;
            DerniereColonne = FeuilleActive.Cells[_PremiereColonne, FeuilleActive.Columns.Count].End(XlDirection.xlToLeft).Column;
            Donnees = FeuilleActive.Range[ConvertirColonneEnLettre(_ColonneCollaborateurs) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne];
            TrierFeuille(FeuilleActive.Range[ConvertirColonneEnLettre(_PremiereColonne) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne],2);
            SupprimerDoublons(Donnees);
        }

        public string ObtenirEntreesEtSortiesDuMois(int index)
        {
            ActiverClasseur();
            DateTime? DateEntreeSortie = null;
            string dateEntreeSortieSplit = "";
            string entreeSortie = Donnees.Cells[index, _ColonneEntreeSortie].Text;
            string datesEntreeSortie = Donnees.Cells[index, _ColonneDatesEntreeSortie].Text;
            Char delimiter = '-';

            if (entreeSortie != "")
            {
                if (entreeSortie == "Le collaborateur a démarré ce mois")
                {
                    //Extraction des caractères concernant la date d'entrée et initialisation de la variable à insérer
                    dateEntreeSortieSplit = datesEntreeSortie.Split(delimiter)[0];
                    entreeSortie = "Entrée le " + dateEntreeSortieSplit;
                    DateEntreeSortie = Convert.ToDateTime(dateEntreeSortieSplit);
                }
                else if (entreeSortie == "Le collaborateur quitte ce mois-ci")
                {
                    //Extraction des caractères concernant la date de sortie et initialisation de la variable à insérer
                    if (datesEntreeSortie.Length > 10)
                    {
                        dateEntreeSortieSplit = datesEntreeSortie.Split(delimiter)[1];
                    }
                    else
                    {
                        dateEntreeSortieSplit = datesEntreeSortie;
                    }
                    entreeSortie = "Sortie le " + dateEntreeSortieSplit;
                    DateEntreeSortie = Convert.ToDateTime(dateEntreeSortieSplit);
                }
            }
            return entreeSortie;
        }
    }
}
