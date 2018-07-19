using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    public class ClasseurAbsences : ClasseurExcel
    {
        public DateTime PremiereDateAbsence { get; set; }
        public DateTime DateDepartAbsence {get; set;}
        public DateTime DateRetourAbsence { get; set; }
        public List<DateTime> JoursOuvresAbsence { get; set; } // NetworkDays(DateDepartAbsence, DateRetourAbsence, JoursFeries)
        public const int _ColonneCollaborateurs = 3;
        public const int _ColonneTypeAbsence = 7;
        public const int _ColonneDepartAbsence = 8;
        public const int _ColonneRetourAbsence = 10;
        public const int _ColonneDetailsAbsence = 13;

        public ClasseurAbsences() { }
        public ClasseurAbsences(string libelleClasseur)
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
            TrierFeuille(FeuilleActive.Range[ConvertirColonneEnLettre(_PremiereColonne) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne],4);
            SupprimerDoublons(FeuilleActive.Range[ConvertirColonneEnLettre(_PremiereColonne) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne]);
        }

        public string ObtenirAbsence(int index)
        {
            ActiverClasseur();
            #region Variables
            string nbAbsence;
            string typeAbsence = FeuilleActive.Cells[index, _ColonneTypeAbsence].Text;
            string departAbsence = FeuilleActive.Cells[index, _ColonneDepartAbsence].Text;
            string retourAbsence = FeuilleActive.Cells[index, _ColonneRetourAbsence].Text;
            
            Periode periode = new Periode(PremiereDateAbsence);
            
            string texteARetourner;
            #endregion

            #region Dates
            // Réécriture des dates dans le format souhaité
            DateTime dateTest;
            if (departAbsence != "Premier jour" && departAbsence != "")
            {
                if (!DateTime.TryParse(Convert.ToDateTime(departAbsence).ToString("dd/MM/yyyy"), out dateTest))
                {
                    departAbsence = DateTime.ParseExact(
                    departAbsence,
                    Periode._Formats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None)
                    .ToString("dd/MM/yyyy");
                }
                DateDepartAbsence = Convert.ToDateTime(departAbsence);
                // Initialisation de la date au premier jour du mois, si inférieure
                if (DateDepartAbsence.Month != periode.DateDebutPeriode.Month)
                    DateDepartAbsence = periode.DateDebutPeriode;
                departAbsence = DateDepartAbsence.ToShortDateString();
            }
            if (retourAbsence != "Dernier jour" && retourAbsence != "")
            {
                if (!DateTime.TryParse(Convert.ToDateTime(retourAbsence).ToString("dd/MM/yyyy"), out dateTest))
                {
                    retourAbsence = DateTime.ParseExact(
                    retourAbsence,
                    Periode._Formats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None)
                    .ToString("dd/MM/yyyy");
                }
                DateRetourAbsence = Convert.ToDateTime(retourAbsence);
                //Initialisation de la date au dernier jour du mois, si supérieure
                if (DateRetourAbsence.Month != periode.DateFinPeriode.Month)
                    DateRetourAbsence = periode.DateFinPeriode;
                retourAbsence = DateRetourAbsence.ToShortDateString();
            }
            #endregion

            #region Traitement
            // Traitement
            nbAbsence = ObtenirNombreJourAbsence(index);
            if(nbAbsence != "")
            {
                if (departAbsence != retourAbsence)
                {
                    texteARetourner = nbAbsence + " du " + departAbsence + " au " + retourAbsence + " (" + typeAbsence + ")";
                }
                else
                {
                    texteARetourner = nbAbsence + " le " + departAbsence + " (" + typeAbsence + ")";
                }
            }
            else
            {
                texteARetourner = "";
            }
            return texteARetourner;
            #endregion
        }

        public string ObtenirNombreJourAbsence(int index)
        {
            ActiverClasseur();
            #region Variables
            string detailsAbsence = FeuilleActive.Cells[index, _ColonneDetailsAbsence].Text;
            string typeAbsence = FeuilleActive.Cells[index, _ColonneTypeAbsence].Text;
            //string nbAbsence;
            Periode periode = new Periode(PremiereDateAbsence);
            #endregion

            #region Gestion des cas de période d'Absence(periode courante ou qui en chevauche une autre)
            /* Réécriture du nombre de jours d'absence selon si l'absence 
             débute avant la période ou s'arrête après la période */
            if (detailsAbsence != "Détail des jours" && detailsAbsence != "")
            {
                //Traitement des absences qui chevauchent deux périodes
                if (detailsAbsence.Contains("|"))
                {
                    string partieGaucheChaine = detailsAbsence.Split('|')[0];
                    string partieDroiteChaine = detailsAbsence.Split('|')[1];

                    if (partieGaucheChaine.Contains(periode.DateDebutPeriode.Month.ToString()))
                    {
                        detailsAbsence = partieGaucheChaine;
                    }
                    else if (partieDroiteChaine.Contains(periode.DateDebutPeriode.Month.ToString()))
                    {
                        detailsAbsence = partieDroiteChaine;
                    }
                }
                detailsAbsence = detailsAbsence.Split(':')[1];
                detailsAbsence = detailsAbsence.Split('j')[0];
                return detailsAbsence.Trim() ;
            }
            else
            {
                return "";
            }
            #endregion
        }
    }
}
