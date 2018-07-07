using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    class ClasseurAbsences : ClasseurExcel
    {
        public DateTime PremiereDateAbsence { get; set; }
        public DateTime DateDepartAbsence {get; set;}
        public DateTime DateRetourAbsence { get; set; }
        public List<DateTime> JoursOuvresAbsence { get; set; } // NetworkDays(DateDepartAbsence, DateRetourAbsence, JoursFeries)
        public const int _ColonneCollaborateurs = 5;
        public const int _ColonneTypeAbsence = 7;
        public const int _ColonneDepartAbsence = 8;
        public const int _ColonneRetourAbsence = 10;
        public const int _ColonneNombreJoursAbsence = 12;
        public const int _ColonneDetailsAbsence = 13;

        public ClasseurAbsences() { }
        public ClasseurAbsences(string libelleClasseur)
        {
            ExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            ExcelApp.Application.DisplayAlerts = true;
            ExcelApp.Visible = true;
            Classeur = ExcelApp.Workbooks.Open(libelleClasseur);
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            DerniereLigne = FeuilleActive.Cells[FeuilleActive.Rows.Count, _PremiereColonne].End(XlDirection.xlUp).Row;
            DerniereColonne = FeuilleActive.Cells[_PremiereLigne-1, FeuilleActive.Columns.Count].End(XlDirection.xlToLeft).Column;
            Donnees = FeuilleActive.Range[ConvertirColonneEnLettre(_PremiereColonne) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne];
            CelluleA1 = FeuilleActive.get_Range("A1", Type.Missing);
            TrierFeuille(4);
            SupprimerDoublons();
        }

        public string ObtenirAbsence(string typeAbsenceRecherche, int index)
        {
            #region Variables
            string nbAbsence = FeuilleActive.Cells[index, _ColonneNombreJoursAbsence].Text;
            string typeAbsence = FeuilleActive.Cells[index, _ColonneTypeAbsence].Text;
            string departAbsence = FeuilleActive.Cells[index, _ColonneDepartAbsence].Text;
            string retourAbsence = FeuilleActive.Cells[index, _ColonneRetourAbsence].Text;
            
            Periode periode = new Periode(PremiereDateAbsence);
            
            string texteARetourner;
            #endregion

            #region Dates
            // Réécriture des dates dans le format souhaité
            if (departAbsence != "Premier jour" && departAbsence != "")
            {
                departAbsence = DateTime.ParseExact(
                    departAbsence,
                    Periode._Formats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None)
                    .ToString("dd/MM/yyyy");
                DateDepartAbsence = Convert.ToDateTime(departAbsence);
                // Initialisation de la date au premier jour du mois, si inférieure
                if (DateDepartAbsence.Month != periode.DateDebutPeriode.Month)
                    DateDepartAbsence = periode.DateDebutPeriode;
                departAbsence = DateDepartAbsence.ToShortDateString();
            }
            if (retourAbsence != "Dernier jour" && retourAbsence != "")
            {
                retourAbsence = DateTime.ParseExact(
                    retourAbsence,
                    Periode._Formats,
                    CultureInfo.InvariantCulture, 
                    DateTimeStyles.None)
                    .ToString("dd/MM/yyyy");
                DateRetourAbsence = Convert.ToDateTime(retourAbsence);
                //Initialisation de la date au dernier jour du mois, si supérieure
                if (DateRetourAbsence.Month != periode.DateFinPeriode.Month)
                    DateRetourAbsence = periode.DateFinPeriode;
                retourAbsence = DateRetourAbsence.ToShortDateString();
            }
            #endregion

            #region Traitement
            // Traitement
            Decimal nbAbsenceDecimal = ObtenirNombreJourAbsence(typeAbsenceRecherche, index);
            nbAbsence = nbAbsenceDecimal.ToString();

            if (departAbsence != retourAbsence)
            {
                texteARetourner = nbAbsence + " du " + departAbsence + " au " + retourAbsence;
            }
            else
            {
                texteARetourner = nbAbsence + " le " + departAbsence;
            }

            if (typeAbsenceRecherche == "Maladie")
            {
                if (typeAbsence == "Maladie non justifiée" ||
                    typeAbsence == "Congé maternité" ||
                    typeAbsence == "Enfant malade")
                    return texteARetourner + " (" + typeAbsence + ")";
                else if (typeAbsence == "Maladie")
                    return texteARetourner;
                return "";
            }
            else if (typeAbsenceRecherche == "Congés_Payés" && typeAbsence == "Congé payé")
            {
                //AjouterLeNbAbsenceAuToTalDeLaColonne(nbAbsence, ClasseurResultats._ColonneCongesPayes);
                return texteARetourner;
            }
            else if (typeAbsenceRecherche == "RTT" && typeAbsence == "RTT salarié" || typeAbsence == "RTT employeur")
            {
                return texteARetourner;
            }
            else if (typeAbsenceRecherche == "Récupération" && typeAbsence == "Récupération du temps de travail")
            {
                return texteARetourner;
            }
            else if (typeAbsenceRecherche == "Formation" && typeAbsence == "Formation")
            {
                return texteARetourner;
            }
            else
            {
                return "";
            }
            #endregion
        }

        public decimal ObtenirNombreJourAbsence(string typeAbsenceRecherche, int index)
        {
            #region Variables
            int indexDepart;
            int indexFin;
            string finChaineARechercherDansColonneDetails = "jour(s)";
            int longueurCoupureDeChaine;
            string detailsAbsence = FeuilleActive.Cells[index, _ColonneDetailsAbsence].Text;
            string typeAbsence = FeuilleActive.Cells[index, _ColonneTypeAbsence].Text;
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
                    string partieGaucheChaine = detailsAbsence.Substring(1, detailsAbsence.IndexOf("|") - 2);
                    string partieDroiteChaine = detailsAbsence.Substring(detailsAbsence.IndexOf("|") + 2, detailsAbsence.IndexOf(detailsAbsence.Substring(detailsAbsence.Length - 1)) + 1);
                    if (partieDroiteChaine.Substring(partieDroiteChaine.Length - 1, 1) == "(")
                        partieDroiteChaine = detailsAbsence.Substring(detailsAbsence.IndexOf("|") + 2, detailsAbsence.IndexOf(detailsAbsence.Substring(detailsAbsence.Length - 1)) + 3);

                    if (partieGaucheChaine.Contains(periode.DateDebutPeriode.Month.ToString()))
                    {

                        indexDepart = partieGaucheChaine.IndexOf(":") + 2;
                        indexFin = partieGaucheChaine.IndexOf(finChaineARechercherDansColonneDetails);
                        longueurCoupureDeChaine = indexFin - indexDepart - 1;
                        detailsAbsence = partieGaucheChaine.Substring(indexDepart, longueurCoupureDeChaine);
                    }
                    else if (partieDroiteChaine.Contains(periode.DateDebutPeriode.Month.ToString()))
                    {
                        indexDepart = partieDroiteChaine.IndexOf(":") + 2;
                        indexFin = partieDroiteChaine.IndexOf(finChaineARechercherDansColonneDetails);
                        longueurCoupureDeChaine = indexFin - indexDepart - 1;
                        detailsAbsence = partieDroiteChaine.Substring(indexDepart, longueurCoupureDeChaine);
                    }
                }
                else
                {
                    indexDepart = detailsAbsence.IndexOf((periode.DateDebutPeriode.Month).ToString());
                    indexFin = detailsAbsence.IndexOf(detailsAbsence.Substring(detailsAbsence.Length - 8));
                    longueurCoupureDeChaine = indexFin - 10;
                    if (detailsAbsence.Substring(10, longueurCoupureDeChaine) == "")
                    {
                        indexDepart = 1;
                        //indexFin = indexFin - 1;
                        longueurCoupureDeChaine = indexFin - 9;
                        detailsAbsence = detailsAbsence.Substring(9, longueurCoupureDeChaine);
                    }
                    else
                    {
                        detailsAbsence = detailsAbsence.Substring(10, longueurCoupureDeChaine);
                    }
                }
                Decimal detailsAbsenceNbJours = Convert.ToDecimal(detailsAbsence);
                #endregion

            #region Traitement
                if (typeAbsenceRecherche == "Maladie" &&
                    typeAbsence == "Maladie non justifiée" ||
                    typeAbsence == "Congé maternité" ||
                    typeAbsence == "Enfant malade" ||
                    typeAbsence == "Maladie")
                {
                    return detailsAbsenceNbJours;
                }
                else if (typeAbsenceRecherche == "Congés_Payés" && typeAbsence == "Congé payé")
                {
                    //AjouterLeNbAbsenceAuToTalDeLaColonne(nbAbsence, ClasseurResultats._ColonneCongesPayes);
                    return detailsAbsenceNbJours;
                }
                else if (typeAbsenceRecherche == "RTT" && typeAbsence == "RTT salarié" || typeAbsence == "RTT employeur")
                {
                    return detailsAbsenceNbJours;
                }
                else if (typeAbsenceRecherche == "Récupération" && typeAbsence == "Récupération du temps de travail")
                {
                    return detailsAbsenceNbJours;
                }
                else if (typeAbsenceRecherche == "Formation" && typeAbsence == "Formation")
                {
                    return detailsAbsenceNbJours;
                }
                else
                {
                    return 0;
                }
            }
            return 0;
            #endregion
        }
    }
}
