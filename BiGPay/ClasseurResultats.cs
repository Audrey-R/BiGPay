using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    class ClasseurResultats : ClasseurExcel
    {
        public const int _ColonneMatricules = 1;
        public const int _ColonneCollaborateurs = 2;
        public const int _ColonneEntreesSorties = 3;
        public const int _ColonneJoursOuvres = 5;
        public const int _ColonneJoursTravailles = 6;
        public const int _ColonneTotalCongesPayes = 7;
        public const int _ColonneCongesPayes = 8;
        public const int _ColonneTotalCongesExceptionnels = 9;
        public const int _ColonneCongesExceptionnels = 10;
        public const int _ColonneTotalRTT = 11;
        public const int _ColonneRTT = 12;
        public const int _ColonneTotalRecup = 13;
        public const int _ColonneRecup = 14;
        public const int _ColonneTotalFormation = 15;
        public const int _ColonneFormation = 16;
        public const int _ColonneTotalMaladie = 17;
        public const int _ColonneMaladie = 18;
        public new const int _PremiereLigne = 7;

        public ClasseurResultats()
        {
            ExcelApp = new Application();
            ExcelApp.Application.DisplayAlerts = true;
            ExcelApp.Visible = true;
            Classeur = ExcelApp.Workbooks.Add();
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            CelluleA1 = FeuilleActive.get_Range("A1", Type.Missing);
        }

        public void RemplirColonneCollaborateurs(ClasseurCollaborateurs classeurCollaborateurs)
        {
            DerniereLigne = 5 + classeurCollaborateurs.DerniereLigne;
            FeuilleActive.Range[
                ConvertirColonneEnLettre(_ColonneCollaborateurs) + _PremiereLigne, 
                ConvertirColonneEnLettre(_ColonneCollaborateurs) + DerniereLigne
            ].Value
            = classeurCollaborateurs.FeuilleActive.Range[
                ConvertirColonneEnLettre(ClasseurCollaborateurs._ColonneCollaborateurs) + ClasseurExcel._PremiereLigne, 
                ConvertirColonneEnLettre(ClasseurCollaborateurs._ColonneCollaborateurs) + classeurCollaborateurs.DerniereLigne
            ].Value;
        }

        public void RemplirColonneMatricules(ClasseurCollaborateurs classeurCollaborateurs)
        {
            FeuilleActive.Range[
                ConvertirColonneEnLettre(_ColonneMatricules) + _PremiereLigne,
                ConvertirColonneEnLettre(_ColonneMatricules) + DerniereLigne
            ].Value
            = classeurCollaborateurs.FeuilleActive.Range[
                ConvertirColonneEnLettre(ClasseurCollaborateurs._ColonneMatricules) + ClasseurExcel._PremiereLigne,
                ConvertirColonneEnLettre(ClasseurCollaborateurs._ColonneMatricules) + classeurCollaborateurs.DerniereLigne
           ].Value;
        }

        public long RechercherCollaborateur<Classeur>(Classeur classeurSource, int index) where Classeur : ClasseurExcel
        {
            //Recherche de la valeur contenue dans la constante ColonneColaborateurs du classeur source
            List<int> constantes = new List<int>();
            foreach (FieldInfo field in typeof(Classeur).GetFields().Where(f => f.Name.StartsWith("_ColonneCollaborateurs")))
            {
                constantes.Add(Convert.ToInt32(field.GetRawConstantValue()));
            }
            int colonneCollaborateursSource = constantes[0];
            classeurSource.Collaborateur = classeurSource.Donnees.Cells[index, colonneCollaborateursSource -1];

            LigneAcompleter = 0;
            
            Collaborateur = FeuilleActive.Cells[_PremiereLigne, _ColonneCollaborateurs]
                            .Find(classeurSource.Collaborateur.Text,
                            Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, Type.Missing,
                            XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing
                            );
            if (Collaborateur.Text == classeurSource.Collaborateur.Text)
            {
                LigneAcompleter = Collaborateur.Row;
            }
            return LigneAcompleter;
        }

        public void RemplirColonneEntreesSorties(long ligneACompleter, int index, ClasseurCollaborateurs classeurCollaborateurs)
        {
            Range celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneEntreesSorties];
            if(classeurCollaborateurs.ObtenirEntreesEtSortiesDuMois(index) != null)
            {
                celluleACompleter.Value = classeurCollaborateurs.ObtenirEntreesEtSortiesDuMois(index).ToString();
            }
        }

        public void CreerPeriodePaie(ClasseurAbsences classeurAbsences)
        {
            if (DerniereLigne > 1)
            {
               //Réécriture de la première date d'absence
               string premiereAbsence = classeurAbsences.FeuilleActive.Cells[ClasseurExcel._PremiereLigne, ClasseurAbsences._ColonneDepartAbsence].Text;
               premiereAbsence = DateTime.ParseExact(
                    premiereAbsence,
                    Periode._Formats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None)
                    .ToString("dd/MM/yyyy");
                //Conversion de la date réécrite et initialisation dans sa variable
                classeurAbsences.PremiereDateAbsence = Convert.ToDateTime(premiereAbsence);
                Periode periode = new Periode(classeurAbsences.PremiereDateAbsence);
                //Remplissage de la cellule concernant l'information du mois en cours dans le classeur paie
                FeuilleActive.Range["B6"].Value = periode.DateDebutPeriode.ToString("MMMM yyyy");
            }
        }

        public void RemplirColonne(string typeAbsenceRecherche, int colonne, long ligneACompleter, int index, ClasseurAbsences classeurAbsences)
        {
            Range celluleACompleter = FeuilleActive.Cells[ligneACompleter, colonne];
            if (classeurAbsences.ObtenirAbsence(typeAbsenceRecherche, index) != "")
            { 
                if (celluleACompleter.Text == "")
                {
                    celluleACompleter.Value = classeurAbsences.ObtenirAbsence(typeAbsenceRecherche, index);
                }
                else
                {
                    celluleACompleter.Value = celluleACompleter.Text + " et " + classeurAbsences.ObtenirAbsence(typeAbsenceRecherche, index);
                }
            }
        }

        public void RemplirTotalColonne(string typeAbsenceRecherche, int colonne, long ligneACompleter, int index, ClasseurAbsences classeurAbsences)
        {
            Range celluleACompleter = FeuilleActive.Cells[ligneACompleter, colonne];
            Decimal nbAbsence;
            Decimal valeurCellule ;
            if (classeurAbsences.ObtenirNombreJourAbsence(typeAbsenceRecherche, index) != 0)
            {
                nbAbsence = classeurAbsences.ObtenirNombreJourAbsence(typeAbsenceRecherche, index);
                if (celluleACompleter.Text == "")
                {
                    celluleACompleter.Value = nbAbsence;
                }
                else
                {
                    valeurCellule = (Decimal)celluleACompleter.Value;
                    celluleACompleter.Value = Decimal.Add(valeurCellule, nbAbsence);
                }
            }
        }
    }
}
