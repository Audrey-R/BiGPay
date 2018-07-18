using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    public class ClasseurResultats : ClasseurExcel
    {
        public new const int _PremiereLigne = 7;
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
        public const int _ColonneCodeAstreinte = 20;
        public const int _ColonneTotalHeureSup_Sem_8_20 = 22;
        public const int _ColonneHeureSup_Sem_8_20 = 23;
        public const int _ColonneTotalHeureSup_Sem_20_8 = 24;
        public const int _ColonneHeureSup_Sem_20_8 = 25;
        public const int _ColonneTotalHeureSup_Sam_8_20 = 26;
        public const int _ColonneHeureSup_Sam_8_20 = 27;
        public const int _ColonneTotalHeureSup_Sam_20_8 = 28;
        public const int _ColonneHeureSup_Sam_20_8 = 29;
        public const int _ColonneTotalHeureSup_DF_8_20 = 30;
        public const int _ColonneHeureSup_DF_8_20 = 31;
        public const int _ColonneTotalHeureSup_DF_20_8 = 32;
        public const int _ColonneHeureSup_DF_20_8 = 33;


        public ClasseurResultats()
        {
            ExcelApp = new Application();
            ExcelApp.Application.DisplayAlerts = false;
            ExcelApp.Visible = false;
            Classeur = ExcelApp.Workbooks.Add();
            Libelle = Classeur.Name;
            FeuilleActive = Classeur.Sheets[1];
            CelluleA1 = FeuilleActive.get_Range("A1", Type.Missing);
        }

        public void RemplirColonneCollaborateurs(ClasseurCollaborateurs classeurCollaborateurs)
        {
            DerniereLigne = 6 + classeurCollaborateurs.DerniereLigne;
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
            if (colonneCollaborateursSource - 1 > 0)
                colonneCollaborateursSource = colonneCollaborateursSource - 1;
            classeurSource.Collaborateur = classeurSource.Donnees.Cells[index, colonneCollaborateursSource];
            LigneAcompleter = 0;
            if (classeurSource.Collaborateur.Text != "Collaborateur")
            {
                Collaborateur = FeuilleActive.Cells[_PremiereLigne, _ColonneCollaborateurs]
                                .Find(classeurSource.Collaborateur.Text,
                                Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, Type.Missing,
                                XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing
                                );
                if (Collaborateur.Text == classeurSource.Collaborateur.Text)
                {
                    LigneAcompleter = Collaborateur.Row;
                }
            }
            return LigneAcompleter;
        }

        public void RemplirColonneEntreesSorties(long ligneACompleter, int index, ClasseurCollaborateurs classeurCollaborateurs)
        {
            Range celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneEntreesSorties];
            if(classeurCollaborateurs.ObtenirEntreesEtSortiesDuMois(index) != "")
            {
                celluleACompleter.Value = classeurCollaborateurs.ObtenirEntreesEtSortiesDuMois(index).ToString();
            }
        }
        
        public void RemplirAbsences(long ligneACompleter, int index, ClasseurAbsences classeurAbsences, Periode periode)
        {
            #region Variables
            string detailsAbsenceAvecTypeAbsence;
            string detailsAbsenceSansTypeAbsence;
            string typeAbsence;
            string valeurAEcrire;
            Range celluleACompleter;
            Range celluleTotalACompleter;
            Decimal nbAbsence;
            Decimal valeurCelluleTotal;
            #endregion

            #region Traitement
            if (classeurAbsences.ObtenirAbsence(index) != "")
            {
                // Extraction des chaines de caratères
                detailsAbsenceAvecTypeAbsence = classeurAbsences.ObtenirAbsence(index).Trim();
                typeAbsence = detailsAbsenceAvecTypeAbsence.Split('(')[1];
                typeAbsence = typeAbsence.Split(')')[0];
                detailsAbsenceSansTypeAbsence = detailsAbsenceAvecTypeAbsence.Split('(')[0].Trim();

                if (typeAbsence != "Type")
                {
                    /* Définition de la cellule à compléter selon le type d'absence retourné
                   et de la valeur à écrire dans cette dernière */
                    if (typeAbsence == "Congé payé")
                    {
                        celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneCongesPayes];
                        valeurAEcrire = detailsAbsenceSansTypeAbsence;
                        celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalCongesPayes];
                    }
                    else if (typeAbsence == "Maladie non justifiée" ||
                        typeAbsence == "Congé maternité" ||
                        typeAbsence == "Enfant malade" ||
                        typeAbsence == "Maladie")
                    {
                        celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneMaladie];
                        valeurAEcrire = detailsAbsenceAvecTypeAbsence;
                        celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalMaladie];
                    }
                    else if (typeAbsence == "RTT salarié" || typeAbsence == "RTT employeur")
                    {
                        celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneRTT];
                        valeurAEcrire = detailsAbsenceSansTypeAbsence;
                        celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalRTT];
                    }
                    else if (typeAbsence == "Récupération du temps de travail")
                    {
                        celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneRecup];
                        valeurAEcrire = detailsAbsenceSansTypeAbsence;
                        celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalRecup];
                    }
                    else if (typeAbsence == "Formation")
                    {
                        celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneFormation];
                        valeurAEcrire = detailsAbsenceSansTypeAbsence;
                        celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalFormation];
                    }
                    else
                    {
                        celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneCongesExceptionnels];
                        valeurAEcrire = detailsAbsenceAvecTypeAbsence;
                        celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalCongesExceptionnels];
                    }
                    // Remplissage de la cellule selon si elle est vide ou non
                    if (celluleACompleter.Text == "")
                    {
                        celluleACompleter.Value = valeurAEcrire;
                    }
                    else
                    {
                        celluleACompleter.Value = celluleACompleter.Text + " et " + valeurAEcrire;
                    }

                    // Remplissage de la cellule du total
                    if (classeurAbsences.ObtenirNombreJourAbsence(index) != "")
                    {
                        valeurAEcrire = classeurAbsences.ObtenirNombreJourAbsence(index);
                        nbAbsence = Convert.ToDecimal(valeurAEcrire);
                        if (celluleTotalACompleter.Text == "")
                        {
                            celluleTotalACompleter.Value = nbAbsence;
                        }
                        else
                        {
                            valeurCelluleTotal = (Decimal)celluleTotalACompleter.Value;
                            celluleTotalACompleter.Value = Decimal.Add(valeurCelluleTotal, nbAbsence);
                        }
                    }
                }
            }
            RemplirJoursTravaillesPeriode(ligneACompleter, index, periode);
            #endregion
        }

        public void RemplirJoursTravaillesPeriode(long ligneACompleter, int index, Periode periode)
        {
            Range celluleJoursOuvres = FeuilleActive.Cells[ligneACompleter, _ColonneJoursOuvres];
            celluleJoursOuvres.Value = periode.NbJoursOuvresPeriode;

            Decimal? totalCongesPayes = (Decimal?)FeuilleActive.Cells[ligneACompleter, _ColonneTotalCongesPayes].Value;
            Decimal? totalCongesExceptionnels = (Decimal?)FeuilleActive.Cells[ligneACompleter, _ColonneTotalCongesExceptionnels].Value;
            Decimal? totalRTT = (Decimal?)FeuilleActive.Cells[ligneACompleter, _ColonneTotalRTT].Value;
            Decimal? totalMaladie = (Decimal?)FeuilleActive.Cells[ligneACompleter, _ColonneTotalMaladie].Value;
            Decimal? totalFormation = (Decimal?)FeuilleActive.Cells[ligneACompleter, _ColonneTotalFormation].Value;

            totalCongesPayes = ReecrireSiNull(totalCongesPayes);
            totalCongesExceptionnels = ReecrireSiNull(totalCongesExceptionnels);
            totalRTT = ReecrireSiNull(totalRTT);
            totalMaladie = ReecrireSiNull(totalMaladie);
            totalFormation = ReecrireSiNull(totalFormation);

            Range celluleJoursTravailles = FeuilleActive.Cells[ligneACompleter, _ColonneJoursTravailles];
            celluleJoursTravailles.Value = periode.NbJoursOuvresPeriode
                                           - totalCongesPayes
                                           - totalCongesExceptionnels
                                           - totalRTT
                                           - totalMaladie
                                           - totalFormation;
        }

        public void RemplirJoursTravaillesPeriode(Periode periode)
        {
            for(int ligneACompleter = 8; ligneACompleter <= DerniereLigne; ligneACompleter++)
            {
                Range celluleJoursOuvres = FeuilleActive.Cells[ligneACompleter, _ColonneJoursOuvres];
                Range celluleJoursTravailles = FeuilleActive.Cells[ligneACompleter, _ColonneJoursTravailles];
                if (celluleJoursOuvres.Text == "")
                {
                    celluleJoursOuvres.Value = periode.NbJoursOuvresPeriode;
                    celluleJoursTravailles.Value = celluleJoursOuvres.Value;
                }
            }
            //Supression du remplissage sur la ligne d'en-tête
            if (FeuilleActive.Cells[7, _ColonneJoursOuvres].Text == periode.NbJoursOuvresPeriode.ToString())
            {
                FeuilleActive.Cells[7, _ColonneJoursOuvres].Value = "";
                FeuilleActive.Cells[7, _ColonneJoursTravailles].Value = "";
            }
        }

        public void RemplirHeuresSupplementaires(long ligneACompleter, int index, ClasseurHeuresSup classeurHeuresSup, Periode periode)
        {
            #region Variables
            string heureSupAvecDestination;
            string destination ="";
            string heureSupSansDestination ="";
            string valeurAEcrire = "";
            Range celluleACompleter;
            Range celluleTotalACompleter;
            Decimal nbHeures;
            Decimal valeurCelluleTotal;
            #endregion

            #region Traitement
            // Extraction des chaines de caratères
            heureSupAvecDestination = classeurHeuresSup.ObtenirHeuresSupplementaires(index, periode).Trim();
            if(heureSupAvecDestination != "Erreur")
            {
                destination = heureSupAvecDestination.Split('|')[0];
                heureSupSansDestination = heureSupAvecDestination.Split('|')[1].Trim();
                valeurAEcrire = heureSupSansDestination;
            }
            // Définition de la cellule à compléter selon le type d'heures supplémentaires retourné
            celluleACompleter = DefinitionDeLaCelluleACompleter(destination, ligneACompleter);
            // Définition de la cellule Total à compléter selon le type d'heures supplémentaires retourné
            celluleTotalACompleter = DefinitionDeLaCelluleTotalACompleter(destination, ligneACompleter);
            
            // Remplissage de la cellule selon si elle est vide ou non
            if (celluleACompleter != null && celluleACompleter.Text == "")
            {
                celluleACompleter.Value = valeurAEcrire;
            }
            else
            {
                celluleACompleter.Value = celluleACompleter.Text + " et " + valeurAEcrire;
            }

            // Remplissage de la cellule du total
            if (classeurHeuresSup.ObtenirNombreHeuresSupplementaires(index) != 0)
            {
                nbHeures = classeurHeuresSup.ObtenirNombreHeuresSupplementaires(index);
                if (celluleTotalACompleter.Text == "")
                {
                    celluleTotalACompleter.Value = nbHeures;
                }
                else
                {
                    valeurCelluleTotal = (Decimal)celluleTotalACompleter.Value;
                    celluleTotalACompleter.Value = Decimal.Add(valeurCelluleTotal, nbHeures);
                }
            }
            #endregion
        }

        private Range DefinitionDeLaCelluleACompleter(string destination, long ligneACompleter)
        {
            Range celluleACompleter;
            // Définition de la cellule à compléter selon le type d'heures supplémentaires retourné
            if (destination == "Sem-8-20")
            {
                celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneHeureSup_Sem_8_20];
             }
            else if (destination == "Sem-20-8")
            {
                celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneHeureSup_Sem_20_8];
            }
            else if (destination == "Sam-8-20")
            {
                celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneHeureSup_Sam_8_20];
            }
            else if (destination == "Sam-20-8")
            {
                celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneHeureSup_Sam_20_8];
            }
            else if (destination == "DF-8-20")
            {
                celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneHeureSup_DF_8_20];
           }
            else if (destination == "DF-20-8")
            {
                celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneHeureSup_DF_20_8];
            }
            else
            {
                celluleACompleter = null;
            }
            return celluleACompleter;
        }

        private Range DefinitionDeLaCelluleTotalACompleter(string destination, long ligneACompleter)
        {
            Range celluleTotalACompleter;
            if (destination == "Sem-8-20")
            {
                celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalHeureSup_Sem_8_20];
            }
            else if (destination == "Sem-20-8")
            {
               celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalHeureSup_Sem_20_8];
            }
            else if (destination == "Sam-8-20")
            {
               celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalHeureSup_Sam_8_20];
            }
            else if (destination == "Sam-20-8")
            {
                celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalHeureSup_Sam_20_8];
            }
            else if (destination == "DF-8-20")
            {
                celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalHeureSup_DF_8_20];
            }
            else if (destination == "DF-20-8")
            {
                celluleTotalACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneTotalHeureSup_DF_20_8];
            }
            else
            {
                celluleTotalACompleter = null;
            }
            return celluleTotalACompleter;
        }

        public void RemplirCodesAstreintes(long ligneACompleter, int index, ClasseurAstreintes classeurAstreintes)
        {
            Range celluleACompleter = FeuilleActive.Cells[ligneACompleter, _ColonneCodeAstreinte];
            celluleACompleter.Value = classeurAstreintes.ObtenirCodeAstreinte(index);
        }

        public void RemplirTicketsAstreintes(Ticket ticket, long ligneACompleter, ClasseurPdf classeurPdf, Periode periode)
        {
            #region Variables
            string heureSupAvecDestination;
            string destination = "";
            string heureSupSansDestination = "";
            string valeurAEcrire = "";
            Range celluleACompleter;
            Range celluleTotalACompleter;
            Decimal nbHeures;
            Decimal valeurCelluleTotal;
            #endregion

            #region Traitement
            // Extraction des chaines de caratères
            heureSupAvecDestination = classeurPdf.ObtenirHeuresSupplementairesTicket(ticket, periode).Trim();
            if (heureSupAvecDestination != "Erreur")
            {
                destination = heureSupAvecDestination.Split('|')[0];
                heureSupSansDestination = heureSupAvecDestination.Split('|')[1].Trim();
                valeurAEcrire = heureSupSansDestination;
            }
            // Définition de la cellule à compléter selon le type d'heures supplémentaires retourné
            celluleACompleter = DefinitionDeLaCelluleACompleter(destination, ligneACompleter);
            // Définition de la cellule Total à compléter selon le type d'heures supplémentaires retourné
            celluleTotalACompleter = DefinitionDeLaCelluleTotalACompleter(destination, ligneACompleter);

            // Remplissage de la cellule selon si elle est vide ou non
            if (celluleACompleter != null && celluleACompleter.Text == "")
            {
                celluleACompleter.Value = valeurAEcrire;
            }
            else
            {
                celluleACompleter.Value = celluleACompleter.Text + " et " + valeurAEcrire;
            }

            // Remplissage de la cellule du total
            if (ticket.NbHeures != 0)
            {
                string nbheures = ticket.NbHeures.ToString("0.0");
                if (nbheures.Split(',')[1] == "0")
                    nbheures = ticket.NbHeures.ToString("0");

                nbHeures = Convert.ToDecimal(nbheures);
                if (celluleTotalACompleter.Text == "")
                {
                    celluleTotalACompleter.Value = nbHeures;
                }
                else
                {
                    valeurCelluleTotal = (Decimal)celluleTotalACompleter.Value;
                    celluleTotalACompleter.Value = Decimal.Add(valeurCelluleTotal, nbHeures);
                }
            }
            #endregion
        }

        public void FormaterClasseur()
        {
            FeuilleActive.Range["B2:B3"].Interior.Color = Color.FromArgb(217, 217, 217);
            FeuilleActive.Cells[2, 2].Value = "Collaborateur";
            FeuilleActive.Range["C2:C3"].Interior.Color = Color.FromArgb(217, 217, 217);
            FeuilleActive.Cells[2, 3].Value = "Entrées / Sorties";
            FeuilleActive.Cells[2, 5].Value = "Jours d'absence";
            FeuilleActive.Range["E2:R3"].Interior.Color = Color.FromArgb(217, 217, 217);
            FeuilleActive.Range["E2:R3"].WrapText = true;
            FeuilleActive.Cells[3, 5].Value = "NB Jours ouvrés M-1";
            FeuilleActive.Range["E3"].WrapText = true;
            FeuilleActive.Cells[3, 6].Value = "NB Jours travaillés M-1";
            FeuilleActive.Range["F3"].WrapText = true;
            FeuilleActive.Range["E7:F"+ DerniereLigne].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            FeuilleActive.Cells[3, 7].Value = "Congés payés";
            FeuilleActive.Cells[3, 9].Value = "Congés exceptionnels";
            FeuilleActive.Cells[3, 11].Value = "RTT";
            FeuilleActive.Cells[3, 13].Value = "Récup";
            FeuilleActive.Cells[3, 15].Value = "Formation";
            FeuilleActive.Cells[3, 17].Value = "Maladie";
            FeuilleActive.Range["T2:T3"].Interior.Color = Color.FromArgb(217, 217, 217);
            FeuilleActive.Cells[2, 20].Value = "Astreintes";
            FeuilleActive.Range["V2:AG3"].Interior.Color = Color.FromArgb(217, 217, 217);
            FeuilleActive.Range["V2:AG3"].WrapText = true;
            FeuilleActive.Cells[2, 22].Value = "Heures supplémentaires";
            FeuilleActive.Cells[3, 22].Value = "Semaine de 8h à 20h";
            FeuilleActive.Cells[3, 24].Value = "Semaine de 20h à 8h";
            FeuilleActive.Cells[3, 26].Value = "Samedi de 8h à 20h";
            FeuilleActive.Cells[3, 28].Value = "Samedi de 20h à 8h";
            FeuilleActive.Cells[3, 30].Value = "Dimanche/Jour férié de 8h à 20h";
            FeuilleActive.Cells[3, 32].Value = "Dimanche/Jour férié de 20h à 8h";
            FeuilleActive.Cells[2, 1].EntireRow.Font.Bold = true;
            FeuilleActive.Cells[2, 1].EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            FeuilleActive.Cells[2, 1].EntireRow.VerticalAlignment = XlHAlign.xlHAlignCenter;
            FeuilleActive.Cells[3, 1].EntireRow.Font.Bold = true;
            FeuilleActive.Cells[3, 1].EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            FeuilleActive.Cells[3, 1].EntireRow.VerticalAlignment = XlHAlign.xlHAlignCenter;
            FeuilleActive.Cells[6, 2].EntireRow.Font.Bold = true;
            FeuilleActive.Cells[6, 2].EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            FeuilleActive.Cells[6, 2].EntireRow.VerticalAlignment = XlHAlign.xlHAlignCenter;
            FeuilleActive.Range["G7:G" + DerniereLigne].EntireColumn.Font.Bold = true;


            FeuilleActive.Range["1:1"].RowHeight = 7.25;
            FeuilleActive.Range["2:2"].RowHeight = 38;
            FeuilleActive.Range["3:3"].RowHeight = 52.50;
            FeuilleActive.Range["4:4"].RowHeight = 6;
            FeuilleActive.Range["5:5"].RowHeight = 0;
            FeuilleActive.Range["6:6"].RowHeight = 25;
            
            FeuilleActive.get_Range("B2", "B3").Merge();
            FeuilleActive.get_Range("C2", "C3").Merge();
            FeuilleActive.get_Range("E2", "R2").Merge();
            FeuilleActive.get_Range("G3", "H3").Merge();
            FeuilleActive.get_Range("I3", "J3").Merge();
            FeuilleActive.get_Range("K3", "L3").Merge();
            FeuilleActive.get_Range("M3", "N3").Merge();
            FeuilleActive.get_Range("O3", "P3").Merge();
            FeuilleActive.get_Range("Q3", "R3").Merge();
            FeuilleActive.get_Range("T2", "T3").Merge();
            FeuilleActive.get_Range("V2", "AG2").Merge();
            FeuilleActive.get_Range("V3", "W3").Merge();
            FeuilleActive.get_Range("X3", "Y3").Merge();
            FeuilleActive.get_Range("Z3", "AA3").Merge();
            FeuilleActive.get_Range("AB3", "AC3").Merge();
            FeuilleActive.get_Range("AD3", "AE3").Merge();
            FeuilleActive.get_Range("AF3", "AG3").Merge();

            FormaterLesBordures(123,161,206, FeuilleActive.Range["B7", "C" + DerniereLigne].Borders);
            FormaterLesBordures(123, 161, 206, FeuilleActive.Range["E7", "R" + DerniereLigne].Borders);
            FormaterLesBordures(123, 161, 206, FeuilleActive.Range["T7", "T" + DerniereLigne].Borders);
            FormaterLesBordures(123, 161, 206, FeuilleActive.Range["V7", "AG" + DerniereLigne].Borders);

            FormaterLaCouleurDefond("B", 255, 242, 204, 252, 228, 214);
            FormaterLaCouleurDefond("E", 255, 242, 204, 252, 228, 214);
            FormaterLaCouleurDefond("F", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("G", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("H", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("I", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("J", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("K", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("L", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("M", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("N", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("O", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("P", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("Q", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("R", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("T", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("V", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("W", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("X", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("Y", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("Z", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("AA", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("AB", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("AC", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("AD", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("AE", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("AF", 242, 242, 242, 217, 217, 217);
            FormaterLaCouleurDefond("AG", 242, 242, 242, 217, 217, 217);

            FormaterUneBordure(FeuilleActive.Range["I7", "I" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["K7", "K" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["M7", "M" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["O7", "O" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["Q7", "Q" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["V7", "V" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["X7", "X" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["Z7", "Z" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["AB7", "AB" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["AD7", "AD" + DerniereLigne].Borders);
            FormaterUneBordure(FeuilleActive.Range["AF7", "AF" + DerniereLigne].Borders);

            FeuilleActive.Columns["A:C"].AutoFit();
            FeuilleActive.Columns["G:R"].AutoFit();
            FeuilleActive.Columns["T"].AutoFit();
            FeuilleActive.Columns["V:AG"].AutoFit();
            FeuilleActive.Columns["A"].ColumnWidth = 2.44;
            FeuilleActive.Columns["D"].ColumnWidth = 2.44;
            FeuilleActive.Columns["F"].ColumnWidth = 2.44;
            FeuilleActive.Columns["G"].ColumnWidth = 4.44;
            FeuilleActive.Range["G7:G" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["I"].ColumnWidth = 4.44;
            FeuilleActive.Range["I7:I" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["K"].ColumnWidth = 4.44;
            FeuilleActive.Range["K7:K" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["M"].ColumnWidth = 4.44;
            FeuilleActive.Range["M7:M" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["O"].ColumnWidth = 4.44;
            FeuilleActive.Range["O7:O" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["Q"].ColumnWidth = 4.44;
            FeuilleActive.Range["Q7:Q" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["S"].ColumnWidth = 2.44;
            FeuilleActive.Range["S7:S" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["U"].ColumnWidth = 2.44;
            FeuilleActive.Range["U7:U" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["V"].ColumnWidth = 4.44;
            FeuilleActive.Range["V7:V" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["X"].ColumnWidth = 4.44;
            FeuilleActive.Range["X7:X" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["Z"].ColumnWidth = 4.44;
            FeuilleActive.Range["Z7:Z" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["AB"].ColumnWidth = 4.44;
            FeuilleActive.Range["AB7:AB" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["AD"].ColumnWidth = 4.44;
            FeuilleActive.Range["AD7:AD" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["AF"].ColumnWidth = 4.44;
            FeuilleActive.Range["AF7:AF" + DerniereLigne].EntireColumn.Font.Bold = true;
            FeuilleActive.Columns["E:F"].ColumnWidth = 13;


            FeuilleActive.Range["7:7"].EntireRow.Delete();
            FeuilleActive.Range[DerniereLigne + 1 + ":" + DerniereLigne + 1].EntireRow.Delete();
            FeuilleActive.Range[DerniereLigne + ":" + DerniereLigne].EntireRow.Delete();

            FormaterLesBordures(255, 255, 255, FeuilleActive.Range["B2", "C3"].Borders);
            FormaterLesBordures(255, 255, 255, FeuilleActive.Range["E2", "R3"].Borders);
            FormaterLesBordures(255, 255, 255, FeuilleActive.Range["V2", "AG3"].Borders);
        }

        private Decimal? ReecrireSiNull(Decimal? valeurAVerifier)
        {
            if (valeurAVerifier == null)
                return 0;
            return valeurAVerifier;
        }

        private void FormaterLaCouleurDefond(string colonne, byte R1, byte G1, byte B1, byte R2, byte G2, byte B2)
        {
            for (int index = 7; index <= DerniereLigne; index++)
            {
                int indexSuivant = index++;
                FeuilleActive.Range[colonne + index].Interior.Color = Color.FromArgb(R1, G1, B1);
                FeuilleActive.Range[colonne + indexSuivant].Interior.Color = Color.FromArgb(R2, G2, B2);
            }
            if (FeuilleActive.Cells[DerniereLigne + 1, 2].Text == "")
                FeuilleActive.Range[colonne + DerniereLigne + 1].Interior.Color = Color.FromArgb(255, 255, 255);
        }

        private void FormaterLesBordures(byte R, byte G, byte B, Borders bordures)
        {
            bordures[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            bordures[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            bordures[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            bordures[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
            bordures[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            bordures[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            bordures[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            bordures[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            bordures[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            bordures[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;
            bordures[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            bordures[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            bordures.Color = Color.FromArgb(R, G, B); ;
        }

        private void FormaterUneBordure(Borders bordures)
        {
            bordures[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDash;
            bordures[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
            bordures.Color = Color.FromArgb(123, 161, 206); ;
        }
    }
}
