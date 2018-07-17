using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using PdfConverterLibrary;

namespace BiGPay
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Dossier_Click(object sender, EventArgs e)
        {
            if (SelectDossier.ShowDialog() == DialogResult.OK)
            {
                #region Parcours du dossier et Vérification de la présence des fichiers
                //Chemins du dossier de fichiers d'extractions
                string cheminDossier = SelectDossier.SelectedPath;
                //Variables de vérification de présence des classeurs à parcourir
                bool collaborateurs = false;
                bool absences = false;
                bool astreintes = false;
                bool heuresSup = false;
                bool weekEndferies= false;
                foreach (string fichier in System.IO.Directory.GetFiles(cheminDossier))
                {
                    if (System.IO.Path.GetExtension(fichier) == ".xlsx")
                    {
                        string nomClasseurExcel = System.IO.Path.GetFileNameWithoutExtension(fichier);
                        //MessageBox.Show(nomClasseurExcel);
                        if (nomClasseurExcel == "Collaborateurs")
                            collaborateurs = true;
                        if (nomClasseurExcel == "Absences")
                            absences = true;
                        if (nomClasseurExcel == "Astreintes")
                            astreintes = true;
                        if (nomClasseurExcel == "Heures_sup")
                            heuresSup = true;
                        if (nomClasseurExcel == "Weekend_Feries")
                            weekEndferies = true;
                    }
                }
                if (collaborateurs == false || absences == false || astreintes == false
                   || heuresSup == false || weekEndferies == false)
                {
                    MessageBox.Show("Vérifiez les noms donnés à vos classeurs, un ou plusieurs sont manquants.");
                }
                #endregion

                #region Traitement
                else
                {
                    #region Initialisation_Classeurs
                    //Création des variables qui contiendront les données des classeurs
                    ClasseurResultats classeurResultats = new ClasseurResultats();
                    ClasseurCollaborateurs classeurCollaborateurs = new ClasseurCollaborateurs(cheminDossier + @"\Collaborateurs.xlsx");
                    ClasseurAbsences classeurAbsences = new ClasseurAbsences(cheminDossier + @"\Absences.xlsx");
                    ClasseurHeuresSup classeurHeuresSup = new ClasseurHeuresSup(cheminDossier + @"\Heures_sup.xlsx");
                    ClasseurAstreintes classeurAstreintes = new ClasseurAstreintes(cheminDossier + @"\Astreintes.xlsx");
                    classeurResultats.ExcelApp.Visible = true;
                    //ClasseurExcel classeurWeekEndFeries = new ClasseurExcel(cheminDossier + @"\Weekend_Feries.xlsx");
                    #endregion

                    #region Remplissage_Collaborateurs
                    //Remplissage de la partie liée aux collaborateurs dans le classeur de résultats
                    classeurResultats.RemplirColonneCollaborateurs(classeurCollaborateurs);
                    classeurResultats.RemplirColonneMatricules(classeurCollaborateurs);
                    for (int index = 2; index <= classeurCollaborateurs.DerniereLigne; index++)
                    {
                        //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                        long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurCollaborateurs, index);
                        if(ligneACompleter > 0)
                        {
                            classeurResultats.RemplirColonneEntreesSorties(ligneACompleter, index, classeurCollaborateurs);
                        }
                    }
                    //Fermeture du classeurCollaborateurs
                    classeurCollaborateurs.Classeur.Close(false, Type.Missing, Type.Missing);
                    #endregion

                    #region Remplissage_Absences
                    ////Remplissage de la partie liée aux absences dans le classeur de résultats
                    Periode periode = Periode._CreerPeriodePaie(classeurAbsences);
                    classeurResultats.FeuilleActive.Range["B6"].Value = periode.DateDebutPeriode.ToString("MMMM yyyy");
                    for (int index = 2; index <= classeurAbsences.DerniereLigne; index++)
                    {
                        //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                        long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurAbsences, index);
                        if (ligneACompleter != 0)
                        {
                            classeurResultats.RemplirAbsences(ligneACompleter, index, classeurAbsences, periode);
                        }
                    }
                    classeurResultats.RemplirJoursTravaillesPeriode(periode);
                    //Fermeture du classeurAbsences
                    classeurAbsences.Classeur.Close(false, Type.Missing, Type.Missing);
                    #endregion

                    #region Remplissage des Heures supplémentaires
                    ////Remplissage de la partie liée aux heures supplémentaires dans le classeur de résultats
                    for (int index = 2; index <= classeurHeuresSup.DerniereLigne; index++)
                    {
                        //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                        long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurHeuresSup, index);
                        if (ligneACompleter != 0)
                        {
                            classeurResultats.RemplirHeuresSupplementaires(ligneACompleter,index, classeurHeuresSup, periode);
                        }
                    }
                    //Fermeture du classeurHeuresSup
                    classeurHeuresSup.Classeur.Close(false, Type.Missing, Type.Missing);
                    #endregion

                    #region Remplissage des codes d'astreintes
                    ////Remplissage de la partie liée aux codes d'astreintes dans le classeur de résultats
                    for (int index = 2; index <= classeurAstreintes.DerniereLigne; index++)
                    {
                        //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                        long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurAstreintes, index);
                        if (ligneACompleter != 0)
                        {
                            classeurResultats.RemplirCodesAstreintes(ligneACompleter, index, classeurAstreintes);
                        }
                    }
                    //Fermeture du classeurCollaborateurs
                    classeurAstreintes.Classeur.Close(false, Type.Missing, Type.Missing);
                    #endregion

                    #region Traitement de CRA pdf pour remplissage tickets asteinte
                    // Recherche du dossier "CRA"
                    foreach (string dossier in System.IO.Directory.GetDirectories(cheminDossier))
                    {
                        if (System.IO.Path.GetFileName(dossier) == "CRA")
                        {
                            foreach (string fichier in System.IO.Directory.GetFiles(dossier))
                            {
                                // Recherche de fichiers pdf dans le dossier
                                if (System.IO.Path.GetExtension(fichier) == ".pdf")
                                {
                                    // Conversion du fichier pdf vers excel pour copier/coller les données
                                    Word word = new Word();
                                    word.OuvrirPdf(fichier);
                                    word.CopierLesDonnees();
                                    Excel excel = new Excel();
                                    excel.CollerLesDonnees();

                                    //Initialisation du classeur pdf qui deviendra le classeur généré ci-dessus
                                    ClasseurPdf classeurPdf = new ClasseurPdf();
                                    classeurPdf.Classeur = excel.Workbook;
                                    classeurPdf.InitialiserClasseur();
                                    //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                                    long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurPdf, 3);
                                    if (ligneACompleter != 0)
                                    {
                                        classeurPdf.ListeTickets = classeurPdf.ObtenirTicketsAstreintes();

                                        foreach (Ticket ticket in classeurPdf.ListeTickets)
                                        {
                                            classeurResultats.RemplirTicketsAstreintes(ticket, ligneACompleter, classeurPdf, periode);
                                        }
                                    }
                                    word.FermerDocumentEtApplication();
                                    //Fermeture du classeurPdf 
                                    classeurPdf.Classeur.Close(false, Type.Missing, Type.Missing);
                                }
                            }
                        }
                    }
                    #endregion
                    
                    // Affichage du résultat
                    classeurResultats.ExcelApp.Visible = true ;
                    //Fermeture du formulaire
                    Close();
                }
                #endregion
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Periode periode = new Periode(new DateTime(2018,12,1));
            MessageBox.Show(periode.NbJoursOuvresPeriode.ToString());
            //List<DateTime> liste = periode.RetournerTousLesJoursFeriesPourLaPeriode(DateTime.Now);
            //periode.ExtraireJourDeSolidarite(liste, periode.LundiDePentecote);
            //MessageBox.Show("Pâques    : " + liste.Contains(periode.LundiDePaques) + "\n" +
            //                 "Ascension : " + liste.Contains(periode.JeudiDeLAscension) + "\n" +
            //                 "Pentecôte : " + liste.Contains(periode.LundiDePentecote) + "\n" +
            //                 "Jour de l'an : " + liste.Contains(periode.JourDeLAn) + "\n" +
            //                 "Fete du travail : " + liste.Contains(periode.FeteDuTravail) + "\n" +
            //                 "8 mai : " + liste.Contains(periode.HuitMai1945) + "\n" +
            //                 "Fete nationale : " + liste.Contains(periode.FeteNationale) + "\n" +
            //                 "Assomption : " + liste.Contains(periode.Assomption) + "\n" +
            //                 "Toussait : " + liste.Contains(periode.Toussaint) + "\n" +
            //                 "Armistice : " + liste.Contains(periode.Armistice) + "\n" +
            //                 "Noel : " + liste.Contains(periode.Noel) + "\n" +
            //                 "DateDebut : " + periode.DateDebutPeriode.ToLongDateString() + "\n" +
            //                 "DateFin : " + periode.DateFinPeriode.ToLongDateString() + "\n");
        }

        
    }
}
