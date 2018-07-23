using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using PdfConverterLibrary;

namespace BiGPay
{
    public partial class Form : System.Windows.Forms.Form
    {
        private Compteur ProgressionCompteur { get; set; }

        public Form()
        {
            InitializeComponent();
            ProgressionCompteur = new Compteur(this);
            progressBar.Maximum = 8;
        }

        private void Dossier_Click(object sender, EventArgs e)
        {
            if (SelectDossier.ShowDialog() == DialogResult.OK)
            {
                ProgressionCompteur.DepartDecompte();
                ProgressionCompteur.Incrementation(null, null);

                #region Parcours du dossier et Vérification de la présence des fichiers
                //Chemins du dossier de fichiers d'extractions
                string cheminDossier = SelectDossier.SelectedPath;
                //Variables de vérification de présence des classeurs à parcourir
                bool collaborateurs = false;
                bool absences = false;
                bool astreintes = false;
                bool heuresSup = false;
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
                    }
                }
                if (collaborateurs == false || absences == false || astreintes == false
                   || heuresSup == false)
                {
                    MessageBox.Show("Vérifiez les noms donnés à vos classeurs, un ou plusieurs sont manquants.");
                }
                #endregion
                
                #region Traitement
                else
                {
                    ProgressionCompteur.Incrementation(null, null);
                    #region Initialisation_Classeurs
                    //Création des variables qui contiendront les données des classeurs
                    ClasseurResultats classeurResultats = new ClasseurResultats();
                    ClasseurCollaborateurs classeurCollaborateurs = new ClasseurCollaborateurs(cheminDossier + @"\Collaborateurs.xlsx");
                    ClasseurAbsences classeurAbsences = new ClasseurAbsences(cheminDossier + @"\Absences.xlsx");
                    ClasseurHeuresSup classeurHeuresSup = new ClasseurHeuresSup(cheminDossier + @"\Heures_sup.xlsx");
                    ClasseurAstreintes classeurAstreintes = new ClasseurAstreintes(cheminDossier + @"\Astreintes.xlsx");
                    #endregion

                    ProgressionCompteur.Incrementation(null, null);
                    #region Remplissage_Collaborateurs
                    //Remplissage de la partie liée aux collaborateurs dans le classeur de résultats
                    classeurResultats.RemplirColonneCollaborateurs(classeurCollaborateurs);
                    classeurResultats.RemplirColonneMatricules(classeurCollaborateurs);
                    classeurResultats.ReecrireNomsCollaborateurs();
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

                    ProgressionCompteur.Incrementation(null, null);
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

                    ProgressionCompteur.Incrementation(null, null);
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

                    ProgressionCompteur.Incrementation(null, null);
                    #region Remplissage des astreintes
                    ////Remplissage de la partie liée aux codes d'astreintes dans le classeur de résultats
                    DialogResult result = MessageBox.Show(classeurAstreintes.ObtenirCollaborateursQuiOntEteDAstreinteSurLaPeriode(), "Extraction des tickets d'astreintes", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                    if(result == DialogResult.No)
                    {
                        classeurAstreintes.FermerTousLesProcessus();
                        Close();
                    }
                    else
                    {
                        Activate();
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

                        ProgressionCompteur.Incrementation(null, null);
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
                                        excel.ExcelApp.Visible = false;
                                        excel.CollerLesDonnees();

                                        //Initialisation du classeur pdf qui deviendra le classeur généré ci-dessus
                                        ClasseurPdf classeurPdf = new ClasseurPdf();
                                        classeurPdf.Classeur = excel.Workbook;
                                        classeurPdf.InitialiserClasseur(System.IO.Path.GetFileNameWithoutExtension(fichier));
                                        //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                                        if (classeurPdf.NomCollaborateur != "" && classeurPdf.NomCollaborateur != "Client")
                                        {
                                            long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurPdf, 3);
                                            if (ligneACompleter != 0)
                                            {
                                                classeurPdf.ListeTickets = classeurPdf.ObtenirTicketsAstreintes();

                                                foreach (Ticket ticket in classeurPdf.ListeTickets)
                                                {
                                                    if(ticket.Collaborateur == classeurResultats.FeuilleActive.Cells[ligneACompleter, 2].Text)
                                                    {
                                                        classeurResultats.RemplirTicketsAstreintes(ticket, ligneACompleter, classeurPdf, periode);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Une erreur est survenue lors de la conversion du CRA pdf '" + System.IO.Path.GetFileNameWithoutExtension(fichier) + "', impossible de récupérer le nom du collaborateur et de traiter ses tickets d'astreinte, si il en a.", "Extraction des tickets d'astreintes", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                                            Activate();
                                        }
                                        //Fermeture du classeurPdf 
                                        word.FermerDocumentEtApplication();
                                        classeurPdf.Classeur.Close(false, Type.Missing, Type.Missing);
                                    }
                                }
                            }
                        }
                        #endregion

                        ProgressionCompteur.Incrementation(null, null);
                        #region Mise en forme du classeur de résultats
                        classeurResultats.FormaterClasseur();
                        #endregion
                        
                        #region Affichage
                        ProgressionCompteur.Incrementation(null, null);
                        classeurResultats.ExcelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                        classeurResultats.ExcelApp.Visible = true;
                        classeurResultats.ActiverClasseur();
                        #endregion
                    }
                    #endregion
                }
                #endregion
                
                //Fermeture du formulaire
                Close();
            }
        }
    }
}
