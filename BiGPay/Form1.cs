using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

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
                else
                {
                    #region Initialisation_Classeurs
                    //Création des variables qui contiendront les données des classeurs
                    ClasseurResultats classeurResultats = new ClasseurResultats();
                    ClasseurCollaborateurs classeurCollaborateurs = new ClasseurCollaborateurs(cheminDossier + @"\Collaborateurs.xlsx");
                    ClasseurAbsences classeurAbsences = new ClasseurAbsences(cheminDossier + @"\Absences.xlsx");
                    //ClasseurExcel classeurAstreintes = new ClasseurExcel(cheminDossier + @"\Astreintes.xlsx");
                    //ClasseurExcel classeurHeuresSup = new ClasseurExcel(cheminDossier + @"\Heures_sup.xlsx");
                    //ClasseurExcel classeurWeekEndFeries = new ClasseurExcel(cheminDossier + @"\Weekend_Feries.xlsx");
                    #endregion

                    #region Remplissage_Collaborateurs
                    //Remplissage de la partie liée aux collaborateurs dans le classeur de résultats
                    classeurResultats.RemplirColonneCollaborateurs(classeurCollaborateurs);
                    classeurResultats.RemplirColonneMatricules(classeurCollaborateurs);
                    for (int index = 1; index <= classeurCollaborateurs.DerniereLigne; index++)
                    {
                        //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                        long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurCollaborateurs, index);
                        if(ligneACompleter != 0)
                        {
                            classeurResultats.RemplirColonneEntreesSorties(ligneACompleter, index, classeurCollaborateurs);
                        }
                    }
                    //Fermeture du classeurCollaborateurs
                    classeurCollaborateurs.Classeur.Close(false, Type.Missing, Type.Missing);
                    #endregion

                    #region Remplissage_Absences
                    //Remplissage de la partie liée aux absences dans le classeur de résultats
                    classeurResultats.CreerPeriodePaie(classeurAbsences);
                    for (int index = 1; index <= classeurAbsences.DerniereLigne; index++)
                    {
                        //Recherche d'une correspondance de nom entre les deux classeurs et récupération du numéro de la ligne
                        long ligneACompleter = classeurResultats.RechercherCollaborateur(classeurAbsences, index);
                        if (ligneACompleter != 0)
                        {
                            classeurResultats.RemplirColonne("Congés_Payés", ClasseurResultats._ColonneCongesPayes, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirTotalColonne("Congés_Payés", ClasseurResultats._ColonneTotalCongesPayes, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirColonne("RTT", ClasseurResultats._ColonneRTT, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirTotalColonne("RTT", ClasseurResultats._ColonneTotalRTT, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirColonne("Formation", ClasseurResultats._ColonneFormation, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirTotalColonne("Formation", ClasseurResultats._ColonneTotalFormation, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirColonne("Maladie", ClasseurResultats._ColonneMaladie, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirTotalColonne("Maladie", ClasseurResultats._ColonneTotalMaladie, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirColonne("Récupération", ClasseurResultats._ColonneRecup, ligneACompleter, index, classeurAbsences);
                            classeurResultats.RemplirTotalColonne("Récupération", ClasseurResultats._ColonneTotalRecup, ligneACompleter, index, classeurAbsences);
                        }
                    }
                    #endregion


                    //Fermeture du formulaire
                    Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Periode periode = new Periode(DateTime.Now);
            List<DateTime> liste = periode.RetournerTousLesJoursFeriesPourLaPeriode(DateTime.Now);
            periode.ExtraireJourDeSolidarite(liste, periode.LundiDePentecote);
            MessageBox.Show("Pâques    : " + liste.Contains(periode.LundiDePaques) + "\n" +
                             "Ascension : " + liste.Contains(periode.JeudiDeLAscension) + "\n" +
                             "Pentecôte : " + liste.Contains(periode.LundiDePentecote) + "\n" +
                             "Jour de l'an : " + liste.Contains(periode.JourDeLAn) + "\n" +
                             "Fete du travail : " + liste.Contains(periode.FeteDuTravail) + "\n" +
                             "8 mai : " + liste.Contains(periode.HuitMai1945) + "\n" +
                             "Fete nationale : " + liste.Contains(periode.FeteNationale) + "\n" +
                             "Assomption : " + liste.Contains(periode.Assomption) + "\n" +
                             "Toussait : " + liste.Contains(periode.Toussaint) + "\n" +
                             "Armistice : " + liste.Contains(periode.Armistice) + "\n" +
                             "Noel : " + liste.Contains(periode.Noel) + "\n" +
                             "DateDebut : " + periode.DateDebutPeriode.ToLongDateString() + "\n" +
                             "DateFin : " + periode.DateFinPeriode.ToLongDateString() + "\n");
        }

        
    }
}
