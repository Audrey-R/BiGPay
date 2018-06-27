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
                    //Création des variables qui contiendront les données des classeurs
                    ClasseurResultats classeurResultats = new ClasseurResultats();
                    ClasseurCollaborateurs classeurCollaborateurs = new ClasseurCollaborateurs(cheminDossier + @"\Collaborateurs.xlsx");
                    //ClasseurAbsences classeurAbsences = new ClasseurAbsences(cheminDossier + @"\Absences.xlsx");
                    //ClasseurExcel classeurAstreintes = new ClasseurExcel(cheminDossier + @"\Astreintes.xlsx");
                    //ClasseurExcel classeurHeuresSup = new ClasseurExcel(cheminDossier + @"\Heures_sup.xlsx");
                    //ClasseurExcel classeurWeekEndFeries = new ClasseurExcel(cheminDossier + @"\Weekend_Feries.xlsx");

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
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Periode periode = new Periode(DateTime.Now);
            //List<DateTime> liste = periode.RetournerTousLesJoursFeriesPourLaPeriode(DateTime.Now);
            //MessageBox.Show("Pâques    : " + periode.LundiDePaques.ToLongDateString() + "\n" +
            //                 "Ascension : " + periode.JeudiDeLAscension.ToLongDateString() + "\n" +
            //                 "Pentecôte : " + periode.LundiDePentecote.ToLongDateString() + "\n" +
            //                 "Jour de l'an : " + periode.JourDeLAn.ToLongDateString() + "\n" +
            //                 "Fete du travail : " + periode.FeteDuTravail.ToLongDateString() + "\n" +
            //                 "8 mai : " + periode.HuitMai1945.ToLongDateString() + "\n" +
            //                 "Fete nationale : " + periode.FeteNationale.ToLongDateString() + "\n" +
            //                 "Assomption : " + periode.Assomption.ToLongDateString() + "\n" +
            //                 "Toussait : " + periode.Toussaint.ToLongDateString() + "\n" +
            //                 "Armistice : " + periode.Armistice.ToLongDateString() + "\n" +
            //                 "Noel : " + periode.Noel.ToLongDateString() + "\n" +
            //                 "DateDebut : " + periode.DateDebutPeriode.ToLongDateString() + "\n"+
            //                 "DateFin : " + periode.DateFinPeriode.ToLongDateString() + "\n");
        }

        
    }
}
