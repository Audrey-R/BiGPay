using System;
using System.Collections.Generic;
using System.Globalization;

namespace BiGPay
{
    public class Periode
    {
        public DateTime DateDebutPeriode { get; set; }
        public DateTime DateFinPeriode { get; set; }
        public List<DateTime> JoursFeries { get; set; }
        public DateTime JourDeLAn, Paques, LundiDePaques, FeteDuTravail, 
                         HuitMai1945, JeudiDeLAscension, LundiDePentecote, 
                         FeteNationale, Assomption, Toussaint, Armistice, 
                         Noel;
        public int NbJoursOuvresPeriode { get; set; }
        public static string[] _Formats = { "MM/dd/yyyy", "MMM-dd-yyyy", "yyyy-MM-dd", "MM-dd-yyyy", "MM-dd-yy", "M/d/yyyy", "MMM dd yyyy", "M/yyyy" };
        
        public Periode(DateTime datePremiereAbsenceEnregistree)
        {
            DateDebutPeriode = new DateTime(datePremiereAbsenceEnregistree.Year, datePremiereAbsenceEnregistree.Month, 1);
            DateFinPeriode = new DateTime(datePremiereAbsenceEnregistree.Year, datePremiereAbsenceEnregistree.Month,DateTime.DaysInMonth(datePremiereAbsenceEnregistree.Year, datePremiereAbsenceEnregistree.Month));
            JoursFeries = RetournerTousLesJoursFeriesPourLaPeriode(datePremiereAbsenceEnregistree);
            //Jour de solidarite à ôter de la liste
            ExtraireJourDeSolidarite(JoursFeries, LundiDePentecote);
            NbJoursOuvresPeriode = CalculerJoursOuvresPeriode(DateDebutPeriode, DateFinPeriode);
        }

        public static Periode _CreerPeriodePaie(ClasseurAbsences classeurAbsences)
        {
            if (classeurAbsences.DerniereLigne > 1)
            {
                //Réécriture de la première date d'absence
                string premiereAbsence = classeurAbsences.FeuilleActive.Cells[ClasseurExcel._PremiereLigne + 1, ClasseurAbsences._ColonneDepartAbsence].Text;
                DateTime dateTest;
                if(!DateTime.TryParse(premiereAbsence, out dateTest))
                {
                     premiereAbsence = DateTime.ParseExact(
                     premiereAbsence,
                     _Formats,
                     CultureInfo.InvariantCulture,
                     DateTimeStyles.None)
                     .ToString("dd/MM/yyyy");
                }
                //Conversion de la date réécrite et initialisation dans sa variable
                classeurAbsences.PremiereDateAbsence = Convert.ToDateTime(premiereAbsence);
                Periode periode = new Periode(classeurAbsences.PremiereDateAbsence);

                return periode;
            }
            return null;
        }

        public List<DateTime> RetournerTousLesJoursFeriesPourLaPeriode(DateTime datePremiereAbsenceEnregistree)
        {
            int Y = datePremiereAbsenceEnregistree.Year;// Annee
            int golden;                                 // Nombre d'or
            int solar;                                  // Correction solaire
            int lunar;                                  // Correction lunaire
            int pfm;                                    // Pleine lune de paques
            int dom;                                    // Nombre dominical
            int easter;                                 // jour de paques
            int tmp;

            // Nombre d'or
            golden = (Y % 19) + 1;
            if (Y <= 1752)            // Calendrier Julien
            {
                // Nombre dominical
                dom = (Y + (int)(Y / 4) + 5) % 7;
                if (dom < 0) dom += 7;
                // Date non corrigee de la pleine lune de paques
                pfm = (3 - (11 * golden) - 7) % 30;
                if (pfm < 0) pfm += 30;
            }
            else                       // Calendrier Gregorien
            {
                // Nombre dominical
                dom = (Y + (int)(Y / 4) - (int)(Y / 100) + (int)(Y / 400)) % 7;
                if (dom < 0) dom += 7;
                // Correction solaire et lunaire
                solar = (int)(Y - 1600) / 100 - (int)(Y - 1600) / 400;
                lunar = (int)(((int)(Y - 1400) / 100) * 8) / 25;
                // Date non corrigee de la pleine lune de paques
                pfm = (3 - (11 * golden) + solar - lunar) % 30;
                if (pfm < 0) pfm += 30;
            }
            // Date corrige de la pleine lune de paques :
            // jours apres le 21 mars (equinoxe de printemps)
            if ((pfm == 29) || (pfm == 28 && golden > 11)) pfm--;

            tmp = (4 - pfm - dom) % 7;
            if (tmp < 0) tmp += 7;

            // Paques en nombre de jour apres le 21 mars
            easter = pfm + tmp + 1;

            if (easter < 11)
            {
                Paques = DateTime.Parse((easter + 21) + "/3/" + Y);
            }
            else
            {
                Paques = DateTime.Parse((easter - 10) + "/4/" + Y);
            }

            //1er janvier
            JourDeLAn = new DateTime(Y, 1, 1);
            //Lundi de Pâques
            LundiDePaques = Paques.AddDays(1);
            //Fête du travail
            FeteDuTravail = new DateTime(Y, 5, 1);
            //8 mai 1945
            HuitMai1945 = new DateTime(Y, 5, 8);
            //Jeudi de l'Ascension
            JeudiDeLAscension = Paques.AddDays(39);
            //Lundi de Pentecôte == Solidarité
            LundiDePentecote = Paques.AddDays(50);
            //Fête Nationale
            FeteNationale = new DateTime(Y, 7, 14);
            //Assomption
            Assomption = new DateTime(Y, 8, 15);
            //Toussaint
            Toussaint = new DateTime(Y, 11, 1);
            //Armistice
            Armistice = new DateTime(Y, 11, 11);
            //Noël
            Noel = new DateTime(Y, 12, 25);

            //Ajout des jours fériés dans une liste
            List<DateTime> listeJoursFeries = new List<DateTime>();
            listeJoursFeries.AddRange(
                new List<DateTime>
                {
                    JourDeLAn, LundiDePaques, FeteDuTravail,
                    HuitMai1945, JeudiDeLAscension, LundiDePentecote,
                    FeteNationale, Assomption, Toussaint, Armistice, Noel
                });

            return listeJoursFeries;
        }

        public void ExtraireJourDeSolidarite(List<DateTime> liste, DateTime jourDeSolidarite)
        {
            if(liste.Contains(jourDeSolidarite))
                liste.Remove(jourDeSolidarite);
        }

        private int CalculerJoursOuvresPeriode(DateTime dateDebutPeriode, DateTime dateFinPeriode)
        {
            int jours = 0;
            while (dateDebutPeriode <= dateFinPeriode)
            {
                if (dateDebutPeriode.DayOfWeek != DayOfWeek.Saturday && dateDebutPeriode.DayOfWeek != DayOfWeek.Sunday)
                {
                    jours++;
                    foreach (DateTime jourFerie in JoursFeries)
                    {
                        if (jourFerie == dateDebutPeriode)
                            jours--;
                    }
                }
                dateDebutPeriode = dateDebutPeriode.AddDays(1);
            }
            return jours;
        }
    }
}
