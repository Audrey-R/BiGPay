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
    public class ClasseurHeuresSup : ClasseurExcel
    {
        public DateTime Date { get; set; }
        public TimeSpan HeureDebut { get; set; }
        public Decimal NbHeures { get; set; }
        public TimeSpan _HeureDepartNuit = new TimeSpan(20, 0, 0);
        public TimeSpan _HeureFinNuit = new TimeSpan(8, 0, 0);
        public TimeSpan _Minuit = new TimeSpan(0, 0, 0);
        public const int _ColonneCollaborateurs = 1;
        public const int _ColonneDate = 3;
        public const int _ColonneHeureDebut = 5;
        public const int _ColonneNbHeures = 6;
        

        public ClasseurHeuresSup() { }
        public ClasseurHeuresSup(string libelleClasseur)
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
            CelluleA1 = FeuilleActive.get_Range("A1", Type.Missing);
            TrierFeuille(FeuilleActive.Range[ConvertirColonneEnLettre(_PremiereColonne) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne], 1);
        }

        private Boolean DateTombeEnSemaine(DateTime dateAtester)
        {
            int jourSemaineDateAtester = (int)dateAtester.DayOfWeek;
            if(jourSemaineDateAtester > 0 && jourSemaineDateAtester < 6)
                return true;
            return false;
        }

        private Boolean DateTombeUnSamedi(DateTime dateAtester)
        {
            int jourSemaineDateAtester = (int)dateAtester.DayOfWeek;
            if (jourSemaineDateAtester == 6)
                return true;
            return false;
        }

        private Boolean DateTombeUnDimancheOuUnJourFerie(DateTime dateAtester, Periode periode)
        {
            List<DateTime> joursFeriesPeriode = periode.JoursFeries;
            int jourSemaineDateAtester = (int)dateAtester.DayOfWeek;
            if (jourSemaineDateAtester == 7)
                return true;
            for(int index = 0; index <= joursFeriesPeriode.Count ; index ++)
            {
                if (dateAtester == joursFeriesPeriode[index])
                    return true;
            }
            return false;
        }

        private Boolean HeureEntre8hEt20h(TimeSpan heureAtester)
        {
            if (heureAtester >= _HeureFinNuit && heureAtester <= _HeureDepartNuit)
                return true;
            return false;
        }

        private Boolean HeureEntre20hEt8h(TimeSpan heureAtester)
        {
            if (heureAtester >= _HeureDepartNuit || heureAtester >= _Minuit && heureAtester <= _HeureFinNuit)
                return true;
            return false;
        }

        public string ObtenirHeuresSupplementaires(int index, Periode periode)
        {
            #region Variables
            string date = FeuilleActive.Cells[index, _ColonneDate].Text;
            string heureDebut = FeuilleActive.Cells[index, _ColonneHeureDebut].Text;
            string nbHeures = FeuilleActive.Cells[index, _ColonneNbHeures].Text;
            nbHeures = Convert.ToDecimal(nbHeures).ToString("0.0");
            if(nbHeures.Split(',')[1] == "0")
                nbHeures = nbHeures.Split(',')[0];
            string heuresSupplementaires;
            #endregion

            #region Traitement
            Date = Convert.ToDateTime(date);
            // Conversion des heures textes en TimeSpan
            string heuresHeureDebut = heureDebut.Split(':')[0];
            string minutesHeureDebut = heureDebut.Split(':')[1];
            HeureDebut = new TimeSpan(Convert.ToInt32(heuresHeureDebut), Convert.ToInt32(minutesHeureDebut), 0);


            heuresSupplementaires = nbHeures + "h le " + Date.ToShortDateString();

            // Test date et heure
            if (DateTombeEnSemaine(Date))
            {
                if (HeureEntre8hEt20h(HeureDebut))
                {
                    heuresSupplementaires =  "Sem-8-20|" + heuresSupplementaires ;
                }
                else if (HeureEntre20hEt8h(HeureDebut))
                {
                    heuresSupplementaires = "Sem-20-8|" + heuresSupplementaires;
                }
                else
                {
                    heuresSupplementaires = "Erreur";
                }
            }
            else if (DateTombeUnSamedi(Date))
            {
                if (HeureEntre8hEt20h(HeureDebut))
                {
                    heuresSupplementaires = "Sam-8-20|" + heuresSupplementaires;
                }
                else if (HeureEntre20hEt8h(HeureDebut))
                {
                    heuresSupplementaires = "Sam-20-8|" + heuresSupplementaires;
                }
                else
                {
                    heuresSupplementaires = "Erreur";
                }
            }
            else if (DateTombeUnDimancheOuUnJourFerie(Date, periode))
            {
                if (HeureEntre8hEt20h(HeureDebut))
                {
                    heuresSupplementaires = "DF-8-20|" + heuresSupplementaires;
                }
                else if (HeureEntre20hEt8h(HeureDebut))
                {
                    heuresSupplementaires = "DF-20-8|" + heuresSupplementaires;
                }
                else
                {
                    heuresSupplementaires = "Erreur";
                }
            }
            else
            {
                heuresSupplementaires = "Erreur";
            }
            return heuresSupplementaires;
            #endregion
        }

        public Decimal ObtenirNombreHeuresSupplementaires(int index)
        {
            string nbHeuresTexte = FeuilleActive.Cells[index, _ColonneNbHeures].Text;
            Decimal nbHeures = Convert.ToDecimal(nbHeuresTexte);
            return nbHeures;
        }
    }
}
