using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    class ClasseurResultats : ClasseurExcel
    {
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

        public void RemplirColonneCollaborateurs(ClasseurExcel classeurCollaborateurs)
        {
            DerniereLigne = 5 + classeurCollaborateurs.DerniereLigne;
            FeuilleActive.Range["A7", "A" + DerniereLigne].Value
            = classeurCollaborateurs.FeuilleActive.Range["C2", "C" + classeurCollaborateurs.DerniereLigne].Value;
        }

        public void RemplirColonneMatricules(ClasseurExcel classeurCollaborateurs)
        {
            FeuilleActive.Range["B7", "B" + DerniereLigne].Value
            = classeurCollaborateurs.FeuilleActive.Range["B2", "B" + classeurCollaborateurs.DerniereLigne].Value;
        }

        public long RechercherCollaborateur(ClasseurExcel classeurCollaborateurs, int index)
        {
            LigneAcompleter = 0;
            classeurCollaborateurs.Collaborateur = classeurCollaborateurs.Donnees.Cells[index, 1];
            Collaborateur = FeuilleActive.Cells[6, 2].Find(classeurCollaborateurs.Collaborateur.Text,
                            Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, Type.Missing,
                            XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing
                            );
            if (Collaborateur.Text == classeurCollaborateurs.Collaborateur.Text)
            {
                LigneAcompleter = Collaborateur.Row;
            }
            return LigneAcompleter;
        }

        public void RemplirColonneEntreesSorties(long ligneACompleter, int index, ClasseurCollaborateurs classeurCollaborateurs)
        {
            Range celluleACompleter = FeuilleActive.Cells[ligneACompleter, 3];
            if(classeurCollaborateurs.ObtenirEntreesEtSortiesDuMois(index) != null)
            {
                celluleACompleter.Value = classeurCollaborateurs.ObtenirEntreesEtSortiesDuMois(index).ToString();
            }
        }
    }
}
