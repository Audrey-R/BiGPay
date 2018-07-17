using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    public class ClasseurPdf : ClasseurExcel
    {
        public const int _ColonneType = 6;
        public const int _ColonneDate = 8;
        public const int _ColonneQuantite = 11;
        public const int _ColonneDetails = 12;
        public const int _ColonneCollaborateurs = 1;
        public string NomCollaborateur { get; set; }
        public List<Ticket> ListeTickets { get; set; }


        public ClasseurPdf() { }

        public void InitialiserClasseur()
        {
            ExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            ExcelApp.Application.DisplayAlerts = true;
            ExcelApp.Visible = true;
            FeuilleActive = Classeur.Sheets[1];
            DerniereLigne = FeuilleActive.Cells[FeuilleActive.Rows.Count, 2].End(XlDirection.xlUp).Row;
            DerniereColonne = FeuilleActive.Cells[_PremiereColonne, FeuilleActive.Columns.Count].End(XlDirection.xlToLeft).Column;
            Donnees = FeuilleActive.Range[ConvertirColonneEnLettre(_ColonneCollaborateurs) + _PremiereLigne, ConvertirColonneEnLettre(DerniereColonne) + DerniereLigne];
            Collaborateur = FeuilleActive.Cells[3, 1];
            NomCollaborateur = Collaborateur.Text;
            NomCollaborateur = NomCollaborateur.Split(':')[1].Trim();
            FeuilleActive.Cells[3, 1].Value = NomCollaborateur;
        }

        public List<Ticket> ObtenirTicketsAstreintes()
        {
            List<Ticket> listeTickets = new List<Ticket>() ;
            for (int index = _PremiereLigne; index <= DerniereLigne; index++)
            {
                if (FeuilleActive.Cells[index, 1].Text == "Astreintes/Tickets")
                {
                    int nouvelIndex = index + 2;
                    for(int indexTicket = nouvelIndex; indexTicket <= DerniereLigne; indexTicket++)
                    {
                        if (FeuilleActive.Cells[indexTicket, _ColonneType].Text == "Tickets")
                        {
                            string dateTexte = FeuilleActive.Cells[indexTicket, _ColonneDate].Text;
                            string nbHeuresTexte = Convert.ToDecimal(FeuilleActive.Cells[indexTicket, _ColonneQuantite].Text).ToString("0.0");
                            string details = FeuilleActive.Cells[indexTicket, _ColonneDetails].Text;
                            string heureDebut = "";
                            //Boucle sur chacun des caractères de la chaîne
                            for (int indexChar = 0; indexChar <= details.Length; indexChar++)
                            {
                                //Recherche du chiffre précédant l'unité d'heure, dans la chaine de caractères
                                int indexCharEstInt, indexCharMoinsUnEstInt;
                                var resultChar = int.TryParse(details[indexChar].ToString(), out indexCharEstInt);
                                var resultCharMoinsUn = false;
                                if (indexChar - 1 >= 0)
                                {
                                    resultCharMoinsUn = int.TryParse(details[indexChar - 1].ToString(), out indexCharMoinsUnEstInt);
                                }
                                    

                                //Si l'unité d'heure utilisée dans le commentaire est ':'
                                if (resultChar && details[indexChar+1] == ':')
                                {
                                    if(resultCharMoinsUn)
                                        heureDebut = details.Substring(indexChar - 1, 2);
                                    heureDebut = details.Substring(indexChar, 1);
                                    break;
                                }
                                //Si l'unité d'heure utilisée dans le commentaire est 'h'
                                else if (resultChar && details[indexChar+1] == 'h')
                                {
                                    if (resultCharMoinsUn)
                                        heureDebut = details.Substring(indexChar - 1, 2);
                                    heureDebut = details.Substring(indexChar, 1);
                                    break;
                                }
                                else
                                {
                                    heureDebut = "Impossible de déterminer l'heure de début du ticket d'astreinte";
                                }
                            }
                            Ticket ticket = new Ticket { Date = Convert.ToDateTime(dateTexte), NbHeures = Convert.ToDecimal(nbHeuresTexte), HeureDebut = new TimeSpan(Convert.ToInt32(heureDebut),0,0)};
                            listeTickets.Add(ticket);
                        }
                    }
                }
            }
            return listeTickets;
        }

        public string ObtenirHeuresSupplementairesTicket(Ticket ticket, Periode periode)
        {
            
                string date = ticket.Date.ToShortDateString();
                string nbheures = ticket.NbHeures.ToString("0.0");
                if(nbheures.Split(',')[1] == "0")
                    nbheures = ticket.NbHeures.ToString("0");
                string ticketRetourne = nbheures + "h le " + date + " (Ticket astreinte)";

                // Test date et heure
                if (ClasseurHeuresSup._DateTombeEnSemaine(ticket.Date))
                {
                    if (ClasseurHeuresSup._HeureEntre8hEt20h(ticket.HeureDebut))
                    {
                        ticketRetourne = "Sem-8-20|" + ticketRetourne;
                    }
                    else if (ClasseurHeuresSup._HeureEntre20hEt8h(ticket.HeureDebut))
                    {
                        ticketRetourne = "Sem-20-8|" + ticketRetourne;
                    }
                    else
                    {
                        ticketRetourne = "Erreur";
                    }
                }
                else if (ClasseurHeuresSup._DateTombeUnSamedi(ticket.Date))
                {
                    if (ClasseurHeuresSup._HeureEntre8hEt20h(ticket.HeureDebut))
                    {
                        ticketRetourne = "Sam-8-20|" + ticketRetourne;
                    }
                    else if (ClasseurHeuresSup._HeureEntre20hEt8h(ticket.HeureDebut))
                    {
                        ticketRetourne = "Sam-20-8|" + ticketRetourne;
                    }
                    else
                    {
                        ticketRetourne = "Erreur";
                    }
                }
                else if (ClasseurHeuresSup._DateTombeUnDimancheOuUnJourFerie(ticket.Date, periode))
                {
                    if (ClasseurHeuresSup._HeureEntre8hEt20h(ticket.HeureDebut))
                    {
                        ticketRetourne = "DF-8-20|" + ticketRetourne;
                    }
                    else if (ClasseurHeuresSup._HeureEntre20hEt8h(ticket.HeureDebut))
                    {
                        ticketRetourne = "DF-20-8|" + ticketRetourne;
                    }
                    else
                    {
                        ticketRetourne = "Erreur";
                    }
                }
                else
                {
                    ticketRetourne = "Erreur";
                }
                return ticketRetourne;
        }
    }
}