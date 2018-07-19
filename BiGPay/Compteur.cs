using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BiGPay
{
    public class Compteur
    {
        public Timer Time = new Timer();
        public int Decompte { get; set; }
        public Form Formulaire { get; set; }
        
        public Compteur(Form formulaire)
        {
            Formulaire = formulaire;
            Decompte = 0;
            Time.Enabled = true;
        }

        public void DepartDecompte()
        {
            Time.Tick += new EventHandler(Incrementation);
        }

        public void Incrementation(object sender, EventArgs e)
        {
            if (Decompte >= 9)
            {
                Time.Enabled = false;
            }
            else
            {
                //do something here
                Decompte++;
                Formulaire.progressBar.Increment(1);
            }
        }
    }
}
