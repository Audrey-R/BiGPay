using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BiGPay
{
    public class Ticket
    {
        public DateTime Date { get; set; }
        public Decimal NbHeures { get; set; }
        public TimeSpan HeureDebut { get; set; }
        public string Collaborateur { get; set; }
    }
}
