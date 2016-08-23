using System;
using System.Collections.Generic;

namespace SystemGestionCourier.Models
{
    public class Courrier
    {
        public int ID { get; set; }
        public string Sujet { get; set; }
        public string Expiditeur { get; set; }
        public string Destinataire { get; set; }
        public string Message { get; set; }
        public DateTime CourDate { get; set; }

        public String Type { get; set; }
    }
}