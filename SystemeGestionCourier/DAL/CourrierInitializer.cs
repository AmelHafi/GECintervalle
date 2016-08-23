using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SystemGestionCourier.Models;

namespace SystemeGestionCourier.DAL
{
    public class CourrierInitializer : System.Data.Entity.DropCreateDatabaseIfModelChanges<CourrierContext>
    {
        protected override void Seed(CourrierContext context)
        {
            var Courriers = new List<Courrier>
            {
            new Courrier{Sujet="sujet1",Destinataire="Alexander",Expiditeur="",Message="message1",CourDate=DateTime.Parse("2005-09-01"),Type="départ"},
            new Courrier{Sujet="sujet2",Destinataire="Alonso",Expiditeur="",Message="message2",CourDate=DateTime.Parse("2002-09-01"),Type="arrivée"},
            new Courrier{Sujet="sujet3",Destinataire="",Expiditeur="Alexander",Message="message3",CourDate=DateTime.Parse("2003-09-01"),Type="départ"},
            new Courrier{Sujet="sujet4",Destinataire="",Expiditeur="Alexander",Message="message4",CourDate=DateTime.Parse("2002-09-01"),Type="arrivée"},
            new Courrier{Sujet="sujet5",Destinataire="Li",Expiditeur="",Message="message5",CourDate=DateTime.Parse("2002-09-01"),Type="interne"},
            new Courrier{Sujet="sujet6",Destinataire="Justice",Message="message6",Expiditeur="",CourDate=DateTime.Parse("2001-09-01"),Type="départ"},
            new Courrier{Sujet="sujet7",Destinataire="Norman",Message="message7",Expiditeur="Alexander",CourDate=DateTime.Parse("2003-09-01"),Type="arrivée"},
            new Courrier{Sujet="sujet8",Destinataire="Olivetto",Message="message8",Expiditeur="Alexander",CourDate=DateTime.Parse("2005-09-01"),Type="interne"}
            };

            Courriers.ForEach(s => context.Courrier.Add(s));
            context.SaveChanges();
            

        }
    }
}