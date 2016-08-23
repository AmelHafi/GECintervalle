
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Linq;
using System.Web;
using SystemGestionCourier.Models;

namespace SystemeGestionCourier.DAL
{
    public class CourrierContext : DbContext
    {
        public CourrierContext() : base("CourrierContext")
        {
        }
        
        public DbSet<Courrier> Courrier { get; set; }
      
       
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
        }
    }
}