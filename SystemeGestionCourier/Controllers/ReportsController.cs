using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SystemeGestionCourier.DAL;

namespace SystemeGestionCourier.Controllers
{
    public class ReportsController : Controller
    {
        private readonly CourrierContext _db = new CourrierContext();


        //
        // GET: /Reports/
        public ActionResult CrystalReport1()
        {
            
            var reportPath = Server.MapPath("~/ReportTemplates/CrystalReport1.rpt");
            var customer = _db.Courrier;
           using (var ReportDocument = new ReportDocument())
            {
               
                ReportDocument.Load(reportPath);
                ReportDocument.SetDataSource(customer);
               /* var response = System.Web.HttpContext.Current.Response;
                  ReportDocument.ExportToHttpResponse(ExportFormatType.PortableDocFormat, response, true, "CrystalReport1");
               */

            }
            return new EmptyResult();
        }
	}
}