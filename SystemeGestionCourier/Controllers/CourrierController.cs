using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using SystemGestionCourier.Models;
using SystemeGestionCourier.DAL;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Configuration;
using System.Xml;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace SystemeGestionCourier.Controllers
{
    public class CourrierController : Controller
    {
        public CourrierContext db = new CourrierContext();

        // GET: /Courrier/
        public ActionResult Index()
        {
            return View(db.Courrier.ToList());
        }

        // GET: /Courrier/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Courrier courrier = db.Courrier.Find(id);
            if (courrier == null)
            {
                return HttpNotFound();
            }
            return View(courrier);
        }

        // GET: /Courrier/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: /Courrier/Create
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Sujet,Expiditeur,Destinataire,Message,CourDate,TypeId")] Courrier courrier)
        {
            if (ModelState.IsValid)
            {
                db.Courrier.Add(courrier);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(courrier);
        }

        // GET: /Courrier/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Courrier courrier = db.Courrier.Find(id);
            if (courrier == null)
            {
                return HttpNotFound();
            }
            return View(courrier);
        }

        // POST: /Courrier/Edit/5
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Sujet,Expiditeur,Destinataire,Message,CourDate,TypeId")] Courrier courrier)
        {
            if (ModelState.IsValid)
            {
                db.Entry(courrier).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(courrier);
        }

        // GET: /Courrier/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Courrier courrier = db.Courrier.Find(id);
            if (courrier == null)
            {
                return HttpNotFound();
            }
            return View(courrier);
        }

        // POST: /Courrier/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Courrier courrier = db.Courrier.Find(id);
            db.Courrier.Remove(courrier);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }


        /**************Import from Excel*********************/

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {

            DataSet ds = new DataSet();

            if (Request.Files["file"].ContentLength > 0)
            {
                string fileExtension = System.IO.Path.GetExtension(Request.Files["file"].FileName);

                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    string FileName = System.IO.Path.GetFileName(Request.Files["file"].FileName);

                    string FolderPath = ConfigurationManager.AppSettings["FolderPath"];



                    string fileLocation = Server.MapPath(FolderPath + FileName);
                    //string fileLocation = Server.MapPath("~/Content/") + Request.Files["file"].FileName;
                    if (System.IO.File.Exists(fileLocation))
                    {

                        System.IO.File.Delete(fileLocation);
                    }
                    Request.Files["file"].SaveAs(fileLocation);
                    string excelConnectionString = string.Empty;
                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    //connection String for xls file format.
                    if (fileExtension == ".xls")
                    {
                        excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    //connection String for xlsx file format.
                    else if (fileExtension == ".xlsx")
                    {

                        excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }
                    //Create Connection to Excel work book and add oledb namespace
                    OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                    excelConnection.Open();
                    DataTable dt = new DataTable();

                    dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dt == null)
                    {
                        return null;
                    }

                    String[] excelSheets = new String[dt.Rows.Count];
                    int t = 0;
                    //excel data saves in temp file here.
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheets[t] = row["TABLE_NAME"].ToString();
                        t++;
                    }
                    OleDbConnection excelConnection1 = new OleDbConnection(excelConnectionString);


                    string query = string.Format("Select * from [{0}]", excelSheets[0]);
                    using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection1))
                    {
                        dataAdapter.Fill(ds);
                    }
                }
                if (fileExtension.ToString().ToLower().Equals(".xml"))
                {
                    string fileLocation = Server.MapPath("~/Content/") + Request.Files["FileUpload"].FileName;
                    if (System.IO.File.Exists(fileLocation))
                    {
                        System.IO.File.Delete(fileLocation);
                    }

                    Request.Files["FileUpload"].SaveAs(fileLocation);
                    XmlTextReader xmlreader = new XmlTextReader(fileLocation);
                    // DataSet ds = new DataSet();
                    ds.ReadXml(xmlreader);
                    xmlreader.Close();
                }

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    /* string conn = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
                     SqlConnection con = new SqlConnection(conn);
                       string query = "Insert into Person(Name,Email,Mobile) Values('" + ds.Tables[0].Rows[i][0].ToString() + "','" + ds.Tables[0].Rows[i][1].ToString() + "','" + ds.Tables[0].Rows[i][2].ToString() + "')";
                       con.Open();
                       SqlCommand cmd = new SqlCommand(query, con);
                       cmd.ExecuteNonQuery();
                       con.Close();*/
                    Courrier messages = new Courrier { Sujet = ds.Tables[0].Rows[i][0].ToString(), Destinataire = ds.Tables[0].Rows[i][1].ToString(), Expiditeur = ds.Tables[0].Rows[i][2].ToString(), Message = ds.Tables[0].Rows[i][3].ToString(), CourDate = DateTime.Parse(ds.Tables[0].Rows[i][4].ToString()), Type = ds.Tables[0].Rows[i][5].ToString() };
                    db.Courrier.Add(messages);
                    db.SaveChanges();

                }


            }
            return RedirectToAction("Index");
        }

        /**************Import from XML*********************/

       [HttpPost]
        public ActionResult Uploadxml(HttpPostedFileBase file)
        {


            if (Request.Files["file"].ContentLength > 0)
            {
                DataSet ds = new DataSet();


                string extension = System.IO.Path.GetExtension(Request.Files["file"].FileName);


                if (extension.ToString().ToLower().Equals(".xml"))
                {
                    string fileLocation = Server.MapPath("~/Content/") + Request.Files["file"].FileName;
                    if (System.IO.File.Exists(fileLocation))
                    {
                        System.IO.File.Delete(fileLocation);
                    }

                    Request.Files["file"].SaveAs(fileLocation);
                    XmlTextReader xmlreader = new XmlTextReader(fileLocation);

                    ds.ReadXml(xmlreader);
                    xmlreader.Close();
                }
           
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    Courrier messages = new Courrier { Sujet = ds.Tables[0].Rows[i][0].ToString(), Destinataire = ds.Tables[0].Rows[i][1].ToString(), Expiditeur = ds.Tables[0].Rows[i][2].ToString(), Message = ds.Tables[0].Rows[i][3].ToString(), CourDate = DateTime.Parse(ds.Tables[0].Rows[i][4].ToString()), Type = ds.Tables[0].Rows[i][5].ToString() };
                    db.Courrier.Add(messages);
                    db.SaveChanges();

                }
                
            }
            return RedirectToAction("Index");
        }
        
        /******reporting*********************/
     /*  public ActionResult CrystalReport1()
       {
           //var reportPath=Server.MapPath
           var reportPath = Server.MapPath("~/ReportTemplates/CrystalReport1.rpt");
           var customer = db.Courrier;
           using (var ReportDocument = new ReportDocument())
           {
               ReportDocument.Load(reportPath);
               ReportDocument.SetDataSource(customer);
               var response = System.Web.HttpContext.Current.Response;
               ReportDocument.ExportToHttpResponse(ExportFormatType.PortableDocFormat, response, true, "CrystalReport1");


           }
           return new EmptyResult();
       }*/
    }
}
