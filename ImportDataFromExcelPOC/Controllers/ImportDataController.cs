using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using System.Web.Mvc;
using System;
using System.Web;
using ImportDataFromExcelPOC.Models;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ImportDataFromExcelPOC.Utility;


namespace ImportDataFromExcelPOC.Controllers
{
    public class ImportDataController : Controller
    {
        //
        // GET: /ImportData/

        // GET: Home
        public ActionResult Index()
        
        {
            //string filePath = System.Configuration.ConfigurationManager.AppSettings["FolderPath"].ToString(); //moved folder path to web.cofig
            string filePath = "D:\\Practice\\PtacNewDealers.xlsx";
           
            ExcelUtility ex = new ExcelUtility();
            ex.ReadData(filePath);
       
            return View();
        }
    }
}






