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


namespace ImportDataFromExcelPOC.Controllers
{
    public class ImportDataController : Controller
    {
        //
        // GET: /ImportData/

        // GET: Home
        public ActionResult Index()
        {
            string filePath = "D:\\PTACDealerList.xlsx";
            
                if (System.IO.File.Exists(filePath))
            { 
                string extension = Path.GetExtension(filePath);
                string conString = string.Empty;
                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                        break;
                }
                

                DataTable dt = new DataTable();

                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            foreach (DataRow row in dt.Rows)
                            {
                                var Description = row["SUBFDESC"].ToString().Trim();
                                var Address = row["SUBFADR1"].ToString().Trim();
                                var City = row["SUBFCITY"].ToString().Trim();
                                var State = row["SUBFSTATE"].ToString().Trim();
                                var ZipCode = row["SUBFZIP"].ToString().Trim();
                                var PhoneNumber = row["PhoneNum"].ToString().Trim();
                                var PhoneNumber2 = row["PhoneNum"].ToString().Trim();

                                conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                                using (SqlConnection con = new SqlConnection(conString))
                                {
                                    String query = "INSERT INTO dbo.PtacDealer_TB (Description,Address,City,State,ZipCode,PhoneNumber,PhoneNumber2) VALUES ('" + Description.Replace("'", "''") + "','" + Address.Replace("'", "''") + "','" + City.Replace("'", "''") + "','" + State + "','" + ZipCode + "','" + PhoneNumber + "','" + PhoneNumber2 + "')";

                                    SqlCommand command = new SqlCommand(query, con);

                                    //var desc =  command.Parameters.Add("@Description", Description);
                                    //  command.Parameters.Add("@Address", Address);
                                    //  command.Parameters.Add("@City", City);
                                    //  command.Parameters.Add("@State", State);
                                    //  command.Parameters.Add("@ZipCode", ZipCode);
                                    //  command.Parameters.Add("@PhoneNumber", PhoneNumber);
                                    //  command.Parameters.Add("@PhoneNumber2", PhoneNumber2);
                                    //con.Open();
                                    //string  result = command.ExecuteNonQuery();


                                    con.Open();
                                    command.ExecuteNonQuery();
                                    con.Close();
                                    //command.ExecuteNonQuery();
                                    connExcel.Close();
                                }


                            }
                        }
                    }
                }
            } 
            return View();
        
        
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {


            //string filePath = string.Empty;
            //if (postedFile != null)
            //{
            //    string path = Server.MapPath("~/Uploads/");
            //    if (!Directory.Exists(path))
            //    {
            //        Directory.CreateDirectory(path);
            //    }

            //    filePath = path + Path.GetFileName(postedFile.FileName);
            //    string extension = Path.GetExtension(postedFile.FileName);
            //    postedFile.SaveAs(filePath);

            //    string conString = string.Empty;
            //    switch (extension)
            //    {
            //        case ".xls": //Excel 97-03.
            //            conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
            //            break;
            //        case ".xlsx": //Excel 07 and above.
            //            conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
            //            break;
            //    }

            //    DataTable dt = new DataTable();

            //    conString = string.Format(conString, filePath);

            //    using (OleDbConnection connExcel = new OleDbConnection(conString))
            //    {
            //        using (OleDbCommand cmdExcel = new OleDbCommand())
            //        {
            //            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
            //            {
            //                cmdExcel.Connection = connExcel;

            //                //Get the name of First Sheet.
            //                connExcel.Open();
            //                DataTable dtExcelSchema;
            //                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            //                connExcel.Close();

            //                //Read Data from First Sheet.
            //                connExcel.Open();
            //                cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
            //                odaExcel.SelectCommand = cmdExcel;
            //                odaExcel.Fill(dt);
            //                foreach (DataRow row in dt.Rows)
            //                {
            //                    var Description = row["SUBFDESC"].ToString();
            //                    var Address = row["SUBFADR1"].ToString();
            //                    var City = row["SUBFCITY"].ToString();
            //                    var State = row["SUBFSTATE"].ToString();
            //                    var ZipCode = row["SUBFZIP"].ToString();
            //                    var PhoneNumber = row["PhoneNum"].ToString();
            //                    var PhoneNumber2 = row["PhoneNum"].ToString();

            //                    conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
            //                    using (SqlConnection con = new SqlConnection(conString))
            //                    {
            //                        String query = "INSERT INTO dbo.PtacDealerNew_TB (Description,Address,City,State,ZipCode,PhoneNumber,PhoneNumber2) VALUES ('" + Description.Replace("'", "''") + "','" + Address.Replace("'", "''") + "','" + City.Replace("'", "''") + "','" + State + "','" + ZipCode + "','" + PhoneNumber + "','" + PhoneNumber2 + "')";

            //                        SqlCommand command = new SqlCommand(query, con);

            //                        //var desc =  command.Parameters.Add("@Description", Description);
            //                        //  command.Parameters.Add("@Address", Address);
            //                        //  command.Parameters.Add("@City", City);
            //                        //  command.Parameters.Add("@State", State);
            //                        //  command.Parameters.Add("@ZipCode", ZipCode);
            //                        //  command.Parameters.Add("@PhoneNumber", PhoneNumber);
            //                        //  command.Parameters.Add("@PhoneNumber2", PhoneNumber2);
            //                        //con.Open();
            //                        //string  result = command.ExecuteNonQuery();


            //                        con.Open();
            //                        command.ExecuteNonQuery();
            //                        con.Close();
            //                        //command.ExecuteNonQuery();
            //                    }


            //                }
            //            }
            //        }
            //    }
            //}
            return View();
        }
    }
}

                




