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


namespace ImportDataFromExcelPOC.Controllers
{
    public class ImportDataController : Controller
    {
        //
        // GET: /ImportData/

        // GET: Home
        public ActionResult Index()
        
        {
            string filePath = "D:\\Practice\\PtacNewDealers.xlsx";

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



                            conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                            using (SqlConnection con = new SqlConnection(conString))
                            {
                                String TruncateQuery = "DELETE FROM PtacDealer_TB";
                                SqlCommand command = new SqlCommand(TruncateQuery, con);

                                con.Open();
                                command.ExecuteNonQuery();
                                con.Close();
                                //Read Data from First Sheet.

                                connExcel.Open();
                                cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                foreach (DataRow row in dt.Rows)
                                {
                                    var Description = row["SUBFDESC"].ToString().Trim();
                                    ConvertToCamelCase(Description);

                                    var Address = row["SUBFADR1"].ToString().Trim();
                                    var City = row["SUBFCITY"].ToString().Trim();
                                    var State = row["SUBFSTATE"].ToString().Trim();
                                    var ZipCode = row["SUBFZIP"].ToString().Trim();
                                    var PhoneNumber = row["PhoneNum"].ToString().Trim();
                                    GetPhoneNumber(PhoneNumber);

                                    var PhoneNumber2 = row["PhoneNum2"].ToString().Trim();

                                    //conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                                    //using (SqlConnection con = new SqlConnection(conString))
                                    //{


                                    String query = "INSERT INTO dbo.PtacDealer_TB (Description,Address,City,State,ZipCode,PhoneNumber,PhoneNumber2) VALUES ('" + Description.Replace("'", "''") + "','" + Address.Replace("'", "''") + "','" + City.Replace("'", "''") + "','" + State + "','" + ZipCode + "','" + PhoneNumber + "','" + PhoneNumber2 + "')";

                                    SqlCommand command1 = new SqlCommand(query, con);

                                    con.Open();
                                    command1.ExecuteNonQuery();
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

        public static string ConvertToCamelCase(string DealerName)
        {

            //var data = DealerName.ToUpperInvariant();
            //TextInfo txtInfo = new CultureInfo("en-us", true).TextInfo;
            //DealerName = txtInfo.ToTitleCase(DealerName);


           // var yourString = "WARD_VS_VITAL_SIGNS".ToLower().Replace("_", " ");
            //TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            //DealerName = info.;
            //Console.WriteLine(DealerName);
            //return DealerName;

            //string textToChange = "WARD_VS_VITAL_SIGNS";
            System.Text.StringBuilder resultBuilder = new System.Text.StringBuilder();

            //foreach (char c in DealerName)
            //{
            //    // Replace anything, but letters and digits, with space
            //    if (!Char.IsLetterOrDigit(c))
            //    {
            //        resultBuilder.Append(" ");
            //    }
            //    else
            //    {
            //        resultBuilder.Append(c);
            //    }
            //}

            string result = resultBuilder.ToString();

            // Make result string all lowercase, because ToTitleCase does not change all uppercase correctly
            result = result.ToLower();

            // Creates a TextInfo based on the "en-US" culture.
            TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

            result = myTI.ToTitleCase(result);
            return result;
        }


        public static string GetPhoneNumber(string  number)
        {
            string data;
            string n = number.ToString();
            if (n.Length == 13)
            {
                data =   n.Substring(0,4)+  n.Substring(4, 4) + n.Substring(5,13) ;
                return data;
            }
            else if (n.Length == 12)
            {
                data = "(" + n.Substring(0,3) + ")" + " "+ n.Substring(4,4);
                    //data = "(" + n.Substring(0, 2) + ") " + n.Substring(3, 6) + n.Substring(8, 11);
                //  data = "(" + n.Substring(0, 3) + ") " + n.Substring(3, 3) + " - " + n.Substring(6, 4);
                return data;
            }
            
            else 
            {
                data = "(" + n.Substring(0, 3) + ")" + " " + n.Substring(4, 6) + " " + n.Substring(7, 10);
                return data;
            }
                //data = number;
                //return data;
        }

    }
}






