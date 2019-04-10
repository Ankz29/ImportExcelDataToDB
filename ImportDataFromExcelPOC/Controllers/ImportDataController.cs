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
                                    ConvertToPascalCase(Description);

                                    var Address = row["SUBFADR1"].ToString().Trim();
                                    var City = row["SUBFCITY"].ToString().Trim();
                                    var State = row["SUBFSTATE"].ToString().Trim();
                                    var ZipCode = row["SUBFZIP"].ToString().Trim();
                                    var PhoneNumber = row["PhoneNum"].ToString().Trim();
                                    var formattedPhoneNumber ="";
                                    if (!string.IsNullOrEmpty(PhoneNumber) || PhoneNumber == "NULL")
                                    {
                                        formattedPhoneNumber = GetPhoneNumber(PhoneNumber);
                                    }

                                    var PhoneNumber2 = row["PhoneNum2"].ToString().Trim();
                                    var formattedPhoneNumber2 = "";
                                    if (!string.IsNullOrEmpty(PhoneNumber2) || PhoneNumber2 == "NULL")
                                    {
                                        formattedPhoneNumber2 = GetPhoneNumber(PhoneNumber2);
                                    }
                                    //conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                                    //using (SqlConnection con = new SqlConnection(conString))
                                    //{


                                    String query = "INSERT INTO dbo.PtacDealer_TB (Description,Address,City,State,ZipCode,PhoneNumber,PhoneNumber2) VALUES ('" + Description.Replace("'", "''") + "','" + Address.Replace("'", "''") + "','" + City.Replace("'", "''") + "','" + State + "','" + ZipCode + "','" + formattedPhoneNumber + "','" + formattedPhoneNumber2 + "')";

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

        //method to convert DealerName to Pascal Case & format DealerName as per requirement//
        public static string ConvertToPascalCase(string DealerName)
        {
            string dealerNameLowerCase;
          
            // Make DealerName string all lowercase, because ToTitleCase does not change all uppercase correctly //
            dealerNameLowerCase = DealerName.ToLower();

            // Creates a TextInfo based on the "en-US" culture//           
           TextInfo myTextInfo = new CultureInfo("en-US", false).TextInfo;
            dealerNameLowerCase = myTextInfo.ToTitleCase(dealerNameLowerCase);
            

            String[] exceptionalCases = new string[8] { "PTAC", "SVC", "SVCS", "LLC", "INC", "A/C", "ACR", "CO" };
            //if(dealerNameLowerCase.Contains(exceptionalCases[]))
            //{

            //}


            return dealerNameLowerCase;
        }



        public static string GetPhoneNumber(string  number)
        {
            string formattedNumber;
            //string justDigits = new string(a.Where(number => char.IsDigit(number)).ToArray());    
            if (number != "NULL")
            {
                string result = Regex.Replace(number, @"^(\+)|\D", "$1");
                 formattedNumber = "(" + result.Substring(0, 3) + ")" + " " + result.Substring(3, 3) + " " + "-" + " " + result.Substring(6, 4);
            }
            else
            {
                return number="";
            }
                return formattedNumber;
        }

    }
}






