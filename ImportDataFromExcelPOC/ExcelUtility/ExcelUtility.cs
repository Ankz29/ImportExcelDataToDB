using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using ImportDataFromExcelPOC.Models;
using System.Globalization;

namespace ImportDataFromExcelPOC.Utility
{
    public class ExcelUtility
    {
        //string filePath = ConfigurationManager.ConnectionStrings["FolderPath"].ConnectionString; //moved folder path to web.cofig
        string conString = string.Empty;
        public string ReadData(string filePath)
        {
             if (System.IO.File.Exists(filePath)) //dont use Full qualified namespaces//
            {
                string extension = Path.GetExtension(filePath);              
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
                            //connExcel.Close();



                            conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                            using (SqlConnection con = new SqlConnection(conString))
                            {
                                String TruncateQuery = "DELETE FROM PtacDealer_TB";
                                SqlCommand command = new SqlCommand(TruncateQuery, con);

                                con.Open();
                                command.ExecuteNonQuery();
                                con.Close();
                                //Read Data from First Sheet.

                                //connExcel.Open();
                                cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                foreach (DataRow row in dt.Rows)
                                {
                                    var importData = new ImportDataModel();

                                    importData.Description = row["SUBFDESC"].ToString().Trim();
                                   var formattedData = ConvertToPascalCase(importData.Description);

                                    importData.Address = row["SUBFADR1"].ToString().Trim();
                                    importData.City = row["SUBFCITY"].ToString().Trim();
                                    importData.State = row["SUBFSTATE"].ToString().Trim();
                                    importData.ZipCode = row["SUBFZIP"].ToString().Trim();
                                    importData.PhoneNumber = row["PhoneNum"].ToString().Trim();
                                    var data = GetPhoneNumber(importData.PhoneNumber);

                                    importData.PhoneNumber2 = row["PhoneNum2"].ToString().Trim();
                                    var data1 = GetPhoneNumber(importData.PhoneNumber2);

                                    String query = "INSERT INTO dbo.PtacDealer_TB (Description,Address,City,State,ZipCode,PhoneNumber,PhoneNumber2) VALUES ('" + formattedData.Replace("'", "''") + "','" + importData.Address.Replace("'", "''") + "','" + importData.City.Replace("'", "''") + "','" + importData.State + "','" + importData.ZipCode + "','" + data + "','" + data1 + "')";

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
             return filePath;
        }

      
        //method to convert DealerName to Pascal Case & format DealerName as per requirement//
        public static string ConvertToPascalCase(string DealerName)
        {
            string dealerNameLowerCase;
            string[] exceptionalCases = new string[] { "PTAC", "SVC", "SVCS", "LLC", "INC", "A/C", "ACR", "CO" };
            // Make DealerName string all lowercase, because ToTitleCase does not change all uppercase correctly //
            dealerNameLowerCase = DealerName.ToLower();

            // Creates a TextInfo based on the "en-US" culture//           
            TextInfo myTextInfo = new CultureInfo("en-US", false).TextInfo;
            dealerNameLowerCase = myTextInfo.ToTitleCase(dealerNameLowerCase);
            //string data;

          var data =  GetFormattedDealerName(exceptionalCases, dealerNameLowerCase);
            
            return data;
        }

        public static string GetFormattedDealerName(string[] exceptionalCases, string dealerNameLowerCase)
        {
            string formattedString = dealerNameLowerCase;
            //string space = dealerNameLowerCase.Remove(WhiteSpaceTrimStringConverter);
            var dealerNames = dealerNameLowerCase.Split(' ');
            //Console.WriteLine(dealerNames[0]+" "+dealerNames[1]+" "+dealerNames[2]);
            foreach (var str in dealerNames)
            {
                bool exists = exceptionalCases.Any(s => s.ToLower().Equals(str.ToLower()));
                if (exists)
                {
                    var tempStr = str.ToUpper();
                    formattedString = formattedString.Replace(str, tempStr);
                }
            }           
            return formattedString;
        }

        public static string GetPhoneNumber(string number)
        {
            string formattedNumber;
            string result = Regex.Replace(number, "[^0-9a-zA-Z]+", "");
            //string justDigits = new string(a.Where(number => char.IsDigit(number)).ToArray());    
            if (result != "NULL" && result.Length >= 10)
            {
                //string result = Regex.Replace(number, @"^(\+)|\D", "$1");
                
                formattedNumber = "(" + result.Substring(0, 3) + ")" + " " + result.Substring(3, 3) + " " + "-" + " " + result.Substring(6, 4);                
            }
            
            else 
            {
               return number = "" ;
            }
            
            return formattedNumber;
        }
    }
}
                  
