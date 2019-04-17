using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace ImportDataFromExcelPOC.Models
{
    public class ImportDataModel
    {
        // private string _description;
        public string Description { get; set; }
        //{
        //    get
        //    {
        //        return ConvertToPascalCase(_description);
        //    }
        //    set
        //    {
        //        _description = value;
        //    }
        //}

        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        // private string _phoneNumber;
        public string PhoneNumber { get; set; }
        //    {
        //        get
        //        {
        //            return GetPhoneNumber(_phoneNumber);
        //        }
        //        set
        //        {
        //            _phoneNumber = value;
        //        }
        //    }

        public string PhoneNumber2 { get; set; }
        //    {
        //        get
        //        {
        //            return GetPhoneNumber(_phoneNumber);
        //        }
        //        set
        //        {
        //            _phoneNumber = value;
        //        }
        //    }
        ////    public static string GetPhoneNumber(string number)
        //    {
        //        string formattedNumber;
        //        //string justDigits = new string(a.Where(number => char.IsDigit(number)).ToArray());    
        //        if (number != "NULL")
        //        {
        //            string result = Regex.Replace(number, @"^(\+)|\D", "$1");
        //            formattedNumber = "(" + result.Substring(0, 3) + ")" + " " + result.Substring(3, 3) + " " + "-" + " " + result.Substring(6, 4);
        //        }
        //        else
        //        {
        //            return number = "";
        //        }
        //        return formattedNumber;
        //    }


        //    public static string ConvertToPascalCase(string DealerName)
        //    {
        //        string dealerNameLowerCase;
        //        string[] exceptionalCases = new string[] { "PTAC", "SVC", "SVCS", "LLC", "INC", "A/C", "ACR", "CO" };
        //        // Make DealerName string all lowercase, because ToTitleCase does not change all uppercase correctly //
        //        dealerNameLowerCase = DealerName.ToLower();

        //        // Creates a TextInfo based on the "en-US" culture//           
        //        TextInfo myTextInfo = new CultureInfo("en-US", false).TextInfo;
        //        dealerNameLowerCase = myTextInfo.ToTitleCase(dealerNameLowerCase);
        //        //string data;

        //      var formattedDealerName =  GetFormattedDealerName(exceptionalCases, dealerNameLowerCase);

        //        return formattedDealerName;
        //    }

        //    public static string GetFormattedDealerName(string[] exceptionalCases, string dealerNameLowerCase)
        //    {
        //        string formattedString = dealerNameLowerCase;
        //        var dealerNames = dealerNameLowerCase.Split(' ');
        //        //Console.WriteLine(dealerNames[0]+" "+dealerNames[1]+" "+dealerNames[2]);
        //        foreach (var str in dealerNames)
        //        {
        //            bool exists = exceptionalCases.Any(s => s.ToLower().Contains(str.ToLower()));
        //            if (exists)
        //            {
        //                var tempStr = str.ToUpper();
        //                formattedString = formattedString.Replace(str, tempStr);
        //            }
        //        }           
        //        return formattedString;
        //    }
        //}


    }
}