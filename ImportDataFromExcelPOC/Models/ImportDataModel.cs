using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportDataFromExcelPOC.Models
{
    public class ImportDataModel
    {
        public string Description { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string PhoneNumber { get; set; }
        public string PhoneNumber2 { get; set; }
    }
}