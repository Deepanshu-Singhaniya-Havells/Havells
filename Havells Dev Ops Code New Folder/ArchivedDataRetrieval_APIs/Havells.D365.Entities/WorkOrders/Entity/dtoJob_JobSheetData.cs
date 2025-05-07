using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.WorkOrders.Entity
{
    public class dtoJob_JobSheetData
    {
        public string WorkOrderId { get; set; }
        public string ProductCategory { get; set; }
        public string Pincode { get; set; }
        public string ProductSubCategory { get; set; }
        public string MobileNumber { get; set; }
        public string JobNumber { get; set; }
        public string FullAddress { get; set; }
        public string ComplaintDescription { get; set; }
        public string CustomerName { get; set; }
        public string CallingNumber { get; set; }
        public string AlternateNumber { get; set; }
        public string JonWarranty { get; set; }
        public string CountryClassification { get; set; }
        public string ActualCharges { get; set; }
        public string Quantity { get; set; }
        public string Owner { get; set; }
        public string SendTRC { get; set; }
        public string VisitDate { get; set; }

    }
}
