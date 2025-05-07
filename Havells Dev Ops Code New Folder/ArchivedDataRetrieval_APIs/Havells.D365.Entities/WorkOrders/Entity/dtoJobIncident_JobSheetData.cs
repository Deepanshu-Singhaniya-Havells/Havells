using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.WorkOrders.Entity
{
   public class dtoJobIncident_JobSheetData
    {
        public string IncidentId { get; set; }
        public string ObservationName { get; set; }
        public string ModelName { get; set; }
        public string AssetNo { get; set; }
        public string IncidentType { get; set; }
        public string invoiceNo { get; set; }
        public string InvoiceDate { get; set; }
        public string ProductDesription { get; set; }
        public string TechnicianRemarks { get; set; }
        public string PurchasedFrom { get; set; }
    }
}
