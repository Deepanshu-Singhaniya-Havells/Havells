using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.WorkOrders.Entity
{
    public class DtoJobSheetData
    {
        public List<dtoJob_JobSheetData> Job { get; set; }
        public List<dtoJobIncident_JobSheetData> IncidentData { get; set; }
        public List<DtoJobSparePart_JobSheetData> SparePartAndProductData { get; set; }
    }
}
