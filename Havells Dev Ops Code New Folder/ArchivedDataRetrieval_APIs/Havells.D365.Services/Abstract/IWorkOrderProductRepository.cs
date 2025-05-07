using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Havells.D365.Entities.WorkorderProduct.Response;

namespace Havells.D365.Services.Abstract
{
    public interface IWorkOrderProductRepository
    {
        WorkOrderProductResponse GetWOProductDetailByID(string WorkOrderProductID);
        WorkOrderProductResponse GetWOProductDetailByIncidentID(string WorkOrderIncidentID);
        WorkOrderProductResponse GetWOProductDetailByJobID(string WorkOrderID);

    }
}
