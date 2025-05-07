
using Havells.D365.Entities.WorkOrderService.Response;
namespace Havells.D365.Services.Abstract
{
    public interface IWOServiceRepository
    {
        WorkOrderServiceResponse GetWOServiceDetailByJobID(string WorkOrderID);
        WorkOrderServiceResponse GetWOServiceDetailByID(string WorkOrderServiceID);
        WorkOrderServiceResponse GetWOServiceDetailByIncidentID(string WorkOrderIncidentID);
    }
}
