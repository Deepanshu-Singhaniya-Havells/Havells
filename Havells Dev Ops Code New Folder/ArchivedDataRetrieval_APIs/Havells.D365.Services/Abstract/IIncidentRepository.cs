using Havells.D365.Entities.Incident.Response;

namespace Havells.D365.Services.Abstract
{
    public interface IIncidentRepository
    {
        IncidentResponse GetIncidentById(string incidentId);
        IncidentResponse GetIncidentByWorkOrderID(string workorderId);
    }
}
