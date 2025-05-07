using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;


namespace D365WebJobs.RnDs1
{
    public class InspectionTypeSpareAndService
    {
        public InspectionTypeResponse PopulateSpareAndService(IOrganizationService _service, Guid _workOrderIncidentId)
        {
            InspectionTypeResponse _retObj = new InspectionTypeResponse();

            Entity _entWOIncident = _service.Retrieve("msdyn_workorderincident", _workOrderIncidentId, new ColumnSet("msdyn_incidenttype", "msdyn_customerasset", "msdyn_workorder"));
            if (_entWOIncident != null)
            {
                #region Populate Service Items
                QueryExpression Query = new QueryExpression(msdyn_incidenttypeservice.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("msdyn_service");
                Query.Criteria.AddFilter(LogicalOperator.And);
                Query.Criteria.AddCondition("msdyn_incidenttype", ConditionOperator.Equal, _entWOIncident.GetAttributeValue<EntityReference>("msdyn_incidenttype").Id);
                Query.Criteria.AddCondition("msdyn_service", ConditionOperator.NotNull);
                EntityCollection _entCol = _service.RetrieveMultiple(Query);

                foreach (Entity ent in _entCol.Entities)
                {
                    _retObj._services.Add(new InspectionTypeObject()
                    {
                        partCode = ent.GetAttributeValue<EntityReference>("msdyn_service").Name,
                        partDescription = ent.GetAttributeValue<EntityReference>("msdyn_service").Name,
                        partAmount = new Money(0),
                        quantity = 1
                    });
                }
                #endregion
                #region Populate Spare Parts
                Query = new QueryExpression(msdyn_incidenttypeservice.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("msdyn_service");
                Query.Criteria.AddFilter(LogicalOperator.And);
                Query.Criteria.AddCondition("msdyn_incidenttype", ConditionOperator.Equal, _entWOIncident.GetAttributeValue<EntityReference>("msdyn_incidenttype").Id);
                Query.Criteria.AddCondition("msdyn_service", ConditionOperator.NotNull);
                _entCol = _service.RetrieveMultiple(Query);

                foreach (Entity ent in _entCol.Entities)
                {
                    _retObj._services.Add(new InspectionTypeObject()
                    {
                        partCode = ent.GetAttributeValue<EntityReference>("msdyn_service").Name,
                        partDescription = ent.GetAttributeValue<EntityReference>("msdyn_service").Name,
                        partAmount = new Money(0),
                        quantity = 1
                    });
                }
                #endregion
            }
            return _retObj;
        }

    public class InspectionTypeResponse
    {
        public  List<InspectionTypeObject> _spareParts { get; set; }
        public List<InspectionTypeObject> _services { get; set; }

        public InspectionTypeResponse()
        {
            this._spareParts = new List<InspectionTypeObject>();
            this._services = new List<InspectionTypeObject>();
        }
    }
        public class InspectionTypeObject
        {
            public decimal quantity { get; set; }
            public string partCode { get; set; }
            public string partDescription { get; set; }
            public Money partAmount { get; set; }
        }
}
