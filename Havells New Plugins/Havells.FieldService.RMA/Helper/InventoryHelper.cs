using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.FieldService.RMA.Helper
{
    public class InventoryHelper
    {
        public static EntityReference GetUserPosition(IOrganizationService _service, EntityReference _userRef)
        {
            EntityReference positionRef = _service.Retrieve(_userRef.LogicalName, _userRef.Id, new ColumnSet("positionid")).GetAttributeValue<EntityReference>("positionid");
            return positionRef;
        }
        public static RetriveBookableResourceResponse RetriveBookableResource(IOrganizationService service, EntityReference _userID)
        {
            RetriveBookableResourceResponse bookableResource = new RetriveBookableResourceResponse();
            try
            {
                QueryExpression query = new QueryExpression("bookableresource");
                query.ColumnSet = new ColumnSet("msdyn_warehouse", "hil_defectivewarehouse");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("userid", ConditionOperator.Equal, _userID.Id);
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count == 1)
                {
                    bookableResource.BooableResource = _entitys[0].ToEntityReference();
                    bookableResource.DefectiveWarehouse = _entitys[0].GetAttributeValue<EntityReference>("hil_defectivewarehouse");
                    bookableResource.FreshWareHouse = _entitys[0].GetAttributeValue<EntityReference>("msdyn_warehouse");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in Reteiving Bookable Resource. " + ex.Message);
            }
            return bookableResource;
        }
    }
    public class RetriveBookableResourceResponse
    {
        public EntityReference BooableResource { get; set; }
        public EntityReference FreshWareHouse { get; set; }
        public EntityReference DefectiveWarehouse { get; set; }

    }
    public class PurchaseOrderType
    {
        public static readonly Guid Emergency = new Guid("784405c2-974c-ee11-be6f-6045bdac526a");
        public static readonly Guid ManualOrder = new Guid("3da37e41-7557-ee11-be6f-6045bdac5292");
        public static readonly Guid MinimumStockLevel = new Guid("9c0a9f6c-9157-ee11-be6e-6045bdaa91c3");
    }
    public class PurchaseOrderModel
    {
        public EntityReference hil_potype { get; set; }
        public EntityReference msdyn_vendor { get; set; }
        public EntityReference msdyn_receivetowarehouse { get; set; }
        public DateTime msdyn_purchaseorderdate { get; set; }
        public EntityReference msdyn_requestedbyresource { get; set; }
        public EntityReference msdyn_orderedby { get; set; }
        public string IV_Usages { get; set; }
        public EntityReference ownerid { get; set; }
        public EntityReference msdyn_workorder { get; set; }

    }
}
