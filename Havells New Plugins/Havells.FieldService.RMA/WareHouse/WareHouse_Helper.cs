using Havells.FieldService.RMA.Helper;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.FieldService.RMA.WareHouse
{
    public class PositonIDs
    {
        public static readonly Guid DSE = new Guid("7d1ecbab-1208-e911-a94d-000d3af0694e");
        public static readonly Guid Franchise = new Guid("4a1aa189-1208-e911-a94d-000d3af0694e");
    }
    public class WareHouse_Helper
    {
        public static void WareHouseOwnerValidation(IOrganizationService service, EntityReference ownerID)
        {
            EntityReference positionRef = InventoryHelper.GetUserPosition(service, ownerID);
            if (positionRef.Id == PositonIDs.DSE || positionRef.Id == PositonIDs.Franchise)
            {

            }
            else
            {
                throw new InvalidPluginExecutionException("Owner must be in positions \"DSE\", \"Franchise\".");
            }
        }
        public static string GetAccountNumber(IOrganizationService service, EntityReference ownerID)
        {
            string autoNumber;
            try
            {
                QueryExpression query = new QueryExpression("account");
                query.ColumnSet = new ColumnSet("accountnumber");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, ownerID.Id);
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count != 1)
                {
                    throw new InvalidPluginExecutionException("No or More then one Channel Partner found for owner.");
                }
                else
                {
                    autoNumber = _entitys[0].GetAttributeValue<string>("accountnumber");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Creation in Auto Number " + ex.Message);
            }
            return autoNumber;
        }
        public static void CheckDuplicateWareHouse(IOrganizationService service, Entity warehouse)
        {
            EntityCollection _entitys = new EntityCollection();
            try
            {
                QueryExpression query = new QueryExpression("msdyn_warehouse");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                if (warehouse.Contains("hil_warehousetype"))
                    query.Criteria.AddCondition("hil_warehousetype", ConditionOperator.Equal, warehouse.GetAttributeValue<OptionSetValue>("hil_warehousetype").Value);
                if (warehouse.Contains("ownerid"))
                    query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, warehouse.GetAttributeValue<EntityReference>("ownerid").Id);
                _entitys = service.RetrieveMultiple(query);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Duplicate detection " + ex.Message);
            }
            if (_entitys.Entities.Count > 1)
            {
                throw new InvalidPluginExecutionException("Duplicate detection more than one warehouse is exist for this warehousesetup.");
            }
        }
    }
}
