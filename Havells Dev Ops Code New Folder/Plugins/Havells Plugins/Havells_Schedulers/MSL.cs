using System;
using System.Net;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.ServiceModel.Description;
using Havells_Plugin;

namespace Havells_Schedulers
{
    public static class MSLPOScheduler
    {
        public static void RunMSLScheduler(IOrganizationService service)
        {
            try
            {
                DateTime UtcNow = DateTime.Now;
                QueryExpression query = new QueryExpression(Account.EntityLogicalName);
                query.Criteria = new FilterExpression(LogicalOperator.Or);
                query.Criteria.AddCondition(new ConditionExpression("customertypecode", ConditionOperator.Equal, 5));
                query.Criteria.AddCondition(new ConditionExpression("customertypecode", ConditionOperator.Equal, 6));
                query.Criteria.AddCondition(new ConditionExpression("customertypecode", ConditionOperator.Equal, 9));
                query.ColumnSet = new ColumnSet(true);
                EntityCollection Found = service.RetrieveMultiple(query);
                foreach (Account Acc in Found.Entities)
                {
                    int Day1;
                    int Day2;
                    if (Acc.Contains("hil_schedule1") && Acc.Contains("hil_schedule2"))
                    {
                        Day1 = (int)Acc["hil_schedule1"];
                        Day2 = (int)Acc["hil_schedule2"];
                        if ((UtcNow.Day == Day1) || (UtcNow.Day == Day2))
                        {
                            //CheckMSL(service, Acc);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        //public static void CheckMSL(IOrganizationService service, Account Acc)
        //{

        //}

        //account-> MSL
        //-> division wise group


        public static void CheckStocksInInventory(IOrganizationService service, Account Acc)
        {
            try
            {
                QueryByAttribute Query = new QueryByAttribute(hil_inventory.EntityLogicalName);//Check Names (Not Confirmed)
                Query.ColumnSet = new ColumnSet(true);
                Query.AddAttributeValue("hil_owneraccount", Acc.AccountId);//Check Names (Not Confirmed)
                EntityCollection Found = service.RetrieveMultiple(Query);
                foreach (hil_inventory inv in Found.Entities)
                {
                    EntityReference Pdt = new EntityReference();
                    int Qty;
                    if ((inv.Contains("hil_part")) && (inv.Contains("hil_availableqty")))
                    {
                        Pdt = (EntityReference)inv.hil_Part;
                        Qty = (int)inv.hil_AvailableQty;
                        IfPartsRequired(service, Acc, Pdt, Qty);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        public static void IfPartsRequired(IOrganizationService service, Account Acc, EntityReference Pdt, int Qty)
        {
            try
            {
                int PoType;
                QueryExpression query = new QueryExpression(hil_minimumstocklevel.EntityLogicalName);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_account", ConditionOperator.Equal, Acc.AccountId));
                query.Criteria.AddCondition(new ConditionExpression("hil_sparepart", ConditionOperator.Equal, Pdt.Id));
                query.ColumnSet = new ColumnSet(true);
                //QueryExpression Query = new QueryExpression();
                //Query.EntityName = hil_minimumstocklevel.EntityLogicalName;
                //ConditionExpression condition1 = new ConditionExpression("hil_accountid", ConditionOperator.Equal, Acc.AccountId);
                //ConditionExpression condition4 = new ConditionExpression("productid", ConditionOperator.Equal, Pdt.Id);
                //FilterExpression filter1 = new FilterExpression(LogicalOperator.Or);
                //Query.ColumnSet = new ColumnSet(true);
                EntityCollection Found = service.RetrieveMultiple(query);
                foreach (hil_minimumstocklevel Stk in Found.Entities)
                {
                    int MSL;
                    int Required = 0;
                    string sCustomerSAPCOde = string.Empty;
                    OptionSetValue AccountType = new OptionSetValue();
                    OptionSetValue opWarrantyStatus = new OptionSetValue(1);//In warranty
                    if (Acc.Contains("customertypecode"))
                    {
                        AccountType = (OptionSetValue)Acc.CustomerTypeCode;
                    }
                    if (Acc.Contains("hil_inwarrantycustomersapcode"))
                    {
                        sCustomerSAPCOde = (string)Acc.hil_InWarrantyCustomerSAPCode;
                    }

                    string sDistributionChannel = Havells_Plugin.HelperPO.getDistributionChannel(new OptionSetValue(1), AccountType, service);
                    if (Stk.Contains("hil_mslquantity"))
                    {
                        MSL = (int)Stk.hil_MSLQuantity;
                        if (MSL > Qty)
                        {
                            Required = MSL - Qty;
                            PoType = 910590000;
                            //Guid PO = Havells_Plugin.HelperPO.CreatePO(service, Required, Acc.Id, Pdt.Id, PoType,AccountType, sCustomerSAPCOde, sDistributionChannel, opWarrantyStatus);
                            //CreatePO(IOrganizationService service, int iQuantity, Guid fsAccountId, Guid fsPartId, int PoType, OptionSetValue opCategory, String sCustomerSAPCOde, String sDistributionChannel, OptionSetValue opWarrantyType);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}