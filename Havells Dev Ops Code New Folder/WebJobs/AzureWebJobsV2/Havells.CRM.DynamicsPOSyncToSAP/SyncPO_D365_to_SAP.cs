using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DynamicsPOSyncToSAP
{
    public class SyncPO_D365_to_SAP
    {
        public static void GetProductRequisition(IOrganizationService service)
        {
            try
            {
                string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                      <entity name=""msdyn_purchaseorder"">
                                        <attribute name=""msdyn_purchaseorderid"" />
                                        <attribute name=""msdyn_name"" />
                                        <attribute name=""hil_ivusage"" />
                                        <attribute name=""msdyn_workorder"" />
                                        <attribute name=""msdyn_purchaseorderdate"" />
                                        <order attribute=""msdyn_name"" descending=""false"" />
                                        <filter type=""and"">
                                          <condition attribute=""msdyn_approvalstatus"" operator=""eq"" value=""690970000"" />
                                          <condition attribute=""hil_ivusage"" operator=""not-null"" />
                                          <condition attribute=""hil_sapsalesorderno"" operator=""null"" />
                                          <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                        </filter>
                                        <link-entity name=""account"" from=""accountid"" to=""msdyn_vendor"" visible=""false"" link-type=""outer"" alias=""account"">
                                          <attribute name=""hil_outwarrantycustomersapcode"" />
                                          <attribute name=""hil_inwarrantycustomersapcode"" />
                                        </link-entity>
                                        <link-entity name=""msdyn_workorder"" from=""msdyn_workorderid"" to=""msdyn_workorder"" visible=""false"" link-type=""outer"" alias=""job"">
                                          <attribute name=""hil_productcategory"" />
                                          <attribute name=""hil_warrantystatus"" />
                                        </link-entity>
                                      </entity>
                                    </fetch>";
                IntegrationConfiguration integrationConfiguration = HelperClass.GetIntegrationConfiguration(service, "CreateSalesOrder");

                PO_SAP_Request SAPrequest = new PO_SAP_Request();
                EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetch));
                String DivisionCode = null;
                EntityReference jobRef = null;
                EntityReference productCat = null;
                OptionSetValue warrantyStatus = null;
                foreach (Entity entity in collection.Entities)
                {
                    try
                    {
                        DivisionCode = string.Empty;
                        jobRef = entity.GetAttributeValue<EntityReference>("msdyn_workorder");
                        productCat = (EntityReference)entity.GetAttributeValue<AliasedValue>("job.hil_productcategory").Value;
                        warrantyStatus = ((OptionSetValue)entity.GetAttributeValue<AliasedValue>("job.hil_warrantystatus").Value);
                        DivisionCode = service.Retrieve(productCat.LogicalName, productCat.Id, new ColumnSet("hil_sapcode")).GetAttributeValue<string>("hil_sapcode");
                        Parent iParent = new Parent();
                        iParent.SPART = DivisionCode;
                        iParent.BSTKD = (string)entity["msdyn_name"];
                        iParent.BSTDK = entity.GetAttributeValue<DateTime>("msdyn_purchaseorderdate").AddMinutes(330).ToString("yyyy-MM-dd");
                        iParent.KUNNR = (string)entity["hil_ivusage"] == "EM" ? (string)entity.GetAttributeValue<AliasedValue>("account.hil_inwarrantycustomersapcode").Value : (string)entity["account.hil_outwarrantycustomersapcode"];
                        iParent.ABRVW = (string)entity["hil_ivusage"];
                        iParent.VKORG = "HIL";

                        if (((string)entity["hil_ivusage"] == "EM") || iParent.KUNNR.StartsWith("E"))
                        {
                            iParent.AUART = "ZRS4";
                        }
                        else if ((warrantyStatus.Value == (int)WarrantyStatus.Out_Warranty || warrantyStatus.Value == (int)WarrantyStatus.Warranty_Void || warrantyStatus.Value == (int)WarrantyStatus.NA_for_Warranty))
                        {
                            iParent.AUART = "ZRS4";
                        }
                        // OutWarranty C Code ZSW6
                        if (iParent.KUNNR.StartsWith("E"))
                        {
                            iParent.VTWEG = getDistributionChannel(warrantyStatus, new OptionSetValue(9), service);// PRLine.hil_DistributionChannel;
                        }
                        else
                        {
                            if (iParent.KUNNR.StartsWith("C"))
                            {
                                iParent.VTWEG = getDistributionChannel(new OptionSetValue(2), new OptionSetValue(6), service);// PRLine.hil_DistributionChannel;
                            }
                            else if (iParent.KUNNR.StartsWith("F"))
                            {
                                iParent.VTWEG = getDistributionChannel(new OptionSetValue(1), new OptionSetValue(6), service);// PRLine.hil_DistributionChannel;
                            }
                        }
                        fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""msdyn_purchaseorderproduct"">
                                    <attribute name=""createdon"" />
                                    <attribute name=""msdyn_quantity"" />
                                    <attribute name=""msdyn_purchaseorder"" />
                                    <attribute name=""msdyn_product"" />
                                    <attribute name=""msdyn_purchaseorderproductid"" />
                                    <order attribute=""msdyn_purchaseorder"" descending=""true"" />
                                    <filter type=""and"">
                                      <condition attribute=""msdyn_purchaseorder"" operator=""eq"" value=""{entity.Id}"" />
                                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                    </filter>
                                  </entity>
                                </fetch>";
                        EntityCollection _POProductColl = service.RetrieveMultiple(new FetchExpression(fetch));
                        List<IT_DATA> childs = new List<IT_DATA>();
                        foreach (Entity POProduct in _POProductColl.Entities)
                        {
                            IT_DATA iChild = new IT_DATA();
                            iChild.MATNR = POProduct.GetAttributeValue<EntityReference>("msdyn_product").Name;
                            iChild.SPART = DivisionCode;
                            iChild.DZMENG = Convert.ToInt32(POProduct.GetAttributeValue<double>("msdyn_quantity"));
                            childs.Add(iChild);
                        }
                        SAPrequest.IM_PROJECT = "D365";
                        SAPrequest.IM_HEADER = iParent;
                        SAPrequest.LT_LINE_ITEM = childs;
                        var Json = JsonConvert.SerializeObject(SAPrequest);

                        Response resp = JsonConvert.DeserializeObject<Response>(HelperClass.CallAPI(integrationConfiguration,Json,"POST"));
                        Console.WriteLine(resp.EX_SALESDOC_NO + " - " + resp.RETURN + " - " + (string)entity["msdyn_name"]);
                        if (resp.EX_SALESDOC_NO != "")
                        {
                            Entity iUpdateHeader = new Entity(entity.LogicalName, entity.Id);
                            iUpdateHeader["hil_syncmessage"] = resp.RETURN;
                            iUpdateHeader["hil_sapsalesorderno"] = resp.EX_SALESDOC_NO.ToString();
                            iUpdateHeader["msdyn_systemstatus"] = new OptionSetValue(690970001);    //submitted
                            service.Update(iUpdateHeader);
                        }
                        else
                        {
                            Entity iUpdateHeader = new Entity(entity.LogicalName, entity.Id);
                            iUpdateHeader["hil_syncmessage"] = resp.RETURN;
                            iUpdateHeader["hil_sapsalesorderno"] = resp.EX_SALESDOC_NO.ToString();
                            service.Update(iUpdateHeader);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error in PO " + (string)entity["msdyn_name"] + " : " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        static String getDistributionChannel(OptionSetValue opWarrantyStatus, OptionSetValue opAccountType, IOrganizationService service)
        {
            String result = String.Empty;
            try
            {
                QueryExpression query = new QueryExpression("hil_integrationconfiguration");
                query.ColumnSet = new ColumnSet("new_distributionchannel");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("new_warrantystatus", ConditionOperator.Equal, opWarrantyStatus.Value)); //Approval
                query.Criteria.AddCondition(new ConditionExpression("new_customertype", ConditionOperator.Equal, opAccountType.Value));
                query.AddOrder("modifiedon", OrderType.Descending);
                EntityCollection collection = service.RetrieveMultiple(query);

                foreach (Entity iobj in collection.Entities)
                {
                    if (iobj.Contains("new_distributionchannel"))
                    {
                        result = (string)iobj["new_distributionchannel"];
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Get Distribution Channel Error : " + ex.Message);
            }
            return result;
        }

    }
}
