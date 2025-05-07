using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.NatureOfComplaint
{
    public class ClsNatureOfComplaint : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string JsonResponse = "";
            string SerialNumber = string.Empty;
            List<NatureOfComplaint> lstNatureOfComplaint = new List<NatureOfComplaint>();

            if (context.InputParameters.Contains("SerialNumber") && context.InputParameters["SerialNumber"] is string)
            {
                if (string.IsNullOrWhiteSpace(Convert.ToString(context.InputParameters["SerialNumber"])))
                {
                    lstNatureOfComplaint.Add(new NatureOfComplaint
                    {
                        ResultStatus = false,
                        ResultMessage = "Product Serial Number is required."
                    });
                    JsonResponse = JsonSerializer.Serialize(lstNatureOfComplaint);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                SerialNumber = context.InputParameters["SerialNumber"].ToString();
                JsonResponse = JsonSerializer.Serialize(GetNatureOfComplaints(service, SerialNumber));
                context.OutputParameters["data"] = JsonResponse;
            }
        }
        public List<NatureOfComplaint> GetNatureOfComplaints(IOrganizationService service, string SerialNumber)
        {
            NatureOfComplaint objNatureOfComplaint;
            List<NatureOfComplaint> lstNatureOfComplaint = new List<NatureOfComplaint>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                Query = new QueryExpression("msdyn_customerasset");
                Query.ColumnSet = new ColumnSet("msdyn_name");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, SerialNumber);
                Query.TopCount = 1;
                entcoll = service.RetrieveMultiple(Query);

                if (entcoll.Entities.Count == 0)
                {
                    objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = "Product Serial Number does not exist." };
                    lstNatureOfComplaint.Add(objNatureOfComplaint);
                    return lstNatureOfComplaint;
                }
                else
                {
                    string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
               <entity name='hil_natureofcomplaint'>
                   <attribute name='hil_name' />
                   <attribute name='hil_natureofcomplaintid' />
                   <order attribute='hil_name' descending='false' />
                   <filter type='and'>
                     <condition attribute='statecode' operator='eq' value='0' />
                   </filter>
                   <link-entity name='product' from='productid' to='hil_relatedproduct' link-type='inner' alias='ae'>
                       <link-entity name='msdyn_customerasset' from='hil_productsubcategory' to='productid' link-type='inner' alias='af'>
                           <filter type='and'>
                               <condition attribute='msdyn_name' operator='eq' value='" + SerialNumber + @"' />
                           </filter>
                       </link-entity>
                   </link-entity>
               </entity>
               </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = "No Nature of Complaint is mapped with Serial Number." };
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                    }
                    else
                    {
                        foreach (Entity ent in entcoll.Entities)
                        {
                            objNatureOfComplaint = new NatureOfComplaint();
                            objNatureOfComplaint.Guid = ent.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                            objNatureOfComplaint.Name = ent.GetAttributeValue<string>("hil_name");
                            objNatureOfComplaint.SerialNumber = SerialNumber;
                            objNatureOfComplaint.ResultStatus = true;
                            objNatureOfComplaint.ResultMessage = "Success";
                            lstNatureOfComplaint.Add(objNatureOfComplaint);
                        }
                    }
                    return lstNatureOfComplaint;
                }
            }
            catch (Exception ex)
            {
                objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = ex.Message };
                lstNatureOfComplaint.Add(objNatureOfComplaint);
                return lstNatureOfComplaint;
            }
        }
    }
    public class NatureOfComplaint
    {
        public string SerialNumber { get; set; }

        public string Name { get; set; }

        public string ProductCategoryName { get; set; }

        public Guid Guid { get; set; }

        public Guid ProductCategoryGuid { get; set; }

        public bool ResultStatus { get; set; }

        public string ResultMessage { get; set; }

    }

}
