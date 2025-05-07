using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Services;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Case
{
    public class CasePreCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                  && context.PrimaryEntityName.ToLower() == "incident" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    string fieldName = string.Empty;
                    string caseNumber = AutoNumber(service, entity, out fieldName);
                    if (fieldName != string.Empty)
                    {
                        Guid _departmentId = entity.GetAttributeValue<EntityReference>("hil_casedepartment").Id;
                        if (_departmentId == CaseConstaints._samparkDepartment)
                            entity[fieldName] = caseNumber;
                    }

                    if (!entity.Contains("hil_casedepartment")) {
                        entity["hil_casedepartment"] = new EntityReference("hil_casedepartment", CaseConstaints._samparkDepartment); //Sampark Call Center
                    }
                    if (!entity.Contains("caseorigincode"))
                    {
                        entity["caseorigincode"] = new OptionSetValue(1); //Phone Call
                    }
                    if (entity.Contains("hil_job"))
                    {
                        ColumnSet set = new ColumnSet("createdon");
                        Entity job = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_job").LogicalName, entity.GetAttributeValue<EntityReference>("hil_job").Id, new ColumnSet("createdon"));
                        DateTime jobcreatedOn = job.GetAttributeValue<DateTime>("createdon");
                        DateTime today = DateTime.Today;
                        int age = today.Subtract(jobcreatedOn).Days;
                        entity["hil_jobage"] = age;
                    }

                    if (entity.Contains("hil_pincode")) {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' top='1' mapping='logical' distinct='false'>
                          <entity name='hil_businessmapping'>
                            <attribute name='hil_branch' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_pincode' operator='eq' value='{entity.GetAttributeValue<EntityReference>("hil_pincode").Id}' />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection _entBranch = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entBranch.Entities.Count > 0) {
                            entity["hil_branch"] = _entBranch.Entities[0].GetAttributeValue<EntityReference>("hil_branch");
                        }
                    }

                    if (entity.Contains("customerid"))
                    {
                        if (entity.GetAttributeValue<EntityReference>("customerid").LogicalName == "account")
                        {
                            Entity account = service.Retrieve("account", entity.GetAttributeValue<EntityReference>("customerid").Id,
                                new ColumnSet("hil_salesoffice", "telephone1", "hil_branch"));
                            if (account.Contains("telephone1"))
                                entity["hil_mobileno"] = account["telephone1"];
                            if (account.Contains("hil_salesoffice"))
                                entity["hil_salesoffice"] = account["hil_salesoffice"];
                            if (account.Contains("hil_branch"))
                                entity["hil_branch"] = account["hil_branch"];
                        }
                        else {
                            Entity contact = service.Retrieve("contact", entity.GetAttributeValue<EntityReference>("customerid").Id,
                                    new ColumnSet("mobilephone"));
                            if (contact.Contains("mobilephone"))
                                entity["hil_mobileno"] = contact["mobilephone"];
                        }
                    }
                    else
                    {
                        entity["customerid"] = new EntityReference("contact", new Guid("fbb613b0-75e2-ed11-8847-6045bdac51bc"));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Case Pre Create Error : " + ex.Message);
            }
        }
        string AutoNumber(IOrganizationService service, Entity entity, out string fieldName)
        {
            //try
            //{
            DateTime Today = DateTime.Now;
            string prefix = Today.Day.ToString().PadLeft(2, '0') + Today.Month.ToString().PadLeft(2, '0') + Today.Year.ToString().PadLeft(4, '0');

            //Retrive Config
            QueryExpression query = new QueryExpression("plt_idgenerator");
            query.ColumnSet = new ColumnSet(false);
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("plt_idgen_name", ConditionOperator.Equal, entity.LogicalName);
            EntityCollection ecAuto = service.RetrieveMultiple(query);
            Entity entAuto = ecAuto[0];

            //Apply Lock
            Entity couterTable = new Entity(entAuto.LogicalName, entAuto.Id);
            couterTable.Attributes["plt_idgen_prefix"] = "lock " + DateTime.Now;
            service.Update(couterTable);

            Entity AutoPost = service.Retrieve(entAuto.LogicalName, entAuto.Id, new ColumnSet("plt_attributename",
                "plt_idgen_fixednumbersize", "plt_idgen_nextnumber", "plt_idgen_incrementby", "plt_idgen_zeropad"));
            int currentrecordcounternumber = AutoPost.GetAttributeValue<int>("plt_idgen_nextnumber");

            QueryExpression Query = new QueryExpression(entity.LogicalName);
            Query.ColumnSet = new ColumnSet("createdon");
            Query.TopCount = 1;
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection entColl = service.RetrieveMultiple(Query);

            int lastCreatedYear = entColl[0].GetAttributeValue<DateTime>("createdon").Year;
            int currentYear = DateTime.Now.Year;
            if (lastCreatedYear < currentYear)
            {
                currentrecordcounternumber = 1;
            }
            string _runningNumber = string.Empty;
            int fixedNumberSize = AutoPost.GetAttributeValue<int>("plt_idgen_fixednumbersize");
            int incrementby = AutoPost.GetAttributeValue<int>("plt_idgen_incrementby");

            if (AutoPost.GetAttributeValue<bool>("plt_idgen_zeropad"))
            {
                _runningNumber = currentrecordcounternumber.ToString().PadLeft(fixedNumberSize, '0');
                tracingService.Trace("_runningNumber fixedNumberSize  " + _runningNumber + "||" + fixedNumberSize);
            }
            else
            {
                _runningNumber = currentrecordcounternumber.ToString();
                tracingService.Trace("_runningNumber " + _runningNumber);
            }
            currentrecordcounternumber = currentrecordcounternumber + incrementby;
            fieldName = AutoPost.GetAttributeValue<string>("plt_attributename");
            //entity[entColl[0].GetAttributeValue<string>("plt_attributename")] = prefix + _runningNumber;
            tracingService.Trace("fieldName " + fieldName);
            //update the config
            Entity newudpateconfig = new Entity(entAuto.LogicalName, entAuto.Id);
            newudpateconfig["plt_idgen_nextnumber"] = currentrecordcounternumber;
            service.Update(newudpateconfig);
            return prefix + _runningNumber;
            //}
            //catch (Exception ex)
            //{
            //    throw new InvalidPluginExecutionException("Case AutoNumber Error : " + ex.Message);
            //}
        }
    }
}
