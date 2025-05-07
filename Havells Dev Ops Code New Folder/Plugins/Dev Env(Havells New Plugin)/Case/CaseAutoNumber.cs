using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Case
{
    internal class CaseAutoNumber : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "incident" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
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
                    int fixedNumberSize = entColl[0].GetAttributeValue<int>("plt_idgen_fixednumbersize");
                    int incrementby = entColl[0].GetAttributeValue<int>("plt_idgen_incrementby");

                    if (entColl[0].GetAttributeValue<bool>("plt_idgen_zeropad"))
                    {
                        _runningNumber = currentrecordcounternumber.ToString().PadLeft(fixedNumberSize);
                    }
                    else
                    {
                        _runningNumber = currentrecordcounternumber.ToString();
                    }
                    currentrecordcounternumber = currentrecordcounternumber + incrementby;

                    entity[entColl[0].GetAttributeValue<string>("plt_attributename")] = prefix + _runningNumber;

                    //update the config
                    Entity newudpateconfig = new Entity(entAuto.LogicalName, entAuto.Id);
                    newudpateconfig["plt_idgen_nextnumber"] = currentrecordcounternumber;
                    service.Update(newudpateconfig);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
    }
}