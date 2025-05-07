using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace HavellsNewPlugin.Case
{
    public class PostCreatePhoneCall : IPlugin
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
                  && context.PrimaryEntityName.ToLower() == "phonecall" && context.Depth == 1 && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("regardingobjectid"))
                    {
                        EntityReference _regardingEntity = entity.GetAttributeValue<EntityReference>("regardingobjectid");
                        if (_regardingEntity.LogicalName == "incident") //Case
                        {
                            Guid CaseOwnerId = service.Retrieve(_regardingEntity.LogicalName, _regardingEntity.Id,
                                new ColumnSet("ownerid")).GetAttributeValue<EntityReference>("ownerid").Id;

                            if (CaseOwnerId == context.InitiatingUserId) //If Owner is Same
                            {
                                //Updatting Case First Response Time
                                Entity _entCase = service.Retrieve(_regardingEntity.LogicalName, _regardingEntity.Id, new ColumnSet("firstresponsesent"));

                                if (!_entCase.Contains("firstresponsesent") || !_entCase.GetAttributeValue<bool>("firstresponsesent"))
                                {
                                    Entity _case = new Entity(_regardingEntity.LogicalName, _regardingEntity.Id);
                                    _case["firstresponsesent"] = true;
                                    _case["hil_firstresponsesenton"] = DateTime.Now.AddMinutes(330);
                                    service.Update(_case);
                                }

                                QueryExpression query = new QueryExpression("hil_grievancehandlingactivity");
                                query.ColumnSet = new ColumnSet("actualstart", "createdon");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, _regardingEntity.Id);
                                query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, CaseOwnerId);
                                query.Criteria.AddCondition("actualstart", ConditionOperator.Null);
                                query.AddOrder("createdon", OrderType.Ascending);
                                query.TopCount = 1;
                                EntityCollection entCallGHA = service.RetrieveMultiple(query);

                                if (entCallGHA.Entities.Count > 0)
                                {
                                    Entity entGHA = new Entity(entCallGHA.Entities[0].LogicalName, entCallGHA.Entities[0].Id);
                                    DateTime _firstResponse = DateTime.Now.AddMinutes(330);
                                    entGHA["actualstart"] = _firstResponse;
                                    service.Update(entGHA);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
    }
}
