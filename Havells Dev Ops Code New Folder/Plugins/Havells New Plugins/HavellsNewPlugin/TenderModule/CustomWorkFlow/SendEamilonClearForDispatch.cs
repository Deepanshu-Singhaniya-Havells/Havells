using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Crm.Sdk.Messages;
using HavellsNewPlugin.TenderModule.MailtoTeams;

namespace HavellsNewPlugin.TenderModule.CustomWorkFlow
{
    public class SendEamilonClearForDispatch : CodeActivity
    {
        public static ITracingService tracingService = null;

        [RequiredArgument]
        [Input("MailBody")]

        public InArgument<String> mailBody { get; set; }

        [RequiredArgument]
        [Input("MailSubject")]

        public InArgument<String> mailsubject { get; set; }

        [RequiredArgument]
        [Input("To")]
        public InArgument<String> to { get; set; }

        [Input("CC")]
        public InArgument<String> cc { get; set; }

        [RequiredArgument]
        [Input("Regarding")]
        [ReferenceTarget("hil_oaheader")]
        public InArgument<EntityReference> regarding { get; set; }

        [RequiredArgument]
        [Input("From")]
        [ReferenceTarget("queue")]
        public InArgument<EntityReference> from { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            try
            {
                var context = executionContext.GetExtension<IWorkflowContext>();
                var serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                tracingService = executionContext.GetExtension<ITracingService>();
                string to = this.to.Get((ActivityContext)executionContext);
                string cc = this.cc.Get((ActivityContext)executionContext);


                EntityReference fromRef = this.from.Get((ActivityContext)executionContext);
                string mailsubject = this.mailsubject.Get((ActivityContext)executionContext);
                string mailBody = this.mailBody.Get((ActivityContext)executionContext);
                EntityReference regardingRef = this.regarding.Get((ActivityContext)executionContext);

                Entity _regardingEntity = service.Retrieve(regardingRef.LogicalName, regardingRef.Id,
                    new ColumnSet(false));

                QueryExpression tenderURL = new QueryExpression("hil_integrationconfiguration");
                tenderURL.ColumnSet = new ColumnSet("hil_url");
                tenderURL.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "TenderAppURL");
                EntityCollection tenderURLCol = service.RetrieveMultiple(tenderURL);
                string URL = tenderURLCol.Entities[0].GetAttributeValue<string>("hil_url");

                string recordURL = URL + _regardingEntity.LogicalName + "&id=" + _regardingEntity.Id;
                tracingService.Trace("recordURL " + recordURL);
                string clickhereContent = "For more details please &nbsp; <a target='_blank' href=" + recordURL + "> Click Here </a>";
                if (mailBody.Contains("#Clickhere"))
                {
                    tracingService.Trace("_regardingEntity.Id " + _regardingEntity.Id);
                    QueryExpression query = new QueryExpression("hil_attachment");
                    query.ColumnSet = new ColumnSet(true);
                    query.Distinct = true;
                    query.Criteria.AddCondition("hil_isdeleted", ConditionOperator.Equal, false);
                    query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, _regardingEntity.Id);
                    EntityCollection oaheaderCol = service.RetrieveMultiple(query);
                    tracingService.Trace("oaheaderCol.Entities.Count  " + oaheaderCol.Entities.Count);
                    if (oaheaderCol.Entities.Count > 0)
                    {
                        mailBody = mailBody.Replace("#Clickhere", clickhereContent);
                    }
                    else
                    {
                        mailBody = mailBody.Replace("#Clickhere", "");
                    }
                }
                SendEmailtoteamsonOA.sendEmailonEventBasis(service, regardingRef, to, cc, mailBody, mailsubject, fromRef, tracingService);
            }
            catch (Exception ex)
            {
                throw new InvalidWorkflowException("Error :- " + ex.Message);
            }
        }
    }
}
