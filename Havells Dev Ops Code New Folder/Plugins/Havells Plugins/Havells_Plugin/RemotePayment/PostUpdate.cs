using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;
using Havells_Plugin;

namespace Havells_Plugin.RemotePayment
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder"
                    && context.MessageName.ToUpper() == "UPDATE")
            {
                try
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Attributes.Contains("hil_resendpaymentlink"))
                    {
                        if (entity.GetAttributeValue<bool>("hil_resendpaymentlink"))
                        {
                            QueryExpression _query = new QueryExpression("hil_paymentstatus");
                            _query.ColumnSet = new ColumnSet("hil_url", "hil_phone");
                            _query.Criteria = new FilterExpression(LogicalOperator.And);
                            _query.Criteria.AddCondition("hil_job", ConditionOperator.Equal, entity.Id);
                            Entity _paymentEntity = service.RetrieveMultiple(_query)[0];
                            if (_paymentEntity != null)
                            {
                                Entity _jobEntity = service.Retrieve("msdyn_workorder", entity.Id, new ColumnSet("hil_customerref", "hil_receiptamount"));
                                Decimal _receiptAmt = Math.Round(_jobEntity.GetAttributeValue<Money>("hil_receiptamount").Value, 2);

                                String _msg = "Dear Customer, Please use the below link to process payment of amount INR " +
                                    _receiptAmt.ToString() + ". " + _paymentEntity.GetAttributeValue<String>("hil_url")
                                    + " - Havells";
                                Entity _sms = new Entity("hil_smsconfiguration");
                                _sms["hil_contact"] = _jobEntity.GetAttributeValue<EntityReference>("hil_customerref");
                                _sms["hil_smstemplate"] = new EntityReference("hil_smstemplates", new Guid("7cb6ee6b-cdac-eb11-8236-0022486eccb7"));
                                _sms["subject"] = "Resend Remort Pay SMS";
                                _sms["hil_message"] = _msg;
                                _sms["hil_mobilenumber"] = _paymentEntity.GetAttributeValue<String>("hil_phone");
                                _sms["regardingobjectid"] = new EntityReference("msdyn_workorder", entity.Id);
                                _sms["hil_direction"] = new OptionSetValue(2);
                                service.Create(_sms);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("Error :" + ex.Message);
                }
            }
        }
    }
}
