using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Linq;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
namespace Havells_Plugin.Notes
{
    public class PostCreate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == Annotation.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    tracingService.Trace("1");
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity ent = (Entity)context.InputParameters["Target"];
                    sendTCRPDFinemail(ent.Id,service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error at Havells_Plugin.Notes.PostCreate.Execute " + ex.Message);
            }
        }


        public static void sendTCRPDFinemail(Guid notesid, IOrganizationService service)
        {
            try
            {
                tracingService.Trace("2");
                OrganizationServiceContext orgContext = new OrganizationServiceContext(service);

                var obj = from _Notes in orgContext.CreateQuery<Annotation>()
                           join _Job in orgContext.CreateQuery<msdyn_workorder>()
                           on _Notes.ObjectId.Id equals _Job.Id
                           where (_Notes.Subject.Contains("TCR for Job") || _Notes.Subject.Contains("Job Estimate"))
                           && (_Notes.AnnotationId.Value==notesid) && (_Notes.Contains("mimetype") && (_Notes.MimeType != null) && (_Notes.MimeType == "application / pdf")
                           && (_Job.hil_SendTCR != null) && (_Job.hil_SendTCR == false))
                           orderby _Notes.CreatedOn descending
                           select new
                           {
                               _Notes,
                               _Notes.ObjectId,
                               _Job.hil_CustomerRef,
                               _Job.msdyn_workorderId,
                               _Job.hil_Email,
                               _Job.msdyn_name
                           };
                foreach (var iobj in obj)
                {
                    string toemail = iobj.hil_Email;
                    tracingService.Trace("3");

                    List<Entity> _toList = new List<Entity>();
                    Entity entityUser = new Entity("activityparty");
                    if (iobj.hil_CustomerRef != null && iobj.hil_CustomerRef.LogicalName == Contact.EntityLogicalName)
                    {
                        Contact contRef = (Contact)service.Retrieve(Contact.EntityLogicalName, iobj.hil_CustomerRef.Id, new ColumnSet("emailaddress1"));
                        if(contRef.EMailAddress1 != null)
                        {
                            Entity Toparty = new Entity("activityparty");
                            Toparty["partyid"] = new EntityReference(iobj.hil_CustomerRef.LogicalName, iobj.hil_CustomerRef.Id);

                            tracingService.Trace("4");
                            OrganizationRequest orgRequest = new OrganizationRequest("hil_SendemailActionTCR");
                            //Passing the dynamic link of email crated as a parameter to action - Vishal
                            orgRequest["Jobnumber"] = iobj.msdyn_name;
                            OrganizationResponse orgResponse = (OrganizationResponse)service.Execute(orgRequest);
                            if (orgResponse.Results != null)
                            {
                                tracingService.Trace("5");
                                //Get the email reference from action which is returned from action - Vishal
                                EntityReference emailRef = (EntityReference)orgResponse.Results["email"];
                                if (emailRef.Id != Guid.Empty)
                                {
                                    tracingService.Trace("6");
                                    Entity email = service.Retrieve(emailRef.LogicalName, emailRef.Id, new ColumnSet("to", "from"));
                                    if (email.Id != Guid.Empty)
                                    {
                                        addattachments(email.Id, iobj._Notes, service);
                                        tracingService.Trace("7");
                                        SendEmailAction(email, service, Toparty);
                                        tracingService.Trace("8");
                                        #region MarkSend_TCR_Yes
                                        msdyn_workorder upJob = new msdyn_workorder();
                                        upJob.Id = iobj.msdyn_workorderId.Value;
                                        upJob.hil_SendTCR = true;
                                        service.Update(upJob);
                                        #endregion
                                        tracingService.Trace("9");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error at Havells_Plugin.Notes.PostCreate.sendTCRPDFinemail " + ex.Message);

            }

        }
        public static void SendEmailAction(Entity email, IOrganizationService service, Entity Toparty)
        {

            try
            {
                email["to"] = new Entity[] { Toparty };
                service.Update(email);
                SendEmailRequest sendEmailreq = new SendEmailRequest
                {
                    EmailId = email.Id,
                    TrackingToken = "",
                    IssueSend = true
                };
                SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailreq);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error at Havells_Plugin.Notes.PostCreate.SendEmailAction " + ex.Message);
            }
        }
        public static void addattachments(Guid emailguid, Annotation notes, IOrganizationService service)
        {
            Entity attachment = new Entity("activitymimeattachment");
            attachment["subject"] = notes.Subject + " act";
            attachment["filename"] = notes.FileName;

            attachment["body"] = notes.DocumentBody;
            attachment["mimetype"] = "application/pdf";
            attachment["attachmentnumber"] = 1;
            attachment["objectid"] = new EntityReference("email", emailguid);
            attachment["objecttypecode"] = "email";
            service.Create(attachment);

        //    service.Delete(notes.LogicalName, notes.Id);
        }


    }

}
