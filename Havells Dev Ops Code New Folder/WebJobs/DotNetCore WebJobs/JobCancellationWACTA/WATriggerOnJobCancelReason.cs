using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Text.RegularExpressions;


namespace JobCancellationWACTA
{
    public class WATriggerOnJobCancelReason
    {
        private readonly ServiceClient service;
        public WATriggerOnJobCancelReason(ServiceClient _service)
        {
            service = _service;
        }
        public void SentDataForRequestForCancellation(string triggerName)
        {
            try
            {
                if (service.IsReady)
                {
                    #region VariableDeclaration
                    EntityCollection? entcoll = null;
                    EntityCollection? entcollWA = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    string _fetchXML = string.Empty;
                    int PageNumber = 1;
                    string DynamicWAMessageEnglish = string.Empty;
                    string DynamicWAMessageHindi = string.Empty;
                    string ProductName = string.Empty;
                    #endregion

                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' page='{PageNumber++}'>
                            <entity name='hil_wacta'>
                            <attribute name='hil_wactaid'/>
                            <attribute name='hil_watriggername'/>
                            <attribute name='createdon'/>
                            <attribute name='hil_whatsappctaresponse'/>
                            <attribute name='hil_wastatusreason'/>
                            <attribute name='hil_validupto'/>
                            <attribute name='hil_triggercount'/>
                            <attribute name='hil_jobid'/>
                            <order attribute='hil_watriggername' descending='false'/>
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_whatsappctaresponse' operator='null' />
                                <condition attribute='hil_triggercount' operator='null' />
                                <condition attribute='hil_watriggername' operator='eq' value='{triggerName}' />
                                <condition attribute='hil_wastatusreason' operator='eq' value='1' />
                            </filter>
                            </entity>
                            </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            _totalCount += entcoll.Entities.Count;
                            Root ParamModelData = null;
                            Template template = null;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    ++_rowCount;
                                    Entity Jobs = service.Retrieve("msdyn_workorder", ent.GetAttributeValue<EntityReference>("hil_jobid").Id, new ColumnSet("hil_mobilenumber", "hil_customerref", "msdyn_name", "hil_productcategory", "hil_jobcancelreason"));
                                    string? MobileNumber = Jobs.Contains("hil_mobilenumber") ? GetLast10Characters(Jobs.GetAttributeValue<string>("hil_mobilenumber")) : null;
                                    string? JobNumber = Jobs.Contains("msdyn_name") ? Jobs.GetAttributeValue<string>("msdyn_name") : null;
                                    string JobGuid = Jobs.Id.ToString();

                                    EntityReference? ProductNameInfo = Jobs.Contains("hil_productcategory") ? Jobs.GetAttributeValue<EntityReference>("hil_productcategory") : null;
                                    if (ProductNameInfo != null)
                                    {
                                        ProductName = ToUpperCamelCase(ProductNameInfo.Name);
                                    }
                                    string ActiveJobNumber = "null";

                                    EntityReference? Customer = Jobs.GetAttributeValue<EntityReference>("hil_customerref");
                                    string UserName = FixName(Customer.Name);

                                    OptionSetValue? JobCancelReason = Jobs.Contains("hil_jobcancelreason") ? Jobs.GetAttributeValue<OptionSetValue>("hil_jobcancelreason") : null;
                                    if (JobCancelReason != null && triggerName != null)
                                    {
                                        string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_jobcancellationwamapping'>
                                            <attribute name='createdon'/>
                                            <attribute name='hil_jobcancelreasonwa'/>
                                            <attribute name='hil_dynamicwamessage'/>
                                            <attribute name='hil_wacampaignname'/>
                                            <attribute name='hil_jobcancellationwamappingid'/>
                                            <attribute name='hil_dynamicwamessagehindi' />
                                            <order attribute='hil_jobcancelreasonwa' descending='false'/>
                                            <filter type='and'>
                                                <condition attribute='statecode' operator='eq' value='0'/>
                                                <condition attribute='hil_wacampaignname' operator='eq' value='{triggerName}'/>
                                                <condition attribute='hil_jobcancelreasonwa' operator='eq' value='{JobCancelReason.Value}'/>
                                            </filter>
                                            </entity>
                                            </fetch>";
                                        entcollWA = service.RetrieveMultiple(new FetchExpression(fetchXML));
                                        if (entcollWA.Entities.Count > 0)
                                        {
                                            DynamicWAMessageEnglish = entcollWA.Entities[0].Contains("hil_dynamicwamessage") ? entcollWA.Entities[0].GetAttributeValue<string>("hil_dynamicwamessage") : "";
                                            DynamicWAMessageHindi = entcollWA.Entities[0].Contains("hil_dynamicwamessage") ? entcollWA.Entities[0].GetAttributeValue<string>("hil_dynamicwamessagehindi") : "";
                                        }
                                    }
                                    Console.WriteLine($"Processing: {_rowCount}/{_totalCount} for Job#{JobNumber} Mobile Number: {MobileNumber}");
                                    Console.WriteLine($"Checking for previous WA CTA Trigger.");

                                    if (!ExistJobIdInWACTA(ent.Id) && UserName != null && MobileNumber != null)
                                    {
                                        ParamModelData = new Root();
                                        ParamModelData.phoneNumber = MobileNumber;
                                        ParamModelData.callbackData = triggerName;

                                        template = GetTemplatebyTriggerName(triggerName, JobNumber, JobGuid, UserName, ProductName, ActiveJobNumber, DynamicWAMessageEnglish, DynamicWAMessageHindi);
                                        ParamModelData.template = template;

                                        var RESULT = UpdateWhatsAppUser(ParamModelData);
                                        if (RESULT.result == "true")
                                        {
                                            Console.WriteLine($"Registered Consumer's mobile number {MobileNumber} on Whatsapp.");

                                            DateTime _validUpto = DateTime.Now.AddDays(1).Date;
                                            ent["hil_triggercount"] = 1;
                                            ent["hil_validupto"] = new DateTime(_validUpto.Year, _validUpto.Month, _validUpto.Day, 07, 30, 00);
                                            service.Update(ent);
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error: While Registering  Consumer's mobile number {MobileNumber} on Whatsapp: {RESULT.message}");
                                            ent["hil_errorlogs"] = RESULT.message;
                                            service.Update(ent);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                    ent["hil_errorlogs"] = ex.Message;
                                    service.Update(ent);
                                }
                            }
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        public void SentDataForDuplicateRequest(string triggerName)
        {
            try
            {
                if (service.IsReady)
                {
                    #region VariableDeclaration
                    EntityCollection entcoll = null;
                    EntityCollection entcollWA = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    string _fetchXML = string.Empty;
                    int PageNumber = 1;
                    string ProductName = string.Empty;
                    string DynamicWAMessageEnglish = string.Empty;
                    string DynamicWAMessageHindi = string.Empty;
                    #endregion

                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' page='{PageNumber++}'>
                            <entity name='hil_wacta'>
                            <attribute name='hil_wactaid'/>
                            <attribute name='hil_watriggername'/>
                            <attribute name='createdon'/>
                            <attribute name='hil_whatsappctaresponse'/>
                            <attribute name='hil_wastatusreason'/>
                            <attribute name='hil_validupto'/>
                            <attribute name='hil_triggercount'/>
                            <attribute name='hil_jobid'/>
                            <order attribute='hil_watriggername' descending='false'/>
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_whatsappctaresponse' operator='null' />
                                <condition attribute='hil_triggercount' operator='null' />
                                <condition attribute='hil_watriggername' operator='eq' value='{triggerName}' />
                                <condition attribute='hil_wastatusreason' operator='eq' value='1' />
                            </filter>
                            </entity>
                            </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            _totalCount += entcoll.Entities.Count;
                            Root ParamModelData = null;
                            Template template = null;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    ++_rowCount;
                                    Entity Jobs = service.Retrieve("msdyn_workorder", ent.GetAttributeValue<EntityReference>("hil_jobid").Id, new ColumnSet("hil_mobilenumber", "hil_customerref", "msdyn_name", "hil_productcategory", "hil_newjobid", "hil_webclosureremarks", "hil_jobcancelreason", "msdyn_substatus"));
                                    string? MobileNumber = Jobs.Contains("hil_mobilenumber") ? GetLast10Characters(Jobs.GetAttributeValue<string>("hil_mobilenumber")) : null;
                                    string? JobNumber = Jobs.Contains("msdyn_name") ? Jobs.GetAttributeValue<string>("msdyn_name") : null;
                                    string JobGuid = Jobs.Id.ToString();

                                    EntityReference? ProductNameInfo = Jobs.Contains("hil_productcategory") ? Jobs.GetAttributeValue<EntityReference>("hil_productcategory") : null;
                                    if (ProductNameInfo != null)
                                    {
                                        ProductName = ToUpperCamelCase(ProductNameInfo.Name);
                                    }
                                    OptionSetValue? JobCancelReason = Jobs.Contains("hil_jobcancelreason") ? Jobs.GetAttributeValue<OptionSetValue>("hil_jobcancelreason") : null;
                                    if (JobCancelReason != null && triggerName != null)
                                    {
                                        string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_jobcancellationwamapping'>
                                            <attribute name='createdon'/>
                                            <attribute name='hil_jobcancelreasonwa'/>
                                            <attribute name='hil_dynamicwamessage'/>
                                            <attribute name='hil_wacampaignname'/>
                                            <attribute name='hil_jobcancellationwamappingid'/>
                                            <attribute name='hil_dynamicwamessagehindi' />
                                            <order attribute='hil_jobcancelreasonwa' descending='false'/>
                                            <filter type='and'>
                                                <condition attribute='statecode' operator='eq' value='0'/>
                                                <condition attribute='hil_wacampaignname' operator='eq' value='{triggerName}'/>
                                                <condition attribute='hil_jobcancelreasonwa' operator='eq' value='{JobCancelReason.Value}'/>
                                            </filter>
                                            </entity>
                                            </fetch>";
                                        entcollWA = service.RetrieveMultiple(new FetchExpression(fetchXML));
                                        if (entcollWA.Entities.Count > 0)
                                        {
                                            DynamicWAMessageEnglish = entcollWA.Entities[0].Contains("hil_dynamicwamessage") ? entcollWA.Entities[0].GetAttributeValue<string>("hil_dynamicwamessage") : "";
                                            DynamicWAMessageHindi = entcollWA.Entities[0].Contains("hil_dynamicwamessage") ? entcollWA.Entities[0].GetAttributeValue<string>("hil_dynamicwamessagehindi") : "";
                                        }
                                    }

                                    string? ActiveJobNumber = Jobs.Contains("hil_newjobid") ? Jobs.GetAttributeValue<EntityReference>("hil_newjobid").Name : "";
                                    if (Jobs.FormattedValues["msdyn_substatus"].ToLower() != "canceled")
                                    {
                                        // Attempt to update the status
                                        Entity updateJob = new Entity("msdyn_workorder", new Guid(ent.GetAttributeValue<EntityReference>("hil_jobid").Id.ToString()));
                                        updateJob["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("1527fa6c-fa0f-e911-a94e-000d3af060a1")); //Cancelled
                                        updateJob["hil_webclosureremarks"] = "Duplicate Job : " + ActiveJobNumber;

                                        try
                                        {
                                            service.Update(updateJob);
                                            Console.WriteLine($"Updated job {JobNumber} status to 'Cancelled'.");
                                        }
                                        catch (Exception updateException)
                                        {
                                            Console.WriteLine($"Error updating job {JobNumber} status: {updateException.Message}");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Job {JobNumber} is already in the 'Cancelled' status or the transition is not allowed.");
                                    }
                                    EntityReference? Customer = Jobs.GetAttributeValue<EntityReference>("hil_customerref");
                                    string UserName = FixName(Customer.Name);
                                    Console.WriteLine($"Processing: {_rowCount}/{_totalCount} for Job#{JobNumber} Mobile Number: {MobileNumber}");
                                    Console.WriteLine($"Checking for previous WA CTA Trigger.");

                                    if (UserName != null && MobileNumber != null)
                                    {
                                        ParamModelData = new Root();
                                        ParamModelData.phoneNumber = MobileNumber;
                                        ParamModelData.callbackData = triggerName;


                                        template = GetTemplatebyTriggerName(triggerName, JobNumber, JobGuid, UserName, ProductName, ActiveJobNumber, DynamicWAMessageEnglish, DynamicWAMessageHindi);
                                        ParamModelData.template = template;
                                        Console.WriteLine($"Processing:{_rowCount}/{_totalCount}  for Mobile Number: {MobileNumber}");

                                        var RESULT = UpdateWhatsAppUser(ParamModelData);
                                        if (RESULT.result == "true")
                                        {
                                            Console.WriteLine($"Registered Consumer's mobile number {MobileNumber} on Whatsapp.");

                                            ent["hil_triggercount"] = 1;
                                            ent["hil_wastatusreason"] = new OptionSetValue(2);
                                            service.Update(ent);
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error: While Registering  Consumer's mobile number {MobileNumber} on Whatsapp: {RESULT.message}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        public void SentDataForReminderWACTA(string triggerName)
        {
            try
            {
                if (service.IsReady)
                {
                    DateTime _validUpto = DateTime.Now.Date;
                    //  DateTime _validUpto = new DateTime(2025, 4, 6); // Replace 2024 with the desired year

                    DateTime _proccessDate = new DateTime(_validUpto.Year, _validUpto.Month, _validUpto.Day, 23, 59, 59);
                    //string _proccessDate = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    Console.WriteLine($"Batch Starts: Proccess Date: {_proccessDate}");

                    #region VariableDeclaration
                    EntityCollection entcoll = null;
                    int _rowCount = 1;
                    int _totalCount = 0;
                    string _fetchXML = string.Empty;
                    EntityCollection entcollWA = null;
                    string DynamicWAMessageHindi = string.Empty;
                    string DynamicWAMessageEnglish = string.Empty;
                    int PageNumber = 1;
                    string ProductName = string.Empty;
                    #endregion

                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' page='{PageNumber++}'>
                            <entity name='hil_wacta'>
                            <attribute name='hil_wactaid'/>
                            <attribute name='hil_watriggername'/>
                            <attribute name='createdon'/>
                            <attribute name='hil_whatsappctaresponse'/>
                            <attribute name='hil_wastatusreason'/>
                            <attribute name='hil_validupto'/>
                            <attribute name='hil_triggercount'/>
                            <attribute name='hil_jobid'/>
                            <order attribute='hil_watriggername' descending='false'/>
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_whatsappctaresponse' operator='null' />
                                <condition attribute='hil_validupto' operator='lt' value='{_proccessDate.ToString("yyyy-MM-dd hh:mm:ss")}' /> 
                                <condition attribute='hil_triggercount' operator='eq' value='1' />
                                <condition attribute='hil_watriggername' operator='eq' value='{triggerName}' />
                                <condition attribute='hil_wastatusreason' operator='eq' value='1' />
                            </filter>
                            </entity>
                            </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            Console.WriteLine($"Total WhatsApp CTA Reminders found to be sent: {entcoll.Entities.Count}");
                            _totalCount += entcoll.Entities.Count;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    int triggerCount = ent.GetAttributeValue<int>("hil_triggercount");
                                    DateTime validUpto = ent.GetAttributeValue<DateTime>("hil_validupto");

                                    Entity Jobs = service.Retrieve("msdyn_workorder", ent.GetAttributeValue<EntityReference>("hil_jobid").Id, new ColumnSet("hil_mobilenumber", "hil_customerref", "msdyn_name", "hil_productcategory", "hil_jobcancelreason"));
                                    string? MobileNumber = Jobs.Contains("hil_mobilenumber") ? GetLast10Characters(Jobs.GetAttributeValue<string>("hil_mobilenumber")) : null;
                                    string? JobNumber = Jobs.Contains("msdyn_name") ? Jobs.GetAttributeValue<string>("msdyn_name") : null;
                                    string JobGuid = Jobs.Id.ToString();

                                    EntityReference? ProductNameInfo = Jobs.Contains("hil_productcategory") ? Jobs.GetAttributeValue<EntityReference>("hil_productcategory") : null;
                                    if (ProductNameInfo != null)
                                    {
                                        ProductName = ToUpperCamelCase(ProductNameInfo.Name);
                                    }
                                    string ActiveJobNumber = "null";
                                    OptionSetValue? JobCancelReason = Jobs.Contains("hil_jobcancelreason") ? Jobs.GetAttributeValue<OptionSetValue>("hil_jobcancelreason") : null;
                                    if (JobCancelReason != null && triggerName != null)
                                    {
                                        string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_jobcancellationwamapping'>
                                            <attribute name='hil_jobcancelreasonwa'/>
                                            <attribute name='hil_dynamicwamessage'/>
                                            <attribute name='hil_wacampaignname'/>
                                            <attribute name='hil_jobcancellationwamappingid'/>
                                            <attribute name='hil_dynamicwamessagehindi' />
                                            <order attribute='hil_jobcancelreasonwa' descending='false'/>
                                            <filter type='and'>
                                                <condition attribute='statecode' operator='eq' value='0'/>
                                                <condition attribute='hil_wacampaignname' operator='eq' value='{triggerName}'/>
                                                <condition attribute='hil_jobcancelreasonwa' operator='eq' value='{JobCancelReason.Value}'/>
                                            </filter>
                                            </entity>
                                            </fetch>";
                                        entcollWA = service.RetrieveMultiple(new FetchExpression(fetchXML));
                                        if (entcollWA.Entities.Count > 0)
                                        {
                                            DynamicWAMessageEnglish = entcollWA.Entities[0].Contains("hil_dynamicwamessage") ? entcollWA.Entities[0].GetAttributeValue<string>("hil_dynamicwamessage") : "";
                                            DynamicWAMessageHindi = entcollWA.Entities[0].Contains("hil_dynamicwamessage") ? entcollWA.Entities[0].GetAttributeValue<string>("hil_dynamicwamessagehindi") : "";
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Error: Job Cancel reason is null");
                                    }
                                    EntityReference? Customer = Jobs.GetAttributeValue<EntityReference>("hil_customerref");
                                    string UserName = Customer.Name;
                                    Console.WriteLine($"Record# {_rowCount++}/{_totalCount} Job# {JobNumber} Mobile Number# {MobileNumber} validUpto: {validUpto}");

                                    if (UserName != null && MobileNumber != null)
                                    {
                                        Root ParamModelData = new Root();
                                        ParamModelData.phoneNumber = MobileNumber;
                                        ParamModelData.callbackData = triggerName;

                                        Template template = GetTemplatebyTriggerName(triggerName, JobNumber, JobGuid, UserName, ProductName, ActiveJobNumber, DynamicWAMessageEnglish, DynamicWAMessageHindi);
                                        ParamModelData.template = template;

                                        var RESULT = UpdateWhatsAppUser(ParamModelData);
                                        if (RESULT.result == "true")
                                        {
                                            triggerCount = triggerCount + 1;
                                            ent["hil_triggercount"] = triggerCount;
                                            ent["hil_validupto"] = validUpto.AddHours(24);
                                            Console.WriteLine($"Reminder: {validUpto.AddHours(24)}");
                                            service.Update(ent);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        private Template GetTemplatebyTriggerName(string triggerName, string JobNumber, string JobGuid, string UserName, string ProductName, string ActiveJobNumber, string DynamicWAMessageEnglish, string DynamicWAMessageHindi)
        {
            Template template = new Template();
            switch (triggerName)
            {
                case "cancel_2cta_h8":
                    {
                        template.name = triggerName;
                        template.headerValues = new List<string> { UserName };
                        template.bodyValues = new List<string> { ProductName, JobNumber };
                        template.buttonPayload = new Dictionary<string, List<string>>
                        {
                            { "0", new List<string> { $"Cancel My SR: {JobNumber} [{triggerName}]"  } } ,
                            { "1", new List<string> { $"Reschedule today for my SR: {JobNumber} [{triggerName}]" } },
                        };
                    }
                    break;
                case "cancel_3cta_jo":
                    {
                        template.name = triggerName;
                        template.headerValues = new List<string> { UserName };
                        template.bodyValues = new List<string> { ProductName, JobNumber, DynamicWAMessageEnglish, DynamicWAMessageHindi };
                        template.buttonPayload = new Dictionary<string, List<string>>
                        {
                            { "0", new List<string> { $"Reschedule today for my SR: {JobNumber} [{triggerName}]"  } } ,
                            { "1", new List<string> { $"Reschedule later for my SR: {JobNumber} [{triggerName}]" } },
                            { "2", new List<string> { $"Cancel my SR: {JobNumber} [{triggerName}]" } },
                        };
                    }
                    break;
                case "cancel_nocta_jp":
                    {
                        template.name = triggerName;
                        template.headerValues = new List<string> { UserName };
                        template.bodyValues = new List<string> { ProductName, JobNumber, ActiveJobNumber };
                        template.buttonPayload = new Dictionary<string, List<string>>
                        {

                        };
                    }
                    break;
                case "invoice_2cta_68":
                    {
                        template.name = triggerName;
                        template.headerValues = new List<string> { UserName };
                        template.bodyValues = new List<string> { ProductName, JobNumber };
                        template.buttonPayload = new Dictionary<string, List<string>>
                        {
                            { "0", new List<string> { $"Upload invoice for GUID: {JobNumber} [{triggerName}]" } } ,
                            { "1", new List<string> { $"Cancel My SR : {JobNumber} [{triggerName}]" } },
                        };
                    }
                    break;
                case "est_na_2cta":
                    {
                        template.name = triggerName;
                        template.headerValues = new List<string> { UserName };
                        template.bodyValues = new List<string> { ProductName, JobNumber };
                        template.buttonPayload = new Dictionary<string, List<string>>
                        {
                            { "0", new List<string> { $"Yes, SR estimate for SR: {JobNumber} [{triggerName}]" } } ,
                            { "1", new List<string> { $"No, SR estimate for SR: {JobNumber} [{triggerName}]" } },
                        };
                    }
                    break;
            }
            return template;
        }
        private string GetLast10Characters(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }
            return input.Length <= 10 ? input : input.Substring(input.Length - 10);
        }
        private string FixName(string input)
        {
            // Remove leading and trailing spaces
            input = input.Trim();

            // Remove leading dot
            if (input.StartsWith("."))
            {
                input = input.Substring(1);
            }

            // Remove trailing dot
            if (input.EndsWith("."))
            {
                input = input.Substring(0, input.Length - 1);
            }

            // Replace multiple dots or spaces with a single one
            input = Regex.Replace(input, @"[.]{2,}", ".");
            input = Regex.Replace(input, @"[ ]{2,}", " ");

            // Remove dot-space and space-dot combinations
            input = Regex.Replace(input, @"[. ]{2}", "");

            // Ensure no dot-space or space-dot combinations
            input = Regex.Replace(input, @"\.\s", ".");
            input = Regex.Replace(input, @"\s\.", " ");

            // Remove any invalid characters
            input = Regex.Replace(input, @"[^A-Za-z. ]", "");

            return input;
        }
        private string ToUpperCamelCase(string str)
        {
            string[] words = str.Split(' ');
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1).ToLower();
            }
            return string.Join("", words);
        }
        private UpdateAppUserRes UpdateWhatsAppUser(Root userDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration("Send WA Trigger");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            Console.WriteLine(data);
            using (var client = new HttpClient())
            {
                HttpContent objcontent = new StringContent(data);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objcontent).Result;
                return JsonConvert.DeserializeObject<UpdateAppUserRes>(result.Content.ReadAsStringAsync().Result);
            }
        }
        private sendMesgRes SendMessage(TrackEvents campaignDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration("JobCancellationWhatsAppCampaign");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string Campaigndata = JsonConvert.SerializeObject(campaignDetails);
            using (var client = new HttpClient())
            {
                HttpContent objCampaign = new StringContent(Campaigndata);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objCampaign).Result;
                return JsonConvert.DeserializeObject<sendMesgRes>(result.Content.ReadAsStringAsync().Result);
            }
        }
        private IntegrationConfig IntegrationConfiguration(string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + " " + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
        private bool ExistJobIdInWACTA(Guid jobId)
        {
            try
            {
                if (service.IsReady)
                {
                    string fetchXML = $@"
                        <fetch top='1'>
                        <entity name='hil_wacta'>
                        <attribute name='hil_jobid' />
                            <filter>
                                 <condition attribute='hil_jobid' operator='eq' value='{jobId}' />
                            </filter>
                        </entity>
                        </fetch>";

                    EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    return result.Entities.Count > 0;
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }
    }

    public class UpdateAppUserRes
    {
        public string result { get; set; }
        public string message { get; set; }
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class TrackEvents
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public string @event { get; set; }
        public trait traits { get; set; }
    }
    public class trait
    {
        public string name { get; set; }
        public string category { get; set; }
        public string prdSerial { get; set; }
    }
    public class sendMesgRes
    {
        public bool result { get; set; }
        public string message { get; set; }
        public string id { get; set; }
    }

    public class Root
    {
        public string countryCode { get; set; } = "+91";
        public string phoneNumber { get; set; }
        public string callbackData { get; set; }
        public string type { get; set; } = "Template";
        public Template template { get; set; }
    }
    public class Template
    {
        public string name { get; set; }
        public string languageCode { get; set; } = "en";
        public List<string> headerValues { get; set; }
        public List<string> bodyValues { get; set; }
        public Dictionary<string, List<string>> buttonPayload { get; set; }
    }
    public class AMCRoot
    {
        public string countryCode { get; set; } = "+91";
        public string phoneNumber { get; set; }
        public string callbackData { get; set; }
        public string type { get; set; } = "Template";
        public AMCTemplate template { get; set; }
    }
    public class AMCTemplate
    {
        public string name { get; set; }
        public string languageCode { get; set; } = "en";
        public List<string> headerValues { get; set; }
        public List<string> bodyValues { get; set; }
    }
}
