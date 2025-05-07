using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.EnviromentMetaDateMigration
{
    public class GlobleOptionSet
    {
        public static void CreateGlobalOptionSet(IOrganizationService service, IOrganizationService _DestinationService)
        {
            WriteLogFile.WriteLog("********************** CREATE GLOBAL OPTIONSET METHOD **********************");
            string optionsSetName = "hil_yesnona;hil_yesno;hil_warrantysubstatus;hil_voltage;hil_vaccinationstatus;hil_typeofproduct;hil_tenderstatus;hil_tenderstakeholder;hil_syncstatus;hil_sourceofcreation;hil_smstemplatetype;hil_slastatus;hil_serviceengineerstatus;hil_serialnumbercount;hil_sawcategoryentrymode;hil_sawapprovalstatus;hil_salutation;hil_returntype;hil_requesttype;hil_recordtype;hil_prtype;hil_performainvoicestatus;hil_paymentterm;hil_paymentstatus;hil_operator;hil_nomineerelationship;hil_maritalstatus;hil_level;hil_leadtype;hil_joberrorsstatus;hil_jobclass;hil_jobclaimstatus;hil_jobadditionalactions;hil_inventorytype;hil_interest;hil_incentivetype;hil_incentivecategory;hil_icutomainwirespecs;hil_hierarchylevel;hil_gdatatype;hil_franchiseecategory;hil_enquirytype;hil_disposition;hil_discounttype;hil_departmentenquiry;hil_customerfeedback;hil_countryclassification;hil_consumernonconsumer;hil_claimstatus;hil_chargetype;hil_category;hil_callcenter;hil_brand;hil_bloodgroup;hil_availabilitystatus;hil_approvalstatus;hil_approvallevel;hil_approvalentitystatus;hil_activitygstslab;hil_aboutsyncappdownload";
            //"new_customertype;hil_yesnona;hil_yesno;hil_warrantysubstatus;hil_voltage;hil_vaccinationstatus;hil_typeofproduct;hil_tenderstatus;hil_tenderstakeholder;hil_syncstatus;hil_sourceofcreation;hil_smstemplatetype;hil_slastatus;hil_serviceengineerstatus;hil_serialnumbercount;hil_sawcategoryentrymode;hil_sawapprovalstatus;hil_salutation;hil_returntype;hil_requesttype;hil_recordtype;hil_prtype;hil_performainvoicestatus;hil_paymentterm;hil_paymentstatus;hil_operator;hil_nomineerelationship;hil_maritalstatus;hil_level;hil_leadtype;hil_joberrorsstatus;hil_jobclass;hil_jobclaimstatus;hil_jobadditionalactions;hil_inventorytype;hil_interest;hil_incentivetype;hil_incentivecategory;hil_icutomainwirespecs;hil_hierarchylevel;hil_gdatatype;hil_franchiseecategory;hil_enquirytype;hil_disposition;hil_discounttype;hil_departmentenquiry;hil_customerfeedback;hil_countryclassification;hil_consumernonconsumer;hil_claimstatus;hil_chargetype;hil_category;hil_callcenter;hil_brand;hil_bloodgroup;hil_availabilitystatus;hil_approvalstatus;hil_approvallevel;hil_approvalentitystatus;hil_activitygstslab;hil_aboutsyncappdownload";
            string[] optionsetnames = optionsSetName.Split(';');
            WriteLogFile.WriteLog("Totla Count " + optionsetnames.Length);
            int done = 0;
            int totla = optionsetnames.Length;
            int error = 0;

            foreach (string option in optionsetnames)
            {
                WriteLogFile.WriteLog("Option Set Name " + option);
                try
                {
                    RetrieveOptionSetRequest retrieveOptionSetRequest =
                        new RetrieveOptionSetRequest
                        {
                            Name = option
                        };
                    // Execute the request.
                    RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);
                    OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
                    WriteLogFile.WriteLog("Optons Set retrived");
                    CreateOptionSetRequest createOptionSetRequest = new CreateOptionSetRequest
                    {
                        OptionSet = retrievedOptionSetMetadata
                    };
                    CreateOptionSetResponse optionsResp = (CreateOptionSetResponse)_DestinationService.Execute(createOptionSetRequest);
                    WriteLogFile.WriteLog("Optons Set Created");
                    PublishXmlRequest pxReq2 = new PublishXmlRequest { ParameterXml = String.Format("<importexportxml><optionsets><optionset>{0}</optionset></optionsets></importexportxml>", option) };
                    _DestinationService.Execute(pxReq2);
                    WriteLogFile.WriteLog("Optons Set published");

                    WriteLogFile.WriteLog(done + " / " + totla);
                    done++;
                }
                catch (Exception ex)
                {
                    error++;
                    WriteLogFile.WriteLog("-------------------------Error " + error + " in Createing OptionSet with name " + option + " Error is :- " + ex.Message);
                }
            }
            WriteLogFile.WriteLog("********************** CREATE GLOBAL OPTIONSET METHOD ENDs **********************");
        }
    }
}
