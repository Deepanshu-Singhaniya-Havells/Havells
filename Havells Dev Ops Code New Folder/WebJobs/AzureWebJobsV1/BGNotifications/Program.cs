using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.ServiceModel.Description;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace BGNotifications
{
    class Program
    {
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            LoadAppSettings();
            if (loginUserGuid != Guid.Empty)
            {
                GenerateBGNotifications(_service);
            }
        }

        #region App Setting Load/CRM Connection
        static void LoadAppSettings()
        {
            try
            {
                _userId = ConfigurationManager.AppSettings["CrmUserId"].ToString();
                _password = ConfigurationManager.AppSettings["CrmUserPassword"].ToString();
                _soapOrganizationServiceUri = ConfigurationManager.AppSettings["CrmSoapOrganizationServiceUri"].ToString();
                ConnectToCRM();
            }
            catch (Exception ex)
            {
                Console.WriteLine("AMCWarrantyCertificate.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
            }
        }
        static void ConnectToCRM()
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = _userId;
                credentials.UserName.Password = _password;
                Uri serviceUri = new Uri(_soapOrganizationServiceUri);
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                proxy.EnableProxyTypes();
                _service = (IOrganizationService)proxy;
                loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
            }
            catch (Exception ex)
            {
                Console.WriteLine("AMCWarrantyCertificate.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion
        static void GenerateBGNotifications(IOrganizationService _service)
        {
            try
            {
                QueryExpression queryExp = new QueryExpression("hil_tenderbankguarantee");
                queryExp.ColumnSet = new ColumnSet("ownerid");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_validityperiod", ConditionOperator.OnOrBefore, DateTime.Today.AddDays(7));
                queryExp.Criteria.AddCondition("hil_expecteddateofcollectingbankguarantee", ConditionOperator.Null);
                queryExp.Criteria.AddCondition("hil_bgnumber", ConditionOperator.NotNull);
                queryExp.Distinct = true;

                EntityCollection entCol = _service.RetrieveMultiple(queryExp);
                int _rowCount = 1;
                foreach (Entity ent in entCol.Entities)
                {
                    var owner = ent.GetAttributeValue<EntityReference>("ownerid");
                    queryExp = new QueryExpression("hil_tenderbankguarantee");
                    queryExp.ColumnSet = new ColumnSet("hil_nameofdepartment", "createdon", "hil_purpose", "hil_validityperiod", "ownerid", "hil_claimperiod", "hil_bgnumber", "hil_bgnodate", "hil_guarnteeamount", "hil_salesoffice", "hil_orderno", "hil_expecteddateofcollectingbankguarantee");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_validityperiod", ConditionOperator.OnOrBefore, DateTime.Today.AddDays(7));
                    queryExp.Criteria.AddCondition("ownerid", ConditionOperator.Equal, owner.Id);
                    queryExp.Criteria.AddCondition("hil_expecteddateofcollectingbankguarantee", ConditionOperator.Null);
                    queryExp.Criteria.AddCondition("hil_bgnumber", ConditionOperator.NotNull);
                    queryExp.AddOrder("hil_validityperiod", OrderType.Ascending);
                    EntityCollection entColBG = _service.RetrieveMultiple(queryExp);
                    if (entColBG.Entities.Count > 0)
                    {
                        string bodtText = mailBody(entColBG);
                        sendEmail(_service, bodtText, owner, entColBG.Entities[0].ToEntityReference(), entColBG.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice").Name);
                    }
                    Console.WriteLine("Current Record Count: " + _rowCount.ToString() + "/" + entCol.Entities.Count.ToString());
                    _rowCount += 1;
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
        static string mailBody(EntityCollection entColBg)
        {
            StringBuilder sb = new StringBuilder();
            string _baseURL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&pagetype=entityrecord&etn=hil_tenderbankguarantee&id=";

            sb.Append("<div><p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span> Dear Sir </span></p>");
            sb.Append("<p style='margin-bottom:.0001pt;'></p>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'>");
            sb.Append("<span>Plz find below here the list of BGs issued in favor of the customers handled by " + entColBg.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice").Name + " branch, which have already lapsed or getting expired shortly. </span>");
            sb.Append("</p><p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;' ><span></span></p>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span>");
            sb.Append("Kindly arrange to get back the expired BGs and arrange the necessary action for the BGs soon getting expired, as an earliest enabling us to close the liability from bank against the same.");
            sb.Append("</span ></p><p style = 'margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;' ></p>");
            sb.Append("<table border=0 cellspacing=0 cellpadding=0 width=0 style='width:703.9pt;margin-left:-.15pt;border-collapse:collapse'>");
            sb.Append("<tr style='height:24.0pt'><td valign=top style='width:121pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><b><u>");
            sb.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;'> BG Number</span></u></b></p></td>");
            sb.Append("<td valign = top style='width:120pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><b><u>");
            sb.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;'> BG Date </span ></u></b></p>");
            sb.Append("</td><td valign=top style='width:97pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom: .0001pt; font-size: 11.0pt; font-family:Calibri,sans-serif;text-align:center' align=center>");
            sb.Append("<b><u><span style='font-size:9.0pt;font-family:Arial,sans-serif;'> Name of Customer </span>");
            sb.Append("</u></b></p></td><td valign=top style='width:91pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:right' align=right>");
            sb.Append("<b><u><span style='font-size:9.0pt;font-family:Arial,sans-serif;'> BG Amount </span>");
            sb.Append("</u></b></p></td><td valign=top style='width:110pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center >");
            sb.Append("<b><u><span style='font-size:9.0pt;font-family:Arial,sans-serif;'> Validity Period </span>");
            sb.Append("</u></b></p></td><td valign=top style='width:82pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center>");
            sb.Append("<b><u><span style='font-size:9.0pt;font-family:Arial,sans-serif;'> Claim Period </span>");
            sb.Append("</u></b></p></td><td valign=top style='width:74pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center>");
            sb.Append("<b><u><span style='font-size:9.0pt;font-family:Arial,sans-serif;'> Type of Bank Guarantee </span>");
            sb.Append("</u></b></p></td><td width=75 valign=top style='width:54.8pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center>");
            sb.Append("<b><u><span style='font-size:9.0pt;font-family:Arial,sans-serif;'> Sales office </span>");
            sb.Append("</u></b></p></td><td valign=top style='width:242pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibr,sans-serif;text-align:center'><b><u>");
            sb.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;'> Order No.</span></u></b></p>");
            sb.Append("</td></tr>");
            StringBuilder sbtr = new StringBuilder();
            decimal _totalBGamount = 0;
            foreach (Entity bg in entColBg.Entities)
            {
                _totalBGamount += (bg.Contains("hil_guarnteeamount") ? Math.Round(bg.GetAttributeValue<Money>("hil_guarnteeamount").Value, 2) : 0);
                sbtr.Append("<tr style='height:24.0pt'><td valign=top style='width:121pt;border:solid windowtext 1.0pt;border-top:none;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'><a href='"+ (_baseURL +  bg.Id) + "' target='_blank'>" + (bg.Contains("hil_bgnumber") ? bg.GetAttributeValue<string>("hil_bgnumber"): "") + "</a></span>");
                sbtr.Append("</p></td>");
                sbtr.Append("<td valign=top style='width:120pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'>" + (bg.Contains("hil_bgnumber") ? bg.GetAttributeValue<DateTime>("hil_bgnodate").ToString():"") + "</span></p></td>");
                sbtr.Append("<td valign=top style='width:97pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'>" + (bg.Contains("hil_nameofdepartment") ? bg.GetAttributeValue<string>("hil_nameofdepartment"):"") + "</span></p></td>");
                sbtr.Append("<td valign=top style='width:91pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'>" + (bg.Contains("hil_guarnteeamount") ? Math.Round(bg.GetAttributeValue<Money>("hil_guarnteeamount").Value,2) : 0) + " </span></p></td>");
                sbtr.Append("<td valign=top style='width:82pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center>");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'>" + (bg.Contains("hil_validityperiod") ? bg.GetAttributeValue<DateTime>("hil_validityperiod").ToString():"") + "</span ></p></td>");
                sbtr.Append("<td valign=top style='width:110pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center>");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'>" + (bg.Contains("hil_claimperiod") ? bg.GetAttributeValue<DateTime>("hil_claimperiod").ToString():"") + "</span></p></td>");
                sbtr.Append("<td valign=top style='width:74pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p align=center style='text-align:center'><span style='font-size:9.0pt;font-family:'Arial',sans-serif;color:black;'> " + (bg.Contains("hil_purpose") ? bg.FormattedValues["hil_purpose"] : "") + "</span>");
                sbtr.Append("</p></td>");
                sbtr.Append("<td valign=top style='width:75pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center >");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'>" + (bg.Contains("hil_salesoffice") ? bg.GetAttributeValue<EntityReference>("hil_salesoffice").Name : "") + " </span></p></td>");
                sbtr.Append("<td valign=top style='width:200pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
                sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
                sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'>" + (bg.Contains("hil_orderno") ? bg.GetAttributeValue<string>("hil_orderno") : "") + "</span>");
                sbtr.Append("</p></td></tr>");
            }
            sbtr.Append("<tr style='height:24.0pt'><td valign=top style='width:121pt;border:solid windowtext 1.0pt;border-top:none;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'><b>Total BG Amount</b></span>");
            sbtr.Append("</p></td>");
            sbtr.Append("<td valign=top style='width:120pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'></span></p></td>");
            sbtr.Append("<td valign=top style='width:97pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'></span></p></td>");
            sbtr.Append("<td valign=top style='width:91pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'><b>" + (_totalBGamount) + "</b></span></p></td>");
            sbtr.Append("<td valign=top style='width:82pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center>");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'></span ></p></td>");
            sbtr.Append("<td valign=top style='width:110pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center>");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'></span></p></td>");
            sbtr.Append("<td valign=top style='width:74pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p align=center style='text-align:center'><span style='font-size:9.0pt;font-family:'Arial',sans-serif;color:black;'></span>");
            sbtr.Append("</p></td>");
            sbtr.Append("<td valign=top style='width:75pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center' align=center >");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'></span></p></td>");
            sbtr.Append("<td valign=top style='width:200pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:yellow;padding:0cm 5.4pt 0cm 5.4pt;height:24.0pt'>");
            sbtr.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;text-align:center'>");
            sbtr.Append("<span style='font-size:9.0pt;font-family:Arial,sans-serif;color:black;'></span>");
            sbtr.Append("</p></td></tr>");
            StringBuilder sbend = new StringBuilder();
            sbend.Append("</table><p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;' >");
            sbend.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span> Regards </span></p>");
            sbend.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span> Team HO </span></p></div> ");

            string resultmail = sb.ToString() + sbtr.ToString() + sbend.ToString();
            return resultmail;
        }
        static void sendEmail(IOrganizationService _service, string mailBody, EntityReference bgOwner,EntityReference _refId,string _salesOfficeName)
        {
            QueryExpression queryExpUser = new QueryExpression("hil_userbranchmapping");
            queryExpUser.ColumnSet = new ColumnSet("hil_zonalhead", "hil_buhead", "hil_branchproducthead");
            queryExpUser.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpUser.Criteria.AddCondition("hil_user", ConditionOperator.Equal, bgOwner.Id);
            EntityCollection entColUser = _service.RetrieveMultiple(queryExpUser);
            if (entColUser.Entities.Count > 0)
            {
                Entity entEmail = new Entity("email");
                entEmail["subject"] = @"Reminder for BGs expired / getting expired pertains to " + _salesOfficeName.ToUpper() + " Branch";
                entEmail["description"] = mailBody;

                #region To Parties
                Entity entTo = null;
                EntityCollection entToList = new EntityCollection();
                entTo = new Entity("activityparty");
                entTo["partyid"] = entColUser.Entities[0].GetAttributeValue<EntityReference>("hil_branchproducthead");
                entToList.Entities.Add(entTo);
                entTo = new Entity("activityparty");
                entTo["partyid"] = bgOwner;
                entToList.Entities.Add(entTo);
                entEmail["to"] = entToList;
                #endregion

                #region CC Parties
                Entity entCC = null;
                EntityCollection entCCList = new EntityCollection();
                entCC = new Entity("activityparty");
                entCC["partyid"] = entColUser.Entities[0].GetAttributeValue<EntityReference>("hil_zonalhead");
                entCCList.Entities.Add(entCC);
                entCC = new Entity("activityparty");
                entCC["partyid"] = entColUser.Entities[0].GetAttributeValue<EntityReference>("hil_buhead");
                entCCList.Entities.Add(entCC);

                //entCC = new Entity("activityparty");
                //entCC["partyid"] = new EntityReference("systemuser", new Guid("e8343896-3644-eb11-bb23-000d3af0563b")); // Nagmani Nath Tiwari
                //entCCList.Entities.Add(entCC);
                
                // BD Team :: BG Reminder CC
                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_bdteammember'>
                <attribute name='hil_bdteammemberid' />
                <order attribute='hil_name' descending='false' />
                <filter type='and'>
                    <condition attribute='hil_team' operator='eq' value='{207DB5CB-F107-ED11-82E6-6045BDAC5A1D}' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
                entColUser = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                foreach (Entity ccEntity in entColUser.Entities)
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = ccEntity.ToEntityReference();
                    entCCList.Entities.Add(entTo);
                }
                entEmail["cc"] = entCCList;
                #endregion

                #region From Party
                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;
                entEmail["regardingobjectid"] = _refId;
                #endregion

                Guid emailId = _service.Create(entEmail);
                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)_service.Execute(sendEmailReq);
            }
        }
    }
}
