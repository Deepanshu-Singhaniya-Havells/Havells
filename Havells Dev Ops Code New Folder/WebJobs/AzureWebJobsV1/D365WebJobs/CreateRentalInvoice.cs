using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class CreateRentalInvoice
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=bc6676d6-1387-4dc4-be89-ba13b08ceb4e;ClientSecret=73P7Q~sWxupzl4j8-B55y5g3QNosxhkjkV6Q2";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://ogre-rental-dev.crm11.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                int i = 1;
                while (true)
                {
                    Console.WriteLine("Processing Invoice #" + i++.ToString());
                    Entity entity = _service.Retrieve("ogre_contract", new Guid("6BBAAA7D-C1A0-EE11-BE37-0022481B6DF4"), new ColumnSet("ogre_depositapplicable", "ogre_account"));
                    string[] bookingIds = "076A9890-C1A0-EE11-BE37-0022481B6DF4".Split(',');
                    Entity _entBooking = _service.Retrieve("ogre_rentalbooking", new Guid(bookingIds[0]), new ColumnSet(true));
                    InvoicingEngine _invEngine = new InvoicingEngine();

                    //_invEngine.CreateDepositInvoice(_service, entity, bookingIds);
                    ReturnObject _retObj = _invEngine.CreateInvoice(_service, _entBooking);
                    Console.WriteLine(_retObj.StatusRemarks);
                }
                //string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //      <entity name='ogre_rentalbooking'>
                //        <attribute name='ogre_rentalbookingid' />
                //        <attribute name='ogre_bookingnumber' />
                //        <attribute name='createdon' />
                //        <order attribute='ogre_bookingnumber' descending='false' />
                //        <filter type='and'>
                //          <condition attribute='ogre_nextinvoicedate' operator='on-or-after' value='2023-12-23' />
                //        </filter>
                //      </entity>
                //    </fetch>";

                //string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //    <entity name='ogre_rentalbooking'>
                //    <attribute name='ogre_bookingnumber' />
                //    <attribute name='ogre_rentalbookingid' />
                //    <attribute name='ogre_nextinvoicedate' />
                //    <attribute name='ogre_bookingnumber' />
                //    <attribute name='createdon' />
                //    <attribute name='ogre_customer' />
                //    <attribute name='ogre_contract' />
                //    <order attribute='ogre_bookingnumber' descending='false' />
                //    <filter type='and'>
                //        <condition attribute='ogre_nextinvoicedate' operator='on' value='2023-12-23' />
                //    </filter>
                //    </entity>
                //</fetch>";

                //EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //foreach (Entity entBooking in entCol.Entities) {
                //    Entity InvoiceEnt = new Entity("ogre_rentalinvoice");
                //    InvoiceEnt["ogre_invoicetype"] = new EntityReference("ogre_rentalinvoicetype", new Guid("63af5955-f9a0-ee11-be37-0022481b6df4"));//Rental Invoice
                //    InvoiceEnt["ogre_contract"] = entBooking.GetAttributeValue<EntityReference>("ogre_contract");
                //    InvoiceEnt["ogre_customer"] = entBooking.GetAttributeValue<EntityReference>("ogre_customer");
                //    InvoiceEnt["ogre_invoicedate"] = DateTime.Now;
                //    InvoiceEnt["ogre_invoicestatus"] = new EntityReference("ogre_invoicestatus", new Guid("a3fc6e8b-f9a0-ee11-be37-0022481b6df4"));//Invoice Status - Draft

                //    EntityReference InvoiceHeaderID = new EntityReference("ogre_rentalinvoice", _service.Create(InvoiceEnt));

                //    string _fetchXMLBookingLines = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //      <entity name='ogre_rentalbookingline'>
                //        <attribute name='ogre_billingfrom' />
                //        <attribute name='ogre_objectid' />
                //        <attribute name='ogre_unit' />
                //        <attribute name='ogre_rate' />
                //        <attribute name='ogre_invoicepattern' />
                //        <attribute name='ogre_contract' />
                //        <attribute name='ogre_chargetype' />
                //        <attribute name='ogre_status' />
                //        <attribute name='ogre_billinguntil' />
                //        <attribute name='ogre_billingtype' />
                //        <attribute name='ogre_billingperiod' />
                //        <attribute name='ogre_billingcycle' />
                //        <attribute name='ogre_quantity' />
                //        <attribute name='ogre_rentalperiod' />
                //        <attribute name='ogre_amount' />
                //        <attribute name='ogre_rentalbookinglineid' />
                //        <filter type='and'>
                //          <condition attribute='statecode' operator='eq' value='0' />
                //          <condition attribute='ogre_rentalbooking' operator='eq' value='{entBooking.Id}' />
                //        </filter>
                //      </entity>
                //    </fetch>";
                //    EntityCollection entColLines = _service.RetrieveMultiple(new FetchExpression(_fetchXMLBookingLines));
                //    foreach (Entity entLine in entColLines.Entities) {
                //        Entity InvoiceLine = new Entity("ogre_rentalinvoiceline");
                //        InvoiceLine["ogre_chargetype"] = entLine.GetAttributeValue<EntityReference>("ogre_chargetype");
                //        InvoiceLine["ogre_quantity"] = entLine.GetAttributeValue<int>("ogre_quantity");
                //        InvoiceLine["ogre_unit"] = entLine.GetAttributeValue<EntityReference>("ogre_unit");
                //        InvoiceLine["ogre_rate"] = entLine.GetAttributeValue<EntityReference>("ogre_unit");
                //        InvoiceLine["ogre_chargedescription"] = "Rental against Booking# " + entBooking.GetAttributeValue<string>("ogre_bookingnumber");
                //        InvoiceLine["ogre_taxtype"] = null;
                //        InvoiceLine["ogre_taxper"] = new decimal(0.00);
                //        InvoiceLine["ogre_invoicefrom"] = entLine.GetAttributeValue<DateTime>("ogre_nextinvoicedate");
                //        InvoiceLine["ogre_invoiceuntil"] = entLine.GetAttributeValue<DateTime>("ogre_nextinvoicedate").AddDays(30);
                //        InvoiceLine["ogre_name"] = "INV-001";
                //        InvoiceLine["ogre_rentalinvoice"] = InvoiceHeaderID;
                //        InvoiceLine["ogre_rentalbooking"] = entBooking.ToEntityReference();
                //        _service.Create(InvoiceLine);
                //    }
                //}
            }
        }

        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
}
