using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using OMSMail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace CreateFieldsFromExcelToD365
{
    public class Solutions
    {
        public static EntityCollection RetrieveSolutions(IOrganizationService service)

        {
            QueryExpression query = new QueryExpression("solution");
            query.ColumnSet = new ColumnSet(true);
            query.Distinct = true;
            query.Criteria.AddCondition("uniquename", ConditionOperator.Equal, "AquaP");


            RetrieveMultipleRequest request = new RetrieveMultipleRequest();

            request.Query = query;

            try
            {
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                EntityCollection results = response.EntityCollection;
                Entity entity = new Entity(results[0].LogicalName, results[0].Id);
                entity["ismanaged"] = false;
                service.Update(entity);
                return results;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception(ex.Message);
            }

        }
        public static void exportSol(IOrganizationService service)
        {
            string outputDir = @"C:\temp\";

            //Creates the Export Request
            ExportSolutionRequest exportRequest = new ExportSolutionRequest();
            exportRequest.Managed = false;
            exportRequest.SolutionName = "EMS";// loginUser.SolutionName;
            
            ExportSolutionResponse exportResponse = (ExportSolutionResponse)service.Execute(exportRequest);

            //Handles the response
            byte[] exportXml = exportResponse.ExportSolutionFile;
            string filename = "AquaP" + "_" + DateTime.Now.ToString() + ".zip";
            File.WriteAllBytes(outputDir + filename, exportXml);

            //Console.WriteLine("Solution Successfully Exported to {0}", outputDir + filename);

        }
    }
}
