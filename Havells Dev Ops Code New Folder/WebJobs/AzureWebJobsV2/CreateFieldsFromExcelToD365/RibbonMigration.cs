using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CreateFieldsFromExcelToD365
{
    public static class RibbonMigration
    {
        public static void mainFunction(IOrganizationService service, IOrganizationService serviceSer)
        {
            string entityNamestring = "account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;phonecall;product;systemuser;task";
            string[] entityNameArr = entityNamestring.Split(';');
            foreach (string entityName in entityNameArr)
            {
                MigrateRibbon(service, serviceSer, entityName);
            }
        }
        public static void MigrateRibbon(IOrganizationService service, IOrganizationService serviceSer, string entityName)
        {
            Console.WriteLine("************* Ribbom migration Started for entity " + entityName + " *************");
            try
            {
                CreateSolution(serviceSer, entityName, Program.PublisherSer, "Ser", Program.exportFolderSer);
                Console.WriteLine("Solution is Created and Exported from Service Env.");
                CreateSolution(service, entityName, Program.PublisherPrd, "Prd", Program.exportFolderPrd);
                Console.WriteLine("Solution is Created and Exported from Prd Env.");
                ChangeSolution(Program.exportFolderPrd + entityName + "Prd\\customizations.xml", Program.exportFolderSer + entityName + "Ser\\customizations.xml");
                Console.WriteLine("XML Changes are Done.");
                ZipFiles(Program.exportFolderSer + entityName + "Ser", "C:\\Users\\35405\\Downloads\\Zip\\result_" + entityName + ".zip");
                Console.WriteLine("Solution Zip File is Created.");
                //ImportSolution(serviceSer, "C:\\Users\\35405\\Downloads\\Zip\\result_" + entityName + ".zip");
                //Console.WriteLine("Solution is Imported to Service Env.");
                //DeleteFile("C:\\Users\\35405\\Downloads\\Zip\\result_" + entityName + ".zip");
                //Console.WriteLine("Temp file is Deleted.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("-------------------Failuer Solution migration Failed for entity " + entityName + " with Error msg " + ex.Message);
            }
            Console.WriteLine("************* Ribbom migration Ended for entity " + entityName + " *************");
        }
        public static void ZipFiles(string SerPath, string outputFilePath, string password = null)
        {
            ZipFile.CreateFromDirectory(SerPath, outputFilePath, CompressionLevel.Fastest, false);

        }

        public static void CreateSolution(IOrganizationService service, string Entityname, Guid publisherId, string prdName, string exportFolder)
        {

            // Create a Solution
            //Define a solution
            Entity solution = new Entity("solution");
            solution["uniquename"] = Entityname + "_NEW";
            solution["friendlyname"] = Entityname + "_NEW";
            solution["publisherid"] = new EntityReference("publisher", publisherId);
            solution["version"] = "1.0";

            Guid _solutionsSampleSolutionId = service.Create(solution);

            // Add an existing Solution Component
            //Add the Account entity to the solution
            RetrieveEntityRequest retrieveForAddAccountRequest = new RetrieveEntityRequest()
            {
                LogicalName = Entityname
            };
            RetrieveEntityResponse retrieveForAddAccountResponse = (RetrieveEntityResponse)service.Execute(retrieveForAddAccountRequest);
            AddSolutionComponentRequest addReq = new AddSolutionComponentRequest()
            {
                ComponentType = 1,
                ComponentId = (Guid)retrieveForAddAccountResponse.EntityMetadata.MetadataId,
                SolutionUniqueName = Entityname
            };
            service.Execute(addReq);

            // Export or package a solution
            //Export an a solution

            ExportSolutionRequest exportSolutionRequest = new ExportSolutionRequest();
            exportSolutionRequest.Managed = false;
            exportSolutionRequest.SolutionName = Entityname;

            ExportSolutionResponse exportSolutionResponse = (ExportSolutionResponse)service.Execute(exportSolutionRequest);

            byte[] exportXml = exportSolutionResponse.ExportSolutionFile;
            string filename = Entityname + prdName + ".zip";
            File.WriteAllBytes(exportFolder + filename, exportXml);

            Console.WriteLine("Solution exported to {0}.", exportFolder + filename);
            //File.WriteAllBytes(Path.GetFullPath(exportFolder + "\\customizations" + Entityname + ".xml"), unzipRibbon(exportXml));
            Console.WriteLine("Solution exported to {0}.", exportFolder + "\\customizations" + Entityname + ".xml");



            ZipFile.ExtractToDirectory(exportFolder + filename, exportFolder + Entityname + prdName);



        }
        public static void ChangeSolution(string prdPath, string SerPath)
        {
            XmlDocument xmlDocprd = new XmlDocument();
            xmlDocprd.Load(prdPath);
            XmlNode RibbonDiffXmlPRD = xmlDocprd.SelectNodes("ImportExportXml/Entities/Entity/RibbonDiffXml").Item(0);
            string text = RibbonDiffXmlPRD.InnerXml;
            XmlDocument xmlDocSer = new XmlDocument();
            xmlDocSer.Load(SerPath);
            xmlDocSer.SelectNodes("ImportExportXml/Entities/Entity/RibbonDiffXml").Item(0).InnerXml = text;
            xmlDocSer.Save(SerPath);


        }


    }
}
