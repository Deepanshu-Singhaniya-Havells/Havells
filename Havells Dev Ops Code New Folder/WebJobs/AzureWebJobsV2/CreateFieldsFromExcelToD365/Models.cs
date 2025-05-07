using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateFieldsFromExcelToD365
{
    public class Models
    {
        public string DisplayName { get; set; }
        public string AttributeType { get; set; }
        public string MaxLength { get; set; }
        public string MinValue { get; set; }
        public string MaxValue { get; set; }
        public string Precision { get; set; }
        public string RequiredLevel { get; set; }
        public string IsAuditEnabled { get; set; }
        public string EntityName { get; set; }
        public string Options { get; set; }
        public string DateFormat { get; set; }
        public string wholeNumberFormat { get; set; }
        public string StringFormat { get; set; }
        public string OptionValue { get; set; }

        //static void getEntityMetaData(IOrganizationService service)
        //{
        //    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
        //    {
        //        EntityFilters = EntityFilters.All,
        //        LogicalName = "account"
        //    };
        //    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
        //    EntityMetadata AccountEntity = retrieveAccountEntityResponse.EntityMetadata;

        //    Console.WriteLine("Account entity metadata:");
        //    Console.WriteLine(AccountEntity.SchemaName);
        //    Console.WriteLine(AccountEntity.DisplayName.UserLocalizedLabel.Label);
        //    Console.WriteLine(AccountEntity.EntityColor);

        //    Console.WriteLine("Account entity attributes:");
        //    foreach (object attribute in AccountEntity.Attributes)
        //    {
        //        AttributeMetadata a = (AttributeMetadata)attribute;
        //        Console.WriteLine(a.LogicalName);
        //    }
        //    Console.ReadLine();

        //}
    }

}
