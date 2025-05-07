using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    //Serial number validation
    [System.Runtime.Serialization.DataContractAttribute()]
    public partial class SerialNumberValidation
    {

        [System.Runtime.Serialization.DataMemberAttribute()]
        public EX_PRD_DET EX_PRD_DET;
    }

    [System.Runtime.Serialization.DataContractAttribute()]
    public class SerialNumberValidationMRN
    {
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string EX_RETURN { get; set; }
        [System.Runtime.Serialization.DataMemberAttribute()]
        public List<EX_PRD_DET> EX_PRD_DET { get; set; }
    }
    // Type created for JSON at <<root>> --> EX_PRD_DET
    [System.Runtime.Serialization.DataContractAttribute()]
    public partial class EX_PRD_DET
    {

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SERIAL_NO;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string MATNR;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string MAKTX;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SPART;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string REGIO;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string VBELN;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string FKDAT;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string KUNAG;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string NAME1;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string WTY_STATUS;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string IS_TYPE;
    }
}
