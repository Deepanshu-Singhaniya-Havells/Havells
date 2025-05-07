using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.ICommon
{
    public interface ICustomLog
    {
        void LogToFile(Exception ex);
    }
}
