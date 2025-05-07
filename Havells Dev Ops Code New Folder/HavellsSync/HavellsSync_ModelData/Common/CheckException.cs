using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.Common
{
    /// <summary>
    /// Check Exception
    /// </summary>
    public class CheckException : Exception
    {
        /// <summary>
        /// Check Exception
        /// </summary>
        public CheckException()
            : base()
        {
        }

        /// <summary>
        /// Check Exception
        /// </summary>
        /// <param name="message">The message</param>
        public CheckException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Check Exception
        /// </summary>
        /// <param name="message">The message</param>
        /// <param name="innerException">the inner exception</param>
        public CheckException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
