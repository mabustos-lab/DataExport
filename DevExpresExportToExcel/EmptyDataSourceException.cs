using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataExport
{
    public class EmptyDataSourceException:Exception
    {
        public EmptyDataSourceException(string message) : base(message) { }
    }
}
