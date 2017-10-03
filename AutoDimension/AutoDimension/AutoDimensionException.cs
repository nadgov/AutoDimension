using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoDimension
{
    #region AutoDimensionException
    public class AutoDimensionException : Exception
    {
        public AutoDimensionException(string message)
            : base(message)
        {
        }
    }
    #endregion
}
