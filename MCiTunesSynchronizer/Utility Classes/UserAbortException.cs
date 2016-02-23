using System;
using System.Collections.Generic;
using System.Text;

namespace MCiTunesSynchronizer
{
    // So we can tell the difference between a known exception and unknown
    public class UserAbortException : Exception
    {
        public UserAbortException(string message)
            : base(message)
        {
        }
    }
}
