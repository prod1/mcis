using System;
using System.Collections.Generic;
using System.Text;

namespace MCiTunesSynchronizer
{
    public enum ChangeTypeEnum
    {
        Added,
        Fixed,
        Removed,
        Updated,
        Improved
    }

    public class VersionChange
    {
        public ChangeTypeEnum ChangeType;
        public string Description;
    }
}
