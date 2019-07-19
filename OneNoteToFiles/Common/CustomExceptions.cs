using System;

namespace OneNoteToFiles.Common
{
    public class ProgramException : Exception
    {
        public ProgramException(string message)
            : base(message)
        {
        }
    }

    public class HierarchyNotFoundException : ProgramException
    {
        public HierarchyNotFoundException(string message)
            : base(message)
        {
        }
    }
}
