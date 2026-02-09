using System.Collections.Generic;

namespace SlicerExport.Core
{
    public class ExportObject
    {
        public SimpleMesh Mesh { get; set; }
        public PartType PartType { get; set; }
        public string Name { get; set; }
        public Dictionary<string, string> PrintSettings { get; set; }
    }
}
