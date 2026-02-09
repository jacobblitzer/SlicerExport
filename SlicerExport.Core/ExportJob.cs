using System.Collections.Generic;

namespace SlicerExport.Core
{
    public class ExportJob
    {
        public List<ExportGroup> Groups { get; set; } = new List<ExportGroup>();
        public SlicerTarget Target { get; set; }
        public string OutputPath { get; set; }
    }
}
