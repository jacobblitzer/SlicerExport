using System;
using Grasshopper.Kernel;

namespace SlicerExport.Grasshopper
{
    public class SlicerExportInfo : GH_AssemblyInfo
    {
        public override string Name => "SlicerExport";
        public override string Description => "Export meshes to 3MF with slicer-specific part type metadata";
        public override Guid Id => new Guid("a1b2c3d4-e5f6-7890-abcd-ef1234567890");
        public override string AuthorName => "SlicerExport";
        public override string AuthorContact => "";
    }
}
