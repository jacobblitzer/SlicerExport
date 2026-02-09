using Grasshopper.Kernel.Types;
using SlicerExport.Core;

namespace SlicerExport.Grasshopper
{
    public class GH_ExportObject : GH_Goo<ExportObject>
    {
        public GH_ExportObject() { }
        public GH_ExportObject(ExportObject obj) : base(obj) { }

        public override bool IsValid => Value != null && Value.Mesh != null;
        public override string TypeName => "Export Object";
        public override string TypeDescription => "Mesh with slicer part type";

        public override IGH_Goo Duplicate()
        {
            return new GH_ExportObject(Value);
        }

        public override string ToString()
        {
            return Value == null ? "Null" : $"{Value.PartType}: {Value.Name ?? "unnamed"}";
        }
    }
}
