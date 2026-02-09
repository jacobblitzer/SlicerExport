using Grasshopper.Kernel.Types;
using SlicerExport.Core;

namespace SlicerExport.Grasshopper
{
    public class GH_ExportGroup : GH_Goo<ExportGroup>
    {
        public GH_ExportGroup() { }
        public GH_ExportGroup(ExportGroup obj) : base(obj) { }

        public override bool IsValid => Value != null && Value.Objects != null && Value.Objects.Count > 0;
        public override string TypeName => "Export Group";
        public override string TypeDescription => "Group of export objects with one Part parent";

        public override IGH_Goo Duplicate()
        {
            return new GH_ExportGroup(Value);
        }

        public override string ToString()
        {
            if (Value == null) return "Null";
            return $"Group: {Value.Name ?? "unnamed"} ({Value.Objects.Count} objects)";
        }
    }
}
