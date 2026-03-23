using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using Grasshopper.Kernel;
using SlicerExport.Core;

namespace SlicerExport.Grasshopper
{
    public class CreateExportGroupComponent : GH_Component
    {
        public CreateExportGroupComponent()
            : base("Create Export Group", "ExGrp",
                "Create an export group from export objects (exactly one Part required)",
                "Params", "Util")
        { }

        public override Guid ComponentGuid => new Guid("c3d4e5f6-a7b8-9012-cdef-123456789012");

        protected override Bitmap Icon
        {
            get
            {
                var stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("SlicerExport.Grasshopper.Resources.BambuStudio_24.png");
                return stream != null ? new Bitmap(stream) : null;
            }
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Objects", "O", "List of export objects", GH_ParamAccess.list);
            pManager.AddTextParameter("Name", "N", "Group name", GH_ParamAccess.item, "Group");
            pManager[1].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("ExportGroup", "G", "Export group", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var gooList = new List<GH_ExportObject>();
            string name = "Group";

            if (!DA.GetDataList(0, gooList)) return;
            DA.GetData(1, ref name);

            var group = new ExportGroup { Name = name };
            foreach (var goo in gooList)
            {
                if (goo?.Value != null)
                    group.Objects.Add(goo.Value);
            }

            if (!group.IsValid(out string error))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, error);
                return;
            }

            DA.SetData(0, new GH_ExportGroup(group));
        }
    }
}
