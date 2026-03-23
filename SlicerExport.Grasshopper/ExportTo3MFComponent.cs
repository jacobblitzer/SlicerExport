using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Special;
using SlicerExport.Core;

namespace SlicerExport.Grasshopper
{
    public class ExportTo3MFComponent : GH_Component
    {
        public ExportTo3MFComponent()
            : base("Export To 3MF", "Ex3MF",
                "Export groups to a 3MF file for a target slicer",
                "Params", "Util")
        { }

        public override Guid ComponentGuid => new Guid("d4e5f6a7-b8c9-0123-defa-234567890123");

        protected override Bitmap Icon
        {
            get
            {
                var stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("SlicerExport.Grasshopper.Resources.PrusaSlicer_24.png");
                return stream != null ? new Bitmap(stream) : null;
            }
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Groups", "G", "List of export groups", GH_ParamAccess.list);
            pManager.AddIntegerParameter("SlicerTarget", "T", "Target slicer (0=PrusaSlicer, 1=OrcaSlicer, 2=BambuStudio)", GH_ParamAccess.item, 0);
            pManager.AddTextParameter("FilePath", "P", "Output file path (.3mf)", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "R", "Set to true to write the file", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Launch", "L", "Launch slicer after export", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("Success", "Ok", "True if export succeeded", GH_ParamAccess.item);
            pManager.AddTextParameter("Message", "Msg", "Status message", GH_ParamAccess.item);
            pManager.AddTextParameter("FilePath", "F", "Written file path", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var gooList = new List<GH_ExportGroup>();
            int targetInt = 0;
            string filePath = "";
            bool run = false;
            bool launch = false;

            if (!DA.GetDataList(0, gooList)) return;
            DA.GetData(1, ref targetInt);
            if (!DA.GetData(2, ref filePath)) return;
            DA.GetData(3, ref run);
            DA.GetData(4, ref launch);

            if (!filePath.EndsWith(".3mf", StringComparison.OrdinalIgnoreCase))
                filePath += ".3mf";

            if (!run)
            {
                DA.SetData(0, false);
                DA.SetData(1, "Set Run to true to export.");
                DA.SetData(2, null);
                return;
            }

            if (targetInt < 0 || targetInt > 2)
            {
                DA.SetData(0, false);
                DA.SetData(1, "SlicerTarget must be 0-2.");
                DA.SetData(2, null);
                return;
            }

            var groups = new List<ExportGroup>();
            foreach (var goo in gooList)
            {
                if (goo?.Value != null)
                    groups.Add(goo.Value);
            }

            if (groups.Count == 0)
            {
                DA.SetData(0, false);
                DA.SetData(1, "No valid groups to export.");
                DA.SetData(2, null);
                return;
            }

            var target = (SlicerTarget)targetInt;

            try
            {
                var job = new ExportJob
                {
                    Groups = groups,
                    Target = target,
                    OutputPath = filePath
                };

                ThreeMfWriter.Write(job);

                DA.SetData(0, true);
                DA.SetData(1, $"Exported {groups.Count} group(s) to {filePath}");
                DA.SetData(2, filePath);

                if (launch)
                {
                    if (!SlicerLauncher.Launch(target, filePath))
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Could not find slicer executable.");
                    }
                }
            }
            catch (Exception ex)
            {
                DA.SetData(0, false);
                DA.SetData(1, $"Export failed: {ex.Message}");
                DA.SetData(2, null);
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
            }
        }

        public override void AddedToDocument(GH_Document document)
        {
            base.AddedToDocument(document);
            AutoAttachValueList(document, 1, new Dictionary<string, string>
            {
                { "PrusaSlicer", "0" },
                { "OrcaSlicer", "1" },
                { "BambuStudio", "2" }
            });
        }

        private void AutoAttachValueList(GH_Document doc, int paramIndex, Dictionary<string, string> values)
        {
            var param = Params.Input[paramIndex];
            if (param.SourceCount > 0) return;

            // Ensure component layout is computed so input grips have valid positions
            Attributes.ExpireLayout();
            Attributes.PerformLayout();

            var valueList = new GH_ValueList();
            valueList.CreateAttributes();
            valueList.NickName = param.NickName;
            valueList.ListItems.Clear();

            foreach (var kvp in values)
            {
                valueList.ListItems.Add(new GH_ValueListItem(kvp.Key, kvp.Value));
            }

            // Add to document and connect first, then compute layout for accurate bounds
            doc.AddObject(valueList, false);
            param.AddSource(valueList);

            valueList.Attributes.ExpireLayout();
            valueList.Attributes.PerformLayout();

            // Position to the left of the input grip
            valueList.Attributes.Pivot = new System.Drawing.PointF(
                param.Attributes.InputGrip.X - valueList.Attributes.Bounds.Width - 30,
                param.Attributes.InputGrip.Y - valueList.Attributes.Bounds.Height / 2);
        }
    }
}
