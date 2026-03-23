using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Special;
using Rhino;
using Rhino.Geometry;
using SlicerExport.Core;
using SlicerExport.Rhino;

namespace SlicerExport.Grasshopper
{
    public class CreateExportObjectComponent : GH_Component
    {
        public CreateExportObjectComponent()
            : base("Create Export Object", "ExObj",
                "Create an export object from a mesh with a part type",
                "Params", "Util")
        { }

        public override Guid ComponentGuid => new Guid("b2c3d4e5-f6a7-8901-bcde-f12345678901");

        protected override Bitmap Icon
        {
            get
            {
                var stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("SlicerExport.Grasshopper.Resources.OrcaSlicer_24.png");
                return stream != null ? new Bitmap(stream) : null;
            }
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddMeshParameter("Mesh", "M", "Mesh geometry", GH_ParamAccess.item);
            pManager.AddIntegerParameter("PartType", "T", "Part type (0=Part, 1=NegativePart, 2=Modifier, 3=SupportBlocker, 4=SupportEnforcer)", GH_ParamAccess.item, 0);
            pManager.AddTextParameter("Name", "N", "Object name", GH_ParamAccess.item, "Object");
            pManager[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("ExportObject", "O", "Export object with mesh and part type", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Mesh mesh = null;
            int partTypeInt = 0;
            string name = "Object";

            if (!DA.GetData(0, ref mesh)) return;
            DA.GetData(1, ref partTypeInt);
            DA.GetData(2, ref name);

            if (mesh == null || !mesh.IsValid)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid mesh.");
                return;
            }

            if (partTypeInt < 0 || partTypeInt > 4)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "PartType must be 0-4.");
                return;
            }

            double scale = RhinoMath.UnitScale(RhinoDoc.ActiveDoc.ModelUnitSystem, UnitSystem.Millimeters);
            var simpleMesh = RhinoMeshConverter.ToSimpleMesh(mesh.DuplicateMesh(), scale);

            var exportObj = new ExportObject
            {
                Mesh = simpleMesh,
                PartType = (PartType)partTypeInt,
                Name = name
            };

            DA.SetData(0, new GH_ExportObject(exportObj));
        }

        public override void AddedToDocument(GH_Document document)
        {
            base.AddedToDocument(document);
            AutoAttachValueList(document, 1, new Dictionary<string, string>
            {
                { "Part", "0" },
                { "NegativePart", "1" },
                { "Modifier", "2" },
                { "SupportBlocker", "3" },
                { "SupportEnforcer", "4" }
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
