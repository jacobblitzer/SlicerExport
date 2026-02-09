using System.Collections.Generic;
using Rhino;
using Rhino.Commands;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Input;
using Rhino.Input.Custom;
using SlicerExport.Core;

namespace SlicerExport.Rhino
{
    public class ExportToSlicerCommand : Command
    {
        public override string EnglishName => "ExportToSlicer";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // 1. Select objects
            var go = new GetObject();
            go.SetCommandPrompt("Select meshes or breps to export");
            go.GeometryFilter = ObjectType.Mesh | ObjectType.Brep;
            go.SubObjectSelect = false;
            go.GetMultiple(1, 0);

            if (go.CommandResult() != Result.Success)
                return go.CommandResult();

            double unitScale = RhinoMath.UnitScale(doc.ModelUnitSystem, UnitSystem.Millimeters);

            var exportObjects = new List<ExportObject>();

            foreach (var objRef in go.Objects())
            {
                Mesh mesh = null;
                var rhinoObj = objRef.Object();

                if (objRef.Mesh() != null)
                {
                    mesh = objRef.Mesh();
                }
                else if (objRef.Brep() != null)
                {
                    var meshes = Mesh.CreateFromBrep(objRef.Brep(), MeshingParameters.Default);
                    if (meshes != null && meshes.Length > 0)
                    {
                        mesh = new Mesh();
                        foreach (var m in meshes)
                            mesh.Append(m);
                    }
                }

                if (mesh == null)
                {
                    RhinoApp.WriteLine($"Skipping object: could not get mesh.");
                    continue;
                }

                // Read PartType from UserString or prompt
                PartType partType = PartType.Part;
                string userPartType = rhinoObj.Attributes.GetUserString("SlicerExport:PartType");
                if (!string.IsNullOrEmpty(userPartType))
                {
                    System.Enum.TryParse(userPartType, out partType);
                }
                else
                {
                    var gpt = new GetOption();
                    gpt.SetCommandPrompt($"Part type for '{rhinoObj.Name ?? rhinoObj.Id.ToString()}'");
                    gpt.AddOption("Part");
                    gpt.AddOption("NegativePart");
                    gpt.AddOption("Modifier");
                    gpt.AddOption("SupportBlocker");
                    gpt.AddOption("SupportEnforcer");
                    gpt.SetDefaultString("Part");
                    gpt.Get();

                    if (gpt.CommandResult() != Result.Success)
                        return gpt.CommandResult();

                    switch (gpt.Option().EnglishName)
                    {
                        case "NegativePart": partType = PartType.NegativePart; break;
                        case "Modifier": partType = PartType.Modifier; break;
                        case "SupportBlocker": partType = PartType.SupportBlocker; break;
                        case "SupportEnforcer": partType = PartType.SupportEnforcer; break;
                        default: partType = PartType.Part; break;
                    }
                }

                var simpleMesh = RhinoMeshConverter.ToSimpleMesh(mesh, unitScale);
                exportObjects.Add(new ExportObject
                {
                    Mesh = simpleMesh,
                    PartType = partType,
                    Name = rhinoObj.Name ?? "Object"
                });
            }

            if (exportObjects.Count == 0)
            {
                RhinoApp.WriteLine("No valid objects to export.");
                return Result.Nothing;
            }

            // 2. Select slicer target
            var gs = new GetOption();
            gs.SetCommandPrompt("Select slicer target");
            gs.AddOption("PrusaSlicer");
            gs.AddOption("OrcaSlicer");
            gs.AddOption("BambuStudio");
            gs.Get();
            if (gs.CommandResult() != Result.Success)
                return gs.CommandResult();

            SlicerTarget target;
            switch (gs.Option().EnglishName)
            {
                case "OrcaSlicer": target = SlicerTarget.OrcaSlicer; break;
                case "BambuStudio": target = SlicerTarget.BambuStudio; break;
                default: target = SlicerTarget.PrusaSlicer; break;
            }

            // 3. File path
            string path = RhinoGet.GetFileName(GetFileNameMode.Save, "export.3mf",
                "3MF Files (*.3mf)|*.3mf", null);
            if (string.IsNullOrEmpty(path))
                return Result.Cancel;

            // 4. Build ExportJob
            var group = new ExportGroup
            {
                Name = System.IO.Path.GetFileNameWithoutExtension(path),
                Objects = exportObjects
            };

            if (!group.IsValid(out string validationError))
            {
                RhinoApp.WriteLine($"Validation error: {validationError}");
                return Result.Failure;
            }

            var job = new ExportJob
            {
                Groups = new List<ExportGroup> { group },
                Target = target,
                OutputPath = path
            };

            // 5. Write
            ThreeMfWriter.Write(job);
            RhinoApp.WriteLine($"Exported to: {path}");

            // 6. Optionally launch slicer
            var gl = new GetOption();
            gl.SetCommandPrompt("Launch slicer?");
            gl.AddOption("Yes");
            gl.AddOption("No");
            gl.SetDefaultString("No");
            gl.Get();

            if (gl.CommandResult() == Result.Success && gl.Option().EnglishName == "Yes")
            {
                if (!SlicerLauncher.Launch(target, path))
                    RhinoApp.WriteLine("Could not find slicer executable.");
            }

            return Result.Success;
        }
    }
}
