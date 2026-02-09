using Rhino;
using Rhino.Commands;
using Rhino.DocObjects;
using Rhino.Input.Custom;
using SlicerExport.Core;

namespace SlicerExport.Rhino
{
    public class TagPartTypeCommand : Command
    {
        public override string EnglishName => "TagPartType";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            var go = new GetObject();
            go.SetCommandPrompt("Select objects to tag with a part type");
            go.GeometryFilter = ObjectType.Mesh | ObjectType.Brep;
            go.SubObjectSelect = false;
            go.GetMultiple(1, 0);

            if (go.CommandResult() != Result.Success)
                return go.CommandResult();

            var gpt = new GetOption();
            gpt.SetCommandPrompt("Select part type");
            gpt.AddOption("Part");
            gpt.AddOption("NegativePart");
            gpt.AddOption("Modifier");
            gpt.AddOption("SupportBlocker");
            gpt.AddOption("SupportEnforcer");
            gpt.Get();

            if (gpt.CommandResult() != Result.Success)
                return gpt.CommandResult();

            string partTypeStr = gpt.Option().EnglishName;

            int count = 0;
            foreach (var objRef in go.Objects())
            {
                var rhinoObj = objRef.Object();
                rhinoObj.Attributes.SetUserString("SlicerExport:PartType", partTypeStr);
                rhinoObj.CommitChanges();
                count++;
            }

            RhinoApp.WriteLine($"Tagged {count} object(s) as {partTypeStr}.");
            return Result.Success;
        }
    }
}
