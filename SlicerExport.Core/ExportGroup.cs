using System.Collections.Generic;

namespace SlicerExport.Core
{
    public class ExportGroup
    {
        public string Name { get; set; }
        public List<ExportObject> Objects { get; set; } = new List<ExportObject>();

        public bool IsValid(out string error)
        {
            int partCount = 0;
            foreach (var obj in Objects)
            {
                if (obj.PartType == PartType.Part)
                    partCount++;
            }

            if (partCount == 0)
            {
                error = "Group needs at least one Part.";
                return false;
            }

            if (partCount > 1)
            {
                error = "Group must have exactly one Part.";
                return false;
            }

            error = null;
            return true;
        }
    }
}
