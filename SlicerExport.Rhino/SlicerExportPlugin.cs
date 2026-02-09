using Rhino.PlugIns;

namespace SlicerExport.Rhino
{
    public class SlicerExportPlugin : PlugIn
    {
        public static SlicerExportPlugin Instance { get; private set; }

        protected override LoadReturnCode OnLoad(ref string errorMessage)
        {
            Instance = this;
            return LoadReturnCode.Success;
        }
    }
}
