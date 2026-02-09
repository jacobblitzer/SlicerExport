namespace SlicerExport.Core
{
    internal static class PartTypeMapping
    {
        // BambuStudio and OrcaSlicer use identical subtype strings
        public static string ToBambuSubtype(PartType partType)
        {
            switch (partType)
            {
                case PartType.Part: return "normal_part";
                case PartType.NegativePart: return "negative_part";
                case PartType.Modifier: return "modifier_part";
                case PartType.SupportBlocker: return "support_blocker";
                case PartType.SupportEnforcer: return "support_enforcer";
                default: return "normal_part";
            }
        }

        public static string ToPrusaVolumeType(PartType partType)
        {
            switch (partType)
            {
                case PartType.Part: return "ModelPart";
                case PartType.NegativePart: return "NegativeVolume";
                case PartType.Modifier: return "ParameterModifier";
                case PartType.SupportBlocker: return "SupportBlocker";
                case PartType.SupportEnforcer: return "SupportEnforcer";
                default: return "ModelPart";
            }
        }

        // BambuStudio/OrcaSlicer: Part mesh -> type="model", all others -> type="other"
        public static string To3mfObjectType(PartType partType)
        {
            return partType == PartType.Part ? "model" : "other";
        }
    }
}
