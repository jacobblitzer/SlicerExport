using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace SlicerExport.Core
{
    public static class ThreeMfWriter
    {
        private const string IdentityTransform3x4 = "1 0 0 0 1 0 0 0 1 0 0 0";
        private const string IdentityMatrix4x4 = "1 0 0 0 0 1 0 0 0 0 1 0 0 0 0 1";

        private const string NsCore = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02";
        private const string NsBambu = "http://schemas.bambulab.com/package/2021";
        private const string NsProduction = "http://schemas.microsoft.com/3dmanufacturing/production/2015/06";
        private const string NsSlic3r = "http://schemas.slic3r.org/3mf/2017/06";
        private const string NsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string NsRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string RelType3dModel = "http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel";

        public static void Write(ExportJob job)
        {
            using (var stream = File.Create(job.OutputPath))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, false, Encoding.UTF8))
            {
                WriteContentTypes(archive, job.Target);
                WriteRels(archive);

                switch (job.Target)
                {
                    case SlicerTarget.BambuStudio:
                        WriteBambuModel(archive, job);
                        WriteBambuModelSettings(archive, job);
                        break;
                    case SlicerTarget.OrcaSlicer:
                        WriteOrcaModel(archive, job);
                        WriteOrcaModelRels(archive, job);
                        WriteOrcaSubModels(archive, job);
                        WriteOrcaModelSettings(archive, job);
                        break;
                    case SlicerTarget.PrusaSlicer:
                        WritePrusaModel(archive, job);
                        WritePrusaModelConfig(archive, job);
                        break;
                }
            }
        }

        // ─── Shared helpers ────────────────────────────────────────────

        private static XmlWriter CreateXmlWriter(Stream stream)
        {
            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = true,
                IndentChars = " ",
                OmitXmlDeclaration = false
            };
            return XmlWriter.Create(stream, settings);
        }

        private static Stream CreateEntry(ZipArchive archive, string path)
        {
            var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
            return entry.Open();
        }

        private static string F(float v)
        {
            return v.ToString("G", CultureInfo.InvariantCulture);
        }

        private static void WriteMeshXml(XmlWriter w, SimpleMesh mesh)
        {
            w.WriteStartElement("mesh");

            w.WriteStartElement("vertices");
            for (int i = 0; i < mesh.VertexCount; i++)
            {
                w.WriteStartElement("vertex");
                w.WriteAttributeString("x", F(mesh.Vertices[i * 3 + 0]));
                w.WriteAttributeString("y", F(mesh.Vertices[i * 3 + 1]));
                w.WriteAttributeString("z", F(mesh.Vertices[i * 3 + 2]));
                w.WriteEndElement();
            }
            w.WriteEndElement(); // vertices

            w.WriteStartElement("triangles");
            for (int i = 0; i < mesh.TriangleCount; i++)
            {
                w.WriteStartElement("triangle");
                w.WriteAttributeString("v1", mesh.Triangles[i * 3 + 0].ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("v2", mesh.Triangles[i * 3 + 1].ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("v3", mesh.Triangles[i * 3 + 2].ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
            }
            w.WriteEndElement(); // triangles

            w.WriteEndElement(); // mesh
        }

        // ─── [Content_Types].xml ──────────────────────────────────────

        private static void WriteContentTypes(ZipArchive archive, SlicerTarget target)
        {
            using (var stream = CreateEntry(archive, "[Content_Types].xml"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("Types", NsContentTypes);

                w.WriteStartElement("Default", NsContentTypes);
                w.WriteAttributeString("Extension", "rels");
                w.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
                w.WriteEndElement();

                w.WriteStartElement("Default", NsContentTypes);
                w.WriteAttributeString("Extension", "model");
                w.WriteAttributeString("ContentType", "application/vnd.ms-package.3dmanufacturing-3dmodel+xml");
                w.WriteEndElement();

                if (target == SlicerTarget.PrusaSlicer)
                {
                    w.WriteStartElement("Default", NsContentTypes);
                    w.WriteAttributeString("Extension", "png");
                    w.WriteAttributeString("ContentType", "image/png");
                    w.WriteEndElement();
                }

                w.WriteEndElement(); // Types
                w.WriteEndDocument();
            }
        }

        // ─── _rels/.rels ──────────────────────────────────────────────

        private static void WriteRels(ZipArchive archive)
        {
            using (var stream = CreateEntry(archive, "_rels/.rels"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("Relationships", NsRelationships);

                w.WriteStartElement("Relationship", NsRelationships);
                w.WriteAttributeString("Target", "/3D/3dmodel.model");
                w.WriteAttributeString("Id", "rel-1");
                w.WriteAttributeString("Type", RelType3dModel);
                w.WriteEndElement();

                w.WriteEndElement(); // Relationships
                w.WriteEndDocument();
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // BambuStudio
        // ═══════════════════════════════════════════════════════════════

        private static void WriteBambuModel(ZipArchive archive, ExportJob job)
        {
            using (var stream = CreateEntry(archive, "3D/3dmodel.model"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("model", NsCore);
                w.WriteAttributeString("unit", "millimeter");
                w.WriteAttributeString("xml", "lang", "http://www.w3.org/XML/1998/namespace", "en-US");
                w.WriteAttributeString("xmlns", "BambuStudio", null, NsBambu);

                w.WriteStartElement("metadata", NsCore);
                w.WriteAttributeString("name", "Application");
                w.WriteString("SlicerExport");
                w.WriteEndElement();

                w.WriteStartElement("metadata", NsCore);
                w.WriteAttributeString("name", "BambuStudio:3mfVersion");
                w.WriteString("1");
                w.WriteEndElement();

                w.WriteStartElement("resources", NsCore);

                int nextId = 1;
                var groupInfos = new List<BambuGroupInfo>();

                foreach (var group in job.Groups)
                {
                    var info = new BambuGroupInfo();
                    info.MeshIds = new List<BambuMeshInfo>();

                    foreach (var obj in group.Objects)
                    {
                        int meshId = nextId++;
                        info.MeshIds.Add(new BambuMeshInfo { Id = meshId, Object = obj });

                        w.WriteStartElement("object", NsCore);
                        w.WriteAttributeString("id", meshId.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("type", PartTypeMapping.To3mfObjectType(obj.PartType));
                        WriteMeshXml(w, obj.Mesh);
                        w.WriteEndElement(); // object
                    }

                    int groupId = nextId++;
                    info.GroupId = groupId;
                    info.GroupName = group.Name;

                    w.WriteStartElement("object", NsCore);
                    w.WriteAttributeString("id", groupId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("type", "model");

                    w.WriteStartElement("components", NsCore);
                    foreach (var mi in info.MeshIds)
                    {
                        w.WriteStartElement("component", NsCore);
                        w.WriteAttributeString("objectid", mi.Id.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("transform", IdentityTransform3x4);
                        w.WriteEndElement();
                    }
                    w.WriteEndElement(); // components

                    w.WriteEndElement(); // object (group)

                    groupInfos.Add(info);
                }

                w.WriteEndElement(); // resources

                w.WriteStartElement("build", NsCore);
                foreach (var info in groupInfos)
                {
                    w.WriteStartElement("item", NsCore);
                    w.WriteAttributeString("objectid", info.GroupId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("printable", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement(); // build

                w.WriteEndElement(); // model
                w.WriteEndDocument();
            }
        }

        private static void WriteBambuModelSettings(ZipArchive archive, ExportJob job)
        {
            WriteBambuOrOrcaModelSettings(archive, job);
        }

        private static void WriteBambuOrOrcaModelSettings(ZipArchive archive, ExportJob job)
        {
            using (var stream = CreateEntry(archive, "Metadata/model_settings.config"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("config");

                int nextId = 1;
                int plateObjectIndex = 0;
                var groupIds = new List<int>();

                foreach (var group in job.Groups)
                {
                    var meshIds = new List<int>();
                    foreach (var obj in group.Objects)
                    {
                        meshIds.Add(nextId++);
                    }

                    int groupId = nextId++;
                    groupIds.Add(groupId);

                    w.WriteStartElement("object");
                    w.WriteAttributeString("id", groupId.ToString(CultureInfo.InvariantCulture));

                    WriteMetadata(w, "name", group.Name ?? "Object");
                    WriteMetadata(w, "extruder", "1");

                    for (int i = 0; i < group.Objects.Count; i++)
                    {
                        var obj = group.Objects[i];
                        w.WriteStartElement("part");
                        w.WriteAttributeString("id", meshIds[i].ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("subtype", PartTypeMapping.ToBambuSubtype(obj.PartType));

                        WriteMetadata(w, "name", obj.Name ?? "Volume");
                        WriteMetadata(w, "matrix", IdentityMatrix4x4);

                        if (obj.PartType != PartType.Part)
                        {
                            WriteMetadata(w, "extruder", "0");
                        }

                        WriteMeshStat(w);

                        w.WriteEndElement(); // part
                    }

                    w.WriteEndElement(); // object
                    plateObjectIndex++;
                }

                // plate element
                w.WriteStartElement("plate");
                WriteMetadata(w, "plater_id", "1");
                WriteMetadata(w, "locked", "false");

                for (int i = 0; i < groupIds.Count; i++)
                {
                    w.WriteStartElement("model_instance");
                    WriteMetadata(w, "object_id", groupIds[i].ToString(CultureInfo.InvariantCulture));
                    WriteMetadata(w, "instance_id", "0");
                    w.WriteEndElement();
                }

                w.WriteEndElement(); // plate

                w.WriteEndElement(); // config
                w.WriteEndDocument();
            }
        }

        private static void WriteMetadata(XmlWriter w, string key, string value)
        {
            w.WriteStartElement("metadata");
            w.WriteAttributeString("key", key);
            w.WriteAttributeString("value", value);
            w.WriteEndElement();
        }

        private static void WriteMeshStat(XmlWriter w)
        {
            w.WriteStartElement("mesh_stat");
            w.WriteAttributeString("edges_fixed", "0");
            w.WriteAttributeString("degenerate_facets", "0");
            w.WriteAttributeString("facets_removed", "0");
            w.WriteAttributeString("facets_reversed", "0");
            w.WriteAttributeString("backwards_edges", "0");
            w.WriteEndElement();
        }

        // ═══════════════════════════════════════════════════════════════
        // OrcaSlicer
        // ═══════════════════════════════════════════════════════════════

        private static string GenerateUuid(int seed)
        {
            // Deterministic UUID based on seed for reproducible output
            return string.Format(CultureInfo.InvariantCulture,
                "{0:x8}-71cb-4c03-9d28-80fed5dfa1dc", seed);
        }

        private static void WriteOrcaModel(ZipArchive archive, ExportJob job)
        {
            using (var stream = CreateEntry(archive, "3D/3dmodel.model"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("model", NsCore);
                w.WriteAttributeString("unit", "millimeter");
                w.WriteAttributeString("xml", "lang", "http://www.w3.org/XML/1998/namespace", "en-US");
                w.WriteAttributeString("xmlns", "BambuStudio", null, NsBambu);
                w.WriteAttributeString("xmlns", "p", null, NsProduction);
                w.WriteAttributeString("requiredextensions", "p");

                w.WriteStartElement("metadata", NsCore);
                w.WriteAttributeString("name", "Application");
                w.WriteString("SlicerExport");
                w.WriteEndElement();

                w.WriteStartElement("metadata", NsCore);
                w.WriteAttributeString("name", "BambuStudio:3mfVersion");
                w.WriteString("1");
                w.WriteEndElement();

                w.WriteStartElement("resources", NsCore);

                // OrcaSlicer: group IDs start after all mesh IDs across all groups
                // Each group gets its own sub-model file with mesh objects starting at id=1
                int groupIdCounter = 1;
                // First pass: count to determine group IDs
                // Each group's sub-model has local mesh IDs starting from 1
                // Group IDs in the main model are sequential

                var groupInfos = new List<OrcaGroupInfo>();
                foreach (var group in job.Groups)
                {
                    var info = new OrcaGroupInfo();
                    info.GroupId = groupIdCounter++;
                    info.GroupName = group.Name ?? "Object";
                    info.SubModelPath = string.Format(CultureInfo.InvariantCulture,
                        "/3D/Objects/{0}_{1}.model", SanitizeFileName(info.GroupName), info.GroupId);
                    info.Objects = group.Objects;
                    groupInfos.Add(info);
                }

                foreach (var info in groupInfos)
                {
                    w.WriteStartElement("object", NsCore);
                    w.WriteAttributeString("id", info.GroupId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("p", "UUID", NsProduction, GenerateUuid(info.GroupId));
                    w.WriteAttributeString("type", "model");

                    w.WriteStartElement("components", NsCore);
                    for (int i = 0; i < info.Objects.Count; i++)
                    {
                        int localMeshId = i + 1;
                        w.WriteStartElement("component", NsCore);
                        w.WriteAttributeString("p", "path", NsProduction, info.SubModelPath);
                        w.WriteAttributeString("objectid", localMeshId.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("p", "UUID", NsProduction,
                            GenerateUuid(info.GroupId * 10000 + localMeshId));
                        w.WriteAttributeString("transform", IdentityTransform3x4);
                        w.WriteEndElement();
                    }
                    w.WriteEndElement(); // components

                    w.WriteEndElement(); // object
                }

                w.WriteEndElement(); // resources

                w.WriteStartElement("build", NsCore);
                w.WriteAttributeString("p", "UUID", NsProduction, GenerateUuid(99999));
                foreach (var info in groupInfos)
                {
                    w.WriteStartElement("item", NsCore);
                    w.WriteAttributeString("objectid", info.GroupId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("p", "UUID", NsProduction, GenerateUuid(info.GroupId + 100000));
                    w.WriteAttributeString("printable", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement(); // build

                w.WriteEndElement(); // model
                w.WriteEndDocument();
            }
        }

        private static void WriteOrcaModelRels(ZipArchive archive, ExportJob job)
        {
            using (var stream = CreateEntry(archive, "3D/_rels/3dmodel.model.rels"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("Relationships", NsRelationships);

                int relId = 1;
                var seen = new HashSet<string>();
                foreach (var group in job.Groups)
                {
                    string groupName = SanitizeFileName(group.Name ?? "Object");
                    int groupId = relId; // group IDs are sequential from 1
                    string subModelPath = string.Format(CultureInfo.InvariantCulture,
                        "/3D/Objects/{0}_{1}.model", groupName, groupId);

                    if (seen.Add(subModelPath))
                    {
                        w.WriteStartElement("Relationship", NsRelationships);
                        w.WriteAttributeString("Target", subModelPath);
                        w.WriteAttributeString("Id", "rel-" + relId.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("Type", RelType3dModel);
                        w.WriteEndElement();
                    }
                    relId++;
                }

                w.WriteEndElement(); // Relationships
                w.WriteEndDocument();
            }
        }

        private static void WriteOrcaSubModels(ZipArchive archive, ExportJob job)
        {
            int groupIdCounter = 1;
            foreach (var group in job.Groups)
            {
                int groupId = groupIdCounter++;
                string groupName = SanitizeFileName(group.Name ?? "Object");
                string entryPath = string.Format(CultureInfo.InvariantCulture,
                    "3D/Objects/{0}_{1}.model", groupName, groupId);

                using (var stream = CreateEntry(archive, entryPath))
                using (var w = CreateXmlWriter(stream))
                {
                    w.WriteStartDocument();
                    w.WriteStartElement("model", NsCore);
                    w.WriteAttributeString("unit", "millimeter");
                    w.WriteAttributeString("xml", "lang", "http://www.w3.org/XML/1998/namespace", "en-US");
                    w.WriteAttributeString("xmlns", "BambuStudio", null, NsBambu);
                    w.WriteAttributeString("xmlns", "p", null, NsProduction);
                    w.WriteAttributeString("requiredextensions", "p");

                    w.WriteStartElement("metadata", NsCore);
                    w.WriteAttributeString("name", "BambuStudio:3mfVersion");
                    w.WriteString("1");
                    w.WriteEndElement();

                    w.WriteStartElement("resources", NsCore);

                    for (int i = 0; i < group.Objects.Count; i++)
                    {
                        var obj = group.Objects[i];
                        int localId = i + 1;

                        w.WriteStartElement("object", NsCore);
                        w.WriteAttributeString("id", localId.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("p", "UUID", NsProduction,
                            GenerateUuid(groupId * 10000 + localId));
                        w.WriteAttributeString("type", "model");
                        WriteMeshXml(w, obj.Mesh);
                        w.WriteEndElement(); // object
                    }

                    w.WriteEndElement(); // resources

                    // Empty build element
                    w.WriteStartElement("build", NsCore);
                    w.WriteEndElement();

                    w.WriteEndElement(); // model
                    w.WriteEndDocument();
                }
            }
        }

        private static void WriteOrcaModelSettings(ZipArchive archive, ExportJob job)
        {
            // OrcaSlicer uses the same model_settings.config format as BambuStudio
            // but the IDs reference the group IDs (which are sequential from 1)
            // and part IDs reference the local mesh IDs in the sub-model files
            using (var stream = CreateEntry(archive, "Metadata/model_settings.config"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("config");

                int groupIdCounter = 1;
                var groupIds = new List<int>();

                foreach (var group in job.Groups)
                {
                    int groupId = groupIdCounter++;
                    groupIds.Add(groupId);

                    w.WriteStartElement("object");
                    w.WriteAttributeString("id", groupId.ToString(CultureInfo.InvariantCulture));

                    WriteMetadata(w, "name", group.Name ?? "Object");
                    WriteMetadata(w, "extruder", "1");

                    for (int i = 0; i < group.Objects.Count; i++)
                    {
                        var obj = group.Objects[i];
                        int localMeshId = i + 1;

                        w.WriteStartElement("part");
                        w.WriteAttributeString("id", localMeshId.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("subtype", PartTypeMapping.ToBambuSubtype(obj.PartType));

                        WriteMetadata(w, "name", obj.Name ?? "Volume");
                        WriteMetadata(w, "matrix", IdentityMatrix4x4);

                        if (obj.PartType != PartType.Part)
                        {
                            WriteMetadata(w, "extruder", "0");
                        }

                        WriteMeshStat(w);

                        w.WriteEndElement(); // part
                    }

                    w.WriteEndElement(); // object
                }

                // plate
                w.WriteStartElement("plate");
                WriteMetadata(w, "plater_id", "1");
                WriteMetadata(w, "locked", "false");

                for (int i = 0; i < groupIds.Count; i++)
                {
                    w.WriteStartElement("model_instance");
                    WriteMetadata(w, "object_id", groupIds[i].ToString(CultureInfo.InvariantCulture));
                    WriteMetadata(w, "instance_id", "0");
                    w.WriteEndElement();
                }

                w.WriteEndElement(); // plate

                w.WriteEndElement(); // config
                w.WriteEndDocument();
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // PrusaSlicer
        // ═══════════════════════════════════════════════════════════════

        private static void WritePrusaModel(ZipArchive archive, ExportJob job)
        {
            using (var stream = CreateEntry(archive, "3D/3dmodel.model"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("model", NsCore);
                w.WriteAttributeString("unit", "millimeter");
                w.WriteAttributeString("xml", "lang", "http://www.w3.org/XML/1998/namespace", "en-US");
                w.WriteAttributeString("xmlns", "slic3rpe", null, NsSlic3r);

                w.WriteStartElement("metadata", NsCore);
                w.WriteAttributeString("name", "slic3rpe:Version3mf");
                w.WriteString("1");
                w.WriteEndElement();

                w.WriteStartElement("metadata", NsCore);
                w.WriteAttributeString("name", "Application");
                w.WriteString("SlicerExport");
                w.WriteEndElement();

                w.WriteStartElement("resources", NsCore);

                // PrusaSlicer: each group becomes one merged object
                int objectId = 1;
                foreach (var group in job.Groups)
                {
                    w.WriteStartElement("object", NsCore);
                    w.WriteAttributeString("id", objectId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("type", "model");

                    w.WriteStartElement("mesh");

                    // Merge all vertices
                    w.WriteStartElement("vertices");
                    foreach (var obj in group.Objects)
                    {
                        for (int i = 0; i < obj.Mesh.VertexCount; i++)
                        {
                            w.WriteStartElement("vertex");
                            w.WriteAttributeString("x", F(obj.Mesh.Vertices[i * 3 + 0]));
                            w.WriteAttributeString("y", F(obj.Mesh.Vertices[i * 3 + 1]));
                            w.WriteAttributeString("z", F(obj.Mesh.Vertices[i * 3 + 2]));
                            w.WriteEndElement();
                        }
                    }
                    w.WriteEndElement(); // vertices

                    // Merge all triangles with offset vertex indices
                    w.WriteStartElement("triangles");
                    int vertexOffset = 0;
                    foreach (var obj in group.Objects)
                    {
                        for (int i = 0; i < obj.Mesh.TriangleCount; i++)
                        {
                            w.WriteStartElement("triangle");
                            w.WriteAttributeString("v1",
                                (obj.Mesh.Triangles[i * 3 + 0] + vertexOffset).ToString(CultureInfo.InvariantCulture));
                            w.WriteAttributeString("v2",
                                (obj.Mesh.Triangles[i * 3 + 1] + vertexOffset).ToString(CultureInfo.InvariantCulture));
                            w.WriteAttributeString("v3",
                                (obj.Mesh.Triangles[i * 3 + 2] + vertexOffset).ToString(CultureInfo.InvariantCulture));
                            w.WriteEndElement();
                        }
                        vertexOffset += obj.Mesh.VertexCount;
                    }
                    w.WriteEndElement(); // triangles

                    w.WriteEndElement(); // mesh
                    w.WriteEndElement(); // object

                    objectId++;
                }

                w.WriteEndElement(); // resources

                w.WriteStartElement("build", NsCore);
                objectId = 1;
                foreach (var group in job.Groups)
                {
                    w.WriteStartElement("item", NsCore);
                    w.WriteAttributeString("objectid", objectId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("printable", "1");
                    w.WriteEndElement();
                    objectId++;
                }
                w.WriteEndElement(); // build

                w.WriteEndElement(); // model
                w.WriteEndDocument();
            }
        }

        private static void WritePrusaModelConfig(ZipArchive archive, ExportJob job)
        {
            using (var stream = CreateEntry(archive, "Metadata/Slic3r_PE_model.config"))
            using (var w = CreateXmlWriter(stream))
            {
                w.WriteStartDocument();
                w.WriteStartElement("config");

                int objectId = 1;
                foreach (var group in job.Groups)
                {
                    w.WriteStartElement("object");
                    w.WriteAttributeString("id", objectId.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("instances_count", "1");

                    // object-level metadata
                    w.WriteStartElement("metadata");
                    w.WriteAttributeString("type", "object");
                    w.WriteAttributeString("key", "name");
                    w.WriteAttributeString("value", group.Name ?? "Object");
                    w.WriteEndElement();

                    // volumes with triangle index ranges
                    int triangleOffset = 0;
                    foreach (var obj in group.Objects)
                    {
                        int firstId = triangleOffset;
                        int lastId = triangleOffset + obj.Mesh.TriangleCount - 1;

                        w.WriteStartElement("volume");
                        w.WriteAttributeString("firstid", firstId.ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("lastid", lastId.ToString(CultureInfo.InvariantCulture));

                        WritePrusaVolumeMetadata(w, "name", obj.Name ?? "Volume");
                        WritePrusaVolumeMetadata(w, "volume_type", PartTypeMapping.ToPrusaVolumeType(obj.PartType));

                        if (obj.PartType == PartType.Modifier)
                        {
                            WritePrusaVolumeMetadata(w, "modifier", "1");
                        }

                        WritePrusaVolumeMetadata(w, "matrix", IdentityMatrix4x4);
                        WritePrusaVolumeMetadata(w, "extruder", "0");

                        // mesh stats
                        w.WriteStartElement("mesh");
                        w.WriteAttributeString("edges_fixed", "0");
                        w.WriteAttributeString("degenerate_facets", "0");
                        w.WriteAttributeString("facets_removed", "0");
                        w.WriteAttributeString("facets_reversed", "0");
                        w.WriteAttributeString("backwards_edges", "0");
                        w.WriteEndElement();

                        w.WriteEndElement(); // volume

                        triangleOffset += obj.Mesh.TriangleCount;
                    }

                    w.WriteEndElement(); // object
                    objectId++;
                }

                w.WriteEndElement(); // config
                w.WriteEndDocument();
            }
        }

        private static void WritePrusaVolumeMetadata(XmlWriter w, string key, string value)
        {
            w.WriteStartElement("metadata");
            w.WriteAttributeString("type", "volume");
            w.WriteAttributeString("key", key);
            w.WriteAttributeString("value", value);
            w.WriteEndElement();
        }

        // ─── Utility ──────────────────────────────────────────────────

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "Object";

            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
            {
                if (char.IsLetterOrDigit(c) || c == '_' || c == '-')
                    sb.Append(c);
                else
                    sb.Append('_');
            }
            return sb.ToString();
        }

        // ─── Internal info types ──────────────────────────────────────

        private class BambuGroupInfo
        {
            public int GroupId;
            public string GroupName;
            public List<BambuMeshInfo> MeshIds;
        }

        private class BambuMeshInfo
        {
            public int Id;
            public ExportObject Object;
        }

        private class OrcaGroupInfo
        {
            public int GroupId;
            public string GroupName;
            public string SubModelPath;
            public List<ExportObject> Objects;
        }
    }
}
