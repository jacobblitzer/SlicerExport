using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using SlicerExport.Core;
using Xunit;

namespace SlicerExport.Tests
{
    public class ThreeMfWriterTests : IDisposable
    {
        private readonly List<string> _tempFiles = new List<string>();

        private string GetTempPath()
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".3mf");
            _tempFiles.Add(path);
            return path;
        }

        public void Dispose()
        {
            foreach (var f in _tempFiles)
            {
                try { File.Delete(f); } catch { }
            }
        }

        /// <summary>
        /// Creates a simple cube mesh with 8 vertices and 12 triangles.
        /// </summary>
        private static SimpleMesh CreateCubeMesh(float size = 10f)
        {
            float h = size / 2f;
            return new SimpleMesh
            {
                Vertices = new float[]
                {
                     h,  h, -h,   h, -h, -h,  -h, -h, -h,  -h,  h, -h,
                     h,  h,  h,  -h,  h,  h,  -h, -h,  h,   h, -h,  h
                },
                Triangles = new int[]
                {
                    0,1,2, 0,2,3, 4,5,6, 4,6,7,
                    0,4,7, 0,7,1, 1,7,6, 1,6,2,
                    2,6,5, 2,5,3, 4,0,3, 4,3,5
                }
            };
        }

        /// <summary>
        /// Creates a simple triangle mesh (1 triangle, 3 vertices).
        /// </summary>
        private static SimpleMesh CreateTriangleMesh()
        {
            return new SimpleMesh
            {
                Vertices = new float[] { 0, 0, 0, 1, 0, 0, 0, 1, 0 },
                Triangles = new int[] { 0, 1, 2 }
            };
        }

        private static ExportGroup CreateAllPartTypesGroup()
        {
            return new ExportGroup
            {
                Name = "TestObject",
                Objects = new List<ExportObject>
                {
                    new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Part, Name = "Body" },
                    new ExportObject { Mesh = CreateCubeMesh(8f), PartType = PartType.NegativePart, Name = "NegVolume" },
                    new ExportObject { Mesh = CreateCubeMesh(6f), PartType = PartType.Modifier, Name = "Modifier" },
                    new ExportObject { Mesh = CreateCubeMesh(4f), PartType = PartType.SupportBlocker, Name = "Blocker" },
                    new ExportObject { Mesh = CreateCubeMesh(2f), PartType = PartType.SupportEnforcer, Name = "Enforcer" }
                }
            };
        }

        private static XDocument LoadEntryXml(ZipArchive archive, string entryName)
        {
            var entry = archive.GetEntry(entryName);
            Assert.NotNull(entry);
            using (var stream = entry.Open())
            {
                return XDocument.Load(stream);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // BambuStudio Tests
        // ═══════════════════════════════════════════════════════════════

        [Fact]
        public void BambuStudio_AllPartTypes()
        {
            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup> { CreateAllPartTypesGroup() },
                Target = SlicerTarget.BambuStudio,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                // Verify 3D/3dmodel.model
                var modelDoc = LoadEntryXml(archive, "3D/3dmodel.model");
                XNamespace ns = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02";

                var objects = modelDoc.Descendants(ns + "object").ToList();
                // 5 mesh objects + 1 group object = 6
                Assert.Equal(6, objects.Count);

                // First object (Part) should be type="model"
                Assert.Equal("model", objects[0].Attribute("type")?.Value);
                Assert.Equal("1", objects[0].Attribute("id")?.Value);

                // Objects 2-5 (non-Part types) should be type="other"
                for (int i = 1; i < 5; i++)
                {
                    Assert.Equal("other", objects[i].Attribute("type")?.Value);
                }

                // Object 6 is the group with components
                var groupObj = objects[5];
                Assert.Equal("model", groupObj.Attribute("type")?.Value);
                Assert.Equal("6", groupObj.Attribute("id")?.Value);
                var components = groupObj.Descendants(ns + "component").ToList();
                Assert.Equal(5, components.Count);

                // Build item references the group
                var buildItem = modelDoc.Descendants(ns + "item").Single();
                Assert.Equal("6", buildItem.Attribute("objectid")?.Value);
                Assert.Equal("1", buildItem.Attribute("printable")?.Value);

                // Verify Metadata/model_settings.config
                var configDoc = LoadEntryXml(archive, "Metadata/model_settings.config");
                var configObj = configDoc.Root.Elements("object").Single();
                Assert.Equal("6", configObj.Attribute("id")?.Value);

                var parts = configObj.Elements("part").ToList();
                Assert.Equal(5, parts.Count);

                Assert.Equal("1", parts[0].Attribute("id")?.Value);
                Assert.Equal("normal_part", parts[0].Attribute("subtype")?.Value);

                Assert.Equal("2", parts[1].Attribute("id")?.Value);
                Assert.Equal("negative_part", parts[1].Attribute("subtype")?.Value);

                Assert.Equal("3", parts[2].Attribute("id")?.Value);
                Assert.Equal("modifier_part", parts[2].Attribute("subtype")?.Value);

                Assert.Equal("4", parts[3].Attribute("id")?.Value);
                Assert.Equal("support_blocker", parts[3].Attribute("subtype")?.Value);

                Assert.Equal("5", parts[4].Attribute("id")?.Value);
                Assert.Equal("support_enforcer", parts[4].Attribute("subtype")?.Value);

                // Non-Part types should have extruder=0
                for (int i = 1; i < 5; i++)
                {
                    var extruder = parts[i].Elements("metadata")
                        .FirstOrDefault(e => e.Attribute("key")?.Value == "extruder");
                    Assert.NotNull(extruder);
                    Assert.Equal("0", extruder.Attribute("value")?.Value);
                }

                // Part should NOT have extruder=0
                var partExtruder = parts[0].Elements("metadata")
                    .FirstOrDefault(e => e.Attribute("key")?.Value == "extruder");
                Assert.Null(partExtruder);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // PrusaSlicer Tests
        // ═══════════════════════════════════════════════════════════════

        [Fact]
        public void PrusaSlicer_AllPartTypes()
        {
            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup> { CreateAllPartTypesGroup() },
                Target = SlicerTarget.PrusaSlicer,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                // Verify 3D/3dmodel.model
                var modelDoc = LoadEntryXml(archive, "3D/3dmodel.model");
                XNamespace ns = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02";

                // Should have exactly ONE object (merged)
                var objects = modelDoc.Descendants(ns + "object").ToList();
                Assert.Single(objects);
                Assert.Equal("model", objects[0].Attribute("type")?.Value);

                // Merged mesh should have all vertices (8*5 = 40)
                var vertices = objects[0].Descendants(ns + "vertex").ToList();
                Assert.Equal(40, vertices.Count);

                // Merged mesh should have all triangles (12*5 = 60)
                var triangles = objects[0].Descendants(ns + "triangle").ToList();
                Assert.Equal(60, triangles.Count);

                // Should have slic3rpe namespace
                XNamespace nsSlic3r = "http://schemas.slic3r.org/3mf/2017/06";
                Assert.NotNull(modelDoc.Root.Attribute(XNamespace.Xmlns + "slic3rpe"));

                // Verify Metadata/Slic3r_PE_model.config
                var configDoc = LoadEntryXml(archive, "Metadata/Slic3r_PE_model.config");
                var configObj = configDoc.Root.Elements("object").Single();
                Assert.Equal("1", configObj.Attribute("id")?.Value);
                Assert.Equal("1", configObj.Attribute("instances_count")?.Value);

                var volumes = configObj.Elements("volume").ToList();
                Assert.Equal(5, volumes.Count);

                // Check volume types
                string GetVolumeType(XElement vol) =>
                    vol.Elements("metadata")
                        .First(e => e.Attribute("key")?.Value == "volume_type")
                        .Attribute("value")?.Value;

                Assert.Equal("ModelPart", GetVolumeType(volumes[0]));
                Assert.Equal("NegativeVolume", GetVolumeType(volumes[1]));
                Assert.Equal("ParameterModifier", GetVolumeType(volumes[2]));
                Assert.Equal("SupportBlocker", GetVolumeType(volumes[3]));
                Assert.Equal("SupportEnforcer", GetVolumeType(volumes[4]));

                // Modifier should have modifier=1
                var modifierMeta = volumes[2].Elements("metadata")
                    .FirstOrDefault(e => e.Attribute("key")?.Value == "modifier");
                Assert.NotNull(modifierMeta);
                Assert.Equal("1", modifierMeta.Attribute("value")?.Value);

                // Non-modifier volumes should NOT have modifier=1
                for (int i = 0; i < 5; i++)
                {
                    if (i == 2) continue;
                    var mod = volumes[i].Elements("metadata")
                        .FirstOrDefault(e => e.Attribute("key")?.Value == "modifier");
                    Assert.Null(mod);
                }
            }
        }

        [Fact]
        public void PrusaSlicer_TriangleIndexRanges()
        {
            // Create objects with known triangle counts: 12, 24, 8
            var mesh12 = CreateCubeMesh(); // 12 triangles
            var mesh24 = new SimpleMesh   // 24 triangles (two cubes concatenated)
            {
                Vertices = new float[8 * 3 * 2],
                Triangles = new int[12 * 3 * 2]
            };
            // Fill mesh24 with two cubes' worth of data
            var cube = CreateCubeMesh();
            Array.Copy(cube.Vertices, 0, mesh24.Vertices, 0, cube.Vertices.Length);
            Array.Copy(cube.Vertices, 0, mesh24.Vertices, cube.Vertices.Length, cube.Vertices.Length);
            Array.Copy(cube.Triangles, 0, mesh24.Triangles, 0, cube.Triangles.Length);
            for (int i = 0; i < cube.Triangles.Length; i++)
                mesh24.Triangles[cube.Triangles.Length + i] = cube.Triangles[i] + 8;

            // 8-triangle mesh: a simple mesh with specific triangle count
            // Use a tetrahedron-like shape with some extra triangles
            var mesh8 = new SimpleMesh
            {
                Vertices = new float[]
                {
                    0,0,0, 1,0,0, 0,1,0, 0,0,1,
                    2,0,0, 2,1,0, 2,0,1, 1,1,1
                },
                Triangles = new int[]
                {
                    0,1,2, 0,2,3, 0,1,3, 1,2,3,
                    4,5,6, 4,6,7, 5,6,7, 4,5,7
                }
            };
            Assert.Equal(8, mesh8.TriangleCount);

            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup>
                {
                    new ExportGroup
                    {
                        Name = "IndexTest",
                        Objects = new List<ExportObject>
                        {
                            new ExportObject { Mesh = mesh12, PartType = PartType.Part, Name = "V1" },
                            new ExportObject { Mesh = mesh24, PartType = PartType.NegativePart, Name = "V2" },
                            new ExportObject { Mesh = mesh8, PartType = PartType.Modifier, Name = "V3" }
                        }
                    }
                },
                Target = SlicerTarget.PrusaSlicer,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                var configDoc = LoadEntryXml(archive, "Metadata/Slic3r_PE_model.config");
                var volumes = configDoc.Root.Element("object").Elements("volume").ToList();
                Assert.Equal(3, volumes.Count);

                // Volume 1: 12 triangles → firstid=0, lastid=11
                Assert.Equal("0", volumes[0].Attribute("firstid")?.Value);
                Assert.Equal("11", volumes[0].Attribute("lastid")?.Value);

                // Volume 2: 24 triangles → firstid=12, lastid=35
                Assert.Equal("12", volumes[1].Attribute("firstid")?.Value);
                Assert.Equal("35", volumes[1].Attribute("lastid")?.Value);

                // Volume 3: 8 triangles → firstid=36, lastid=43
                Assert.Equal("36", volumes[2].Attribute("firstid")?.Value);
                Assert.Equal("43", volumes[2].Attribute("lastid")?.Value);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // OrcaSlicer Tests
        // ═══════════════════════════════════════════════════════════════

        [Fact]
        public void OrcaSlicer_AllPartTypes()
        {
            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup> { CreateAllPartTypesGroup() },
                Target = SlicerTarget.OrcaSlicer,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                XNamespace ns = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02";
                XNamespace nsP = "http://schemas.microsoft.com/3dmanufacturing/production/2015/06";

                // Verify main model has p: namespace and requiredextensions
                var modelDoc = LoadEntryXml(archive, "3D/3dmodel.model");
                Assert.NotNull(modelDoc.Root.Attribute(XNamespace.Xmlns + "p"));
                Assert.Equal("p", modelDoc.Root.Attribute("requiredextensions")?.Value);

                // Main model should have a group object with components pointing to sub-model
                var groupObj = modelDoc.Descendants(ns + "object").Single();
                Assert.NotNull(groupObj.Attribute(nsP + "UUID"));

                var components = groupObj.Descendants(ns + "component").ToList();
                Assert.Equal(5, components.Count);

                // All components should have p:path pointing to sub-model
                foreach (var comp in components)
                {
                    Assert.NotNull(comp.Attribute(nsP + "path"));
                    Assert.NotNull(comp.Attribute(nsP + "UUID"));
                    Assert.Contains("/3D/Objects/", comp.Attribute(nsP + "path")?.Value);
                }

                // Verify 3D/_rels/3dmodel.model.rels exists
                var relsDoc = LoadEntryXml(archive, "3D/_rels/3dmodel.model.rels");
                XNamespace nsRels = "http://schemas.openxmlformats.org/package/2006/relationships";
                var relationships = relsDoc.Descendants(nsRels + "Relationship").ToList();
                Assert.True(relationships.Count >= 1);
                Assert.Contains("/3D/Objects/", relationships[0].Attribute("Target")?.Value);

                // Verify sub-model file exists and has mesh objects
                var subModelPath = components[0].Attribute(nsP + "path")?.Value;
                // Remove leading / for zip entry name
                var entryName = subModelPath.TrimStart('/');
                var subModelDoc = LoadEntryXml(archive, entryName);
                var subObjects = subModelDoc.Descendants(ns + "object").ToList();
                Assert.Equal(5, subObjects.Count);

                // Each sub-object should have mesh data and p:UUID
                foreach (var obj in subObjects)
                {
                    Assert.NotNull(obj.Attribute(nsP + "UUID"));
                    Assert.NotNull(obj.Element(ns + "mesh"));
                }

                // Verify model_settings.config (same format as BambuStudio)
                var configDoc = LoadEntryXml(archive, "Metadata/model_settings.config");
                var configObj = configDoc.Root.Elements("object").Single();
                var parts = configObj.Elements("part").ToList();
                Assert.Equal(5, parts.Count);

                Assert.Equal("normal_part", parts[0].Attribute("subtype")?.Value);
                Assert.Equal("negative_part", parts[1].Attribute("subtype")?.Value);
                Assert.Equal("modifier_part", parts[2].Attribute("subtype")?.Value);
                Assert.Equal("support_blocker", parts[3].Attribute("subtype")?.Value);
                Assert.Equal("support_enforcer", parts[4].Attribute("subtype")?.Value);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // Validation Tests
        // ═══════════════════════════════════════════════════════════════

        [Fact]
        public void Validation_NoPartFails()
        {
            var group = new ExportGroup
            {
                Name = "NoPart",
                Objects = new List<ExportObject>
                {
                    new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Modifier, Name = "Mod" }
                }
            };

            Assert.False(group.IsValid(out string error));
            Assert.Contains("at least one Part", error);
        }

        [Fact]
        public void Validation_TwoPartsFails()
        {
            var group = new ExportGroup
            {
                Name = "TwoParts",
                Objects = new List<ExportObject>
                {
                    new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Part, Name = "P1" },
                    new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Part, Name = "P2" }
                }
            };

            Assert.False(group.IsValid(out string error));
            Assert.Contains("exactly one Part", error);
        }

        [Fact]
        public void Validation_SinglePartPasses()
        {
            var group = new ExportGroup
            {
                Name = "Valid",
                Objects = new List<ExportObject>
                {
                    new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Part, Name = "Body" },
                    new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Modifier, Name = "Mod" }
                }
            };

            Assert.True(group.IsValid(out string error));
            Assert.Null(error);
        }

        // ═══════════════════════════════════════════════════════════════
        // Multi-Group Tests
        // ═══════════════════════════════════════════════════════════════

        [Fact]
        public void MultiGroup_TwoGroups_BambuStudio()
        {
            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup>
                {
                    new ExportGroup
                    {
                        Name = "Group1",
                        Objects = new List<ExportObject>
                        {
                            new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Part, Name = "Body1" }
                        }
                    },
                    new ExportGroup
                    {
                        Name = "Group2",
                        Objects = new List<ExportObject>
                        {
                            new ExportObject { Mesh = CreateCubeMesh(5f), PartType = PartType.Part, Name = "Body2" },
                            new ExportObject { Mesh = CreateCubeMesh(3f), PartType = PartType.Modifier, Name = "Mod2" }
                        }
                    }
                },
                Target = SlicerTarget.BambuStudio,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                XNamespace ns = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02";
                var modelDoc = LoadEntryXml(archive, "3D/3dmodel.model");

                // Group1: 1 mesh obj (id=1) + 1 group obj (id=2)
                // Group2: 2 mesh objs (id=3,4) + 1 group obj (id=5)
                // Total: 5 objects
                var objects = modelDoc.Descendants(ns + "object").ToList();
                Assert.Equal(5, objects.Count);

                // Two build items (one per group)
                var buildItems = modelDoc.Descendants(ns + "item").ToList();
                Assert.Equal(2, buildItems.Count);
                Assert.Equal("2", buildItems[0].Attribute("objectid")?.Value);
                Assert.Equal("5", buildItems[1].Attribute("objectid")?.Value);

                // Config should have two objects
                var configDoc = LoadEntryXml(archive, "Metadata/model_settings.config");
                var configObjs = configDoc.Root.Elements("object").ToList();
                Assert.Equal(2, configObjs.Count);
            }
        }

        [Fact]
        public void MultiGroup_TwoGroups_PrusaSlicer()
        {
            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup>
                {
                    new ExportGroup
                    {
                        Name = "Group1",
                        Objects = new List<ExportObject>
                        {
                            new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Part, Name = "Body1" }
                        }
                    },
                    new ExportGroup
                    {
                        Name = "Group2",
                        Objects = new List<ExportObject>
                        {
                            new ExportObject { Mesh = CreateCubeMesh(5f), PartType = PartType.Part, Name = "Body2" }
                        }
                    }
                },
                Target = SlicerTarget.PrusaSlicer,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                XNamespace ns = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02";
                var modelDoc = LoadEntryXml(archive, "3D/3dmodel.model");

                // Two merged objects (one per group)
                var objects = modelDoc.Descendants(ns + "object").ToList();
                Assert.Equal(2, objects.Count);

                // Two build items
                var buildItems = modelDoc.Descendants(ns + "item").ToList();
                Assert.Equal(2, buildItems.Count);

                // Two config objects
                var configDoc = LoadEntryXml(archive, "Metadata/Slic3r_PE_model.config");
                var configObjs = configDoc.Root.Elements("object").ToList();
                Assert.Equal(2, configObjs.Count);
                Assert.Equal("1", configObjs[0].Attribute("id")?.Value);
                Assert.Equal("2", configObjs[1].Attribute("id")?.Value);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // Structure / Archive Tests
        // ═══════════════════════════════════════════════════════════════

        [Theory]
        [InlineData(SlicerTarget.BambuStudio)]
        [InlineData(SlicerTarget.OrcaSlicer)]
        [InlineData(SlicerTarget.PrusaSlicer)]
        public void Archive_HasRequiredEntries(SlicerTarget target)
        {
            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup> { CreateAllPartTypesGroup() },
                Target = target,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                Assert.NotNull(archive.GetEntry("[Content_Types].xml"));
                Assert.NotNull(archive.GetEntry("_rels/.rels"));
                Assert.NotNull(archive.GetEntry("3D/3dmodel.model"));

                if (target == SlicerTarget.PrusaSlicer)
                {
                    Assert.NotNull(archive.GetEntry("Metadata/Slic3r_PE_model.config"));
                }
                else
                {
                    Assert.NotNull(archive.GetEntry("Metadata/model_settings.config"));
                }

                if (target == SlicerTarget.OrcaSlicer)
                {
                    Assert.NotNull(archive.GetEntry("3D/_rels/3dmodel.model.rels"));
                    // Should have at least one sub-model
                    Assert.True(archive.Entries.Any(e => e.FullName.StartsWith("3D/Objects/")));
                }
            }
        }

        [Theory]
        [InlineData(SlicerTarget.BambuStudio)]
        [InlineData(SlicerTarget.OrcaSlicer)]
        [InlineData(SlicerTarget.PrusaSlicer)]
        public void ContentTypes_HasCorrectDefaults(SlicerTarget target)
        {
            var path = GetTempPath();
            var job = new ExportJob
            {
                Groups = new List<ExportGroup>
                {
                    new ExportGroup
                    {
                        Name = "CT",
                        Objects = new List<ExportObject>
                        {
                            new ExportObject { Mesh = CreateCubeMesh(), PartType = PartType.Part, Name = "Body" }
                        }
                    }
                },
                Target = target,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                var doc = LoadEntryXml(archive, "[Content_Types].xml");
                XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
                var defaults = doc.Descendants(ns + "Default").ToList();

                Assert.Contains(defaults, d =>
                    d.Attribute("Extension")?.Value == "rels" &&
                    d.Attribute("ContentType")?.Value == "application/vnd.openxmlformats-package.relationships+xml");

                Assert.Contains(defaults, d =>
                    d.Attribute("Extension")?.Value == "model" &&
                    d.Attribute("ContentType")?.Value == "application/vnd.ms-package.3dmanufacturing-3dmodel+xml");

                if (target == SlicerTarget.PrusaSlicer)
                {
                    Assert.Contains(defaults, d =>
                        d.Attribute("Extension")?.Value == "png" &&
                        d.Attribute("ContentType")?.Value == "image/png");
                }
            }
        }

        [Fact]
        public void PrusaSlicer_MergedVertexIndicesAreCorrect()
        {
            // Verify that triangle vertex indices in the merged mesh are correctly offset
            var path = GetTempPath();
            var mesh1 = new SimpleMesh
            {
                Vertices = new float[] { 0, 0, 0, 1, 0, 0, 0, 1, 0 }, // 3 vertices
                Triangles = new int[] { 0, 1, 2 } // 1 triangle
            };
            var mesh2 = new SimpleMesh
            {
                Vertices = new float[] { 2, 0, 0, 3, 0, 0, 2, 1, 0 }, // 3 vertices
                Triangles = new int[] { 0, 1, 2 } // 1 triangle, local indices
            };

            var job = new ExportJob
            {
                Groups = new List<ExportGroup>
                {
                    new ExportGroup
                    {
                        Name = "MergeTest",
                        Objects = new List<ExportObject>
                        {
                            new ExportObject { Mesh = mesh1, PartType = PartType.Part, Name = "M1" },
                            new ExportObject { Mesh = mesh2, PartType = PartType.Modifier, Name = "M2" }
                        }
                    }
                },
                Target = SlicerTarget.PrusaSlicer,
                OutputPath = path
            };

            ThreeMfWriter.Write(job);

            using (var archive = ZipFile.OpenRead(path))
            {
                XNamespace ns = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02";
                var modelDoc = LoadEntryXml(archive, "3D/3dmodel.model");
                var triangles = modelDoc.Descendants(ns + "triangle").ToList();
                Assert.Equal(2, triangles.Count);

                // First triangle: v1=0, v2=1, v3=2 (mesh1, no offset)
                Assert.Equal("0", triangles[0].Attribute("v1")?.Value);
                Assert.Equal("1", triangles[0].Attribute("v2")?.Value);
                Assert.Equal("2", triangles[0].Attribute("v3")?.Value);

                // Second triangle: v1=3, v2=4, v3=5 (mesh2, offset by 3)
                Assert.Equal("3", triangles[1].Attribute("v1")?.Value);
                Assert.Equal("4", triangles[1].Attribute("v2")?.Value);
                Assert.Equal("5", triangles[1].Attribute("v3")?.Value);

                // Total vertices should be 6
                var vertices = modelDoc.Descendants(ns + "vertex").ToList();
                Assert.Equal(6, vertices.Count);
            }
        }
    }
}
