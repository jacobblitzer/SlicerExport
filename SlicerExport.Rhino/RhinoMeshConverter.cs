using Rhino.Geometry;
using SlicerExport.Core;

namespace SlicerExport.Rhino
{
    public static class RhinoMeshConverter
    {
        public static SimpleMesh ToSimpleMesh(Mesh mesh, double unitScaleToMm)
        {
            mesh.Faces.ConvertQuadsToTriangles();
            mesh.Compact();

            var vertices = new float[mesh.Vertices.Count * 3];
            for (int i = 0; i < mesh.Vertices.Count; i++)
            {
                vertices[i * 3 + 0] = (float)(mesh.Vertices[i].X * unitScaleToMm);
                vertices[i * 3 + 1] = (float)(mesh.Vertices[i].Y * unitScaleToMm);
                vertices[i * 3 + 2] = (float)(mesh.Vertices[i].Z * unitScaleToMm);
            }

            var triangles = new int[mesh.Faces.Count * 3];
            for (int i = 0; i < mesh.Faces.Count; i++)
            {
                triangles[i * 3 + 0] = mesh.Faces[i].A;
                triangles[i * 3 + 1] = mesh.Faces[i].B;
                triangles[i * 3 + 2] = mesh.Faces[i].C;
            }

            return new SimpleMesh { Vertices = vertices, Triangles = triangles };
        }
    }
}
