namespace SlicerExport.Core
{
    public class SimpleMesh
    {
        public float[] Vertices { get; set; }   // [x0,y0,z0, x1,y1,z1, ...] in mm
        public int[] Triangles { get; set; }    // [i0,j0,k0, i1,j1,k1, ...]
        public int VertexCount => Vertices.Length / 3;
        public int TriangleCount => Triangles.Length / 3;
    }
}
