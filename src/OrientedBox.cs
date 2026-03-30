using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace Plant3DClipBox;

internal enum BoxAxis
{
    X,
    Y,
    Z,
}

internal enum ResizeMode
{
    NegativeFace,
    BothFaces,
    PositiveFace,
}

internal readonly struct OrientedBox
{
    private const double MinHalfLength = 0.01;

    public Point3d Center { get; }

    public Vector3d XAxis { get; }

    public Vector3d YAxis { get; }

    public Vector3d ZAxis { get; }

    public double HalfX { get; }

    public double HalfY { get; }

    public double HalfZ { get; }

    public OrientedBox(
        Point3d center,
        Vector3d xAxis,
        Vector3d yAxis,
        Vector3d zAxis,
        double halfX,
        double halfY,
        double halfZ)
    {
        var x = NormalizeOrFallback(xAxis, Vector3d.XAxis);

        var yProjected = yAxis - x.MultiplyBy(yAxis.DotProduct(x));
        if (yProjected.Length < Tolerance.Global.EqualVector)
        {
            var fallbackSeed = Math.Abs(x.DotProduct(Vector3d.ZAxis)) < 0.99
                ? Vector3d.ZAxis
                : Vector3d.YAxis;
            yProjected = fallbackSeed - x.MultiplyBy(fallbackSeed.DotProduct(x));
        }

        var y = NormalizeOrFallback(yProjected, Vector3d.YAxis);
        var z = NormalizeOrFallback(x.CrossProduct(y), Vector3d.ZAxis);

        if (z.DotProduct(zAxis) < 0.0)
        {
            z = z.Negate();
            y = NormalizeOrFallback(z.CrossProduct(x), Vector3d.YAxis);
        }
        else
        {
            y = NormalizeOrFallback(z.CrossProduct(x), Vector3d.YAxis);
        }

        Center = center;
        XAxis = x;
        YAxis = y;
        ZAxis = z;
        HalfX = NormalizeHalfLength(halfX);
        HalfY = NormalizeHalfLength(halfY);
        HalfZ = NormalizeHalfLength(halfZ);
    }

    public bool IsValid =>
        IsFinite(HalfX) && IsFinite(HalfY) && IsFinite(HalfZ) &&
        HalfX > 0.0 && HalfY > 0.0 && HalfZ > 0.0;

    public double SizeX => HalfX * 2.0;

    public double SizeY => HalfY * 2.0;

    public double SizeZ => HalfZ * 2.0;

    public Matrix3d BlockTransform => Matrix3d.AlignCoordinateSystem(
        Point3d.Origin,
        Vector3d.XAxis,
        Vector3d.YAxis,
        Vector3d.ZAxis,
        Center,
        XAxis.MultiplyBy(HalfX),
        YAxis.MultiplyBy(HalfY),
        ZAxis.MultiplyBy(HalfZ));

    public Point3d ToLocal(Point3d worldPoint)
    {
        var offset = worldPoint - Center;
        return new Point3d(
            offset.DotProduct(XAxis),
            offset.DotProduct(YAxis),
            offset.DotProduct(ZAxis));
    }

    public Point3d ToWorld(Point3d localPoint)
    {
        return Center
               + XAxis.MultiplyBy(localPoint.X)
               + YAxis.MultiplyBy(localPoint.Y)
               + ZAxis.MultiplyBy(localPoint.Z);
    }

    public Point3d[] GetCorners()
    {
        var x = XAxis.MultiplyBy(HalfX);
        var y = YAxis.MultiplyBy(HalfY);
        var z = ZAxis.MultiplyBy(HalfZ);

        return new[]
        {
            Center - x - y - z,
            Center + x - y - z,
            Center + x + y - z,
            Center - x + y - z,
            Center - x - y + z,
            Center + x - y + z,
            Center + x + y + z,
            Center - x + y + z,
        };
    }

    public Point3dCollection BuildSectionPolygon(bool clockwise)
    {
        var localPoints = clockwise
            ? new[]
            {
                new Point3d(-HalfX, -HalfY, 0.0),
                new Point3d(-HalfX, HalfY, 0.0),
                new Point3d(HalfX, HalfY, 0.0),
                new Point3d(HalfX, -HalfY, 0.0),
                new Point3d(-HalfX, -HalfY, 0.0),
            }
            : new[]
            {
                new Point3d(-HalfX, -HalfY, 0.0),
                new Point3d(HalfX, -HalfY, 0.0),
                new Point3d(HalfX, HalfY, 0.0),
                new Point3d(-HalfX, HalfY, 0.0),
                new Point3d(-HalfX, -HalfY, 0.0),
            };

        var polygon = new Point3dCollection();
        foreach (var point in localPoints)
        {
            polygon.Add(ToWorld(point));
        }

        return polygon;
    }

    public bool Intersects(Extents3d extents)
    {
        var minX = double.PositiveInfinity;
        var minY = double.PositiveInfinity;
        var minZ = double.PositiveInfinity;
        var maxX = double.NegativeInfinity;
        var maxY = double.NegativeInfinity;
        var maxZ = double.NegativeInfinity;

        foreach (var corner in EnumerateExtentsCorners(extents))
        {
            var local = ToLocal(corner);
            minX = Math.Min(minX, local.X);
            minY = Math.Min(minY, local.Y);
            minZ = Math.Min(minZ, local.Z);
            maxX = Math.Max(maxX, local.X);
            maxY = Math.Max(maxY, local.Y);
            maxZ = Math.Max(maxZ, local.Z);
        }

        return !(maxX < -HalfX || minX > HalfX ||
                 maxY < -HalfY || minY > HalfY ||
                 maxZ < -HalfZ || minZ > HalfZ);
    }

    public OrientedBox Resize(BoxAxis axis, ResizeMode mode, double delta)
    {
        var center = Center;
        var halfX = HalfX;
        var halfY = HalfY;
        var halfZ = HalfZ;

        switch (axis)
        {
            case BoxAxis.X:
                ResizeAxis(ref center, ref halfX, XAxis, mode, delta);
                break;

            case BoxAxis.Y:
                ResizeAxis(ref center, ref halfY, YAxis, mode, delta);
                break;

            case BoxAxis.Z:
                ResizeAxis(ref center, ref halfZ, ZAxis, mode, delta);
                break;
        }

        return new OrientedBox(center, XAxis, YAxis, ZAxis, halfX, halfY, halfZ);
    }

    public OrientedBox ResizeAll(double delta)
    {
        var resized = Resize(BoxAxis.X, ResizeMode.BothFaces, delta);
        resized = resized.Resize(BoxAxis.Y, ResizeMode.BothFaces, delta);
        resized = resized.Resize(BoxAxis.Z, ResizeMode.BothFaces, delta);
        return resized;
    }

    public OrientedBox Move(BoxAxis axis, double delta)
    {
        return axis switch
        {
            BoxAxis.X => Move(XAxis.MultiplyBy(delta)),
            BoxAxis.Y => Move(YAxis.MultiplyBy(delta)),
            BoxAxis.Z => Move(ZAxis.MultiplyBy(delta)),
            _ => this,
        };
    }

    public OrientedBox Move(Vector3d translation)
    {
        return new OrientedBox(Center + translation, XAxis, YAxis, ZAxis, HalfX, HalfY, HalfZ);
    }

    public string ToSummaryString(int decimals = 2)
    {
        var format = "F" + decimals;
        return $"Center: {Center.X.ToString(format)}, {Center.Y.ToString(format)}, {Center.Z.ToString(format)}{Environment.NewLine}" +
               $"Size:   {SizeX.ToString(format)} x {SizeY.ToString(format)} x {SizeZ.ToString(format)}";
    }

    public static OrientedBox FromBlockTransform(Matrix3d blockTransform)
    {
        var center = Point3d.Origin.TransformBy(blockTransform);
        var scaledX = Vector3d.XAxis.TransformBy(blockTransform);
        var scaledY = Vector3d.YAxis.TransformBy(blockTransform);
        var scaledZ = Vector3d.ZAxis.TransformBy(blockTransform);

        return new OrientedBox(
            center,
            scaledX.GetNormal(),
            scaledY.GetNormal(),
            scaledZ.GetNormal(),
            scaledX.Length,
            scaledY.Length,
            scaledZ.Length);
    }

    public static OrientedBox FitToProjectedExtents(
        IEnumerable<Extents3d> extentsCollection,
        Vector3d xAxis,
        Vector3d yAxis,
        Vector3d zAxis,
        double padding)
    {
        var x = NormalizeOrFallback(xAxis, Vector3d.XAxis);
        var y = NormalizeOrFallback(yAxis, Vector3d.YAxis);
        var z = NormalizeOrFallback(zAxis, Vector3d.ZAxis);

        var minX = double.PositiveInfinity;
        var minY = double.PositiveInfinity;
        var minZ = double.PositiveInfinity;
        var maxX = double.NegativeInfinity;
        var maxY = double.NegativeInfinity;
        var maxZ = double.NegativeInfinity;
        var found = false;

        foreach (var extents in extentsCollection)
        {
            foreach (var corner in EnumerateExtentsCorners(extents))
            {
                var vector = corner.GetAsVector();
                var px = vector.DotProduct(x);
                var py = vector.DotProduct(y);
                var pz = vector.DotProduct(z);

                minX = Math.Min(minX, px);
                minY = Math.Min(minY, py);
                minZ = Math.Min(minZ, pz);
                maxX = Math.Max(maxX, px);
                maxY = Math.Max(maxY, py);
                maxZ = Math.Max(maxZ, pz);
                found = true;
            }
        }

        if (!found)
        {
            throw new InvalidOperationException("No extents were available to fit the clip box.");
        }

        var halfX = ((maxX - minX) * 0.5) + Math.Max(0.0, padding);
        var halfY = ((maxY - minY) * 0.5) + Math.Max(0.0, padding);
        var halfZ = ((maxZ - minZ) * 0.5) + Math.Max(0.0, padding);

        var center = Point3d.Origin
                     + x.MultiplyBy((minX + maxX) * 0.5)
                     + y.MultiplyBy((minY + maxY) * 0.5)
                     + z.MultiplyBy((minZ + maxZ) * 0.5);

        return new OrientedBox(center, x, y, z, halfX, halfY, halfZ);
    }

    public static IEnumerable<Point3d> EnumerateExtentsCorners(Extents3d extents)
    {
        var min = extents.MinPoint;
        var max = extents.MaxPoint;

        yield return new Point3d(min.X, min.Y, min.Z);
        yield return new Point3d(max.X, min.Y, min.Z);
        yield return new Point3d(max.X, max.Y, min.Z);
        yield return new Point3d(min.X, max.Y, min.Z);
        yield return new Point3d(min.X, min.Y, max.Z);
        yield return new Point3d(max.X, min.Y, max.Z);
        yield return new Point3d(max.X, max.Y, max.Z);
        yield return new Point3d(min.X, max.Y, max.Z);
    }

    private static void ResizeAxis(
        ref Point3d center,
        ref double halfLength,
        Vector3d axisVector,
        ResizeMode mode,
        double delta)
    {
        switch (mode)
        {
            case ResizeMode.BothFaces:
                halfLength = NormalizeHalfLength(halfLength + delta);
                break;

            case ResizeMode.PositiveFace:
            {
                var newHalf = NormalizeHalfLength(halfLength + (delta * 0.5));
                var actualDelta = (newHalf - halfLength) * 2.0;
                halfLength = newHalf;
                center = center + axisVector.MultiplyBy(actualDelta * 0.5);
                break;
            }

            case ResizeMode.NegativeFace:
            {
                var newHalf = NormalizeHalfLength(halfLength + (delta * 0.5));
                var actualDelta = (newHalf - halfLength) * 2.0;
                halfLength = newHalf;
                center = center - axisVector.MultiplyBy(actualDelta * 0.5);
                break;
            }
        }
    }

    private static Vector3d NormalizeOrFallback(Vector3d vector, Vector3d fallback)
    {
        return vector.Length < Tolerance.Global.EqualVector
            ? fallback
            : vector.GetNormal();
    }

    private static double NormalizeHalfLength(double value)
    {
        return Math.Max(MinHalfLength, Math.Abs(value));
    }

    private static bool IsFinite(double value)
    {
        return !double.IsNaN(value) && !double.IsInfinity(value);
    }
}
