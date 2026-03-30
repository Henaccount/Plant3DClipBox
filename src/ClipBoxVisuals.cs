using System;
using Autodesk.AutoCAD.Colors;
using AcColor = Autodesk.AutoCAD.Colors.Color;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace Plant3DClipBox;

internal static class ClipBoxVisuals
{
    public static ObjectId ResolveExistingBoxReference(Database db, Transaction tr, ObjectId candidateId)
    {
        if (IsClipBoxReference(tr, candidateId))
        {
            return candidateId;
        }

        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        foreach (ObjectId id in modelSpace)
        {
            if (IsClipBoxReference(tr, id))
            {
                return id;
            }
        }

        return ObjectId.Null;
    }

    public static ObjectId UpsertBoxReference(Database db, Transaction tr, ObjectId existingId, OrientedBox box, bool visible)
    {
        var blockId = EnsureBlockDefinition(db, tr);
        var layerId = EnsureLayer(db, tr);
        var resolvedId = ResolveExistingBoxReference(db, tr, existingId);

        if (!resolvedId.IsNull &&
            tr.GetObject(resolvedId, OpenMode.ForRead, false) is BlockReference existingReference)
        {
            existingReference.UpgradeOpen();
            ApplyVisualSettings(existingReference, layerId, visible);
            existingReference.BlockTransform = box.BlockTransform;
            return resolvedId;
        }

        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        var reference = new BlockReference(Point3d.Origin, blockId);
        ApplyVisualSettings(reference, layerId, visible);
        reference.BlockTransform = box.BlockTransform;

        var id = modelSpace.AppendEntity(reference);
        tr.AddNewlyCreatedDBObject(reference, true);
        return id;
    }

    public static bool TryGetBox(Transaction tr, ObjectId boxId, out OrientedBox box)
    {
        box = default;

        if (!IsClipBoxReference(tr, boxId))
        {
            return false;
        }

        var blockReference = (BlockReference)tr.GetObject(boxId, OpenMode.ForRead);
        box = OrientedBox.FromBlockTransform(blockReference.BlockTransform);
        return box.IsValid;
    }

    public static void SetBoxVisible(Transaction tr, ObjectId boxId, bool visible)
    {
        if (!IsClipBoxReference(tr, boxId))
        {
            return;
        }

        var blockReference = (BlockReference)tr.GetObject(boxId, OpenMode.ForRead);
        blockReference.UpgradeOpen();
        blockReference.Visible = visible;
    }

    private static void ApplyVisualSettings(BlockReference blockReference, ObjectId layerId, bool visible)
    {
        blockReference.LayerId = layerId;
        blockReference.ColorIndex = 1;
        blockReference.LineWeight = LineWeight.LineWeight050;
        blockReference.Visible = visible;
    }

    private static bool IsClipBoxReference(Transaction tr, ObjectId id)
    {
        if (id.IsNull || id.IsErased)
        {
            return false;
        }

        if (tr.GetObject(id, OpenMode.ForRead, false) is not BlockReference blockReference)
        {
            return false;
        }

        if (tr.GetObject(blockReference.BlockTableRecord, OpenMode.ForRead, false) is not BlockTableRecord definition)
        {
            return false;
        }

        return string.Equals(
            definition.Name,
            ClipBoxConstants.UnitCubeBlockName,
            StringComparison.OrdinalIgnoreCase);
    }

    private static ObjectId EnsureLayer(Database db, Transaction tr)
    {
        var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
        if (layerTable.Has(ClipBoxConstants.ClipBoxLayerName))
        {
            return layerTable[ClipBoxConstants.ClipBoxLayerName];
        }

        layerTable.UpgradeOpen();

        var layer = new LayerTableRecord
        {
            Name = ClipBoxConstants.ClipBoxLayerName,
            IsPlottable = false,
            Color = AcColor.FromColorIndex(ColorMethod.ByAci, 1),
        };

        var id = layerTable.Add(layer);
        tr.AddNewlyCreatedDBObject(layer, true);
        return id;
    }

    private static ObjectId EnsureBlockDefinition(Database db, Transaction tr)
    {
        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        if (blockTable.Has(ClipBoxConstants.UnitCubeBlockName))
        {
            return blockTable[ClipBoxConstants.UnitCubeBlockName];
        }

        blockTable.UpgradeOpen();

        var blockRecord = new BlockTableRecord
        {
            Name = ClipBoxConstants.UnitCubeBlockName,
            Origin = Point3d.Origin,
        };

        var id = blockTable.Add(blockRecord);
        tr.AddNewlyCreatedDBObject(blockRecord, true);

        AddCubeEdges(blockRecord, tr);
        return id;
    }

    private static void AddCubeEdges(BlockTableRecord blockRecord, Transaction tr)
    {
        var p000 = new Point3d(-1.0, -1.0, -1.0);
        var p100 = new Point3d(1.0, -1.0, -1.0);
        var p110 = new Point3d(1.0, 1.0, -1.0);
        var p010 = new Point3d(-1.0, 1.0, -1.0);
        var p001 = new Point3d(-1.0, -1.0, 1.0);
        var p101 = new Point3d(1.0, -1.0, 1.0);
        var p111 = new Point3d(1.0, 1.0, 1.0);
        var p011 = new Point3d(-1.0, 1.0, 1.0);

        var segments = new[]
        {
            (p000, p100), (p100, p110), (p110, p010), (p010, p000),
            (p001, p101), (p101, p111), (p111, p011), (p011, p001),
            (p000, p001), (p100, p101), (p110, p111), (p010, p011),
        };

        foreach (var (start, end) in segments)
        {
            var line = new Line(start, end)
            {
                Layer = "0",
                ColorIndex = 0,
            };

            blockRecord.AppendEntity(line);
            tr.AddNewlyCreatedDBObject(line, true);
        }
    }
}
