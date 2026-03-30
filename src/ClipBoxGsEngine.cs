using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcFilters = Autodesk.AutoCAD.DatabaseServices.Filters;

namespace Plant3DClipBox;

internal static class ClipBoxGsEngine
{
    private const string FilterDictionaryName = "ACAD_FILTER";
    private const string SpatialFilterEntryName = "SPATIAL";
    private const double MinClipDepth = 1.0e-6;
    private const double BoxContainmentTolerance = 1.0e-8;
    private const double VisibilityMarginAbsolute = 1.0e-6;
    private const double VisibilityMarginRelative = 1.0e-6;

    private static readonly object SyncRoot = new();
    private static readonly Dictionary<Database, OrientedBox> ActiveBoxes = new();
    private static readonly Dictionary<Database, HashSet<ObjectId>> HiddenByUs = new();
    private static readonly Dictionary<Database, List<XrefClipSnapshot>> XrefChanges = new();
    private static readonly HashSet<Database> HookedDatabases = new();

    private sealed class XrefClipSnapshot
    {
        public ObjectId BlockReferenceId { get; set; } = ObjectId.Null;
        public bool OriginalVisible { get; set; }
        public bool HadSpatialFilter { get; set; }
        public ObjectId WorkingSpatialFilterId { get; set; } = ObjectId.Null;
        public AcFilters.SpatialFilterDefinition OriginalDefinition { get; set; }
        public bool OriginalInverted { get; set; }
    }

    public static bool Enable(Document document, OrientedBox box)
    {
        if (document is null || !box.IsValid)
        {
            return false;
        }

        lock (SyncRoot)
        {
            ActiveBoxes[document.Database] = box;
            EnsureDocumentHooks(document);
        }

        return ApplyCurrentBox(document);
    }

    public static void Disable(Document document)
    {
        if (document is null)
        {
            return;
        }

        lock (SyncRoot)
        {
            ActiveBoxes.Remove(document.Database);
        }

        RestoreHidden(document, showClipBox: true);
        RefreshScreen(document);
    }

    public static void Shutdown()
    {
        try
        {
            foreach (Document document in AcAp.DocumentManager)
            {
                try
                {
                    RestoreHidden(document, showClipBox: true);
                }
                catch
                {
                    // Ignore cleanup failures during shutdown.
                }
            }
        }
        catch
        {
            // Ignore document enumeration failures during shutdown.
        }

        lock (SyncRoot)
        {
            ActiveBoxes.Clear();
            HiddenByUs.Clear();
            XrefChanges.Clear();
        }
    }

    private static bool ApplyCurrentBox(Document document)
    {
        if (!TryGetActiveBox(document.Database, out var box))
        {
            return false;
        }

        HideClipBox(document);
        var applied = ApplyVisibilityFilter(document, box);
        RefreshScreen(document);
        return applied;
    }

    private static void EnsureDocumentHooks(Document document)
    {
        if (!HookedDatabases.Add(document.Database))
        {
            return;
        }

        document.CommandEnded += (_, _) => OnDocumentCommandEnded(document);
        document.BeginDocumentClose += (_, _) => OnBeforeDocumentClose(document);
        document.Database.BeginSave += (_, _) => OnBeforeDatabaseSave(document);
        document.Database.SaveComplete += (_, _) => OnAfterDatabaseSave(document);
    }

    private static void OnDocumentCommandEnded(Document document)
    {
        if (!TryGetActiveBox(document.Database, out _))
        {
            return;
        }

        ApplyCurrentBox(document);
    }

    private static void OnBeforeDatabaseSave(Document document)
    {
        if (!TryGetActiveBox(document.Database, out _))
        {
            return;
        }

        RestoreHidden(document, showClipBox: true);
    }

    private static void OnAfterDatabaseSave(Document document)
    {
        if (!TryGetActiveBox(document.Database, out _))
        {
            return;
        }

        ApplyCurrentBox(document);
    }

    private static void OnBeforeDocumentClose(Document document)
    {
        lock (SyncRoot)
        {
            ActiveBoxes.Remove(document.Database);
        }

        RestoreHidden(document, showClipBox: true);
    }

    private static bool TryGetActiveBox(Database database, out OrientedBox box)
    {
        lock (SyncRoot)
        {
            return ActiveBoxes.TryGetValue(database, out box) && box.IsValid;
        }
    }

    private static bool ApplyVisibilityFilter(Document document, OrientedBox box)
    {
        if (!RestoreHidden(document, showClipBox: false))
        {
            return false;
        }

        var newlyHidden = new HashSet<ObjectId>();
        var xrefChanges = new List<XrefClipSnapshot>();

        try
        {
            using var documentLock = MaybeLockDocument(document);
            using var transaction = document.Database.TransactionManager.StartOpenCloseTransaction();

            var blockTable = (BlockTable)transaction.GetObject(document.Database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in modelSpace)
            {
                if (id.IsNull || id.IsErased || IsClipBoxReference(transaction, id))
                {
                    continue;
                }

                if (transaction.GetObject(id, OpenMode.ForRead, false) is not Entity entity)
                {
                    continue;
                }

                if (entity is BlockReference blockReference && IsExternalReference(transaction, blockReference))
                {
                    ApplyXrefClip(transaction, blockReference, box, xrefChanges);
                    continue;
                }

                if (ShouldKeepVisible(entity, box))
                {
                    continue;
                }

                if (!entity.Visible)
                {
                    continue;
                }

                try
                {
                    entity.UpgradeOpen();
                    entity.Visible = false;
                    entity.RecordGraphicsModified(true);
                    newlyHidden.Add(id);
                }
                catch
                {
                    // Leave entities visible if they cannot be hidden.
                }
            }

            transaction.Commit();

            lock (SyncRoot)
            {
                HiddenByUs[document.Database] = newlyHidden;
                XrefChanges[document.Database] = xrefChanges;
            }

            return true;
        }
        catch
        {
            lock (SyncRoot)
            {
                HiddenByUs.Remove(document.Database);
                XrefChanges.Remove(document.Database);
            }

            return false;
        }
    }

    private static bool RestoreHidden(Document document, bool showClipBox)
    {
        HashSet<ObjectId>? idsToRestore;
        List<XrefClipSnapshot>? xrefsToRestore;

        lock (SyncRoot)
        {
            idsToRestore = HiddenByUs.TryGetValue(document.Database, out var hidden)
                ? new HashSet<ObjectId>(hidden)
                : null;

            xrefsToRestore = XrefChanges.TryGetValue(document.Database, out var xrefs)
                ? new List<XrefClipSnapshot>(xrefs)
                : null;
        }

        var needsWork =
            (idsToRestore is not null && idsToRestore.Count > 0) ||
            (xrefsToRestore is not null && xrefsToRestore.Count > 0) ||
            showClipBox;

        if (!needsWork)
        {
            return true;
        }

        var committed = false;
        var remainingIds = new HashSet<ObjectId>();
        var remainingXrefs = new List<XrefClipSnapshot>();

        try
        {
            using var documentLock = MaybeLockDocument(document);
            using var transaction = document.Database.TransactionManager.StartOpenCloseTransaction();

            if (idsToRestore is not null)
            {
                foreach (var id in idsToRestore)
                {
                    if (!RestoreHiddenEntity(transaction, id))
                    {
                        remainingIds.Add(id);
                    }
                }
            }

            if (xrefsToRestore is not null)
            {
                foreach (var snapshot in xrefsToRestore)
                {
                    if (!RestoreXrefClip(transaction, snapshot))
                    {
                        remainingXrefs.Add(snapshot);
                    }
                }
            }

            if (showClipBox)
            {
                try
                {
                    var clipBoxId = ClipBoxVisuals.ResolveExistingBoxReference(document.Database, transaction, ObjectId.Null);
                    ClipBoxVisuals.SetBoxVisible(transaction, clipBoxId, true);
                }
                catch
                {
                    // The box is only a visual helper. Do not block restore if it cannot be shown.
                }
            }

            transaction.Commit();
            committed = true;
        }
        catch
        {
            return false;
        }

        if (committed)
        {
            lock (SyncRoot)
            {
                if (remainingIds.Count > 0)
                {
                    HiddenByUs[document.Database] = remainingIds;
                }
                else
                {
                    HiddenByUs.Remove(document.Database);
                }

                if (remainingXrefs.Count > 0)
                {
                    XrefChanges[document.Database] = remainingXrefs;
                }
                else
                {
                    XrefChanges.Remove(document.Database);
                }
            }
        }

        return committed;
    }

    private static bool RestoreHiddenEntity(Transaction transaction, ObjectId id)
    {
        try
        {
            if (id.IsNull || id.IsErased)
            {
                return true;
            }

            if (transaction.GetObject(id, OpenMode.ForRead, false) is not Entity entity)
            {
                return true;
            }

            if (entity.Visible)
            {
                return true;
            }

            entity.UpgradeOpen();
            entity.Visible = true;
            entity.RecordGraphicsModified(true);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void HideClipBox(Document document)
    {
        try
        {
            using var documentLock = MaybeLockDocument(document);
            using var transaction = document.Database.TransactionManager.StartOpenCloseTransaction();
            var clipBoxId = ClipBoxVisuals.ResolveExistingBoxReference(document.Database, transaction, ObjectId.Null);
            ClipBoxVisuals.SetBoxVisible(transaction, clipBoxId, false);
            transaction.Commit();
        }
        catch
        {
            // Ignore visibility failures. The box is only a visual aid.
        }
    }

    private static void ApplyXrefClip(
        Transaction transaction,
        BlockReference blockReference,
        OrientedBox box,
        List<XrefClipSnapshot> xrefChanges)
    {
        if (!blockReference.Visible)
        {
            return;
        }

        if (!TryBuildXrefSpatialFilterDefinition(blockReference, box, out var filterDefinition))
        {
            return;
        }

        var snapshot = CaptureXrefSnapshot(transaction, blockReference);
        if (!ReplaceSpatialFilter(transaction, blockReference, filterDefinition, out var filterId))
        {
            return;
        }

        if (!blockReference.IsWriteEnabled)
        {
            blockReference.UpgradeOpen();
        }

        blockReference.RecordGraphicsModified(true);
        snapshot.WorkingSpatialFilterId = filterId;
        xrefChanges.Add(snapshot);
    }

    private static bool RestoreXrefClip(Transaction transaction, XrefClipSnapshot snapshot)
    {
        try
        {
            if (snapshot.BlockReferenceId.IsNull || snapshot.BlockReferenceId.IsErased)
            {
                return true;
            }

            if (transaction.GetObject(snapshot.BlockReferenceId, OpenMode.ForRead, false) is not BlockReference blockReference)
            {
                return true;
            }

            if (snapshot.HadSpatialFilter)
            {
                if (!TryGetSpatialFilter(transaction, blockReference, OpenMode.ForWrite, out _, out var filter, out _))
                {
                    if (!EnsureSpatialFilter(transaction, blockReference, out filter, out _))
                    {
                        return false;
                    }
                }

                filter.Definition = snapshot.OriginalDefinition;
                filter.Inverted = snapshot.OriginalInverted;
            }
            else
            {
                if (!RemoveSpatialFilter(transaction, blockReference, snapshot.WorkingSpatialFilterId))
                {
                    return false;
                }
            }

            if (blockReference.Visible != snapshot.OriginalVisible)
            {
                if (!blockReference.IsWriteEnabled)
                {
                    blockReference.UpgradeOpen();
                }

                blockReference.Visible = snapshot.OriginalVisible;
            }

            if (!blockReference.IsWriteEnabled)
            {
                blockReference.UpgradeOpen();
            }

            blockReference.RecordGraphicsModified(true);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static XrefClipSnapshot CaptureXrefSnapshot(Transaction transaction, BlockReference blockReference)
    {
        var snapshot = new XrefClipSnapshot
        {
            BlockReferenceId = blockReference.ObjectId,
            OriginalVisible = blockReference.Visible,
        };

        if (TryGetSpatialFilter(transaction, blockReference, OpenMode.ForRead, out _, out var filter, out var filterId) && filter is not null)
        {
            snapshot.HadSpatialFilter = true;
            snapshot.WorkingSpatialFilterId = filterId;
            snapshot.OriginalDefinition = filter.Definition;
            snapshot.OriginalInverted = filter.Inverted;
        }

        return snapshot;
    }

    private static bool TryBuildXrefSpatialFilterDefinition(
        BlockReference blockReference,
        OrientedBox box,
        out AcFilters.SpatialFilterDefinition definition)
    {
        definition = default;

        try
        {
            var inverse = blockReference.BlockTransform.Inverse();

            var centerLocal = box.Center.TransformBy(inverse);
            var frontCenterLocal = box.ToWorld(new Point3d(0.0, 0.0, box.HalfZ)).TransformBy(inverse);
            var backCenterLocal = box.ToWorld(new Point3d(0.0, 0.0, -box.HalfZ)).TransformBy(inverse);

            var p00 = box.ToWorld(new Point3d(-box.HalfX, -box.HalfY, 0.0)).TransformBy(inverse);
            var p10 = box.ToWorld(new Point3d(box.HalfX, -box.HalfY, 0.0)).TransformBy(inverse);
            var p11 = box.ToWorld(new Point3d(box.HalfX, box.HalfY, 0.0)).TransformBy(inverse);
            var p01 = box.ToWorld(new Point3d(-box.HalfX, box.HalfY, 0.0)).TransformBy(inverse);

            var edgeX = p10 - p00;
            var edgeY = p01 - p00;
            if (edgeX.Length < Tolerance.Global.EqualVector || edgeY.Length < Tolerance.Global.EqualVector)
            {
                return false;
            }

            var normalLocal = edgeX.CrossProduct(edgeY);
            if (normalLocal.Length < Tolerance.Global.EqualVector)
            {
                return false;
            }

            normalLocal = normalLocal.GetNormal();

            if ((frontCenterLocal - centerLocal).DotProduct(normalLocal) < 0.0)
            {
                normalLocal = normalLocal.Negate();
            }

            var worldToPlane = Matrix3d.WorldToPlane(normalLocal);
            var centerPlane = centerLocal.TransformBy(worldToPlane);
            var p00Plane = p00.TransformBy(worldToPlane);
            var p10Plane = p10.TransformBy(worldToPlane);
            var p11Plane = p11.TransformBy(worldToPlane);
            var p01Plane = p01.TransformBy(worldToPlane);

            var boundary = new Point2dCollection
            {
                new Point2d(p00Plane.X, p00Plane.Y),
                new Point2d(p10Plane.X, p10Plane.Y),
                new Point2d(p11Plane.X, p11Plane.Y),
                new Point2d(p01Plane.X, p01Plane.Y),
            };

            var frontClip = (frontCenterLocal - centerLocal).DotProduct(normalLocal);
            var backClip = (backCenterLocal - centerLocal).DotProduct(normalLocal);

            if (frontClip < MinClipDepth)
            {
                frontClip = MinClipDepth;
            }

            if (backClip > -MinClipDepth)
            {
                backClip = -MinClipDepth;
            }

            definition = new AcFilters.SpatialFilterDefinition(
                boundary,
                normalLocal,
                centerPlane.Z,
                frontClip,
                backClip,
                true);

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryGetSpatialFilter(
        Transaction transaction,
        BlockReference blockReference,
        OpenMode openMode,
        out DBDictionary? filterDictionary,
        out AcFilters.SpatialFilter? filter,
        out ObjectId filterId)
    {
        filterDictionary = null;
        filter = null;
        filterId = ObjectId.Null;

        try
        {
            if (blockReference.ExtensionDictionary.IsNull)
            {
                return false;
            }

            var dictOpenMode = openMode == OpenMode.ForWrite ? OpenMode.ForWrite : OpenMode.ForRead;
            if (transaction.GetObject(blockReference.ExtensionDictionary, dictOpenMode, false) is not DBDictionary extDict)
            {
                return false;
            }

            if (!extDict.Contains(FilterDictionaryName))
            {
                return false;
            }

            filterDictionary = (DBDictionary)transaction.GetObject(extDict.GetAt(FilterDictionaryName), dictOpenMode, false);
            if (!filterDictionary.Contains(SpatialFilterEntryName))
            {
                return false;
            }

            filterId = filterDictionary.GetAt(SpatialFilterEntryName);
            filter = transaction.GetObject(filterId, openMode, false) as AcFilters.SpatialFilter;
            return filter is not null;
        }
        catch
        {
            filterDictionary = null;
            filter = null;
            filterId = ObjectId.Null;
            return false;
        }
    }

    private static bool EnsureSpatialFilter(
        Transaction transaction,
        BlockReference blockReference,
        out AcFilters.SpatialFilter filter,
        out ObjectId filterId)
    {
        if (TryGetSpatialFilter(transaction, blockReference, OpenMode.ForWrite, out _, out var existingFilter, out var existingFilterId) && existingFilter is not null)
        {
            filter = existingFilter;
            filterId = existingFilterId;
            return true;
        }

        try
        {
            if (blockReference.ExtensionDictionary.IsNull)
            {
                blockReference.UpgradeOpen();
                blockReference.CreateExtensionDictionary();
            }

            var extDict = (DBDictionary)transaction.GetObject(blockReference.ExtensionDictionary, OpenMode.ForWrite);
            DBDictionary filterDict;

            if (extDict.Contains(FilterDictionaryName))
            {
                filterDict = (DBDictionary)transaction.GetObject(extDict.GetAt(FilterDictionaryName), OpenMode.ForWrite);

                if (filterDict.Contains(SpatialFilterEntryName))
                {
                    var staleId = filterDict.GetAt(SpatialFilterEntryName);
                    filterDict.Remove(SpatialFilterEntryName);

                    if (!staleId.IsNull && !staleId.IsErased && transaction.GetObject(staleId, OpenMode.ForWrite, false) is DBObject staleObject)
                    {
                        staleObject.Erase();
                    }
                }
            }
            else
            {
                filterDict = new DBDictionary();
                extDict.SetAt(FilterDictionaryName, filterDict);
                transaction.AddNewlyCreatedDBObject(filterDict, true);
            }

            var newFilter = new AcFilters.SpatialFilter();
            filterId = filterDict.SetAt(SpatialFilterEntryName, newFilter);
            transaction.AddNewlyCreatedDBObject(newFilter, true);

            filter = newFilter;
            return true;
        }
        catch
        {
            filter = null!;
            filterId = ObjectId.Null;
            return false;
        }
    }

    private static bool ReplaceSpatialFilter(
        Transaction transaction,
        BlockReference blockReference,
        AcFilters.SpatialFilterDefinition definition,
        out ObjectId filterId)
    {
        filterId = ObjectId.Null;

        try
        {
            if (blockReference.ExtensionDictionary.IsNull)
            {
                blockReference.UpgradeOpen();
                blockReference.CreateExtensionDictionary();
            }

            var extDict = (DBDictionary)transaction.GetObject(blockReference.ExtensionDictionary, OpenMode.ForWrite);
            DBDictionary filterDict;

            if (extDict.Contains(FilterDictionaryName))
            {
                filterDict = (DBDictionary)transaction.GetObject(extDict.GetAt(FilterDictionaryName), OpenMode.ForWrite);

                if (filterDict.Contains(SpatialFilterEntryName))
                {
                    var existingId = filterDict.GetAt(SpatialFilterEntryName);
                    filterDict.Remove(SpatialFilterEntryName);

                    if (!existingId.IsNull && !existingId.IsErased && transaction.GetObject(existingId, OpenMode.ForWrite, false) is DBObject existingObject)
                    {
                        existingObject.Erase();
                    }
                }
            }
            else
            {
                filterDict = new DBDictionary();
                extDict.SetAt(FilterDictionaryName, filterDict);
                transaction.AddNewlyCreatedDBObject(filterDict, true);
            }

            var filter = new AcFilters.SpatialFilter();
            filterId = filterDict.SetAt(SpatialFilterEntryName, filter);
            transaction.AddNewlyCreatedDBObject(filter, true);

            // Assign the definition only after the filter is database-resident on the xref/block.
            // This gives AutoCAD a fresh spatial filter object for each apply, which is more reliable
            // when the clip depth or the clip volume changes.
            filter.Definition = definition;
            filter.Inverted = false;
            return true;
        }
        catch
        {
            filterId = ObjectId.Null;
            return false;
        }
    }

    private static bool RemoveSpatialFilter(Transaction transaction, BlockReference blockReference, ObjectId preferredFilterId)
    {
        try
        {
            if (TryGetSpatialFilter(transaction, blockReference, OpenMode.ForWrite, out var filterDictionary, out var filter, out var filterId) &&
                filterDictionary is not null &&
                filter is not null)
            {
                filterDictionary.Remove(SpatialFilterEntryName);

                if (!filterId.IsNull && !filterId.IsErased)
                {
                    filter.Erase();
                }

                return true;
            }

            if (!preferredFilterId.IsNull && !preferredFilterId.IsErased &&
                transaction.GetObject(preferredFilterId, OpenMode.ForWrite, false) is DBObject filterObject)
            {
                filterObject.Erase();
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool ShouldKeepVisible(Entity entity, OrientedBox box)
    {
        if (!TryGetEntityExtents(entity, out var extents))
        {
            return true;
        }

        extents = ExpandExtents(extents, ComputeVisibilityMargin(box, extents));

        if (IsExtentsFullyInsideBox(box, extents))
        {
            return true;
        }

        return box.Intersects(extents);
    }

    private static bool IsExtentsFullyInsideBox(OrientedBox box, Extents3d extents)
    {
        foreach (var corner in OrientedBox.EnumerateExtentsCorners(extents))
        {
            var local = box.ToLocal(corner);
            if (local.X < -box.HalfX - BoxContainmentTolerance || local.X > box.HalfX + BoxContainmentTolerance ||
                local.Y < -box.HalfY - BoxContainmentTolerance || local.Y > box.HalfY + BoxContainmentTolerance ||
                local.Z < -box.HalfZ - BoxContainmentTolerance || local.Z > box.HalfZ + BoxContainmentTolerance)
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetEntityExtents(Entity entity, out Extents3d extents)
    {
        extents = default;

        try
        {
            if (entity is BlockReference blockReference)
            {
                try
                {
                    extents = blockReference.GeometryExtentsBestFit();
                    return true;
                }
                catch
                {
                    // Fall back to GeometricExtents below.
                }
            }

            extents = entity.GeometricExtents;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static double ComputeVisibilityMargin(OrientedBox box, Extents3d extents)
    {
        var extentsSizeX = Math.Abs(extents.MaxPoint.X - extents.MinPoint.X);
        var extentsSizeY = Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y);
        var extentsSizeZ = Math.Abs(extents.MaxPoint.Z - extents.MinPoint.Z);

        var maxDimension = Math.Max(
            Math.Max(box.SizeX, Math.Max(box.SizeY, box.SizeZ)),
            Math.Max(extentsSizeX, Math.Max(extentsSizeY, extentsSizeZ)));

        return Math.Max(VisibilityMarginAbsolute, maxDimension * VisibilityMarginRelative);
    }

    private static Extents3d ExpandExtents(Extents3d extents, double margin)
    {
        if (margin <= 0.0)
        {
            return extents;
        }

        return new Extents3d(
            new Point3d(
                extents.MinPoint.X - margin,
                extents.MinPoint.Y - margin,
                extents.MinPoint.Z - margin),
            new Point3d(
                extents.MaxPoint.X + margin,
                extents.MaxPoint.Y + margin,
                extents.MaxPoint.Z + margin));
    }

    private static bool IsExternalReference(Transaction transaction, BlockReference blockReference)
    {
        if (transaction.GetObject(blockReference.BlockTableRecord, OpenMode.ForRead, false) is not BlockTableRecord definition)
        {
            return false;
        }

        return definition.IsFromExternalReference || definition.IsFromOverlayReference;
    }

    private static bool IsClipBoxReference(Transaction transaction, ObjectId id)
    {
        if (id.IsNull || id.IsErased)
        {
            return false;
        }

        if (transaction.GetObject(id, OpenMode.ForRead, false) is not BlockReference blockReference)
        {
            return false;
        }

        if (transaction.GetObject(blockReference.BlockTableRecord, OpenMode.ForRead, false) is not BlockTableRecord definition)
        {
            return false;
        }

        return string.Equals(
            definition.Name,
            ClipBoxConstants.UnitCubeBlockName,
            StringComparison.OrdinalIgnoreCase);
    }

    private static DocumentLock? MaybeLockDocument(Document document)
    {
        return AcAp.DocumentManager.IsApplicationContext
            ? document.LockDocument()
            : null;
    }

    private static void RefreshScreen(Document document)
    {
        try
        {
            document.Editor.UpdateScreen();
            document.Editor.Regen();
        }
        catch
        {
            // Ignore redraw failures in restrictive contexts.
        }
    }
}
