using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plant3DClipBox;

internal sealed class ClipBoxController
{
    private readonly Dictionary<Database, ClipBoxState> _states = new();

    public event EventHandler? StateChanged;

    public void Refresh()
    {
        RaiseStateChanged();
    }

    public bool HasCurrentBox()
    {
        return TryGetCurrentBox(out _);
    }

    public bool IsClippingEnabled()
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return false;
        }

        return GetState(document).ClippingEnabled;
    }

    public string GetCurrentSummary()
    {
        if (!TryGetCurrentBox(out var box))
        {
            return "No clip box in the current drawing.";
        }

        return box.ToSummaryString();
    }

    public IReadOnlyList<string> GetSavedStateNames()
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return Array.Empty<string>();
        }

        try
        {
            using var documentLock = MaybeLockDocument(document);
            return ClipBoxStorage.GetNames(document.Database);
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    public void FitSelection(bool promptIfNeeded)
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return;
        }

        var editor = document.Editor;
        var ids = GetSelectionIds(document, promptIfNeeded);
        if (ids.Length == 0)
        {
            if (!promptIfNeeded)
            {
                WriteMessage(editor, "Preselect one or more objects and click Fit selection, or run P3DCLIPBOXFIT from the command line.");
            }

            RaiseStateChanged();
            return;
        }

        try
        {
            using var documentLock = MaybeLockDocument(document);
            using var transaction = document.Database.TransactionManager.StartTransaction();

            var state = GetState(document);
            state.VisualId = ClipBoxVisuals.ResolveExistingBoxReference(document.Database, transaction, state.VisualId);

            Vector3d xAxis;
            Vector3d yAxis;
            Vector3d zAxis;

            if (ClipBoxVisuals.TryGetBox(transaction, state.VisualId, out var existingBox))
            {
                xAxis = existingBox.XAxis;
                yAxis = existingBox.YAxis;
                zAxis = existingBox.ZAxis;
            }
            else
            {
                GetAxesFromCurrentUcs(editor, out xAxis, out yAxis, out zAxis);
            }

            var extentsList = new List<Extents3d>();
            foreach (var id in ids)
            {
                if (id == state.VisualId)
                {
                    continue;
                }

                if (TryGetEntityExtents(transaction, id, out var extents))
                {
                    extentsList.Add(extents);
                }
            }

            if (extentsList.Count == 0)
            {
                WriteMessage(editor, "No usable extents were found in the current selection.");
                return;
            }

            var box = OrientedBox.FitToProjectedExtents(extentsList, xAxis, yAxis, zAxis, padding: 0.0);
            state.VisualId = ClipBoxVisuals.UpsertBoxReference(
                document.Database,
                transaction,
                state.VisualId,
                box,
                visible: !state.ClippingEnabled);

            transaction.Commit();

            if (state.ClippingEnabled)
            {
                ApplyClipping(document, state, box);
            }
            else
            {
                RequestScreenUpdate(document);
            }
        }
        catch (System.Exception ex)
        {
            WriteMessage(editor, $"Fit selection failed: {ex.Message}");
        }
        finally
        {
            RaiseStateChanged();
        }
    }

    public void ResizeAxis(BoxAxis axis, ResizeMode mode, double delta)
    {
        UpdateCurrentBox(
            box => box.Resize(axis, mode, delta),
            "No clip box found. Fit a selection first.");
    }

    public void ResizeAll(double delta)
    {
        UpdateCurrentBox(
            box => box.ResizeAll(delta),
            "No clip box found. Fit a selection first.");
    }

    public void MoveAxis(BoxAxis axis, double delta)
    {
        UpdateCurrentBox(
            box => box.Move(axis, delta),
            "No clip box found. Fit a selection first.",
            failureVerb: "Move");
    }

    public void SetClipping(bool enabled)
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return;
        }

        var editor = document.Editor;
        var state = GetState(document);

        try
        {
            using var documentLock = MaybeLockDocument(document);

            OrientedBox box = default;
            var hasBox = false;

            using (var transaction = document.Database.TransactionManager.StartTransaction())
            {
                state.VisualId = ClipBoxVisuals.ResolveExistingBoxReference(document.Database, transaction, state.VisualId);
                hasBox = ClipBoxVisuals.TryGetBox(transaction, state.VisualId, out box);

                if (enabled)
                {
                    if (!hasBox)
                    {
                        state.ClippingEnabled = false;
                        WriteMessage(editor, "No clip box found. Fit a selection first.");
                        return;
                    }

                    ClipBoxVisuals.SetBoxVisible(transaction, state.VisualId, false);
                }
                else if (hasBox)
                {
                    ClipBoxVisuals.SetBoxVisible(transaction, state.VisualId, true);
                }

                transaction.Commit();
            }

            if (enabled)
            {
                if (ClipBoxGsEngine.Enable(document, box))
                {
                    state.ClippingEnabled = true;
                }
                else
                {
                    state.ClippingEnabled = false;

                    using var rollbackTransaction = document.Database.TransactionManager.StartTransaction();
                    if (hasBox)
                    {
                        ClipBoxVisuals.SetBoxVisible(rollbackTransaction, state.VisualId, true);
                    }
                    rollbackTransaction.Commit();

                    WriteMessage(editor, "Unable to enable clip-box clipping for the current box.");
                }
            }
            else
            {
                ClipBoxGsEngine.Disable(document);
                state.ClippingEnabled = false;
                RequestScreenUpdate(document);
            }
        }
        catch (System.Exception ex)
        {
            WriteMessage(editor, $"Clipping toggle failed: {ex.Message}");
        }
        finally
        {
            RaiseStateChanged();
        }
    }

    public void SaveCurrentState(string name)
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return;
        }

        if (!TryGetCurrentBox(document, out var box))
        {
            WriteMessage(document.Editor, "No clip box found. Fit a selection first.");
            RaiseStateChanged();
            return;
        }

        var normalizedName = ClipBoxStorage.NormalizeName(name);

        try
        {
            using var documentLock = MaybeLockDocument(document);
            ClipBoxStorage.Save(document.Database, normalizedName, box);
            WriteMessage(document.Editor, $"Clip box state saved as '{normalizedName}'.");
        }
        catch (System.Exception ex)
        {
            WriteMessage(document.Editor, $"Save failed: {ex.Message}");
        }
        finally
        {
            RaiseStateChanged();
        }
    }

    public void LoadState(string name)
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return;
        }

        var editor = document.Editor;
        var state = GetState(document);
        var normalizedName = ClipBoxStorage.NormalizeName(name);

        try
        {
            using var documentLock = MaybeLockDocument(document);

            if (!ClipBoxStorage.TryLoad(document.Database, normalizedName, out var box))
            {
                WriteMessage(editor, $"No saved clip box named '{normalizedName}' was found.");
                return;
            }

            using (var transaction = document.Database.TransactionManager.StartTransaction())
            {
                state.VisualId = ClipBoxVisuals.UpsertBoxReference(
                    document.Database,
                    transaction,
                    state.VisualId,
                    box,
                    visible: !state.ClippingEnabled);

                transaction.Commit();
            }

            if (state.ClippingEnabled)
            {
                ApplyClipping(document, state, box);
            }
            else
            {
                RequestScreenUpdate(document);
            }
        }
        catch (System.Exception ex)
        {
            WriteMessage(editor, $"Load failed: {ex.Message}");
        }
        finally
        {
            RaiseStateChanged();
        }
    }

    public void DeleteState(string name)
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return;
        }

        var normalizedName = ClipBoxStorage.NormalizeName(name);

        try
        {
            using var documentLock = MaybeLockDocument(document);
            ClipBoxStorage.Delete(document.Database, normalizedName);
        }
        catch (System.Exception ex)
        {
            WriteMessage(document.Editor, $"Delete failed: {ex.Message}");
        }
        finally
        {
            RaiseStateChanged();
        }
    }

    private void UpdateCurrentBox(
        Func<OrientedBox, OrientedBox> updater,
        string noBoxMessage,
        string failureVerb = "Resize")
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            return;
        }

        var editor = document.Editor;
        var state = GetState(document);

        try
        {
            using var documentLock = MaybeLockDocument(document);

            OrientedBox updatedBox;
            using (var transaction = document.Database.TransactionManager.StartTransaction())
            {
                state.VisualId = ClipBoxVisuals.ResolveExistingBoxReference(document.Database, transaction, state.VisualId);
                if (!ClipBoxVisuals.TryGetBox(transaction, state.VisualId, out var box))
                {
                    WriteMessage(editor, noBoxMessage);
                    return;
                }

                updatedBox = updater(box);
                state.VisualId = ClipBoxVisuals.UpsertBoxReference(
                    document.Database,
                    transaction,
                    state.VisualId,
                    updatedBox,
                    visible: !state.ClippingEnabled);

                transaction.Commit();
            }

            if (state.ClippingEnabled)
            {
                ApplyClipping(document, state, updatedBox);
            }
            else
            {
                RequestScreenUpdate(document);
            }
        }
        catch (System.Exception ex)
        {
            WriteMessage(editor, $"{failureVerb} failed: {ex.Message}");
        }
        finally
        {
            RaiseStateChanged();
        }
    }

    private void ApplyClipping(Document document, ClipBoxState state, OrientedBox box)
    {
        if (ClipBoxGsEngine.Enable(document, box))
        {
            state.ClippingEnabled = true;
            return;
        }

        state.ClippingEnabled = false;

        using var transaction = document.Database.TransactionManager.StartTransaction();
        ClipBoxVisuals.SetBoxVisible(transaction, state.VisualId, true);
        transaction.Commit();

        WriteMessage(document.Editor, "Unable to enable clip-box clipping for the current box.");
    }

    private bool TryGetCurrentBox(out OrientedBox box)
    {
        var document = GetActiveDocument();
        if (document is null)
        {
            box = default;
            return false;
        }

        return TryGetCurrentBox(document, out box);
    }

    private bool TryGetCurrentBox(Document document, out OrientedBox box)
    {
        box = default;

        try
        {
            using var documentLock = MaybeLockDocument(document);
            using var transaction = document.Database.TransactionManager.StartTransaction();

            var state = GetState(document);
            state.VisualId = ClipBoxVisuals.ResolveExistingBoxReference(document.Database, transaction, state.VisualId);
            return ClipBoxVisuals.TryGetBox(transaction, state.VisualId, out box);
        }
        catch
        {
            return false;
        }
    }

    private static ObjectId[] GetSelectionIds(Document document, bool promptIfNeeded)
    {
        var editor = document.Editor;

        var implied = editor.SelectImplied();
        if (implied.Status == PromptStatus.OK && implied.Value is not null)
        {
            return implied.Value.GetObjectIds();
        }

        if (!promptIfNeeded)
        {
            return Array.Empty<ObjectId>();
        }

        var selection = editor.GetSelection();
        if (selection.Status == PromptStatus.OK && selection.Value is not null)
        {
            return selection.Value.GetObjectIds();
        }

        return Array.Empty<ObjectId>();
    }

    private static void GetAxesFromCurrentUcs(Editor editor, out Vector3d xAxis, out Vector3d yAxis, out Vector3d zAxis)
    {
        var ucs = editor.CurrentUserCoordinateSystem;
        xAxis = Vector3d.XAxis.TransformBy(ucs).GetNormal();
        yAxis = Vector3d.YAxis.TransformBy(ucs).GetNormal();
        zAxis = Vector3d.ZAxis.TransformBy(ucs).GetNormal();
    }

    private static bool TryGetEntityExtents(Transaction transaction, ObjectId id, out Extents3d extents)
    {
        extents = default;

        if (id.IsNull || id.IsErased)
        {
            return false;
        }

        if (transaction.GetObject(id, OpenMode.ForRead, false) is not Entity entity)
        {
            return false;
        }

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

    private static Document? GetActiveDocument()
    {
        return AcAp.DocumentManager.MdiActiveDocument;
    }

    private ClipBoxState GetState(Document document)
    {
        if (!_states.TryGetValue(document.Database, out var state))
        {
            state = new ClipBoxState();
            _states[document.Database] = state;
        }

        return state;
    }

    private static DocumentLock? MaybeLockDocument(Document document)
    {
        return AcAp.DocumentManager.IsApplicationContext
            ? document.LockDocument()
            : null;
    }

    private static void RequestScreenUpdate(Document document)
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

    private static void WriteMessage(Editor editor, string message)
    {
        try
        {
            editor.WriteMessage(Environment.NewLine + message);
        }
        catch
        {
            // Ignore messaging failures.
        }
    }

    private void RaiseStateChanged()
    {
        StateChanged?.Invoke(this, EventArgs.Empty);
    }
}
