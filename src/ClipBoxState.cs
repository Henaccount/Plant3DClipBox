using Autodesk.AutoCAD.DatabaseServices;

namespace Plant3DClipBox;

internal sealed class ClipBoxState
{
    public ObjectId VisualId { get; set; } = ObjectId.Null;

    public bool ClippingEnabled { get; set; }
}
