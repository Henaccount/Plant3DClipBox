using System.Drawing;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;

[assembly: ExtensionApplication(typeof(Plant3DClipBox.ClipBoxPlugin))]
[assembly: CommandClass(typeof(Plant3DClipBox.ClipBoxPlugin))]

namespace Plant3DClipBox;

public sealed class ClipBoxPlugin : IExtensionApplication
{
    private static readonly ClipBoxController Controller = new();
    private static PaletteSet? _palette;
    private static ClipBoxControl? _control;

    public void Initialize()
    {
    }

    public void Terminate()
    {
        ClipBoxGsEngine.Shutdown();
        _palette?.Dispose();
        _control?.Dispose();
    }

    [CommandMethod("P3DCLIPBOX", CommandFlags.Session)]
    public void ShowPaletteCommand()
    {
        EnsurePalette();
        Controller.Refresh();
        _palette!.Visible = true;
    }

    [CommandMethod("P3DCLIPBOXFIT", CommandFlags.UsePickSet)]
    public void FitSelectionCommand()
    {
        Controller.FitSelection(promptIfNeeded: true);
    }

    [CommandMethod("P3DCLIPBOXON")]
    public void ClipOnCommand()
    {
        Controller.SetClipping(true);
    }

    [CommandMethod("P3DCLIPBOXOFF")]
    public void ClipOffCommand()
    {
        Controller.SetClipping(false);
    }

    [CommandMethod("P3DCLIPBOXREFRESH")]
    public void RefreshCommand()
    {
        Controller.Refresh();
    }

    private static void EnsurePalette()
    {
        if (_palette is not null && _control is not null)
        {
            return;
        }

        _control = new ClipBoxControl(Controller);
        _palette = new PaletteSet(ClipBoxConstants.PaletteTitle, ClipBoxConstants.PaletteGuid)
        {
            MinimumSize = new Size(470, 440),
            Size = new Size(520, 660),
            KeepFocus = false,
        };

        _palette.Add("Clip Box", _control);
    }
}
