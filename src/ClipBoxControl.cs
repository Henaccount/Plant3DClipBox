using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Plant3DClipBox;

internal sealed class ClipBoxControl : UserControl
{
    private readonly ClipBoxController _controller;
    private readonly CheckBox _chkClipping;
    private readonly CheckBox _chkNegative;
    private readonly CheckBox _chkMove;
    private readonly Label _lblSummary;
    private readonly Label _lblStatus;
    private readonly ComboBox _cmbStates;
    private readonly NumericUpDown _numX;
    private readonly NumericUpDown _numY;
    private readonly NumericUpDown _numZ;
    private readonly NumericUpDown _numAll;
    private readonly List<NumericUpDown> _deltaInputs = new();
    private readonly List<Button> _moveButtons = new();
    private readonly List<Button> _resizeOnlyButtons = new();

    private bool _updating;
    private bool _normalizingValues;
    private readonly ToolTip _toolTip;

    public ClipBoxControl(ClipBoxController controller)
    {
        _controller = controller ?? throw new ArgumentNullException(nameof(controller));
        _controller.StateChanged += Controller_StateChanged;

        Dock = DockStyle.Fill;
        AutoScroll = true;
        Padding = new Padding(8);

        _toolTip = new ToolTip
        {
            AutoPopDelay = 20000,
            InitialDelay = 250,
            ReshowDelay = 100,
            ShowAlways = true,
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 6),
        };

        var btnFit = CreateButton("Fit selection", (_, _) => Execute(() => _controller.FitSelection(promptIfNeeded: false)));
        var btnRefresh = CreateButton("Refresh", (_, _) => RefreshFromController());
        buttonPanel.Controls.Add(btnFit);
        buttonPanel.Controls.Add(btnRefresh);
        root.Controls.Add(buttonPanel, 0, 0);

        _chkClipping = new CheckBox
        {
            AutoSize = true,
            Text = "clipping on",
            Margin = new Padding(0, 0, 0, 6),
        };
        _chkClipping.CheckedChanged += (_, _) =>
        {
            if (_updating)
            {
                return;
            }

            Execute(() => _controller.SetClipping(_chkClipping.Checked));
        };
        root.Controls.Add(_chkClipping, 0, 1);

        _lblSummary = new Label
        {
            AutoSize = false,
            Height = 88,
            Dock = DockStyle.Top,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(8),
            AutoEllipsis = true,
            Cursor = Cursors.Help,
            Text = "No clip box in the current drawing.",
            Margin = new Padding(0, 0, 0, 6),
        };
        root.Controls.Add(_lblSummary, 0, 2);

        _lblStatus = new Label
        {
            AutoSize = false,
            Height = 96,
            Dock = DockStyle.Top,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(8),
            AutoEllipsis = true,
            Cursor = Cursors.Help,
            Text = string.Empty,
            Margin = new Padding(0, 0, 0, 6),
        };
        root.Controls.Add(_lblStatus, 0, 3);

        var grpResize = new GroupBox
        {
            Text = "Resize / Move",
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 6),
        };
        var resizeHost = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            Padding = new Padding(8),
        };
        resizeHost.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        var modePanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 8),
        };
        _chkNegative = new CheckBox
        {
            AutoSize = true,
            Text = "-",
            Margin = new Padding(0, 2, 10, 2),
        };
        _chkNegative.CheckedChanged += (_, _) => ApplySignModeToInputs();
        modePanel.Controls.Add(_chkNegative);

        _chkMove = new CheckBox
        {
            AutoSize = true,
            Text = "move",
            Margin = new Padding(0, 2, 10, 2),
        };
        _chkMove.CheckedChanged += (_, _) => UpdateModeUi();
        modePanel.Controls.Add(_chkMove);
        resizeHost.Controls.Add(modePanel, 0, 0);

        var resizeTable = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 5,
            AutoSize = true,
        };
        resizeTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58F));
        resizeTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
        resizeTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58F));//58
        resizeTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 74F));//74
        resizeTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58F));//58

        _numX = CreateDeltaControl();
        _numY = CreateDeltaControl();
        _numZ = CreateDeltaControl();
        _numAll = CreateDeltaControl();

        RegisterDeltaInput(_numX);
        RegisterDeltaInput(_numY);
        RegisterDeltaInput(_numZ);
        RegisterDeltaInput(_numAll);

        AddResizeRow(resizeTable, 0, "X", _numX, BoxAxis.X);
        AddResizeRow(resizeTable, 1, "Y", _numY, BoxAxis.Y);
        AddResizeRow(resizeTable, 2, "Z", _numZ, BoxAxis.Z);

        resizeTable.Controls.Add(CreateAxisLabel("All"), 0, 3);
        resizeTable.Controls.Add(_numAll, 1, 3);
        var btnAll = CreateButton("Apply", (_, _) => Execute(() => _controller.ResizeAll((double)_numAll.Value)));
        btnAll.Dock = DockStyle.Fill;
        resizeTable.Controls.Add(btnAll, 2, 3);
        resizeTable.SetColumnSpan(btnAll, 3);
        _resizeOnlyButtons.Add(btnAll);

        resizeHost.Controls.Add(resizeTable, 0, 1);
        grpResize.Controls.Add(resizeHost);
        root.Controls.Add(grpResize, 0, 4);

        var grpStates = new GroupBox
        {
            Text = "Saved boxes",
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 6),
        };
        var statesTable = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            Padding = new Padding(8),
        };
        statesTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        _cmbStates = new ComboBox
        {
            Dock = DockStyle.Top,
            DropDownStyle = ComboBoxStyle.DropDown,
            Margin = new Padding(0, 0, 0, 6),
        };
        statesTable.Controls.Add(_cmbStates, 0, 0);

        var statesButtonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            WrapContents = true,
        };
        statesButtonPanel.Controls.Add(CreateButton("Save", (_, _) => Execute(() => _controller.SaveCurrentState(GetStateName()))));
        statesButtonPanel.Controls.Add(CreateButton("Load", (_, _) => Execute(() => _controller.LoadState(GetStateName()))));
        statesButtonPanel.Controls.Add(CreateButton("Delete", (_, _) => Execute(() => _controller.DeleteState(GetStateName()))));
        grpStates.Controls.Add(statesTable);
        statesTable.Controls.Add(statesButtonPanel, 0, 1);
        root.Controls.Add(grpStates, 0, 5);

        var hint = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(440, 0),
            Text = "Turn clipping off and use native MOVE, 3DROTATE, or SCALE on the visible box to change its position or orientation. With move mode on, the highlighted outer X/Y/Z buttons shift the box along its local axes.",
            Margin = new Padding(0, 0, 0, 6),
        };
        root.Controls.Add(hint, 0, 6);

        Controls.Add(root);
        RefreshFromController();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _controller.StateChanged -= Controller_StateChanged;
            _toolTip.Dispose();
        }

        base.Dispose(disposing);
    }

    public void RefreshFromController()
    {
        if (InvokeRequired)
        {
            BeginInvoke((Action)RefreshFromController);
            return;
        }

        _updating = true;
        try
        {
            var hasBox = _controller.HasCurrentBox();
            var clippingOn = _controller.IsClippingEnabled();

            _chkClipping.Checked = clippingOn;
            _chkClipping.Enabled = hasBox;
            _lblSummary.Text = _controller.GetCurrentSummary();
            _lblStatus.Text = BuildStatusText(hasBox, clippingOn);
            _toolTip.SetToolTip(_lblSummary, _lblSummary.Text);
            _toolTip.SetToolTip(_lblStatus, _lblStatus.Text);

            var currentText = _cmbStates.Text;
            _cmbStates.BeginUpdate();
            _cmbStates.Items.Clear();
            foreach (var name in _controller.GetSavedStateNames())
            {
                _cmbStates.Items.Add(name);
            }
            _cmbStates.EndUpdate();
            _cmbStates.Text = currentText;
        }
        finally
        {
            _updating = false;
        }

        UpdateModeUi();
        ApplySignModeToInputs();
    }

    private void Controller_StateChanged(object? sender, EventArgs e)
    {
        RefreshFromController();
    }

    private void Execute(Action action)
    {
        action();
        RefreshFromController();
    }

    private void AddResizeRow(TableLayoutPanel table, int row, string label, NumericUpDown input, BoxAxis axis)
    {
        table.Controls.Add(CreateAxisLabel(label), 0, row);
        table.Controls.Add(input, 1, row);

        var btnNegative = CreateGridButton("<", 52, (_, _) => Execute(() => ExecuteAxisButton(axis, input, negativeDirection: true, bothFaces: false)));
        var btnBoth = CreateGridButton("<>", 68, (_, _) => Execute(() => ExecuteAxisButton(axis, input, negativeDirection: false, bothFaces: true)));
        var btnPositive = CreateGridButton(">", 52, (_, _) => Execute(() => ExecuteAxisButton(axis, input, negativeDirection: false, bothFaces: false)));

        table.Controls.Add(btnNegative, 2, row);
        table.Controls.Add(btnBoth, 3, row);
        table.Controls.Add(btnPositive, 4, row);

        _moveButtons.Add(btnNegative);
        _moveButtons.Add(btnPositive);
        _resizeOnlyButtons.Add(btnBoth);
    }

    private void ExecuteAxisButton(BoxAxis axis, NumericUpDown input, bool negativeDirection, bool bothFaces)
    {
        if (_chkMove.Checked)
        {
            if (bothFaces)
            {
                return;
            }

            var distance = Math.Abs((double)input.Value);
            var signedDistance = negativeDirection ? -distance : distance;
            _controller.MoveAxis(axis, signedDistance);
            return;
        }

        if (bothFaces)
        {
            _controller.ResizeAxis(axis, ResizeMode.BothFaces, (double)input.Value);
            return;
        }

        _controller.ResizeAxis(
            axis,
            negativeDirection ? ResizeMode.NegativeFace : ResizeMode.PositiveFace,
            (double)input.Value);
    }

    private void RegisterDeltaInput(NumericUpDown input)
    {
        _deltaInputs.Add(input);
        input.ValueChanged += DeltaInput_ValueChanged;
    }

    private void DeltaInput_ValueChanged(object? sender, EventArgs e)
    {
        if (_normalizingValues || sender is not NumericUpDown input)
        {
            return;
        }

        NormalizeInputSign(input);
    }

    private void ApplySignModeToInputs()
    {
        if (_normalizingValues)
        {
            return;
        }

        _normalizingValues = true;
        try
        {
            foreach (var input in _deltaInputs)
            {
                NormalizeInputSignCore(input);
            }
        }
        finally
        {
            _normalizingValues = false;
        }
    }

    private void NormalizeInputSign(NumericUpDown input)
    {
        _normalizingValues = true;
        try
        {
            NormalizeInputSignCore(input);
        }
        finally
        {
            _normalizingValues = false;
        }
    }

    private void NormalizeInputSignCore(NumericUpDown input)
    {
        var absoluteValue = Math.Abs(input.Value);
        input.Value = _chkNegative.Checked ? -absoluteValue : absoluteValue;
    }

    private void UpdateModeUi()
    {
        var moveMode = _chkMove.Checked;

        foreach (var button in _moveButtons)
        {
            button.Enabled = true;
            button.UseVisualStyleBackColor = !moveMode;
            button.BackColor = moveMode ? Color.LightGoldenrodYellow : SystemColors.Control;
            button.FlatStyle = moveMode ? FlatStyle.Popup : FlatStyle.Standard;
        }

        foreach (var button in _resizeOnlyButtons)
        {
            button.Enabled = !moveMode;
            button.UseVisualStyleBackColor = true;
            button.BackColor = SystemColors.Control;
            button.FlatStyle = FlatStyle.Standard;
        }

        _lblStatus.Text = BuildStatusText(_controller.HasCurrentBox(), _controller.IsClippingEnabled());
        _toolTip.SetToolTip(_lblStatus, _lblStatus.Text);
    }

    private string BuildStatusText(bool hasBox, bool clippingOn)
    {
        if (!hasBox)
        {
            return "Create a box with Fit selection, then resize, move, or save it.";
        }

        var movementText = _chkMove.Checked
            ? "Move mode is on. The highlighted outer X/Y/Z buttons shift the box along its local axes. The middle '<>' and 'All' resize buttons are disabled in this mode."
            : "Resize mode is on. Use the entered values to grow or shrink the box. Check '-' to force all delta fields negative for quick shrinking.";

        var clippingText = clippingOn
            ? "Clipping is on. Native entities outside the box are hidden, and xrefs are clipped to the clip-box volume."
            : "Clipping is off. The box is visible and can be moved, rotated, or scaled with native commands.";

        return clippingText + Environment.NewLine + movementText;
    }

    private static NumericUpDown CreateDeltaControl()
    {
        return new NumericUpDown
        {
            DecimalPlaces = 3,
            Increment = 10,
            Minimum = -1000000,
            Maximum = 1000000,
            Value = 100,
            ThousandsSeparator = true,
            Dock = DockStyle.Fill,
            MinimumSize = new Size(138, 0),
            TextAlign = HorizontalAlignment.Right,
        };
    }

    private static Button CreateButton(string text, EventHandler handler)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            Margin = new Padding(2),
        };
        button.Click += handler;
        return button;
    }

    private static Button CreateGridButton(string text, int minimumWidth, EventHandler handler)
    {
        var button = CreateButton(text, handler);
        button.AutoSize = false;
        button.Dock = DockStyle.Fill;
        button.MinimumSize = new Size(minimumWidth, 0);
        return button;
    }

    private static Label CreateAxisLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = false,
            Dock = DockStyle.Fill,
            MinimumSize = new Size(48, 0),
            TextAlign = ContentAlignment.MiddleLeft,
        };
    }

    private string GetStateName()
    {
        return string.IsNullOrWhiteSpace(_cmbStates.Text) ? "Box" : _cmbStates.Text.Trim();
    }
}
