using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Plant3DClipBox;

internal sealed class ProjectRefResultsForm : Form
{
    private readonly BindingList<ProjectRefRow> _rows;
    private readonly DataGridView _grid;

    public ProjectRefResultsForm(IReadOnlyList<ProjectRefCheckService.ProjectRefCandidate> candidates)
    {
        if (candidates is null)
        {
            throw new ArgumentNullException(nameof(candidates));
        }

        Text = "Missing project xrefs";
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.SizableToolWindow;
        MinimumSize = new Size(760, 380);
        Size = new Size(980, 560);

        _rows = new BindingList<ProjectRefRow>(
            candidates
                .Select(candidate => new ProjectRefRow(candidate))
                .ToList());

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(10),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var intro = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(940, 0),
            Margin = new Padding(0, 0, 0, 8),
            Text =
                "The files below belong to the current Plant 3D project's Plant 3D drawings, are not currently attached as xrefs, and are likely relevant for the current clip box. " +
                "Select the files you want to load as overlay xrefs at 0,0,0.",
        };
        root.Controls.Add(intro, 0, 0);

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoGenerateColumns = false,
            MultiSelect = true,
            ReadOnly = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            DataSource = _rows,
            Margin = new Padding(0, 0, 0, 8),
        };
        _grid.CurrentCellDirtyStateChanged += Grid_CurrentCellDirtyStateChanged;
        _grid.CellDoubleClick += Grid_CellDoubleClick;

        _grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            DataPropertyName = nameof(ProjectRefRow.Selected),
            HeaderText = "Load",
            Width = 54,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            SortMode = DataGridViewColumnSortMode.Automatic,
        });

        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(ProjectRefRow.FileName),
            HeaderText = "File",
            Width = 180,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
            ReadOnly = true,
        });

        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(ProjectRefRow.ProjectPartName),
            HeaderText = "Project part",
            Width = 120,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
            ReadOnly = true,
        });

        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(ProjectRefRow.Reason),
            HeaderText = "Reason",
            Width = 230,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 28F,
            ReadOnly = true,
        });

        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(ProjectRefRow.ProjectRelativePath),
            HeaderText = "Project path",
            Width = 360,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 42F,
            ReadOnly = true,
        });

        root.Controls.Add(_grid, 0, 1);

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            WrapContents = true,
        };

        var btnLoad = new Button
        {
            Text = "Load selected",
            AutoSize = true,
            DialogResult = DialogResult.OK,
            Margin = new Padding(6, 0, 0, 0),
        };
        buttonPanel.Controls.Add(btnLoad);

        var btnCancel = new Button
        {
            Text = "Cancel",
            AutoSize = true,
            DialogResult = DialogResult.Cancel,
            Margin = new Padding(6, 0, 0, 0),
        };
        buttonPanel.Controls.Add(btnCancel);

        var btnClear = new Button
        {
            Text = "Clear",
            AutoSize = true,
            Margin = new Padding(6, 0, 0, 0),
        };
        btnClear.Click += (_, _) => SetAllSelections(false);
        buttonPanel.Controls.Add(btnClear);

        var btnSelectAll = new Button
        {
            Text = "Select all",
            AutoSize = true,
            Margin = new Padding(6, 0, 0, 0),
        };
        btnSelectAll.Click += (_, _) => SetAllSelections(true);
        buttonPanel.Controls.Add(btnSelectAll);

        root.Controls.Add(buttonPanel, 0, 2);

        Controls.Add(root);
        AcceptButton = btnLoad;
        CancelButton = btnCancel;
    }

    public List<ProjectRefCheckService.ProjectRefCandidate> GetSelectedCandidates()
    {
        return _rows
            .Where(row => row.Selected)
            .Select(row => new ProjectRefCheckService.ProjectRefCandidate
            {
                AbsoluteFileName = row.AbsoluteFileName,
                FileName = row.FileName,
                ProjectPartName = row.ProjectPartName,
                ProjectRelativePath = row.ProjectRelativePath,
                Reason = row.Reason,
            })
            .ToList();
    }

    private void SetAllSelections(bool selected)
    {
        foreach (var row in _rows)
        {
            row.Selected = selected;
        }

        _grid.Refresh();
    }

    private void Grid_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
    {
        if (_grid.IsCurrentCellDirty)
        {
            _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void Grid_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _rows.Count)
        {
            return;
        }

        _rows[e.RowIndex].Selected = !_rows[e.RowIndex].Selected;
        _grid.Refresh();
    }


    private sealed class ProjectRefRow
    {
        public ProjectRefRow(ProjectRefCheckService.ProjectRefCandidate candidate)
        {
            Selected = true;
            AbsoluteFileName = candidate.AbsoluteFileName;
            FileName = candidate.FileName;
            ProjectPartName = candidate.ProjectPartName;
            ProjectRelativePath = candidate.ProjectRelativePath;
            Reason = candidate.Reason;
        }

        public bool Selected { get; set; }

        public string AbsoluteFileName { get; }

        public string FileName { get; }

        public string ProjectPartName { get; }

        public string ProjectRelativePath { get; }

        public string Reason { get; }
    }
}
