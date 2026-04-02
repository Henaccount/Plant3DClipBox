using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plant3DClipBox;

internal static class ProjectRefCheckService
{
    internal sealed class ProjectRefCheckResult
    {
        public bool Applicable { get; set; }

        public bool Completed { get; set; }

        public bool ChangedDrawing { get; set; }

        public int MissingCacheCount { get; set; }

        public int ScannedDrawingCount { get; set; }

        public int ProposedCount { get; set; }

        public int LoadedCount { get; set; }

        public int FailedLoadCount { get; set; }

        public string Message { get; set; } = string.Empty;
    }

    internal sealed class ProjectRefCandidate
    {
        public string FileName { get; set; } = string.Empty;

        public string AbsoluteFileName { get; set; } = string.Empty;

        public string ProjectPartName { get; set; } = string.Empty;

        public string ProjectRelativePath { get; set; } = string.Empty;

        public string Reason { get; set; } = string.Empty;
    }

    private sealed class PlantProjectInfo
    {
        public string RootDirectory { get; set; } = string.Empty;

        public bool IsCollaborationProject { get; set; }

        public List<ProjectDrawingInfo> Drawings { get; } = new();
    }

    private sealed class ProjectPartInfo
    {
        public string PartName { get; set; } = string.Empty;

        public string ProjectDirectory { get; set; } = string.Empty;

        public string ProjectDwgDirectory { get; set; } = string.Empty;

        public List<ProjectDrawingInfo> Drawings { get; } = new();
    }

    private sealed class ProjectDrawingInfo
    {
        public string AbsoluteFileName { get; set; } = string.Empty;

        public string ProjectPartName { get; set; } = string.Empty;
    }

    private sealed class ProjectPathResolver
    {
        private readonly List<string> _baseDirectories;
        private readonly List<string> _searchRoots;
        private Dictionary<string, List<string>>? _filesByName;
        private List<string>? _indexedFiles;

        public ProjectPathResolver(IEnumerable<string> baseDirectories)
        {
            _baseDirectories = baseDirectories
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(NormalizeFullPath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            _searchRoots = _baseDirectories
                .Where(Directory.Exists)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public string ResolveDrawingPath(IEnumerable<string> rawValues)
        {
            var values = rawValues
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim().Trim('"'))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var firstGuess = string.Empty;
            foreach (var rawValue in values)
            {
                var exactMatch = TryResolveExactOrCombined(rawValue, out var guess);
                if (!string.IsNullOrWhiteSpace(exactMatch))
                {
                    return exactMatch;
                }

                if (string.IsNullOrWhiteSpace(firstGuess) && !string.IsNullOrWhiteSpace(guess))
                {
                    firstGuess = guess;
                }
            }

            foreach (var rawValue in values)
            {
                var fallback = TryResolveBySearch(rawValue);
                if (!string.IsNullOrWhiteSpace(fallback))
                {
                    return fallback;
                }
            }

            return firstGuess;
        }

        private string TryResolveExactOrCombined(string rawValue, out string guess)
        {
            guess = string.Empty;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return string.Empty;
            }

            try
            {
                if (Path.IsPathRooted(rawValue))
                {
                    var absolute = NormalizeFullPath(rawValue);
                    guess = absolute;
                    return File.Exists(absolute) ? absolute : string.Empty;
                }

                foreach (var baseDirectory in _baseDirectories)
                {
                    var candidate = NormalizeFullPath(Path.Combine(baseDirectory, rawValue));
                    if (string.IsNullOrWhiteSpace(guess))
                    {
                        guess = candidate;
                    }

                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }

                var normalized = NormalizeFullPath(rawValue);
                if (string.IsNullOrWhiteSpace(guess))
                {
                    guess = normalized;
                }

                return File.Exists(normalized) ? normalized : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string TryResolveBySearch(string rawValue)
        {
            var lookupPath = NormalizeRelativeLookupPath(rawValue);
            if (string.IsNullOrWhiteSpace(lookupPath))
            {
                return string.Empty;
            }

            EnsureIndex();
            if (_indexedFiles is null || _filesByName is null || _indexedFiles.Count == 0)
            {
                return string.Empty;
            }

            if (lookupPath.IndexOf(Path.DirectorySeparatorChar) >= 0)
            {
                var matches = _indexedFiles
                    .Where(file => HasRelativeSuffix(file, lookupPath))
                    .OrderBy(file => file.Length)
                    .ToList();

                if (matches.Count > 0)
                {
                    return matches[0];
                }
            }

            var fileName = Path.GetFileName(lookupPath);
            if (!string.IsNullOrWhiteSpace(fileName) && _filesByName.TryGetValue(fileName, out var fileMatches))
            {
                if (fileMatches.Count == 1)
                {
                    return fileMatches[0];
                }
            }

            return string.Empty;
        }

        private void EnsureIndex()
        {
            if (_indexedFiles is not null && _filesByName is not null)
            {
                return;
            }

            _indexedFiles = new List<string>();
            _filesByName = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var root in _searchRoots)
            {
                try
                {
                    foreach (var file in Directory.EnumerateFiles(root, "*.dwg", SearchOption.AllDirectories))
                    {
                        var normalized = NormalizeFullPath(file);
                        if (string.IsNullOrWhiteSpace(normalized) || !seen.Add(normalized))
                        {
                            continue;
                        }

                        _indexedFiles.Add(normalized);
                        var fileName = Path.GetFileName(normalized);
                        if (string.IsNullOrWhiteSpace(fileName))
                        {
                            continue;
                        }

                        if (!_filesByName.TryGetValue(fileName, out var list))
                        {
                            list = new List<string>();
                            _filesByName[fileName] = list;
                        }

                        list.Add(normalized);
                    }
                }
                catch
                {
                    // Best effort only.
                }
            }
        }

        private static string NormalizeRelativeLookupPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var normalized = path.Trim().Trim('"')
                .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

            while (normalized.StartsWith("." + Path.DirectorySeparatorChar, StringComparison.Ordinal))
            {
                normalized = normalized.Substring(2);
            }

            return normalized.TrimStart(Path.DirectorySeparatorChar);
        }

        private static bool HasRelativeSuffix(string absolutePath, string relativePath)
        {
            var normalizedAbsolute = absolutePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            var normalizedRelative = NormalizeRelativeLookupPath(relativePath);
            if (string.IsNullOrWhiteSpace(normalizedRelative))
            {
                return false;
            }

            return normalizedAbsolute.EndsWith(Path.DirectorySeparatorChar + normalizedRelative, StringComparison.OrdinalIgnoreCase) ||
                   normalizedAbsolute.EndsWith(normalizedRelative, StringComparison.OrdinalIgnoreCase);
        }
    }

    private sealed class ScanOutcome
    {
        public bool Intersects { get; set; }

        public bool AnalysisFailed { get; set; }

        public bool ProposeConservatively { get; set; }

        public string Reason { get; set; } = string.Empty;
    }

    private const double ScanTolerance = 1.0e-8;

    public static ProjectRefCheckResult RunInteractive(Document document, OrientedBox clipBox)
    {
        var result = new ProjectRefCheckResult();

        if (document is null || !clipBox.IsValid)
        {
            result.Message = "No valid clip box is available for project reference checking.";
            return result;
        }

        if (!TryGetCurrentPlantProject(document, out var projectInfo, out var notApplicableMessage))
        {
            result.Message = notApplicableMessage;
            return result;
        }

        result.Applicable = true;

        var currentDrawingPath = NormalizeFullPath(document.Name);
        var projectDrawings = projectInfo.Drawings
            .Where(d => !string.IsNullOrWhiteSpace(d.AbsoluteFileName))
            .Where(d => !PathsEqual(d.AbsoluteFileName, currentDrawingPath))
            .ToList();

        if (projectDrawings.Count == 0)
        {
            result.Completed = true;
            result.Message = "No other Plant 3D drawings were found in the active Plant 3D project.";
            return result;
        }

        var availableDrawings = new List<ProjectDrawingInfo>();
        var unavailableDrawings = new List<ProjectDrawingInfo>();

        foreach (var drawing in projectDrawings)
        {
            if (File.Exists(drawing.AbsoluteFileName))
            {
                availableDrawings.Add(drawing);
            }
            else
            {
                unavailableDrawings.Add(drawing);
            }
        }

        result.MissingCacheCount = unavailableDrawings.Count;
        if (projectInfo.IsCollaborationProject && unavailableDrawings.Count > 0)
        {
            var continueScan = AskToContinueWithCachedFiles(unavailableDrawings);
            if (!continueScan)
            {
                result.Message = $"Reference check cancelled. {unavailableDrawings.Count} Plant 3D drawing(s) are not available in the local cache.";
                return result;
            }
        }

        if (availableDrawings.Count == 0)
        {
            result.Completed = true;
            result.Message = projectInfo.IsCollaborationProject
                ? "No cached Plant 3D drawings are available for scanning."
                : "No accessible Plant 3D drawings are available for scanning.";
            return result;
        }

        var existingXrefPaths = CollectExistingXrefPaths(document.Database);
        var candidates = new List<ProjectRefCandidate>();

        Cursor? previousCursor = null;
        try
        {
            previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            foreach (var drawing in availableDrawings)
            {
                if (existingXrefPaths.Contains(NormalizeFullPath(drawing.AbsoluteFileName)))
                {
                    continue;
                }

                var scanOutcome = AnalyzeDrawingAgainstClipBox(drawing.AbsoluteFileName, clipBox);
                if (!scanOutcome.Intersects && !scanOutcome.ProposeConservatively)
                {
                    continue;
                }

                candidates.Add(new ProjectRefCandidate
                {
                    FileName = Path.GetFileName(drawing.AbsoluteFileName),
                    AbsoluteFileName = drawing.AbsoluteFileName,
                    ProjectPartName = drawing.ProjectPartName,
                    ProjectRelativePath = MakeProjectRelativePath(projectInfo.RootDirectory, drawing.AbsoluteFileName),
                    Reason = !string.IsNullOrWhiteSpace(scanOutcome.Reason)
                        ? scanOutcome.Reason
                        : (scanOutcome.Intersects
                            ? "Intersects the clip box"
                            : "Could not analyze cleanly; proposed conservatively"),
                });
            }
        }
        finally
        {
            Cursor.Current = previousCursor ?? Cursors.Default;
        }

        result.Completed = true;
        result.ScannedDrawingCount = availableDrawings.Count;
        result.ProposedCount = candidates.Count;

        if (candidates.Count == 0)
        {
            result.Message = AppendAvailabilityNote(
                "No missing project xrefs were found for the current clip box.",
                unavailableDrawings.Count,
                projectInfo.IsCollaborationProject);
            return result;
        }

        using var dialog = new ProjectRefResultsForm(candidates);
        var dialogResult = AcAp.ShowModalDialog(dialog);
        if (dialogResult != DialogResult.OK)
        {
            result.Message = AppendAvailabilityNote(
                $"Reference check finished. {candidates.Count} missing project xref(s) were found, but none were added.",
                unavailableDrawings.Count,
                projectInfo.IsCollaborationProject);
            return result;
        }

        var selected = dialog.GetSelectedCandidates();
        if (selected.Count == 0)
        {
            result.Message = AppendAvailabilityNote(
                $"Reference check finished. {candidates.Count} missing project xref(s) were found, but none were selected for loading.",
                unavailableDrawings.Count,
                projectInfo.IsCollaborationProject);
            return result;
        }

        var loadOutcome = LoadSelectedReferences(document, selected);
        result.ChangedDrawing = loadOutcome.loaded > 0;
        result.LoadedCount = loadOutcome.loaded;
        result.FailedLoadCount = loadOutcome.failed;

        if (loadOutcome.loaded > 0 && loadOutcome.failed == 0)
        {
            result.Message = $"Loaded {loadOutcome.loaded} missing project xref(s) as overlay references at 0,0,0.";
        }
        else if (loadOutcome.loaded > 0)
        {
            result.Message = $"Loaded {loadOutcome.loaded} missing project xref(s). {loadOutcome.failed} xref(s) could not be attached.";
        }
        else
        {
            result.Message = $"No missing project xrefs were attached. {loadOutcome.failed} selected xref(s) could not be loaded.";
        }

        result.Message = AppendAvailabilityNote(result.Message, unavailableDrawings.Count, projectInfo.IsCollaborationProject);
        return result;
    }

    private static bool AskToContinueWithCachedFiles(IReadOnlyList<ProjectDrawingInfo> missingCache)
    {
        var preview = string.Join(
            Environment.NewLine,
            missingCache
                .Take(8)
                .Select(d => "- " + Path.GetFileName(d.AbsoluteFileName)));

        if (missingCache.Count > 8)
        {
            preview += Environment.NewLine + $"- ... and {missingCache.Count - 8} more";
        }

        var message =
            $"{missingCache.Count} Plant 3D drawing(s) listed under Plant 3D Drawings in Project Manager are not available in the local cache." + Environment.NewLine + Environment.NewLine +
            "Download the relevant files first for a complete check, or continue with the cached files only." + Environment.NewLine + Environment.NewLine +
            preview + Environment.NewLine + Environment.NewLine +
            "Click Yes to continue with the cached files only. Click No to cancel.";

        return MessageBox.Show(
                   message,
                   "Clip Box - Missing cached project files",
                   MessageBoxButtons.YesNo,
                   MessageBoxIcon.Warning,
                   MessageBoxDefaultButton.Button2) == DialogResult.Yes;
    }

    private static string AppendAvailabilityNote(string message, int unavailableCount, bool collaborationProject)
    {
        if (unavailableCount <= 0)
        {
            return message;
        }

        var note = collaborationProject
            ? $" {unavailableCount} Plant 3D drawing(s) were not in the local cache, so the scan used only the cached subset."
            : $" {unavailableCount} listed Plant 3D drawing(s) were not accessible and were skipped.";

        return string.IsNullOrWhiteSpace(message)
            ? note.Trim()
            : message + note;
    }

    private static bool TryGetCurrentPlantProject(
        Document document,
        out PlantProjectInfo projectInfo,
        out string message)
    {
        projectInfo = new PlantProjectInfo();
        message = string.Empty;

        var plantApplicationType = FindType(
            "Autodesk.ProcessPower.PlantInstance.PlantApplication",
            "PnPProjectManagerMgd");

        if (plantApplicationType is null)
        {
            message = "chk.ref is available only in AutoCAD Plant 3D projects.";
            return false;
        }

        var currentProject = GetStaticPropertyValue(plantApplicationType, "CurrentProject");
        if (currentProject is null)
        {
            message = "No active Plant 3D project was detected.";
            return false;
        }

        var projectFileDirectory = GetProjectFileDirectory(currentProject);
        if (string.IsNullOrWhiteSpace(projectFileDirectory))
        {
            projectFileDirectory = SafeGetDirectoryName(document.Name);
        }

        var projectParts = ExtractProjectParts(currentProject, projectFileDirectory);
        if (projectParts.Count == 0)
        {
            message = "No project drawings were found in the active Plant 3D project.";
            return false;
        }

        var targetParts = SelectPlantProjectParts(document, projectParts);
        var drawings = MergeDrawings(targetParts);
        if (drawings.Count == 0)
        {
            message = "No Plant 3D drawings were found in the active Plant 3D project scope for the current drawing.";
            return false;
        }

        var collaborationRoot = FindCollaborationProjectRoot(document, drawings, currentProject);
        var isCollaborationProject = false;
        if (!string.IsNullOrWhiteSpace(collaborationRoot))
        {
            var slnkPath = Path.Combine(collaborationRoot, "DocumentServer.slnk");
            var slnkText = SafeReadAllText(slnkPath);
            isCollaborationProject = LooksLikeCollaborationDocumentServer(slnkText, collaborationRoot);
        }

        projectInfo.RootDirectory = isCollaborationProject
            ? collaborationRoot
            : FindProjectRoot(document, currentProject, projectParts);
        projectInfo.IsCollaborationProject = isCollaborationProject;
        projectInfo.Drawings.AddRange(drawings);
        return true;
    }

    private static List<ProjectPartInfo> ExtractProjectParts(object currentProject, string projectRootHint)
    {
        var parts = new List<ProjectPartInfo>();
        var projectParts = GetPropertyValue(currentProject, "ProjectParts") as IEnumerable;
        if (projectParts is null)
        {
            return parts;
        }

        foreach (var part in projectParts)
        {
            if (part is null)
            {
                continue;
            }

            var projectPartName = ResolveProjectPartName(currentProject, part);
            var rawProjectDwgDirectory =
                Convert.ToString(GetPropertyValue(part, "ProjectDwgDirectory")) ??
                Convert.ToString(GetPropertyValue(part, "ProjectDirectory")) ??
                string.Empty;
            var rawProjectDirectory =
                Convert.ToString(GetPropertyValue(part, "ProjectDirectory")) ??
                rawProjectDwgDirectory;

            var info = new ProjectPartInfo
            {
                PartName = projectPartName,
                ProjectDirectory = ResolveProjectPath(projectRootHint, rawProjectDirectory),
                ProjectDwgDirectory = ResolveProjectPath(projectRootHint, rawProjectDwgDirectory),
            };

            var pathResolver = new ProjectPathResolver(new[]
            {
                info.ProjectDwgDirectory,
                info.ProjectDirectory,
                projectRootHint,
            });

            var getDrawingsMethod = part.GetType().GetMethod("GetPnPDrawingFiles", BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
            var drawingCollection = getDrawingsMethod?.Invoke(part, null) as IEnumerable;
            if (drawingCollection is not null)
            {
                foreach (var drawing in drawingCollection)
                {
                    if (drawing is null)
                    {
                        continue;
                    }

                    var rawPathCandidates = CollectDrawingPathCandidates(drawing);
                    if (rawPathCandidates.Count == 0)
                    {
                        continue;
                    }

                    var normalized = pathResolver.ResolveDrawingPath(rawPathCandidates);
                    if (string.IsNullOrWhiteSpace(normalized))
                    {
                        continue;
                    }

                    if (!info.Drawings.Any(d => PathsEqual(d.AbsoluteFileName, normalized)))
                    {
                        info.Drawings.Add(new ProjectDrawingInfo
                        {
                            AbsoluteFileName = normalized,
                            ProjectPartName = projectPartName,
                        });
                    }
                }
            }

            if (info.Drawings.Count > 0 || !string.IsNullOrWhiteSpace(info.ProjectDirectory) || !string.IsNullOrWhiteSpace(info.ProjectDwgDirectory))
            {
                parts.Add(info);
            }
        }

        return parts;
    }

    private static List<ProjectPartInfo> SelectPlantProjectParts(Document document, IReadOnlyList<ProjectPartInfo> allParts)
    {
        var currentDrawingPath = NormalizeFullPath(document.Name);
        var currentDrawingParts = allParts
            .Where(part => part.Drawings.Any(d => PathsEqual(d.AbsoluteFileName, currentDrawingPath)))
            .ToList();

        if (currentDrawingParts.Count > 0)
        {
            return currentDrawingParts
                .Where(part => IsLikelyPlantModelPart(part, currentDrawingPath))
                .ToList();
        }

        return allParts
            .Where(part => IsLikelyPlantModelPart(part, currentDrawingPath))
            .ToList();
    }

    private static List<string> CollectDrawingPathCandidates(object drawing)
    {
        var candidates = new List<string>();
        var propertyNames = new[]
        {
            "AbsoluteFileName",
            "LocalFileName",
            "RelativeFileName",
            "FullPath",
            "FilePath",
            "DwgFileName",
            "FileName",
            "RelativePath",
            "Path",
        };

        foreach (var propertyName in propertyNames)
        {
            var value = Convert.ToString(GetPropertyValue(drawing, propertyName));
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            candidates.Add(value);
        }

        return candidates
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string GetProjectFileDirectory(object currentProject)
    {
        var projectFileName =
            Convert.ToString(GetPropertyValue(currentProject, "FileName")) ??
            Convert.ToString(GetPropertyValue(currentProject, "ProjectFileName")) ??
            string.Empty;

        return SafeGetDirectoryName(projectFileName);
    }

    private static string ResolveProjectPath(string projectRootHint, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            if (Path.IsPathRooted(path))
            {
                return NormalizeFullPath(path);
            }

            if (!string.IsNullOrWhiteSpace(projectRootHint))
            {
                return NormalizeFullPath(Path.Combine(projectRootHint, path));
            }

            return NormalizeFullPath(path);
        }
        catch
        {
            return NormalizeFullPath(path);
        }
    }

    private static List<ProjectDrawingInfo> MergeDrawings(IReadOnlyList<ProjectPartInfo> parts)
    {
        var drawings = new Dictionary<string, ProjectDrawingInfo>(StringComparer.OrdinalIgnoreCase);

        foreach (var part in parts)
        {
            foreach (var drawing in part.Drawings)
            {
                if (string.IsNullOrWhiteSpace(drawing.AbsoluteFileName))
                {
                    continue;
                }

                var normalized = NormalizeFullPath(drawing.AbsoluteFileName);
                if (drawings.ContainsKey(normalized))
                {
                    continue;
                }

                drawings[normalized] = new ProjectDrawingInfo
                {
                    AbsoluteFileName = normalized,
                    ProjectPartName = drawing.ProjectPartName,
                };
            }
        }

        return drawings.Values
            .OrderBy(d => d.AbsoluteFileName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsLikelyPlantModelPart(ProjectPartInfo part, string currentDrawingPath)
    {
        if (part.Drawings.Count == 0)
        {
            return false;
        }

        if (LooksLikeExcludedProjectPart(part.PartName))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(currentDrawingPath) &&
            part.Drawings.Any(d => PathsEqual(d.AbsoluteFileName, currentDrawingPath)))
        {
            return true;
        }

        var partDirectory = string.IsNullOrWhiteSpace(part.ProjectDwgDirectory)
            ? part.ProjectDirectory
            : part.ProjectDwgDirectory;
        var currentDrawingDirectory = SafeGetDirectoryName(currentDrawingPath);

        if (!string.IsNullOrWhiteSpace(partDirectory) && !string.IsNullOrWhiteSpace(currentDrawingDirectory))
        {
            if (IsSameOrParentDirectory(partDirectory, currentDrawingDirectory) ||
                IsSameOrParentDirectory(currentDrawingDirectory, partDirectory))
            {
                return true;
            }
        }

        var name = (part.PartName ?? string.Empty).Trim();
        return name.IndexOf("piping", StringComparison.OrdinalIgnoreCase) >= 0 ||
               name.IndexOf("plant", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool LooksLikeExcludedProjectPart(string partName)
    {
        if (string.IsNullOrWhiteSpace(partName))
        {
            return false;
        }

        return partName.IndexOf("pnid", StringComparison.OrdinalIgnoreCase) >= 0 ||
               partName.IndexOf("p&id", StringComparison.OrdinalIgnoreCase) >= 0 ||
               partName.IndexOf("ortho", StringComparison.OrdinalIgnoreCase) >= 0 ||
               partName.IndexOf("orthographic", StringComparison.OrdinalIgnoreCase) >= 0 ||
               partName.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 ||
               partName.IndexOf("isometric", StringComparison.OrdinalIgnoreCase) >= 0 ||
               partName.IndexOf("related", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string FindProjectRoot(
        Document document,
        object currentProject,
        IReadOnlyList<ProjectPartInfo> projectParts)
    {
        var projectFileName =
            Convert.ToString(GetPropertyValue(currentProject, "FileName")) ??
            Convert.ToString(GetPropertyValue(currentProject, "ProjectFileName")) ??
            string.Empty;

        var projectFileDirectory = SafeGetDirectoryName(projectFileName);
        if (!string.IsNullOrWhiteSpace(projectFileDirectory))
        {
            return projectFileDirectory;
        }

        var candidateDirectories = new List<string>();
        foreach (var part in projectParts)
        {
            if (!string.IsNullOrWhiteSpace(part.ProjectDwgDirectory))
            {
                candidateDirectories.Add(part.ProjectDwgDirectory);
            }

            if (!string.IsNullOrWhiteSpace(part.ProjectDirectory))
            {
                candidateDirectories.Add(part.ProjectDirectory);
            }
        }

        if (!string.IsNullOrWhiteSpace(document.Name))
        {
            var currentDir = SafeGetDirectoryName(document.Name);
            if (!string.IsNullOrWhiteSpace(currentDir))
            {
                candidateDirectories.Add(currentDir);
            }
        }

        var commonAncestor = FindCommonAncestor(candidateDirectories);
        if (!string.IsNullOrWhiteSpace(commonAncestor))
        {
            return commonAncestor;
        }

        return candidateDirectories
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(NormalizeFullPath)
            .FirstOrDefault() ?? string.Empty;
    }

    private static string FindCommonAncestor(IEnumerable<string> directories)
    {
        var normalized = directories
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(NormalizeFullPath)
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (normalized.Count == 0)
        {
            return string.Empty;
        }

        var ancestor = normalized[0];
        while (!string.IsNullOrWhiteSpace(ancestor))
        {
            if (normalized.All(path => IsSameOrParentDirectory(ancestor, path)))
            {
                return ancestor;
            }

            var parent = SafeGetDirectoryName(ancestor);
            if (string.IsNullOrWhiteSpace(parent) || PathsEqual(parent, ancestor))
            {
                break;
            }

            ancestor = parent;
        }

        return string.Empty;
    }

    private static bool IsSameOrParentDirectory(string parentCandidate, string childCandidate)
    {
        var parent = AppendDirectorySeparator(NormalizeFullPath(parentCandidate));
        var child = AppendDirectorySeparator(NormalizeFullPath(childCandidate));

        if (string.IsNullOrWhiteSpace(parent) || string.IsNullOrWhiteSpace(child))
        {
            return false;
        }

        return child.StartsWith(parent, StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveProjectPartName(object currentProject, object projectPart)
    {
        var method = currentProject.GetType().GetMethod(
            "ProjectPartName",
            BindingFlags.Public | BindingFlags.Instance,
            null,
            new[] { projectPart.GetType() },
            null);

        if (method is not null)
        {
            try
            {
                var value = method.Invoke(currentProject, new[] { projectPart });
                if (value is string projectPartName && !string.IsNullOrWhiteSpace(projectPartName))
                {
                    return projectPartName;
                }
            }
            catch
            {
                // Fall back to generic names below.
            }
        }

        var partName = Convert.ToString(GetPropertyValue(projectPart, "PartName"));
        if (!string.IsNullOrWhiteSpace(partName))
        {
            return partName;
        }

        var projectName = Convert.ToString(GetPropertyValue(projectPart, "ProjectName"));
        if (!string.IsNullOrWhiteSpace(projectName))
        {
            return projectName;
        }

        return projectPart.GetType().Name;
    }

    private static string FindCollaborationProjectRoot(
        Document document,
        IReadOnlyList<ProjectDrawingInfo> drawings,
        object currentProject)
    {
        var candidateDirectories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var drawing in drawings)
        {
            var dir = SafeGetDirectoryName(drawing.AbsoluteFileName);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                candidateDirectories.Add(dir);
            }
        }

        var projectParts = GetPropertyValue(currentProject, "ProjectParts") as IEnumerable;
        if (projectParts is not null)
        {
            foreach (var part in projectParts)
            {
                var projectDirectory = Convert.ToString(GetPropertyValue(part!, "ProjectDirectory"));
                if (!string.IsNullOrWhiteSpace(projectDirectory))
                {
                    candidateDirectories.Add(NormalizeFullPath(projectDirectory));
                }

                var projectDwgDirectory = Convert.ToString(GetPropertyValue(part!, "ProjectDwgDirectory"));
                if (!string.IsNullOrWhiteSpace(projectDwgDirectory))
                {
                    candidateDirectories.Add(NormalizeFullPath(projectDwgDirectory));
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(document.Name))
        {
            var currentDir = SafeGetDirectoryName(document.Name);
            if (!string.IsNullOrWhiteSpace(currentDir))
            {
                candidateDirectories.Add(currentDir);
            }
        }

        var roots = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var directory in candidateDirectories)
        {
            var root = FindAncestorContainingFile(directory, "DocumentServer.slnk");
            if (string.IsNullOrWhiteSpace(root))
            {
                continue;
            }

            if (roots.ContainsKey(root))
            {
                roots[root]++;
            }
            else
            {
                roots[root] = 1;
            }
        }

        if (roots.Count == 0)
        {
            return string.Empty;
        }

        return roots
            .OrderByDescending(kvp => kvp.Value)
            .ThenBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase)
            .Select(kvp => kvp.Key)
            .FirstOrDefault() ?? string.Empty;
    }

    private static string? FindTypeMarker(string text, params string[] markers)
    {
        foreach (var marker in markers)
        {
            if (text.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return marker;
            }
        }

        return null;
    }

    private static bool LooksLikeCollaborationDocumentServer(string slnkText, string rootDirectory)
    {
        if (!string.IsNullOrWhiteSpace(slnkText))
        {
            var hasProjectType = FindTypeMarker(slnkText, "Project Type", "ProjectType") is not null;
            var hasAccOrBim = FindTypeMarker(slnkText, "ACC", "BIM") is not null;
            var hasWorkspace = FindTypeMarker(slnkText, "WorkspaceID", "WorkspaceId") is not null;
            var hasHub = FindTypeMarker(slnkText, "Project Hub ID", "ProjectHubId", "HubId") is not null;

            if ((hasProjectType && hasAccOrBim) || hasWorkspace || hasHub)
            {
                return true;
            }
        }

        return rootDirectory.IndexOf("CollaborationCache", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static HashSet<string> CollectExistingXrefPaths(Database database)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            var xrefGraph = database.GetHostDwgXrefGraph(true);
            for (var i = 0; i < xrefGraph.NumNodes; i++)
            {
                var node = xrefGraph.GetXrefNode(i) as XrefGraphNode;
                if (node is null)
                {
                    continue;
                }

                var nodeDatabase = node.Database;
                if (nodeDatabase is null)
                {
                    continue;
                }

                var fileName = NormalizeFullPath(nodeDatabase.Filename);
                if (string.IsNullOrWhiteSpace(fileName) || PathsEqual(fileName, database.Filename))
                {
                    continue;
                }

                result.Add(fileName);
            }
        }
        catch
        {
            // Best effort only.
        }

        return result;
    }

    private static ScanOutcome AnalyzeDrawingAgainstClipBox(string fileName, OrientedBox clipBox)
    {
        var openDocument = FindOpenDocumentByPath(fileName);
        if (openDocument is not null)
        {
            return AnalyzeDatabaseAgainstClipBox(openDocument.Database, clipBox);
        }

        try
        {
            using var sideDb = new Database(false, true);
            sideDb.ReadDwgFile(fileName, FileOpenMode.OpenForReadAndReadShare, false, string.Empty);
            return AnalyzeDatabaseAgainstClipBox(sideDb, clipBox);
        }
        catch
        {
            return new ScanOutcome
            {
                AnalysisFailed = true,
                ProposeConservatively = true,
                Reason = "Could not open drawing for scan; proposed conservatively",
            };
        }
    }

    private static ScanOutcome AnalyzeDatabaseAgainstClipBox(Database database, OrientedBox clipBox)
    {
        var outcome = new ScanOutcome();
        var hasHeaderExtents = TryGetHeaderModelExtents(database, out var headerExtents);
        if (hasHeaderExtents && !IntersectsClipBox(clipBox, headerExtents))
        {
            return outcome;
        }

        try
        {
            using var transaction = database.TransactionManager.StartOpenCloseTransaction();
            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            var sawAnyEntity = false;
            var sawKnownEntity = false;
            var hadUnknownEntity = false;

            foreach (ObjectId id in modelSpace)
            {
                if (id.IsNull || id.IsErased)
                {
                    continue;
                }

                Entity? entity;
                try
                {
                    entity = transaction.GetObject(id, OpenMode.ForRead, false) as Entity;
                }
                catch
                {
                    hadUnknownEntity = true;
                    continue;
                }

                if (entity is null)
                {
                    continue;
                }

                if (IsExcludedFromSideDbScan(entity, transaction))
                {
                    continue;
                }

                sawAnyEntity = true;
                if (!TryGetEntityExtents(entity, out var entityExtents))
                {
                    hadUnknownEntity = true;
                    continue;
                }

                sawKnownEntity = true;
                if (IntersectsClipBox(clipBox, entityExtents))
                {
                    outcome.Intersects = true;
                    outcome.Reason = "Intersects the clip box";
                    return outcome;
                }
            }

            if (!sawAnyEntity)
            {
                return outcome;
            }

            if (hadUnknownEntity && !sawKnownEntity)
            {
                outcome.AnalysisFailed = true;
                outcome.ProposeConservatively = true;
                outcome.Reason = hasHeaderExtents
                    ? "No measurable local entity extents were available; proposed conservatively"
                    : "No measurable local entity extents were available and drawing extents were unavailable; proposed conservatively";
            }
        }
        catch
        {
            outcome.AnalysisFailed = true;
            outcome.ProposeConservatively = true;
            outcome.Reason = "Could not analyze drawing geometry cleanly; proposed conservatively";
        }

        return outcome;
    }

    private static bool TryGetHeaderModelExtents(Database database, out Extents3d extents)
    {
        extents = default;

        try
        {
            var min = database.Extmin;
            var max = database.Extmax;
            if (!IsFinite(min) || !IsFinite(max))
            {
                return false;
            }

            if (min.X > max.X || min.Y > max.Y || min.Z > max.Z)
            {
                return false;
            }

            extents = new Extents3d(min, max);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool IntersectsClipBox(OrientedBox clipBox, Extents3d extents)
    {
        extents = ExpandExtents(extents, ScanTolerance);
        return clipBox.Intersects(extents);
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
                    // Fall back below.
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

    private static bool IsExcludedFromSideDbScan(Entity entity, Transaction transaction)
    {
        if (entity is not BlockReference blockReference)
        {
            return false;
        }

        try
        {
            if (transaction.GetObject(blockReference.BlockTableRecord, OpenMode.ForRead, false) is not BlockTableRecord blockTableRecord)
            {
                return string.Equals(entity.Layer, ClipBoxConstants.ClipBoxLayerName, StringComparison.OrdinalIgnoreCase);
            }

            if (string.Equals(blockTableRecord.Name, ClipBoxConstants.UnitCubeBlockName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(entity.Layer, ClipBoxConstants.ClipBoxLayerName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return blockTableRecord.IsFromExternalReference || blockTableRecord.IsFromOverlayReference;
        }
        catch
        {
            return string.Equals(entity.Layer, ClipBoxConstants.ClipBoxLayerName, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static Extents3d ExpandExtents(Extents3d extents, double margin)
    {
        return new Extents3d(
            new Point3d(extents.MinPoint.X - margin, extents.MinPoint.Y - margin, extents.MinPoint.Z - margin),
            new Point3d(extents.MaxPoint.X + margin, extents.MaxPoint.Y + margin, extents.MaxPoint.Z + margin));
    }

    private static (int loaded, int failed) LoadSelectedReferences(Document document, IReadOnlyList<ProjectRefCandidate> selected)
    {
        var loaded = 0;
        var failed = 0;

        using var documentLock = MaybeLockDocument(document);
        using var transaction = document.Database.TransactionManager.StartTransaction();

        var blockTable = (BlockTable)transaction.GetObject(document.Database.BlockTableId, OpenMode.ForRead);
        var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        var hostDirectory = SafeGetDirectoryName(document.Name);

        foreach (var candidate in selected)
        {
            try
            {
                var xrefPath = MakeRelativePath(hostDirectory, candidate.AbsoluteFileName);
                var blockName = MakeUniqueXrefBlockName(blockTable, candidate.FileName);
                var xrefBlockId = document.Database.OverlayXref(xrefPath, blockName);
                if (xrefBlockId.IsNull)
                {
                    failed++;
                    continue;
                }

                var blockReference = new BlockReference(Point3d.Origin, xrefBlockId);
                modelSpace.AppendEntity(blockReference);
                transaction.AddNewlyCreatedDBObject(blockReference, true);
                loaded++;
            }
            catch
            {
                failed++;
            }
        }

        transaction.Commit();
        return (loaded, failed);
    }

    private static string MakeUniqueXrefBlockName(BlockTable blockTable, string fileName)
    {
        var baseName = SanitizeBlockName(Path.GetFileNameWithoutExtension(fileName));
        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = "XREF";
        }

        var uniqueName = baseName;
        var index = 1;
        while (blockTable.Has(uniqueName))
        {
            uniqueName = baseName + "_" + index;
            index++;
        }

        return uniqueName;
    }

    private static string SanitizeBlockName(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return string.Empty;
        }

        var invalidChars = new[] { '<', '>', '/', '\\', '"', ':', ';', '?', '*', '|', ',', '=', '`' };
        var chars = input.Trim().ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            if (invalidChars.Contains(chars[i]) || char.IsControl(chars[i]))
            {
                chars[i] = '_';
            }
        }

        return new string(chars);
    }

    private static Document? FindOpenDocumentByPath(string? fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return null;
        }

        var normalized = NormalizeFullPath(fileName);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return null;
        }

        try
        {
            return AcAp.DocumentManager
                .Cast<Document>()
                .FirstOrDefault(document => !string.IsNullOrWhiteSpace(document?.Name) && PathsEqual(document.Name, normalized));
        }
        catch
        {
            return null;
        }
    }

    private static string NormalizeFullPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        catch
        {
            return path.Trim();
        }
    }

    private static string SafeGetDirectoryName(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return NormalizeFullPath(Path.GetDirectoryName(path));
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeReadAllText(string path)
    {
        try
        {
            return File.Exists(path) ? File.ReadAllText(path) : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string FindAncestorContainingFile(string startDirectory, string fileName)
    {
        try
        {
            var directory = new DirectoryInfo(startDirectory);
            while (directory is not null)
            {
                var candidate = Path.Combine(directory.FullName, fileName);
                if (File.Exists(candidate))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }
        }
        catch
        {
            // Ignore path traversal issues.
        }

        return string.Empty;
    }

    private static string MakeProjectRelativePath(string rootDirectory, string absoluteFileName)
    {
        if (string.IsNullOrWhiteSpace(rootDirectory))
        {
            return absoluteFileName;
        }

        return MakeRelativePath(rootDirectory, absoluteFileName);
    }

    private static string MakeRelativePath(string? fromDirectory, string absoluteFileName)
    {
        if (string.IsNullOrWhiteSpace(fromDirectory))
        {
            return absoluteFileName;
        }

        try
        {
            var fromPath = AppendDirectorySeparator(NormalizeFullPath(fromDirectory));
            var toPath = NormalizeFullPath(absoluteFileName);
            if (string.IsNullOrWhiteSpace(fromPath) || string.IsNullOrWhiteSpace(toPath))
            {
                return absoluteFileName;
            }

            var fromUri = new Uri(fromPath, UriKind.Absolute);
            var toUri = new Uri(toPath, UriKind.Absolute);
            if (!string.Equals(fromUri.Scheme, toUri.Scheme, StringComparison.OrdinalIgnoreCase))
            {
                return absoluteFileName;
            }

            var relativeUri = fromUri.MakeRelativeUri(toUri);
            var relativePath = Uri.UnescapeDataString(relativeUri.ToString()).Replace('/', Path.DirectorySeparatorChar);
            return string.IsNullOrWhiteSpace(relativePath) ? absoluteFileName : relativePath;
        }
        catch
        {
            return absoluteFileName;
        }
    }

    private static string AppendDirectorySeparator(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return path;
        }

        if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
            path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal))
        {
            return path;
        }

        return path + Path.DirectorySeparatorChar;
    }

    private static bool PathsEqual(string left, string right)
    {
        return string.Equals(NormalizeFullPath(left), NormalizeFullPath(right), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFinite(Point3d point)
    {
        return !(double.IsNaN(point.X) || double.IsNaN(point.Y) || double.IsNaN(point.Z) ||
                 double.IsInfinity(point.X) || double.IsInfinity(point.Y) || double.IsInfinity(point.Z));
    }

    private static Type? FindType(string fullName, params string[] assemblyNames)
    {
        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
            var type = assembly.GetType(fullName, false);
            if (type is not null)
            {
                return type;
            }
        }

        foreach (var assemblyName in assemblyNames)
        {
            try
            {
                var assembly = Assembly.Load(assemblyName);
                var type = assembly.GetType(fullName, false);
                if (type is not null)
                {
                    return type;
                }
            }
            catch
            {
                // Ignore load errors. The caller will handle a null result.
            }
        }

        return null;
    }

    private static object? GetStaticPropertyValue(Type type, string propertyName)
    {
        try
        {
            var property = type.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Static);
            return property?.GetValue(null, null);
        }
        catch
        {
            return null;
        }
    }

    private static object? GetPropertyValue(object target, string propertyName)
    {
        try
        {
            var property = target.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
            return property?.GetValue(target, null);
        }
        catch
        {
            return null;
        }
    }

    private static DocumentLock? MaybeLockDocument(Document document)
    {
        return AcAp.DocumentManager.IsApplicationContext
            ? document.LockDocument()
            : null;
    }
}
