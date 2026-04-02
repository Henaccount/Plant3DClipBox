using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace Plant3DClipBox;

internal static class ClipBoxStorage
{
    private static readonly char[] InvalidNameChars =
    {
        '|', '*', '\\', ':', ';', '<', '>', '?', '"', ',', '=', '`', '/',
    };

    private static readonly object SyncRoot = new();

    public static IReadOnlyList<string> GetNames(Database database)
    {
        TryEnsureLegacyStatesMigrated(database);

        lock (SyncRoot)
        {
            var store = ReadGlobalStore();
            return store.States.Keys
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }

    public static void Save(Database database, string name, OrientedBox box)
    {
        var key = SanitizeName(name);
        TryEnsureLegacyStatesMigrated(database);

        lock (SyncRoot)
        {
            var store = ReadGlobalStore();
            store.DeletedNames.Remove(key);
            store.States[key] = box;
            WriteGlobalStore(store);
        }
    }

    public static bool TryLoad(Database database, string name, out OrientedBox box)
    {
        box = default;
        var key = SanitizeName(name);
        TryEnsureLegacyStatesMigrated(database);

        lock (SyncRoot)
        {
            var store = ReadGlobalStore();
            if (store.States.TryGetValue(key, out box))
            {
                return box.IsValid;
            }

            if (store.DeletedNames.Contains(key))
            {
                box = default;
                return false;
            }
        }

        return TryLoadLegacy(database, key, out box);
    }

    public static void Delete(Database database, string name)
    {
        var key = SanitizeName(name);

        lock (SyncRoot)
        {
            var store = ReadGlobalStore();
            store.States.Remove(key);
            store.DeletedNames.Add(key);
            WriteGlobalStore(store);
        }

        TryDeleteLegacy(database, key);
    }

    public static string NormalizeName(string name)
    {
        return SanitizeName(name);
    }

    private static void TryEnsureLegacyStatesMigrated(Database database)
    {
        try
        {
            EnsureLegacyStatesMigrated(database);
        }
        catch
        {
            // Keep save/load usable even if legacy migration cannot run.
        }
    }

    private static void EnsureLegacyStatesMigrated(Database database)
    {
        if (database is null)
        {
            return;
        }

        var legacyStates = ReadAllLegacyStates(database);
        if (legacyStates.Count == 0)
        {
            return;
        }

        lock (SyncRoot)
        {
            var store = ReadGlobalStore();
            var changed = false;

            foreach (var pair in legacyStates)
            {
                if (store.DeletedNames.Contains(pair.Key) || store.States.ContainsKey(pair.Key))
                {
                    continue;
                }

                store.States[pair.Key] = pair.Value;
                changed = true;
            }

            if (changed)
            {
                WriteGlobalStore(store);
            }
        }
    }

    private static Dictionary<string, OrientedBox> ReadAllLegacyStates(Database database)
    {
        var result = new Dictionary<string, OrientedBox>(StringComparer.OrdinalIgnoreCase);

        using var transaction = database.TransactionManager.StartTransaction();
        var root = OpenLegacyRootDictionary(transaction, database, createIfMissing: false);
        if (root is null)
        {
            return result;
        }

        foreach (DBDictionaryEntry entry in root)
        {
            if (transaction.GetObject(entry.Value, OpenMode.ForRead, false) is not Xrecord xrecord)
            {
                continue;
            }

            using var data = xrecord.Data;
            if (TryParseBox(data?.AsArray(), out var box))
            {
                result[entry.Key] = box;
            }
        }

        return result;
    }

    private static bool TryLoadLegacy(Database database, string key, out OrientedBox box)
    {
        box = default;

        try
        {
            using var transaction = database.TransactionManager.StartTransaction();
            var root = OpenLegacyRootDictionary(transaction, database, createIfMissing: false);
            if (root is null || !root.Contains(key))
            {
                return false;
            }

            if (transaction.GetObject(root.GetAt(key), OpenMode.ForRead, false) is not Xrecord xrecord)
            {
                return false;
            }

            using var data = xrecord.Data;
            return TryParseBox(data?.AsArray(), out box);
        }
        catch
        {
            box = default;
            return false;
        }
    }

    private static void TryDeleteLegacy(Database database, string key)
    {
        if (database is null)
        {
            return;
        }

        try
        {
            using var transaction = database.TransactionManager.StartTransaction();
            var root = OpenLegacyRootDictionary(transaction, database, createIfMissing: false);
            if (root is null || !root.Contains(key))
            {
                return;
            }

            root.UpgradeOpen();
            var id = root.GetAt(key);
            root.Remove(key);

            if (!id.IsNull && !id.IsErased)
            {
                var dbObject = transaction.GetObject(id, OpenMode.ForWrite, false);
                dbObject?.Erase();
            }

            transaction.Commit();
        }
        catch
        {
            // Best-effort cleanup only.
        }
    }

    private static bool TryParseBox(TypedValue[]? typedValues, out OrientedBox box)
    {
        box = default;
        if (typedValues is null || typedValues.Length < 16)
        {
            return false;
        }

        var index = 0;
        _ = typedValues[index++].Value?.ToString();

        double NextReal()
        {
            var value = typedValues[index++].Value;
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        try
        {
            box = new OrientedBox(
                new Point3d(NextReal(), NextReal(), NextReal()),
                new Vector3d(NextReal(), NextReal(), NextReal()),
                new Vector3d(NextReal(), NextReal(), NextReal()),
                new Vector3d(NextReal(), NextReal(), NextReal()),
                NextReal(),
                NextReal(),
                NextReal());

            return box.IsValid;
        }
        catch
        {
            box = default;
            return false;
        }
    }

    private static DBDictionary? OpenLegacyRootDictionary(Transaction transaction, Database database, bool createIfMissing)
    {
        var namedObjects = (DBDictionary)transaction.GetObject(database.NamedObjectsDictionaryId, OpenMode.ForRead);
        if (namedObjects.Contains(ClipBoxConstants.StateDictionaryName))
        {
            return (DBDictionary)transaction.GetObject(namedObjects.GetAt(ClipBoxConstants.StateDictionaryName), OpenMode.ForRead);
        }

        if (!createIfMissing)
        {
            return null;
        }

        namedObjects.UpgradeOpen();
        var root = new DBDictionary();
        namedObjects.SetAt(ClipBoxConstants.StateDictionaryName, root);
        transaction.AddNewlyCreatedDBObject(root, true);
        return root;
    }

    private static GlobalStoreSnapshot ReadGlobalStore()
    {
        var snapshot = new GlobalStoreSnapshot();
        var path = GetGlobalStorePath();
        if (!File.Exists(path))
        {
            return snapshot;
        }

        try
        {
            var document = XDocument.Load(path);
            var root = document.Root;
            if (root is null)
            {
                return snapshot;
            }

            foreach (var deletedElement in root.Elements("deleted"))
            {
                var name = SanitizeName((string?)deletedElement.Attribute("name") ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    snapshot.DeletedNames.Add(name);
                }
            }

            foreach (var stateElement in root.Elements("state"))
            {
                var name = SanitizeName((string?)stateElement.Attribute("name") ?? string.Empty);
                if (string.IsNullOrWhiteSpace(name) || snapshot.DeletedNames.Contains(name))
                {
                    continue;
                }

                if (!TryReadXmlBox(stateElement, out var box))
                {
                    continue;
                }

                snapshot.States[name] = box;
            }
        }
        catch
        {
            return new GlobalStoreSnapshot();
        }

        return snapshot;
    }

    private static void WriteGlobalStore(GlobalStoreSnapshot snapshot)
    {
        var path = GetGlobalStorePath();
        var directory = Path.GetDirectoryName(path) ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var root = new XElement("clipBoxStates",
            new XAttribute("version", "2"));

        foreach (var name in snapshot.States.Keys.OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
        {
            var box = snapshot.States[name];
            root.Add(new XElement("state",
                new XAttribute("name", name),
                new XAttribute("cx", Format(box.Center.X)),
                new XAttribute("cy", Format(box.Center.Y)),
                new XAttribute("cz", Format(box.Center.Z)),
                new XAttribute("xx", Format(box.XAxis.X)),
                new XAttribute("xy", Format(box.XAxis.Y)),
                new XAttribute("xz", Format(box.XAxis.Z)),
                new XAttribute("yx", Format(box.YAxis.X)),
                new XAttribute("yy", Format(box.YAxis.Y)),
                new XAttribute("yz", Format(box.YAxis.Z)),
                new XAttribute("zx", Format(box.ZAxis.X)),
                new XAttribute("zy", Format(box.ZAxis.Y)),
                new XAttribute("zz", Format(box.ZAxis.Z)),
                new XAttribute("hx", Format(box.HalfX)),
                new XAttribute("hy", Format(box.HalfY)),
                new XAttribute("hz", Format(box.HalfZ))));
        }

        foreach (var name in snapshot.DeletedNames
            .Where(name => !snapshot.States.ContainsKey(name))
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
        {
            root.Add(new XElement("deleted", new XAttribute("name", name)));
        }

        var document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        var tempPath = path + ".tmp";
        document.Save(tempPath);

        if (File.Exists(path))
        {
            File.Copy(tempPath, path, true);
            File.Delete(tempPath);
        }
        else
        {
            File.Move(tempPath, path);
        }
    }

    private static bool TryReadXmlBox(XElement stateElement, out OrientedBox box)
    {
        box = default;

        if (!TryReadDouble(stateElement, "cx", out var cx) ||
            !TryReadDouble(stateElement, "cy", out var cy) ||
            !TryReadDouble(stateElement, "cz", out var cz) ||
            !TryReadDouble(stateElement, "xx", out var xx) ||
            !TryReadDouble(stateElement, "xy", out var xy) ||
            !TryReadDouble(stateElement, "xz", out var xz) ||
            !TryReadDouble(stateElement, "yx", out var yx) ||
            !TryReadDouble(stateElement, "yy", out var yy) ||
            !TryReadDouble(stateElement, "yz", out var yz) ||
            !TryReadDouble(stateElement, "zx", out var zx) ||
            !TryReadDouble(stateElement, "zy", out var zy) ||
            !TryReadDouble(stateElement, "zz", out var zz) ||
            !TryReadDouble(stateElement, "hx", out var hx) ||
            !TryReadDouble(stateElement, "hy", out var hy) ||
            !TryReadDouble(stateElement, "hz", out var hz))
        {
            return false;
        }

        box = new OrientedBox(
            new Point3d(cx, cy, cz),
            new Vector3d(xx, xy, xz),
            new Vector3d(yx, yy, yz),
            new Vector3d(zx, zy, zz),
            hx,
            hy,
            hz);

        return box.IsValid;
    }

    private static bool TryReadDouble(XElement element, string attributeName, out double value)
    {
        value = 0.0;
        var raw = (string?)element.Attribute(attributeName);
        return !string.IsNullOrWhiteSpace(raw) &&
               double.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
    }

    private static string GetGlobalStorePath()
    {
        var baseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        if (string.IsNullOrWhiteSpace(baseDirectory))
        {
            baseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        }

        if (string.IsNullOrWhiteSpace(baseDirectory))
        {
            baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        }

        return Path.Combine(baseDirectory, "Autodesk", "Plant3DClipBox", "SavedClipBoxes.xml");
    }

    private static string Format(double value)
    {
        return value.ToString("R", CultureInfo.InvariantCulture);
    }

    private static string SanitizeName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return "Box";
        }

        var trimmed = name.Trim();
        var builder = new StringBuilder(trimmed.Length);
        foreach (var character in trimmed)
        {
            if (char.IsControl(character) || Array.IndexOf(InvalidNameChars, character) >= 0)
            {
                builder.Append('_');
            }
            else
            {
                builder.Append(character);
            }
        }

        var sanitized = builder.ToString().Trim();
        return string.IsNullOrWhiteSpace(sanitized) ? "Box" : sanitized;
    }

    private sealed class GlobalStoreSnapshot
    {
        public Dictionary<string, OrientedBox> States { get; } = new(StringComparer.OrdinalIgnoreCase);

        public HashSet<string> DeletedNames { get; } = new(StringComparer.OrdinalIgnoreCase);
    }
}
