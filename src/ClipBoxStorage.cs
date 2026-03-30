using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace Plant3DClipBox;

internal static class ClipBoxStorage
{
    private static readonly char[] InvalidNameChars =
    {
        '|', '*', '\\', ':', ';', '<', '>', '?', '"', ',', '=', '`', '/',
    };

    public static IReadOnlyList<string> GetNames(Database database)
    {
        var names = new List<string>();

        using var transaction = database.TransactionManager.StartTransaction();
        var root = OpenRootDictionary(transaction, database, createIfMissing: false);
        if (root is null)
        {
            return names;
        }

        foreach (DBDictionaryEntry entry in root)
        {
            names.Add(entry.Key);
        }

        names.Sort(StringComparer.OrdinalIgnoreCase);
        return names;
    }

    public static void Save(Database database, string name, OrientedBox box)
    {
        var key = SanitizeName(name);

        using var transaction = database.TransactionManager.StartTransaction();
        var root = OpenRootDictionary(transaction, database, createIfMissing: true)!;

        var values = new[]
        {
            new TypedValue((int)DxfCode.Text, "V1"),
            new TypedValue((int)DxfCode.Real, box.Center.X),
            new TypedValue((int)DxfCode.Real, box.Center.Y),
            new TypedValue((int)DxfCode.Real, box.Center.Z),
            new TypedValue((int)DxfCode.Real, box.XAxis.X),
            new TypedValue((int)DxfCode.Real, box.XAxis.Y),
            new TypedValue((int)DxfCode.Real, box.XAxis.Z),
            new TypedValue((int)DxfCode.Real, box.YAxis.X),
            new TypedValue((int)DxfCode.Real, box.YAxis.Y),
            new TypedValue((int)DxfCode.Real, box.YAxis.Z),
            new TypedValue((int)DxfCode.Real, box.ZAxis.X),
            new TypedValue((int)DxfCode.Real, box.ZAxis.Y),
            new TypedValue((int)DxfCode.Real, box.ZAxis.Z),
            new TypedValue((int)DxfCode.Real, box.HalfX),
            new TypedValue((int)DxfCode.Real, box.HalfY),
            new TypedValue((int)DxfCode.Real, box.HalfZ),
        };

        var resultBuffer = new ResultBuffer(values);

        if (root.Contains(key))
        {
            var xrecord = (Xrecord)transaction.GetObject(root.GetAt(key), OpenMode.ForWrite);
            xrecord.Data = resultBuffer;
        }
        else
        {
            root.UpgradeOpen();
            var xrecord = new Xrecord
            {
                Data = resultBuffer,
            };

            root.SetAt(key, xrecord);
            transaction.AddNewlyCreatedDBObject(xrecord, true);
        }

        transaction.Commit();
    }

    public static bool TryLoad(Database database, string name, out OrientedBox box)
    {
        box = default;
        var key = SanitizeName(name);

        using var transaction = database.TransactionManager.StartTransaction();
        var root = OpenRootDictionary(transaction, database, createIfMissing: false);
        if (root is null || !root.Contains(key))
        {
            return false;
        }

        var xrecord = transaction.GetObject(root.GetAt(key), OpenMode.ForRead, false) as Xrecord;
        if (xrecord is null)
        {
            return false;
        }

        using var data = xrecord.Data;
        var typedValues = data?.AsArray();
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

    public static void Delete(Database database, string name)
    {
        var key = SanitizeName(name);

        using var transaction = database.TransactionManager.StartTransaction();
        var root = OpenRootDictionary(transaction, database, createIfMissing: false);
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

    public static string NormalizeName(string name)
    {
        return SanitizeName(name);
    }

    private static DBDictionary? OpenRootDictionary(Transaction transaction, Database database, bool createIfMissing)
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
}
