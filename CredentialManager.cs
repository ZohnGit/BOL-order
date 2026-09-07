using System.Runtime.InteropServices;
using System.Text;

namespace BolOrderExporter;

internal static class CredentialManager
{
    private const string Prefix = "BolOrderExporter:";
    private const uint CredTypeGeneric = 1;
    private const uint CredPersistLocalMachine = 2;

    public static IReadOnlyList<SavedCredential> LoadAll()
    {
        var result = new List<SavedCredential>();
        if (!CredEnumerate(Prefix + "*", 0, out var count, out var credentials))
        {
            return result;
        }

        try
        {
            for (var i = 0; i < count; i++)
            {
                var pointer = Marshal.ReadIntPtr(credentials, i * IntPtr.Size);
                var credential = Marshal.PtrToStructure<NativeCredential>(pointer);
                if (string.IsNullOrWhiteSpace(credential.UserName)) continue;

                var password = string.Empty;
                if (credential.CredentialBlob != IntPtr.Zero && credential.CredentialBlobSize > 0)
                {
                    var bytes = new byte[credential.CredentialBlobSize];
                    Marshal.Copy(credential.CredentialBlob, bytes, 0, bytes.Length);
                    password = Encoding.Unicode.GetString(bytes).TrimEnd('\0');
                }

                result.Add(new SavedCredential(credential.UserName, password));
            }
        }
        finally
        {
            CredFree(credentials);
        }

        return result.OrderBy(x => x.User, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public static void Save(string user, string password)
    {
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        var blob = Marshal.AllocCoTaskMem(passwordBytes.Length);
        try
        {
            Marshal.Copy(passwordBytes, 0, blob, passwordBytes.Length);
            var credential = new NativeCredential
            {
                Type = CredTypeGeneric,
                TargetName = TargetName(user),
                CredentialBlobSize = (uint)passwordBytes.Length,
                CredentialBlob = blob,
                Persist = CredPersistLocalMachine,
                UserName = user
            };

            if (!CredWrite(ref credential, 0))
            {
                throw new InvalidOperationException($"无法保存账号信息（Windows 错误 {Marshal.GetLastWin32Error()}）。");
            }
        }
        finally
        {
            for (var i = 0; i < passwordBytes.Length; i++) Marshal.WriteByte(blob, i, 0);
            Marshal.FreeCoTaskMem(blob);
            Array.Clear(passwordBytes);
        }
    }

    public static void Delete(string user)
    {
        if (!CredDelete(TargetName(user), CredTypeGeneric, 0))
        {
            var error = Marshal.GetLastWin32Error();
            if (error != 1168)
            {
                throw new InvalidOperationException($"无法删除账号信息（Windows 错误 {error}）。");
            }
        }
    }

    private static string TargetName(string user)
    {
        var safe = Convert.ToBase64String(Encoding.UTF8.GetBytes(user))
            .Replace('/', '_')
            .Replace('+', '-')
            .TrimEnd('=');
        return Prefix + safe;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public uint Flags;
        public uint Type;
        public string TargetName;
        public string? Comment;
        public long LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public string? TargetAlias;
        public string UserName;
    }

    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredWrite(ref NativeCredential credential, uint flags);

    [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredDelete(string target, uint type, uint flags);

    [DllImport("advapi32.dll", EntryPoint = "CredEnumerateW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredEnumerate(string filter, uint flags, out int count, out IntPtr credentials);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern void CredFree(IntPtr buffer);
}
