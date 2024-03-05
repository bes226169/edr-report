using System.ComponentModel;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace EDR_Report
{
    public class NetworkCopier
    {
        private IConfiguration conf { get; set; }
        private string? defaultDomain { get; set; }
        private string? defaultUserId { get; set; }
        private string? defaultPasswd { get; set; }
        private string? defaultPath { get; set; }
        public NetworkCopier()
        {
            conf = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            GetNASSettings();
            GetDefaultPath();
            if (string.IsNullOrEmpty(defaultPath)) throw new DirectoryNotFoundException($"路徑設定錯誤: {defaultPath}");
        }
        public NetworkCopier(string? domain, string? userid, string? passwd, string? path)
        {
            conf = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            if (string.IsNullOrEmpty(domain) || string.IsNullOrEmpty(userid) || string.IsNullOrEmpty(passwd))
            {
                GetNASSettings();
            }
            else
            {
                defaultDomain = domain;
                defaultUserId = userid;
                defaultPasswd = passwd;
            }
            if (string.IsNullOrEmpty(path))
            {
                GetDefaultPath();
            }
            else
            {
                defaultPath = path;
            }
            if (string.IsNullOrEmpty(defaultPath)) throw new DirectoryNotFoundException($"路徑設定錯誤: {defaultPath}");
        }
        private void GetNASSettings()
        {
            defaultDomain = conf.GetValue<string?>("NAS:Domain");
            defaultUserId = conf.GetValue<string?>("NAS:User");
            defaultPasswd = conf.GetValue<string?>("NAS:Password");
        }
        private void GetDefaultPath()
        {
            var li = new DBFunc().query<dynamic>("erp", "SELECT UPLOADPATH || FILETRANS_PATH NAS_TMP FROM DBENV_INFO");
            if (li.Count() > 0) defaultPath = li.First().NAS_TMP;
        }
        private string NotNull(string? s)
        {
            if (string.IsNullOrEmpty(s))
                throw new ApplicationException("未設定預設登入身分");
            return s;
        }
        private string DefaultDomain => NotNull(defaultDomain);
        private string DefaultUserId => NotNull(defaultUserId);
        private string DefaultPassword => NotNull(defaultPasswd);
        /// <summary>
        /// 組合在 NAS 中的完整路徑
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string GetSharePath(string path)
        {
            var m = Regex.Match(path, @"^\\\\[^\\]+");
            if (m.Success) return path;
            return $"{defaultPath}\\{path}";
        }
        /// <summary>
        /// 複製檔案(from Local to NAS)
        /// </summary>
        /// <param name="srcPath"></param>
        /// <param name="dstPath"></param>
        /// <param name="domain"></param>
        /// <param name="userId"></param>
        /// <param name="passwd"></param>
        /// <returns></returns>
        public (string src, string dst) CopyFromLocal(string srcPath, string dstPath, string? domain = null, string? userId = null, string? passwd = null)
        {
            using (new NetworkConnection(defaultPath,
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                dstPath = GetSharePath(dstPath);
                var folder = Path.GetDirectoryName(dstPath);
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
                File.Copy(srcPath, dstPath);
            }
            return (srcPath, dstPath);
        }
        /// <summary>
        /// 複製檔案(from NAS to Local)
        /// </summary>
        /// <param name="srcPath"></param>
        /// <param name="dstPath"></param>
        /// <param name="domain"></param>
        /// <param name="userId"></param>
        /// <param name="passwd"></param>
        /// <returns></returns>
        public (string src, string dst) CopyFromNAS(string srcPath, string dstPath, string? domain = null, string? userId = null, string? passwd = null)
        {
            using (new NetworkConnection(defaultPath,
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                srcPath = GetSharePath(srcPath);
                var folder = Path.GetDirectoryName(srcPath);
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
                File.Copy(GetSharePath(srcPath), dstPath);
            }
            return (srcPath, dstPath);
        }
        /// <summary>
        /// 直接將 IFormFile 上傳到 NAS
        /// </summary>
        /// <param name="file"></param>
        /// <param name="folder"></param>
        /// <param name="fileName"></param>
        /// <param name="domain"></param>
        /// <param name="userId"></param>
        /// <param name="passwd"></param>
        /// <returns></returns>
        public string StreamSave(IFormFile file, string folder, string fileName, string? domain = null, string? userId = null, string? passwd = null)
        {
            var path = $"{folder}\\{fileName}";
            using (new NetworkConnection(defaultPath, 
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                folder = GetSharePath(folder);
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
                path = $"{folder}\\{fileName}";
                if (File.Exists(path)) File.Delete(path);
                using (var stream = new FileStream(path, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
            }
            return path;
        }
        /// <summary>
        /// 取得 NAS 上的檔案內容
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="domain"></param>
        /// <param name="userId"></param>
        /// <param name="passwd"></param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        public byte[] GetFile(string filePath, string? domain = null, string? userId = null, string? passwd = null)
        {
            using (new NetworkConnection(defaultPath,
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                filePath = GetSharePath(filePath);
                if (File.Exists(filePath))
                {
                    var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    byte[] b = new byte[fs.Length];
                    fs.Read(b, 0, b.Length);
                    fs.Close();
                    return b;
                }
            }
            throw new FileNotFoundException($"檔案不存在: {filePath}");
        }
        /// <summary>
        /// 刪除 NAS 上的檔案
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="domain"></param>
        /// <param name="userId"></param>
        /// <param name="passwd"></param>
        public void DeleteFile(string filePath, string? domain = null, string? userId = null, string? passwd = null)
        {
            using (new NetworkConnection(defaultPath,
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                filePath = GetSharePath(filePath);
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
        }
        /// <summary>
        /// 檢查檔案是否存在 NAS
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="domain"></param>
        /// <param name="userId"></param>
        /// <param name="passwd"></param>
        /// <returns></returns>
        public bool FileExists(string filePath, string? domain = null, string? userId = null, string? passwd = null)
        {
            var rt = false;
            using (new NetworkConnection(defaultPath, 
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                rt = File.Exists(GetSharePath(filePath));
            }
            return rt;
        }
        public void CreateFolder(string srcPath, string? domain = null, string? userId = null, string? passwd = null)
        {
            using (new NetworkConnection(defaultPath, 
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                srcPath = GetSharePath(srcPath);
                if (!Directory.Exists(srcPath))
                    Directory.CreateDirectory(srcPath);
            }
        }
        public string[] DirFiles(string path, string pattern, string? domain = null, string? userId = null, string? passwd = null)
        {
            using (new NetworkConnection(defaultPath,
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain)))
            {
                return Directory.GetFiles(GetSharePath(path), pattern);
            }
        }

        public NetworkConnection GetConnectionContext(string path, string? domain = null, string? userId = null, string? passwd = null) => 
            new NetworkConnection(GetSharePath(path),
                new NetworkCredential(userId ?? DefaultUserId, passwd ?? DefaultPassword, domain ?? DefaultDomain));

    }
    //引用來源: https://stackoverflow.com/a/1197430/288936
    public class NetworkConnection : IDisposable
    {
        string _networkName;
        public NetworkConnection(string networkName, NetworkCredential credentials)
        {
            _networkName = networkName;
            var netResource = new NetResource()
            {
                Scope = ResourceScope.GlobalNetwork,
                ResourceType = ResourceType.Disk,
                DisplayType = ResourceDisplaytype.Share,
                RemoteName = networkName
            };
            var userName = string.IsNullOrEmpty(credentials.Domain)
                ? credentials.UserName
                : string.Format(@"{0}\{1}", credentials.Domain, credentials.UserName);
            var result = WNetAddConnection2(
                netResource,
                credentials.Password,
                userName,
                0);
            if (result != 0)
            {
                throw new Win32Exception(result);
            }
        }
        ~NetworkConnection()
        {
            Dispose(false);
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            WNetCancelConnection2(_networkName, 0, true);
        }
        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(NetResource netResource, string password, string username, int flags);
        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);
    }
    [StructLayout(LayoutKind.Sequential)]
    public class NetResource
    {
        public ResourceScope Scope;
        public ResourceType ResourceType;
        public ResourceDisplaytype DisplayType;
        public int Usage;
        public string LocalName;
        public string RemoteName;
        public string Comment;
        public string Provider;
    }
    public enum ResourceScope : int
    {
        Connected = 1,
        GlobalNetwork,
        Remembered,
        Recent,
        Context
    };
    public enum ResourceType : int
    {
        Any = 0,
        Disk = 1,
        Print = 2,
        Reserved = 8,
    }
    public enum ResourceDisplaytype : int
    {
        Generic = 0x0,
        Domain = 0x01,
        Server = 0x02,
        Share = 0x03,
        File = 0x04,
        Group = 0x05,
        Network = 0x06,
        Root = 0x07,
        Shareadmin = 0x08,
        Directory = 0x09,
        Tree = 0x0a,
        Ndscontainer = 0x0b
    }
}
