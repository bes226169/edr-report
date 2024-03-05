using Lib.BesSSO;
using System.Text;
using System.Text.Json;

namespace Lib
{
    public class BesSSOLib
    {
        /// <summary>
        /// 當設定檔載入後，回傳true。
        /// </summary>
        public bool CanUse { get => Configs != null; }
        /// <summary>
        /// 是否從Client端Google帳號登入。
        /// 若是使用QR Code登入，回傳false。
        /// </summary>
        public bool GoogleIsLogin { get; private set; } = false;
        /// <summary>
        /// 設定檔
        /// </summary>
        public BesSSOModel Configs { get; set; }
        /// <summary>
        /// State參數，可以用來核對連線
        /// </summary>
        public string State { get; set; } = string.Empty;
        private HttpClient client { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public BesSSOLib() 
        {
            Configs = new();
            client = new();
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="_configs"></param>
        public BesSSOLib(BesSSOModel _configs)
        {
            Configs = _configs;
            Configs.Path = Configs.Path.TrimEnd('/');
            client = new();
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="path"></param>
        /// <param name="client_id"></param>
        /// <param name="client_secret"></param>
        /// <param name="scope"></param>
        public BesSSOLib(string path, string client_id, string client_secret, string scope)
        {
            Configs = new(path.TrimEnd('/'), client_id, client_secret, scope);
            client = new();
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="path"></param>
        /// <param name="client_id"></param>
        /// <param name="client_secret"></param>
        /// <param name="scope"></param>
        public BesSSOLib(string path, string client_id, string client_secret, string[] scope)
        {
            Configs = new(path.TrimEnd('/'), client_id, client_secret, scope);
            client = new();
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="path"></param>
        /// <param name="client_id"></param>
        /// <param name="client_secret"></param>
        /// <param name="scope"></param>
        public BesSSOLib(string path, string client_id, string client_secret, List<string> scope)
        {
            Configs = new(path.TrimEnd('/'), client_id, client_secret, scope);
            client = new();
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="path"></param>
        /// <param name="client_id"></param>
        /// <param name="client_secret"></param>
        /// <param name="scope"></param>
        public BesSSOLib(string path, string client_id, string client_secret, IEnumerable<string> scope)
        {
            Configs = new(path.TrimEnd('/'), client_id, client_secret, scope);
            client = new();
        }
        /// <summary>
        /// 取得登入頁網址
        /// </summary>
        /// <param name="redirect_uri"></param>
        /// <param name="state"></param>
        /// <returns></returns>
        public string AuthUrl(string redirect_uri, string? state = null)
        {
            if (!string.IsNullOrEmpty(state)) State = state;
            else State = DateTime.Now.Ticks.ToString();
            return $"{Configs.Path}/auth" +
                "?grant_types=authorization_code" +
                "&response_type=id_token" +
                $"&client_id={Configs.ClientID}" +
                $"&client_secret={Configs.ClientSecret}" +
                $"&scope={Uri.EscapeDataString(Configs.Scope)}" +
                $"&redirect_uri={redirect_uri}" +
                $"&state={State}";
        }
        /// <summary>
        /// 取得QR Code登入頁網址
        /// </summary>
        /// <param name="redirect_uri"></param>
        /// <param name="state"></param>
        /// <returns></returns>
        public string QRCodeAuthUrl(string redirect_uri, string? state = null)
        {
            if (!string.IsNullOrEmpty(state)) State = state;
            else State = DateTime.Now.Ticks.ToString();
            return $"{Configs.Path}/sseauth" +
                "?grant_types=authorization_code" +
                "&response_type=id_token" +
                $"&client_id={Configs.ClientID}" +
                $"&client_secret={Configs.ClientSecret}" +
                $"&scope={Uri.EscapeDataString(Configs.Scope)}" +
                $"&redirect_uri={redirect_uri}" +
                $"&state={State}";
        }
        /// <summary>
        /// 取得登出網址
        /// </summary>
        /// <param name="id_token"></param>
        /// <param name="redirect_uri"></param>
        /// <param name="logout_with_sso"></param>
        /// <param name="logout_with_google"></param>
        /// <returns></returns>
        public string LogoutUrl(string id_token, string redirect_uri, bool logout_with_sso = false, bool logout_with_google = false) => $"{Configs.Path}/logout" +
            $"?id_token={Uri.EscapeDataString(id_token)}" +
            $"&redirect_uri={Uri.EscapeDataString(redirect_uri)}" +
            $"&f={(logout_with_sso ? "1" : "0")}" + 
            $"&g={(logout_with_google ? "1" : "0")}";
        /// <summary>
        /// 取得使用者資料
        /// </summary>
        /// <param name="data"></param>
        /// <param name="auto_refresh_token"></param>
        /// <returns></returns>
        public async Task<ResponseModel?> GetProfile(ResponseModel req, bool auto_refresh_token = true)
        {
            if (req.g == true) GoogleIsLogin = true;
            var data = new Dictionary<string, string?>() 
            {
                { "access_token", req.access_token },
                { "refresh_token", auto_refresh_token ? req.refresh_token : null }
            };
            var content = new StringContent(JsonSerializer.Serialize(data), Encoding.UTF8, "application/json");
            client.DefaultRequestHeaders.Authorization = new("Bearer", req.id_token);
            var res = await client.PostAsync($"{Configs.Path}/profile", content);
            var body = await res.Content.ReadAsStringAsync();
            var rt = JsonSerializer.Deserialize<ResponseModel>(body);
            client.Dispose();
            return rt;
        }
        /// <summary>
        /// 更新Token
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public async Task<ResponseModel?> RefreshToken(ResponseModel req)
        {
            var data = new Dictionary<string, string?>()
            {
                { "refresh_token", req.refresh_token }
            };
            var content = new StringContent(JsonSerializer.Serialize(data), Encoding.UTF8, "application/json");
            client.DefaultRequestHeaders.Authorization = new("Bearer", req.id_token);
            var res = await client.PostAsync($"{Configs.Path}/refresh_token", content);
            var body = await res.Content.ReadAsStringAsync();
            var rt = JsonSerializer.Deserialize<ResponseModel>(body);
            client.Dispose();
            return rt;
        }
    }
}