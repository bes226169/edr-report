namespace Lib.BesSSO
{
    public class BesSSOModel
    {
        public string Path { get; set; }
        public string ClientID { get; set; }
        public string ClientSecret { get; set; }
        public string Scope { get; set; }
        public BesSSOModel()
        {
            Path = string.Empty;
            ClientID = string.Empty;
            ClientSecret = string.Empty;
            Scope = "openid";
        }
        public BesSSOModel(string path, string client_id, string client_secret, string scope)
        {
            Path = path;
            ClientID = client_id;
            ClientSecret = client_secret;
            Scope = scope;
        }
        public BesSSOModel(string path, string client_id, string client_secret, string[] scope)
        {
            Path = path;
            ClientID = client_id;
            ClientSecret = client_secret;
            Scope = string.Join(' ', scope);
        }
        public BesSSOModel(string path, string client_id, string client_secret, List<string> scope)
        {
            Path = path;
            ClientID = client_id;
            ClientSecret = client_secret;
            Scope = string.Join(' ', scope);
        }
        public BesSSOModel(string path, string client_id, string client_secret, IEnumerable<string> scope)
        {
            Path = path;
            ClientID = client_id;
            ClientSecret = client_secret;
            Scope = string.Join(' ', scope);
        }
    }
}
