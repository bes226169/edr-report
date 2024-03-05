namespace Lib.BesSSO
{
    public class ResponseModel
    {
        public string? access_token { get; set; }
        public string? token_type { get; set; }
        public string? id_token { get; set; }
        public int? expires_in { get; set; }
        public string? refresh_token { get; set; }
        public string? error { get; set; }
        public string? error_description { get; set; }
        public bool? g { get; set; }
        public int? status_code { get; set; }
        public ProfileModel? profile { get; set; }
        public bool IsSuccess { get => string.IsNullOrEmpty(error); }
        public ResponseModel() { }
    }
}
