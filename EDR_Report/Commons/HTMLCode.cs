using System.Text;

namespace EDR_Report
{
    public class HTMLCode
    {
        private StringBuilder text { get; set; }
        public HTMLCode()
        {
            text = new();
        }
        /// <summary>
        /// div tag
        /// </summary>
        /// <param name="str"></param>
        /// <param name="style"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public HTMLCode Div(string? str, Style? style = null, Dictionary<string, string>? attr = null)
        {
            Tag("div", str, style, attr);
            return this;
        }
        /// <summary>
        /// p tag
        /// </summary>
        /// <param name="str"></param>
        /// <param name="style"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public HTMLCode P(string? str, Style? style = null, Dictionary<string, string>? attr = null)
        {
            Tag("p", str, style, attr);
            return this;
        }
        /// <summary>
        /// span tag
        /// </summary>
        /// <param name="str"></param>
        /// <param name="style"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public HTMLCode Span(string? str, Style? style = null, Dictionary<string, string>? attr = null)
        {
            Tag("span", str, style, attr);
            return this;
        }
        /// <summary>
        /// label tag
        /// </summary>
        /// <param name="str"></param>
        /// <param name="style"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public HTMLCode Label(string? str, Style? style = null, Dictionary<string, string>? attr = null)
        {
            Tag("label", str, style, attr);
            return this;
        }
        /// <summary>
        /// br tag(換行)
        /// </summary>
        /// <returns></returns>
        public HTMLCode Br()
        {
            text.Append("<br>");
            return this;
        }
        /// <summary>
        /// hr tag(水平線)
        /// </summary>
        /// <returns></returns>
        public HTMLCode Hr()
        {
            text.Append("<hr>");
            return this;
        }
        /// <summary>
        /// 一般文字(允許HTML)
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public HTMLCode Text(string? str)
        {
            text.Append(str);
            return this;
        }
        /// <summary>
        /// 產生 tag
        /// </summary>
        /// <param name="tagName"></param>
        /// <param name="str"></param>
        /// <param name="style"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public HTMLCode Tag(string tagName, string? str, Style? style = null, Dictionary<string, string>? attr = null)
        {
            var a = string.Empty;
            if (attr != null && attr.Count > 0)
            {
                foreach (var kvp in attr) a += $" {kvp.Key}=\"{kvp.Value??string.Empty}\"";
            }
            text.Append($"<{tagName} style=\"{(style == null ? "" : style.ToString())}\"{a??string.Empty}>{str??string.Empty}</{tagName}>");
            return this;
        }
        /// <summary>
        /// 插入到新的Tag裡面
        /// </summary>
        /// <param name="tagName"></param>
        /// <param name="style"></param>
        /// <param name="attr"></param>
        /// <param name="html"></param>
        /// <returns></returns>
        public HTMLCode Append(string tagName, Style? style = null, Dictionary<string, string>? attr = null, params HTMLCode[] html)
        {
            if (html == null || html.Length == 0) Tag(tagName, string.Empty, style, attr);
            else
            {
                var t = new StringBuilder();
                foreach (var h in html)
                {
                    if (h == null) continue;
                    t.Append(h.ToString());
                }
                Tag(tagName, t.ToString(), style, attr);
            }
            return this;
        }
        /// <summary>
        /// 直接插入到後面
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        public HTMLCode Append(params HTMLCode[] html)
        {
            if (html != null && html.Length > 0)
            {
                foreach (var h in html)
                {
                    if (h == null) continue;
                    Text(h.ToString());
                }
            }
            return this;
        }
        /// <summary>
        /// 取得 HTML Code
        /// </summary>
        /// <returns></returns>
        public override string ToString() => text.ToString();
        public class Style
        {
            /// <summary>文字顏色(可使用色碼、rgb、rgba)</summary>
            public string? FontColor { get; set; }
            /// <summary>背景色(可使用色碼、rgb、rgba)</summary>
            public string? BackgroundColor { get; set; }
            /// <summary>文字大小(px)</summary>
            public int? FontSize { get; set; }
            /// <summary>斜體</summary>
            public bool? Italic { get; set; }
            /// <summary>粗體</summary>
            public bool? Bold { get; set; }
            /// <summary>自訂樣式</summary>
            public string? Custom { get; set; }
            /// <summary>水平對齊</summary>
            public string? TextAlign { get; set; }
            /// <summary>垂直對齊</summary>
            public string? VerticalAlign { get; set; }
            public override string ToString()
            {
                var rt = string.Empty;
                if (!string.IsNullOrEmpty(FontColor)) rt += $"color:{FontColor};";
                if (!string.IsNullOrEmpty(BackgroundColor)) rt += $"background:{BackgroundColor};";
                if (FontSize.HasValue && FontSize > 0) rt += $"font-size:{FontSize}px;";
                if (Italic.HasValue && Italic.Value) rt += "font-style: italic;";
                if (Bold.HasValue && Bold.Value) rt += "font-weight: bold;";
                if (!string.IsNullOrEmpty(TextAlign)) rt += $"text-align:{TextAlign};";
                if (!string.IsNullOrEmpty(VerticalAlign)) rt += $"vertical-align:{VerticalAlign};";
                if (!string.IsNullOrEmpty(Custom)) rt += $"{Custom};";
                return rt;
            }
        }
        /// <summary>水平對齊</summary>
        public static class TextAlign
        {
            /// <summary>水平對齊:靠左</summary>
            public static readonly string Left = "left";
            /// <summary>水平對齊:置中</summary>
            public static readonly string Center = "center";
            /// <summary>水平對齊:靠右</summary>
            public static readonly string Right = "right";
        }
        /// <summary>垂直對齊</summary>
        public static class VerticalAlign
        {
            /// <summary>垂直對齊:置頂</summary>
            public static readonly string Top = "top";
            /// <summary>垂直對齊:置中</summary>
            public static readonly string Middle = "middle";
            /// <summary>垂直對齊:置底</summary>
            public static readonly string Bottom = "bottom";
        }
        /// <summary>
        /// 色碼
        /// </summary>
        public static class Color
        {
            /// <summary>#F0F8FF</summary>
            public static readonly string AliceBlue = "aliceblue";
            /// <summary>#FAEBD7</summary>
            public static readonly string AntiqueWhite = "antiquewhite";
            /// <summary>#00FFFF</summary>
            public static readonly string Aqua = "aqua";
            /// <summary>#7FFFD4</summary>
            public static readonly string Aquamarine = "aquamarine";
            /// <summary>#F0FFFF</summary>
            public static readonly string Azure = "azure";
            /// <summary>#F5F5DC</summary>
            public static readonly string Beige = "beige";
            /// <summary>#FFE4C4</summary>
            public static readonly string Bisque = "bisque";
            /// <summary>#000000</summary>
            public static readonly string Black = "black";
            /// <summary>#FFEBCD</summary>
            public static readonly string BlanchedAlmond = "blanchedalmond";
            /// <summary>#0000FF</summary>
            public static readonly string Blue = "blue";
            /// <summary>#8A2BE2</summary>
            public static readonly string BlueViolet = "blueviolet";
            /// <summary>#A52A2A</summary>
            public static readonly string Brown = "brown";
            /// <summary>#DEB887</summary>
            public static readonly string BurlyWood = "burlywood";
            /// <summary>#5F9EA0</summary>
            public static readonly string CadetBlue = "cadetblue";
            /// <summary>#7FFF00</summary>
            public static readonly string Chartreuse = "chartreuse";
            /// <summary>#D2691E</summary>
            public static readonly string Chocolate = "chocolate";
            /// <summary>#FF7F50</summary>
            public static readonly string Coral = "coral";
            /// <summary>#6495ED</summary>
            public static readonly string CornflowerBlue = "cornflowerblue";
            /// <summary>#FFF8DC</summary>
            public static readonly string Cornsilk = "cornsilk";
            /// <summary>#DC143C</summary>
            public static readonly string Crimson = "crimson";
            /// <summary>#00FFFF</summary>
            public static readonly string Cyan = "cyan";
            /// <summary>#00008B</summary>
            public static readonly string DarkBlue = "darkblue";
            /// <summary>#008B8B</summary>
            public static readonly string DarkCyan = "darkcyan";
            /// <summary>#B8860B</summary>
            public static readonly string DarkGoldenRod = "darkgoldenrod";
            /// <summary>#A9A9A9</summary>
            public static readonly string DarkGray = "darkgray";
            /// <summary>#006400</summary>
            public static readonly string DarkGreen = "darkgreen";
            /// <summary>#BDB76B</summary>
            public static readonly string DarkKhaki = "darkkhaki";
            /// <summary>#8B008B</summary>
            public static readonly string DarkMagenta = "darkmagenta";
            /// <summary>#556B2F</summary>
            public static readonly string DarkOliveGreen = "darkolivegreen";
            /// <summary>#FF8C00</summary>
            public static readonly string DarkOrange = "darkorange";
            /// <summary>#9932CC</summary>
            public static readonly string DarkOrchid = "darkorchid";
            /// <summary>#8B0000</summary>
            public static readonly string DarkRed = "darkred";
            /// <summary>#E9967A</summary>
            public static readonly string DarkSalmon = "darksalmon";
            /// <summary>#8FBC8F</summary>
            public static readonly string DarkSeaGreen = "darkseagreen";
            /// <summary>#483D8B</summary>
            public static readonly string DarkSlateBlue = "darkslateblue";
            /// <summary>#2F4F4F</summary>
            public static readonly string DarkSlateGray = "darkslategray";
            /// <summary>#00CED1</summary>
            public static readonly string DarkTurquoise = "darkturquoise";
            /// <summary>#9400D3</summary>
            public static readonly string DarkViolet = "darkviolet";
            /// <summary>#FF1493</summary>
            public static readonly string DeepPink = "deeppink";
            /// <summary>#00BFFF</summary>
            public static readonly string DeepSkyBlue = "deepskyblue";
            /// <summary>#696969</summary>
            public static readonly string DimGray = "dimgray";
            /// <summary>#1E90FF</summary>
            public static readonly string DodgerBlue = "dodgerblue";
            /// <summary>#B22222</summary>
            public static readonly string FireBrick = "firebrick";
            /// <summary>#FFFAF0</summary>
            public static readonly string FloralWhite = "floralwhite";
            /// <summary>#228B22</summary>
            public static readonly string ForestGreen = "forestgreen";
            /// <summary>#FF00FF</summary>
            public static readonly string Fuchsia = "fuchsia";
            /// <summary>#DCDCDC</summary>
            public static readonly string Gainsboro = "gainsboro";
            /// <summary>#F8F8FF</summary>
            public static readonly string GhostWhite = "ghostwhite";
            /// <summary>#FFD700</summary>
            public static readonly string Gold = "gold";
            /// <summary>#DAA520</summary>
            public static readonly string GoldenRod = "goldenrod";
            /// <summary>#808080</summary>
            public static readonly string Gray = "gray";
            /// <summary>#008000</summary>
            public static readonly string Green = "green";
            /// <summary>#ADFF2F</summary>
            public static readonly string GreenYellow = "greenyellow";
            /// <summary>#F0FFF0</summary>
            public static readonly string HoneyDew = "honeydew";
            /// <summary>#FF69B4</summary>
            public static readonly string HotPink = "hotpink";
            /// <summary>#CD5C5C</summary>
            public static readonly string IndianRed = "indianred";
            /// <summary>#4B0082</summary>
            public static readonly string Indigo = "indigo";
            /// <summary>#FFFFF0</summary>
            public static readonly string Ivory = "ivory";
            /// <summary>#F0E68C</summary>
            public static readonly string Khaki = "khaki";
            /// <summary>#E6E6FA</summary>
            public static readonly string Lavender = "lavender";
            /// <summary>#FFF0F5</summary>
            public static readonly string LavenderBlush = "lavenderblush";
            /// <summary>#7CFC00</summary>
            public static readonly string LawnGreen = "lawngreen";
            /// <summary>#FFFACD</summary>
            public static readonly string LemonChiffon = "lemonchiffon";
            /// <summary>#ADD8E6</summary>
            public static readonly string LightBlue = "lightblue";
            /// <summary>#F08080</summary>
            public static readonly string LightCoral = "lightcoral";
            /// <summary>#E0FFFF</summary>
            public static readonly string LightCyan = "lightcyan";
            /// <summary>#FAFAD2</summary>
            public static readonly string LightGoldenRodYellow = "lightgoldenrodyellow";
            /// <summary>#D3D3D3</summary>
            public static readonly string LightGray = "lightgray";
            /// <summary>#90EE90</summary>
            public static readonly string LightGreen = "lightgreen";
            /// <summary>#FFB6C1</summary>
            public static readonly string LightPink = "lightpink";
            /// <summary>#FFA07A</summary>
            public static readonly string LightSalmon = "lightsalmon";
            /// <summary>#20B2AA</summary>
            public static readonly string LightSeaGreen = "lightseagreen";
            /// <summary>#87CEFA</summary>
            public static readonly string LightSkyBlue = "lightskyblue";
            /// <summary>#778899</summary>
            public static readonly string LightSlateGray = "lightslategray";
            /// <summary>#B0C4DE</summary>
            public static readonly string LightSteelBlue = "lightsteelblue";
            /// <summary>#FFFFE0</summary>
            public static readonly string LightYellow = "lightyellow";
            /// <summary>#00FF00</summary>
            public static readonly string Lime = "lime";
            /// <summary>#32CD32</summary>
            public static readonly string LimeGreen = "limegreen";
            /// <summary>#FAF0E6</summary>
            public static readonly string Linen = "linen";
            /// <summary>#FF00FF</summary>
            public static readonly string Magenta = "magenta";
            /// <summary>#800000</summary>
            public static readonly string Maroon = "maroon";
            /// <summary>#66CDAA</summary>
            public static readonly string MediumAquaMarine = "mediumaquamarine";
            /// <summary>#0000CD</summary>
            public static readonly string MediumBlue = "mediumblue";
            /// <summary>#BA55D3</summary>
            public static readonly string MediumOrchid = "mediumorchid";
            /// <summary>#9370DB</summary>
            public static readonly string MediumPurple = "mediumpurple";
            /// <summary>#3CB371</summary>
            public static readonly string MediumSeaGreen = "mediumseagreen";
            /// <summary>#7B68EE</summary>
            public static readonly string MediumSlateBlue = "mediumslateblue";
            /// <summary>#00FA9A</summary>
            public static readonly string MediumSpringGreen = "mediumspringgreen";
            /// <summary>#48D1CC</summary>
            public static readonly string MediumTurquoise = "mediumturquoise";
            /// <summary>#C71585</summary>
            public static readonly string MediumVioletRed = "mediumvioletred";
            /// <summary>#191970</summary>
            public static readonly string MidnightBlue = "midnightblue";
            /// <summary>#F5FFFA</summary>
            public static readonly string MintCream = "mintcream";
            /// <summary>#FFE4E1</summary>
            public static readonly string MistyRose = "mistyrose";
            /// <summary>#FFE4B5</summary>
            public static readonly string Moccasin = "moccasin";
            /// <summary>#FFDEAD</summary>
            public static readonly string NavajoWhite = "navajowhite";
            /// <summary>#000080</summary>
            public static readonly string Navy = "navy";
            /// <summary>#FDF5E6</summary>
            public static readonly string OldLace = "oldlace";
            /// <summary>#808000</summary>
            public static readonly string Olive = "olive";
            /// <summary>#6B8E23</summary>
            public static readonly string OliveDrab = "olivedrab";
            /// <summary>#FFA500</summary>
            public static readonly string Orange = "orange";
            /// <summary>#FF4500</summary>
            public static readonly string OrangeRed = "orangered";
            /// <summary>#DA70D6</summary>
            public static readonly string Orchid = "orchid";
            /// <summary>#EEE8AA</summary>
            public static readonly string PaleGoldenRod = "palegoldenrod";
            /// <summary>#98FB98</summary>
            public static readonly string PaleGreen = "palegreen";
            /// <summary>#AFEEEE</summary>
            public static readonly string PaleTurquoise = "paleturquoise";
            /// <summary>#DB7093</summary>
            public static readonly string PaleVioletRed = "palevioletred";
            /// <summary>#FFEFD5</summary>
            public static readonly string PapayaWhip = "papayawhip";
            /// <summary>#FFDAB9</summary>
            public static readonly string PeachPuff = "peachpuff";
            /// <summary>#CD853F</summary>
            public static readonly string Peru = "peru";
            /// <summary>#FFC0CB</summary>
            public static readonly string Pink = "pink";
            /// <summary>#DDA0DD</summary>
            public static readonly string Plum = "plum";
            /// <summary>#B0E0E6</summary>
            public static readonly string PowderBlue = "powderblue";
            /// <summary>#800080</summary>
            public static readonly string Purple = "purple";
            /// <summary>#FF0000</summary>
            public static readonly string Red = "red";
            /// <summary>#BC8F8F</summary>
            public static readonly string RosyBrown = "rosybrown";
            /// <summary>#4169E1</summary>
            public static readonly string RoyalBlue = "royalblue";
            /// <summary>#8B4513</summary>
            public static readonly string SaddleBrown = "saddlebrown";
            /// <summary>#FA8072</summary>
            public static readonly string Salmon = "salmon";
            /// <summary>#F4A460</summary>
            public static readonly string SandyBrown = "sandybrown";
            /// <summary>#2E8B57</summary>
            public static readonly string SeaGreen = "seagreen";
            /// <summary>#FFF5EE</summary>
            public static readonly string SeaShell = "seashell";
            /// <summary>#A0522D</summary>
            public static readonly string Sienna = "sienna";
            /// <summary>#C0C0C0</summary>
            public static readonly string Silver = "silver";
            /// <summary>#87CEEB</summary>
            public static readonly string SkyBlue = "skyblue";
            /// <summary>#6A5ACD</summary>
            public static readonly string SlateBlue = "slateblue";
            /// <summary>#708090</summary>
            public static readonly string SlateGray = "slategray";
            /// <summary>#FFFAFA</summary>
            public static readonly string Snow = "snow";
            /// <summary>#00FF7F</summary>
            public static readonly string SpringGreen = "springgreen";
            /// <summary>#4682B4</summary>
            public static readonly string SteelBlue = "steelblue";
            /// <summary>#D2B48C</summary>
            public static readonly string Tan = "tan";
            /// <summary>#008080</summary>
            public static readonly string Teal = "teal";
            /// <summary>#D8BFD8</summary>
            public static readonly string Thistle = "thistle";
            /// <summary>#FF6347</summary>
            public static readonly string Tomato = "tomato";
            /// <summary>#40E0D0</summary>
            public static readonly string Turquoise = "turquoise";
            /// <summary>#EE82EE</summary>
            public static readonly string Violet = "violet";
            /// <summary>#F5DEB3</summary>
            public static readonly string Wheat = "wheat";
            /// <summary>#FFFFFF</summary>
            public static readonly string White = "white";
            /// <summary>#F5F5F5</summary>
            public static readonly string WhiteSmoke = "whitesmoke";
            /// <summary>#FFFF00</summary>
            public static readonly string Yellow = "yellow";
            /// <summary>#9ACD32</summary>
            public static readonly string YellowGreen = "yellowgreen";

        }
    }
}
