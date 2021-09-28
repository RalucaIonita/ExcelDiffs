using System.IO;
using System.Text.Json.Serialization;

namespace RootLogic.Models
{
    public class ExcelByteArrayResponse
    {
        public string ByteArray { get; set; }
        public string MimeType { get; set; }
        [JsonIgnore] public MemoryStream Stream { get; set; }
    }
}