using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Internal;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RootLogic.Helpers
{
    public static class FileHelper
    {
        public static MemoryStream BuildStream(string path)
        {
            
            var content = File.ReadAllBytes(path);
            var memoryStream = new MemoryStream(content);
            return memoryStream;
        }

        public static void WriteToFile(this List<string> list, string path)
        {
            File.WriteAllLines(path, list);
        }



    }
}
