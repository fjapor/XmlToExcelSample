using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLToExcellConverter.Messages
{
    public class FileNameOperationResponse
    {
        public FileNameOperationResponse(bool success,string fileName = "")
        {
            Success = success;
            FileName = fileName;
        }

        public bool Success { get; private set; }
        public string FileName { get; private set; }
    }
}
