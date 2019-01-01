using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography;

namespace SharePointOnline_MigrationTool.Logic
{
    class FileHash
    {
        public string filePath { get; set; }

        public FileHash(string filePath)
        {
            this.filePath = filePath;
        }

        public string CreateHash()
        {

            byte[] buffer;
            int byteRead;
            long size;
            long totalByteRead = 0;

            using (Stream file = File.OpenRead(filePath))
            {

                size = file.Length;

                using (HashAlgorithm hasher = MD5.Create())
                {
                    do
                    {
                        buffer = new byte[4096];

                        byteRead = file.Read(buffer, 0, buffer.Length);

                        totalByteRead += byteRead;

                        hasher.TransformBlock(buffer, 0, byteRead, null, 0);

                    }
                    while (byteRead != 0);

                    hasher.TransformFinalBlock(buffer, 0, 0);

                    return MakeHashString(hasher.Hash);

                }     
            }
        }

        private static string MakeHashString(byte[] hashBytes)
        {
            StringBuilder hash = new StringBuilder(32);

            foreach (byte b in hashBytes)
                hash.Append(b.ToString("X2").ToLower());

            return hash.ToString();
        }
    }
}
