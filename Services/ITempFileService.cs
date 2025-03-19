using System;
using System.IO;

namespace DocumentConverterApplication.Services
{
    public interface ITempFileService
    {
        string GetTempDirectory();
        void EnsureDirectoryExists(string filePath);
        void CleanupTempFiles(TimeSpan olderThan);
    }
}