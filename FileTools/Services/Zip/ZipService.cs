using SharpCompress.Archives;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;
using SharpCompress.Readers;

namespace FileTools.Services.Zip
{
    public class ZipService
    {
        public void ExtractHere(string inputFullpath, bool overwriteFiles)
        {
            try
            {
                if (!File.Exists(inputFullpath))
                    throw new FileNotFoundException($"El archivo '{inputFullpath}' no se encontró en la ubicación especificada.");

                using (ZipArchive zipArchive = ZipArchive.Open(inputFullpath))
                {
                    string inputPath = Path.GetDirectoryName(inputFullpath);
                    List<ZipArchiveEntry> files = zipArchive.Entries.Where(entry => !entry.IsDirectory).ToList();

                    foreach (ZipArchiveEntry file in files)
                    {
                        string outputFullpath = Path.Combine(inputPath, file.Key);

                        ExtractionOptions options = new ExtractionOptions() { ExtractFullPath = true, Overwrite = overwriteFiles };

                        if (overwriteFiles)
                            file.WriteToDirectory(inputPath, options);
                        else
                            if (!File.Exists(outputFullpath))
                            file.WriteToDirectory(inputPath, options);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ExtractHere(string inputFullpath, bool overwriteFiles, string password)
        {
            try
            {
                if (!File.Exists(inputFullpath))
                    throw new FileNotFoundException($"El archivo '{inputFullpath}' no se encontró en la ubicación especificada.");

                using (ZipArchive zipArchive = ZipArchive.Open(inputFullpath, new ReaderOptions { Password = password }))
                {
                    string inputPath = Path.GetDirectoryName(inputFullpath);
                    List<ZipArchiveEntry> files = zipArchive.Entries.Where(entry => !entry.IsDirectory).ToList();

                    foreach (ZipArchiveEntry file in files)
                    {
                        string outputFullpath = Path.Combine(inputPath, file.Key);

                        ExtractionOptions options = new ExtractionOptions() { ExtractFullPath = true, Overwrite = overwriteFiles };

                        if (overwriteFiles)
                            file.WriteToDirectory(inputPath, options);
                        else
                            if (!File.Exists(outputFullpath))
                            file.WriteToDirectory(inputPath, options);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ExtractIn(string inputFullpath, bool overwriteFiles)
        {
            try
            {
                if (!File.Exists(inputFullpath))
                    throw new FileNotFoundException($"El archivo '{inputFullpath}' no se encontró en la ubicación especificada.");

                using (ZipArchive zipArchive = ZipArchive.Open(inputFullpath))
                {
                    List<ZipArchiveEntry> files = zipArchive.Entries.Where(entry => !entry.IsDirectory).ToList();
                    string outputPath = Path.Combine(Path.GetDirectoryName(inputFullpath), Path.GetFileNameWithoutExtension(inputFullpath));

                    if (!Directory.Exists(outputPath))
                        Directory.CreateDirectory(outputPath);

                    foreach (ZipArchiveEntry file in files)
                    {
                        string outputFullpath = Path.Combine(outputPath, file.Key);

                        ExtractionOptions options = new ExtractionOptions() { ExtractFullPath = true, Overwrite = overwriteFiles };

                        if (overwriteFiles)
                            file.WriteToDirectory(outputPath, options);
                        else
                            if (!File.Exists(outputFullpath))
                            file.WriteToDirectory(outputPath, options);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ExtractIn(string inputFullpath, bool overwriteFiles, string password)
        {
            try
            {
                if (!File.Exists(inputFullpath))
                    throw new FileNotFoundException($"El archivo '{inputFullpath}' no se encontró en la ubicación especificada.");

                using (ZipArchive zipArchive = ZipArchive.Open(inputFullpath, new ReaderOptions { Password = password }))
                {
                    List<ZipArchiveEntry> files = zipArchive.Entries.Where(entry => !entry.IsDirectory).ToList();
                    string outputPath = Path.Combine(Path.GetDirectoryName(inputFullpath), Path.GetFileNameWithoutExtension(inputFullpath));

                    if (!Directory.Exists(outputPath))
                        Directory.CreateDirectory(outputPath);

                    foreach (ZipArchiveEntry file in files)
                    {
                        string outputFullpath = Path.Combine(outputPath, file.Key);

                        ExtractionOptions options = new ExtractionOptions() { ExtractFullPath = true, Overwrite = overwriteFiles };

                        if (overwriteFiles)
                            file.WriteToDirectory(outputPath, options);
                        else
                            if (!File.Exists(outputFullpath))
                            file.WriteToDirectory(outputPath, options);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ExtractIn(string inputFullpath, string outputPath, bool overwriteFiles)
        {
            try
            {
                if (!File.Exists(inputFullpath))
                    throw new FileNotFoundException($"El archivo '{inputFullpath}' no se encontró en la ubicación especificada.");

                if (!Directory.Exists(outputPath))
                    throw new DirectoryNotFoundException($"El directorio '{outputPath}' no existe.");

                using (ZipArchive zipArchive = ZipArchive.Open(inputFullpath))
                {
                    List<ZipArchiveEntry> files = zipArchive.Entries.Where(entry => !entry.IsDirectory).ToList();

                    foreach (ZipArchiveEntry file in files)
                    {
                        string outputFullpath = Path.Combine(outputPath, file.Key);

                        ExtractionOptions options = new ExtractionOptions() { ExtractFullPath = true, Overwrite = overwriteFiles };

                        if (overwriteFiles)
                            file.WriteToDirectory(outputPath, options);
                        else
                            if (!File.Exists(outputFullpath))
                            file.WriteToDirectory(outputPath, options);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ExtractIn(string inputFullpath, string outputPath, bool overwriteFiles, string password)
        {
            try
            {
                if (!File.Exists(inputFullpath))
                    throw new FileNotFoundException($"El archivo '{inputFullpath}' no se encontró en la ubicación especificada.");

                if (!Directory.Exists(outputPath))
                    throw new DirectoryNotFoundException($"El directorio '{outputPath}' no existe.");

                using (ZipArchive zipArchive = ZipArchive.Open(inputFullpath, new ReaderOptions { Password = password }))
                {
                    List<ZipArchiveEntry> files = zipArchive.Entries.Where(entry => !entry.IsDirectory).ToList();
                    foreach (ZipArchiveEntry file in files)
                    {
                        string outputFullpath = Path.Combine(outputPath, file.Key);

                        ExtractionOptions options = new ExtractionOptions() { ExtractFullPath = true, Overwrite = overwriteFiles };

                        if (overwriteFiles)
                            file.WriteToDirectory(outputPath, options);
                        else
                            if (!File.Exists(outputFullpath))
                            file.WriteToDirectory(outputPath, options);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}