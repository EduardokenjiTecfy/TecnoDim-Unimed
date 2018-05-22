using Aspose.Pdf;
using GdPicture;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using TecnoDimOcr.OcrGdpicture;

namespace TecnoDimOcr
{
    partial class ServiceOcrTecnodim : ServiceBase
    {
        private System.Timers.Timer timer;
        public ServiceOcrTecnodim()
        {
            Aspose.Words.License license = new Aspose.Words.License();
           
            license.SetLicense(new FileStream(@"C:\Users\leonardo.domeneghett\Desktop\TECNODIM\TecnoDimOcr\TecnoDimOcr\Aspose\Aspose.Words.lic", FileMode.OpenOrCreate));
            Aspose.Words.Document document = new Aspose.Words.Document(@"C:\Users\leonardo.domeneghett\Desktop\DOCS VERA CRUZ\BAVK\Tabela_Cabesp2016.xlsx");
            var asasa = document.GetText();
            File.AppendAllText(@"C:\Users\leonardo.domeneghett\Desktop\DOCS VERA CRUZ\teste1.txt", asasa);
            oGdPicturePDF.SetLicenseNumber("4118106456693265856441854");
            oGdPictureImaging.SetLicenseNumber("4118106456693265856441854");
            InitializeComponent();
            this.ServiceName = "ServiceOcrTecnodim";
        }
        private GdPictureImaging oGdPictureImaging = new GdPictureImaging();
        private GdPicturePDF oGdPicturePDF = new GdPicturePDF();


        public bool IsFileLocked(string filename)
        {


            bool Locked = false;
            try
            {
                FileStream fs =
                    File.Open(filename, FileMode.OpenOrCreate,
                    FileAccess.ReadWrite, FileShare.None);
                fs.Close();
            }
            catch (IOException ex)
            {
                Locked = true;
            }
            return Locked;
        }
        private static List<FileInfo> files = new List<FileInfo>();

       
         public   void sss()
        {
            string item = ConfigurationManager.AppSettings["pastaCAPTACAO"];
            try
            {
                string[] files = Directory.GetFiles(item);
                for (int i = 0; i < (int)files.Length; i++)
                {
                    string str = files[i];
                    if ((Path.GetExtension(str).Trim().ToLower() == ".tif" ? true : Path.GetExtension(str).Trim().ToLower() == ".tiff"))
                    {
                        while (true)
                        {
                            if (!this.IsFileLocked(str))
                            {
                                break;
                            }
                        }
                        string item1 = ConfigurationManager.AppSettings["pastaBACKUP"];
                        if (File.Exists(string.Concat(item1, "\\", Path.GetFileName(str))))
                        {
                            File.Delete(string.Concat(item1, "\\", Path.GetFileName(str)));
                        }
                        while (true)
                        {
                            if (!this.IsFileLocked(str))
                            {
                                break;
                            }
                        }
                        File.Copy(str, string.Concat(item1, "\\", Path.GetFileName(str)));
                        (new Ocr()).splitTiff(this.oGdPictureImaging, str);
                    }
                }
            }
            catch (Exception exception1)
            {
                Exception exception = exception1;
                using (FileStream fileStream = File.Create("logErro.txt"))
                {
                    byte[] bytes = (new UTF8Encoding(true)).GetBytes(exception.Message);
                    fileStream.Write(bytes, 0, (int)bytes.Length);
                }
            }
            FileSystemWatcher fileSystemWatcher = new FileSystemWatcher()
            {
                Path = item,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.LastAccess
            };
            fileSystemWatcher.Created += new FileSystemEventHandler(this.executar);
            fileSystemWatcher.IncludeSubdirectories = true;
            fileSystemWatcher.EnableRaisingEvents = true;
        }
        protected override void OnStart(string[] args)
        {
            string item = ConfigurationManager.AppSettings["pastaCAPTACAO"];
            try
            {
                string[] files = Directory.GetFiles(item);
                for (int i = 0; i < (int)files.Length; i++)
                {
                    string str = files[i];
                    if ((Path.GetExtension(str).Trim().ToLower() == ".tif" ? true : Path.GetExtension(str).Trim().ToLower() == ".tiff"))
                    {
                        while (true)
                        {
                            if (!this.IsFileLocked(str))
                            {
                                break;
                            }
                        }
                        string item1 = ConfigurationManager.AppSettings["pastaBACKUP"];
                        if (File.Exists(string.Concat(item1, "\\", Path.GetFileName(str))))
                        {
                            File.Delete(string.Concat(item1, "\\", Path.GetFileName(str)));
                        }
                        while (true)
                        {
                            if (!this.IsFileLocked(str))
                            {
                                break;
                            }
                        }
                        File.Copy(str, string.Concat(item1, "\\", Path.GetFileName(str)));
                        (new Ocr()).splitTiff(this.oGdPictureImaging, str);
                    }
                }
            }
            catch (Exception exception1)
            {
                Exception exception = exception1;
                using (FileStream fileStream = File.Create("logErro.txt"))
                {
                    byte[] bytes = (new UTF8Encoding(true)).GetBytes(exception.Message);
                    fileStream.Write(bytes, 0, (int)bytes.Length);
                }
            }
            FileSystemWatcher fileSystemWatcher = new FileSystemWatcher()
            {
                Path = item,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.LastAccess
            };
            fileSystemWatcher.Created += new FileSystemEventHandler(this.executar);
            fileSystemWatcher.IncludeSubdirectories = true;
            fileSystemWatcher.EnableRaisingEvents = true;
        }

        protected override void OnStop()
        {
            this.timer.Stop();
            this.timer = null;
        }


        public void executar(object sender, FileSystemEventArgs e)
        {
            string item = ConfigurationManager.AppSettings["pastaBACKUP"];
            if ((Path.GetExtension(e.FullPath).Trim().ToLower() == ".tif" ? true : Path.GetExtension(e.FullPath).Trim().ToLower() == ".tiff"))
            {
                while (true)
                {
                    if (!this.IsFileLocked(e.FullPath))
                    {
                        break;
                    }
                }
                if (File.Exists(string.Concat(item, "\\", Path.GetFileName(e.FullPath))))
                {
                    File.Delete(string.Concat(item, "\\", Path.GetFileName(e.FullPath)));
                }
                while (true)
                {
                    if (!this.IsFileLocked(e.FullPath))
                    {
                        break;
                    }
                }
                File.Copy(e.FullPath, string.Concat(item, "\\", Path.GetFileName(e.FullPath)));
                (new Ocr()).splitTiff(this.oGdPictureImaging, e.FullPath);
            }
        }



    }
}
