using GdPicture;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;
using ZXing;
using System.Drawing;
using ZXing.Common;
using System.Threading;

namespace TecnoDimOcr.OcrGdpicture
{
    public class Ocr
    {
        public Ocr()
        {
        }

        public static string castTopdf(string file, GdPictureImaging oGdPictureImaging, GdPicturePDF oGdPicturePDF)
        {
            string str = "";
            oGdPictureImaging.TiffOpenMultiPageForWrite(false);
            int num = oGdPictureImaging.CreateGdPictureImageFromFile(file);
            if (num != 0)
            {
                oGdPicturePDF.NewPDF();
                if (oGdPictureImaging.TiffIsMultiPage(num))
                {
                    int num1 = oGdPictureImaging.TiffGetPageCount(num);
                    bool flag = true;
                    int num2 = 1;
                    while (num2 <= num1)
                    {
                        oGdPictureImaging.TiffSelectPage(num, num2);
                        oGdPicturePDF.AddImageFromGdPictureImage(num, false, true);
                        if (oGdPicturePDF.GetStat() == 0)
                        {
                            num2++;
                        }
                        else
                        {
                            flag = false;
                            break;
                        }
                    }
                    if (flag)
                    {
                        str = file.Replace(Path.GetExtension(file), ".pdf");
                        oGdPicturePDF.SaveToFile(file.Replace(Path.GetExtension(file), ".pdf"));
                        if (oGdPicturePDF.GetStat() != 0)
                        {
                        }
                    }
                    oGdPicturePDF.CloseDocument();
                    oGdPictureImaging.ReleaseGdPictureImage(num);
                }
                else
                {
                    oGdPicturePDF.AddImageFromGdPictureImage(num, false, true);
                    if (oGdPicturePDF.GetStat() == 0)
                    {
                        str = file.Replace(Path.GetExtension(file), ".pdf");
                        if (oGdPicturePDF.SaveToFile(file.Replace(Path.GetExtension(file), ".pdf")) != 0)
                        {
                        }
                    }
                    oGdPicturePDF.CloseDocument();
                    oGdPictureImaging.ReleaseGdPictureImage(num);
                }
            }
            File.Delete(file);
            return str;
        }

        private System.Object lockThis = new System.Object();
        public void splitTiff(GdPictureImaging _gdPictureImaging, string file)
        {
            lock (lockThis)
            {

                string diretoriodeEnvio = ConfigurationManager.AppSettings["pastaSAIDA"];
                string diretorioMerge = ConfigurationManager.AppSettings["pastaMERGE"];
                string diretorioerro = ConfigurationManager.AppSettings["pastaERRO"];
                string diretoriobackup = ConfigurationManager.AppSettings["pastaBACKUP"];
                string pastazuada = ConfigurationManager.AppSettings["pastazuada"];
                string pastaJoin = ConfigurationManager.AppSettings["pastaJoin"];

                String identificacaoArquivo = Guid.NewGuid().ToString();
                string merge = diretorioMerge + @"\" + identificacaoArquivo;
                string pastajoinArq = pastaJoin + @"\" + identificacaoArquivo;
                if (!Directory.Exists(merge))
                {
                    Directory.CreateDirectory(merge);
                }

                if (!Directory.Exists(pastajoinArq))
                {
                    Directory.CreateDirectory(pastajoinArq);
                }
                int TiffImageID = _gdPictureImaging.TiffCreateMultiPageFromFile(file);

                int ImageCount = _gdPictureImaging.TiffGetPageCount(TiffImageID);
                string DocmentName = "";
                int ContadorPaginas = 0;
                String CodUnimed = "", NumBeneficiario, CodPrestador, DataAtendimento, NumNotaGuia;
                List<string> lista = new List<string>();
                try
                {

                    for (int i = 1; i <= ImageCount; i++)
                    {

                        var tif = _gdPictureImaging.TiffSelectPage(TiffImageID, i);
                        if (tif == 0)
                        {
                            List<String> barcode = new List<string>();
                            ContadorPaginas++;



                            if (_gdPictureImaging.BarcodeQRReaderDoScan(TiffImageID) == GdPicture.GdPictureStatus.OK)
                            {


                                var str = pastajoinArq + @"\" + Guid.NewGuid() + ".tif";
                                _gdPictureImaging.SaveAsTIFF(TiffImageID, str, TiffCompression.TiffCompressionCCITT4);
                                // load a bitmap

                                var readers = new ZXing.QrCode.QRCodeReader();
                                var barcodeBitmsap = (Bitmap)Bitmap.FromFile(str);
                                LuminanceSource sources = new BitmapLuminanceSource(barcodeBitmsap);
                                BinaryBitmap binBitmap = new BinaryBitmap(new GlobalHistogramBinarizer(sources));
                                var resulst = readers.decode(binBitmap);
                                barcodeBitmsap.Dispose();


                                if (resulst != null)
                                {
                                    barcode.Add(resulst.Text);
                                }

                                if (barcode.Count == 0)
                                {
                                    var result = _gdPictureImaging.BarcodeQRReaderGetBarcodeCount();
                                    for (int bar = 1; bar <= result; bar++)
                                    {
                                        barcode.Add(_gdPictureImaging.BarcodeQRReaderGetBarcodeValue(bar));
                                    }
                                }








                                if (barcode.Count > 0)
                                {


                                    foreach (var item in barcode)
                                    {
                                        if (item.Length == 54)
                                        {
                                            if (DocmentName == "")
                                            {
                                                CodUnimed = item.Substring(0, 4);
                                                NumBeneficiario = item.Substring(4, 13);
                                                CodPrestador = item.Substring(17, 9).TrimStart('0');
                                                DataAtendimento = item.Substring(26, 8);
                                                NumNotaGuia = item.Substring(34, 20).TrimStart('0');
                                                DocmentName = "Producao Medica_" + CodUnimed + "_" + NumBeneficiario + "_" + CodPrestador + "_" + DataAtendimento + "_" + NumNotaGuia + "_GuiasAtendimento.tif";
                                            }

                                            else
                                            {
                                                if (lista.Count > 0)
                                                {


                                                    ///join dos arquivos 
                                                    _gdPictureImaging.TiffMergeFiles(lista.ToArray(), merge + @"\" + DocmentName, TiffCompression.TiffCompressionCCITT4);
                                                    while (true)
                                                    {
                                                        if (!IsFileLocked(merge + @"\" + DocmentName))
                                                        {

                                                            break;
                                                        }
                                                    }
                                                    if (!File.Exists(diretoriodeEnvio + @"\" + DocmentName))
                                                        File.Move(merge + @"\" + DocmentName, diretoriodeEnvio + @"\" + DocmentName);
                                                    lista = new List<string>();


                                                    CodUnimed = item.Substring(0, 4);
                                                    NumBeneficiario = item.Substring(4, 13);
                                                    CodPrestador = item.Substring(17, 9).TrimStart('0');
                                                    DataAtendimento = item.Substring(26, 8);
                                                    NumNotaGuia = item.Substring(34, 20).TrimStart('0');
                                                    DocmentName = "Producao Medica_" + CodUnimed + "_" + NumBeneficiario + "_" + CodPrestador + "_" + DataAtendimento + "_" + NumNotaGuia + "_GuiasAtendimento.tif";// "Producao Medica_" + CodUnimed + "_" + NumBeneficiario + "_" + CodPrestador + "_" + DataAtendimento + "_" + NumNotaGuia + "_GuiasAtendimento.tif";

                                                }
                                            }



                                        }
                                    }
                                    var intDestDocID = _gdPictureImaging.CreateClonedGdPictureImageI(TiffImageID);
                                    var guid = Guid.NewGuid();
                                    var nome = merge + @"\" + guid + ".tif";
                                    _gdPictureImaging.SaveAsTIFF(intDestDocID, nome, TiffCompression.TiffCompressionCCITT4);
                                    lista.Add(nome);



                                }
                                else
                                {





                                    var intDestDocID = _gdPictureImaging.CreateClonedGdPictureImageI(TiffImageID);
                                    var guid = Guid.NewGuid();
                                    var nome = merge + @"\" + guid + ".tif";
                                    _gdPictureImaging.SaveAsTIFF(intDestDocID, nome, TiffCompression.TiffCompressionCCITT4);
                                    lista.Add(nome);

                                }
                            }


                            if (ContadorPaginas == 1 && (barcode.Where(_s => _s.Length > 54).Count() > 0 || barcode.Where(_s => _s.Length > 54).Count() < 0))
                            {
                                var intDestDocID = _gdPictureImaging.CreateClonedGdPictureImageI(TiffImageID);
                                var guid = Guid.NewGuid();
                                var nome = pastazuada + @"\" + guid + ".tif";
                                _gdPictureImaging.SaveAsTIFF(intDestDocID, nome, TiffCompression.TiffCompressionCCITT4);
                                lista.Add(nome);
                            }

                            if (ContadorPaginas == ImageCount)
                            {
                                if (DocmentName == "")
                                {
                                    _gdPictureImaging.ReleaseGdPictureImage(TiffImageID);
                                    while (true)
                                    {
                                        if (!IsFileLocked(file))
                                        {

                                            break;
                                        }
                                    }
                                    if (!File.Exists(pastazuada + @"\" + Path.GetFileName(file)))
                                        File.Move(file, pastazuada + @"\" + Path.GetFileName(file));
                                }
                                else
                                {
                                    var d = _gdPictureImaging.TiffMergeFiles(lista.ToArray(), diretorioMerge + @"\" + DocmentName, TiffCompression.TiffCompressionCCITT4);

                                    while (true)
                                    {
                                        if (!IsFileLocked(diretorioMerge + @"\" + DocmentName))
                                        {
                                            //if (File.Exists(diretoriodeEnvio + @"\" + DocmentName))
                                            //{
                                            //    File.Delete(diretoriodeEnvio + @"\" + DocmentName);
                                            //}
                                            if (!File.Exists(diretoriodeEnvio + @"\" + DocmentName))
                                            {
                                                File.Move(diretorioMerge + @"\" + DocmentName, diretoriodeEnvio + @"\" + DocmentName);
                                            }
                                            else
                                            {
                                                int contador = 0;
                                                var nomeArq = diretorioMerge + @"\" + DocmentName.Replace(".tif", "") + "(" + contador + ")" + ".tif";
                                                while (File.Exists(nomeArq))
                                                {
                                                    contador++;
                                                    nomeArq = diretorioMerge + @"\" + DocmentName.Replace(".tif", "") + "(" + contador + ")" + ".tif";

                                                }

                                                File.Move(diretorioMerge + @"\" + DocmentName, diretoriodeEnvio + @"\" + nomeArq);

                                            }

                                            break;
                                        }
                                    }
                                }
                                lista = new List<string>();

                            }
                        }



                    }


                }

                catch (Exception ex)
                {
                    using (FileStream fs = File.Create("c:\\errrrrrrr.txt"))
                    {
                        Byte[] info = new UTF8Encoding(true).GetBytes(ex.Message);
                        // Add some information to the file.
                        fs.Write(info, 0, info.Length);
                    }
                    throw new Exception(ex.Message);
                }
                finally
                {
                    _gdPictureImaging.ReleaseGdPictureImage(TiffImageID);

                    while (true)
                    {
                        if (!IsFileLocked(file))
                        {
                            File.Delete(file);
                            break;
                        }
                    }
                    deleteMerge(merge);
                    deleteMerge(pastajoinArq);

                }

            }
        }

        public void deleteMerge(String pasta)
        {

            foreach (var item in Directory.GetFiles(pasta))
            {
                if (!IsFileLocked(item))
                {
                    File.Delete(item);

                }
            }
            Directory.Delete(pasta);

        }



        public static string GerarDocumentoPesquisavelPdf(GdPictureImaging _gdPictureImaging, GdPicturePDF _gdPicturePDF, string documento, bool pdfa = true, string idioma = "por", string titulo = null, string autor = null, string assunto = null, string palavrasChaves = null, string criador = null, int dpi = 250)
        {
            if (Path.GetExtension(documento) != ".pdf")
            {
                if (Path.GetExtension(documento).ToLower() == ".tiff")
                {
                    File.Move(documento, Path.GetDirectoryName(documento) + "\\" + Path.GetFileNameWithoutExtension(documento) + ".tif");
                    documento = Path.GetDirectoryName(documento) + "\\" + Path.GetFileNameWithoutExtension(documento) + ".tif";
                }
                documento = Ocr.castTopdf(documento, _gdPictureImaging, _gdPicturePDF);
                var d = _gdPicturePDF.LoadFromFile(documento, false);
            }
            else
            {
                var d = _gdPicturePDF.LoadFromFile(documento, false);
            }
            int num = 0;
            var pasta = Guid.NewGuid().ToString();

            string str = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");


            using (FileStream fs = File.Create("c:\\lodg.txt"))
            {
                Byte[] info = new UTF8Encoding(true).GetBytes(str);
                // Add some information to the file.
                fs.Write(info, 0, info.Length);
            }
            //  Console.WriteLine(ex.Message);

            string str1 = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
            string str2 = string.Concat(str1, "\\", Path.GetFileName(documento));
            string folder = Guid.NewGuid().ToString();
            int pageCount = _gdPicturePDF.GetPageCount();
            for (int i = 1; i <= pageCount; i++)
            {
                Directory.CreateDirectory(str1 + "\\" + pasta);
                _gdPicturePDF.SelectPage(i);
                int gdPictureImageEx = _gdPicturePDF.RenderPageToGdPictureImageEx((float)dpi, true);
                if (gdPictureImageEx != 0)
                {
                    num = _gdPictureImaging.PdfOCRStart(str1 + "\\" + pasta + "\\" + i.ToString() + ".pdf", pdfa, titulo, autor, assunto, palavrasChaves, criador);
                    _gdPictureImaging.PdfAddGdPictureImageToPdfOCR(num, gdPictureImageEx, idioma, str, "");
                    _gdPictureImaging.ReleaseGdPictureImage(gdPictureImageEx);
                    _gdPictureImaging.PdfOCRStop(num);
                }
            }

            _gdPicturePDF.CloseDocument();
            File.Delete(documento);

            GdPictureStatus status = _gdPicturePDF.MergeDocuments(Directory.GetFiles(str1 + "\\" + pasta), str2);


            DirectoryInfo dir = new DirectoryInfo(str1 + "\\" + pasta);

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            Directory.Delete(str1 + "\\" + pasta);
            return str2;
        }

        private static string GetCurrentDirectory()
        {
            string absolutePath = (new Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath;
            string fullName = (new DirectoryInfo(Path.GetDirectoryName(absolutePath))).FullName;
            return Uri.UnescapeDataString(fullName);
        }

        public static bool IsFileLocked(string filePath)
        {
            bool flag;
            try
            {
                using (FileStream fileStream = File.Open(filePath, FileMode.Open))
                {
                }
            }
            catch (IOException oException)
            {
                int hRForException = Marshal.GetHRForException(oException) & 65535;
                flag = (hRForException == 32 ? true : hRForException == 33);
                return flag;
            }
            flag = false;
            return flag;
        }
    }
}