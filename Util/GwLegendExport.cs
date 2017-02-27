using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.esriSystem;
using System.Drawing;
using System.IO;
using ESRI.ArcGIS.Geometry;
using System.Drawing.Imaging;
using Newtonsoft.Json;

namespace MXDLegendExport
{
    class GwLegendExport
    {
        private static int resolution = 96;//分辨率
        public static int Resolution
        {
            get
            {
                return resolution;
            }
            set
            {
                resolution = value;
            }
        }
        public static bool GenerateJson(ILegend pLegend, string outputpath)
        {
            try
            {

                for (int i = 0; i < pLegend.ItemCount; i++)
                {
                    var item = pLegend.get_Item(i);
                    ILayer player = item.Layer;
                    string layername = player.Name;
                    List<Dictionary<string, string>> layerGroup = new List<Dictionary<string, string>>();
                    

                    if (string.IsNullOrEmpty(layername.Trim()))
                    {
                        layername = "UnKnown";
                    }
                    ILegendInfo pLegendInfo = player as ILegendInfo;
                    for (int j = 0; j < pLegendInfo.LegendGroupCount; j++)
                    {
                        var legendgroup = pLegendInfo.LegendGroup[j];
                        for (int k = 0; k < legendgroup.ClassCount; k++)
                        {
                            Dictionary<string, string> labelGroup = new Dictionary<string, string>();
                            var legendclass = legendgroup.Class[k];
                            var symbol = legendclass.Symbol;
                            IStyleGalleryItem pStyleGalleryItem = new ServerStyleGalleryItemClass();
                            IClone pClone = symbol as IClone;
                            pStyleGalleryItem.Item = pClone.Clone();
                            var bitmap = GetImageFromSymbol(symbol, 64, 64);
                            try
                            {
                                string strBitmap = ImgToBase64String(bitmap);
                                labelGroup.Add("label", legendclass.Label);
                                labelGroup.Add("image", strBitmap);
                                layerGroup.Add(labelGroup);
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                    string strAll = JsonConvert.SerializeObject(layerGroup);
                    string filename = GenerateFileName(outputpath + "\\" + layername + ".txt");
                    WriteToFile(filename, strAll);                  
                }
            }
            catch
            {
                return false;
            }

            return true;
        }
        
        public static bool GenerateLabelImage(ILegend pLegend, string outputpath)
        {
            try
            {
                for (int i = 0; i < pLegend.ItemCount; i++)
                {
                    var item = pLegend.get_Item(i);
                    ILayer player = item.Layer;
                    ILegendInfo pLegendInfo = player as ILegendInfo;
                    for (int j = 0; j < pLegendInfo.LegendGroupCount; j++)
                    {
                        var legendgroup = pLegendInfo.LegendGroup[j];
                        for (int k = 0; k < legendgroup.ClassCount; k++)
                        {
                            var legendclass = legendgroup.Class[k];
                            var symbol = legendclass.Symbol;
                            IStyleGalleryItem pStyleGalleryItem = new ServerStyleGalleryItemClass();
                            IClone pClone = symbol as IClone;
                            pStyleGalleryItem.Item = pClone.Clone();
                            var bitmap = GetImageFromSymbol(symbol, 64, 64);

                            string strLayerName = player.Name;
                            if (string.IsNullOrEmpty(strLayerName))
                            {
                                strLayerName = "UnKnown";
                            }
                            string strLabelname = legendclass.Label;
                            if (string.IsNullOrEmpty(strLabelname))
                            {
                                strLabelname = strLayerName;
                            }
                            string filename = GenerateFileName(outputpath + "\\" + strLayerName + "\\" + strLabelname + ".bmp");
                            try
                            {
                                bitmap.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        public static bool GenerateWholeImage(IMapSurround mapsurround, double width, double height, string filename)
        {
            //var querySize = (IQuerySize)mapsurround;
            double widthUnit = width;
            double heightUnit = height;
            //querySize.QuerySize(ref widthUnit, ref heightUnit);
            //var widthUnit = widthPoints;
            //var heightUnit = heightPoints;
            double widthPUnit = widthUnit;
            var heightPUnit = widthPUnit * (heightUnit / widthUnit);
            var pixelsWidth = (int)(Math.Ceiling(widthPUnit * resolution));
            var pixelsHeight = (int)(Math.Ceiling(heightPUnit *  resolution));
            try
            {
                using(var bitmap = new Bitmap(pixelsWidth, pixelsHeight))
                {
                    bitmap.SetResolution(resolution, resolution);
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.Clear(Color.White);

                        var simpleDisplay = new SimpleDisplayClass();
                        simpleDisplay.DisplayTransformation.Resolution = resolution;
                        simpleDisplay.DisplayTransformation.Units = esriUnits.esriInches;
                        //simpleDisplay.DisplayTransformation.ScaleRatio = widthUnit;
                        simpleDisplay.DisplayTransformation.Bounds = new EnvelopeClass
                        {
                            XMin = 0,
                            YMin = 0,
                            XMax = widthPUnit,//pixelsWidth,//
                            YMax = heightPUnit//pixelsHeight//
                        };
                        simpleDisplay.DisplayTransformation.VisibleBounds = simpleDisplay.DisplayTransformation.Bounds;
                        var tagRect = new tagRECT { left = 0, top = 0, right = pixelsWidth, bottom = pixelsHeight };
                        simpleDisplay.DisplayTransformation.set_DeviceFrame(ref tagRect);
                        simpleDisplay.StartDrawing(graphics.GetHdc().ToInt32(), 0);

                        var drawBounds = new EnvelopeClass
                        {
                            XMin = 0,
                            YMin = 0,
                            XMax = widthUnit,//pixelsWidth,//widthPUnit,//
                            YMax = heightUnit//pixelsHeight//
                        };
                        mapsurround.Draw(simpleDisplay, null, simpleDisplay.DisplayTransformation.Bounds);
                        simpleDisplay.FinishDrawing();
                    }
                    string nfilename = GenerateFileName(filename);
                    bitmap.Save(nfilename, ImageFormat.Bmp);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GenerateFileName(string filename)
        {
            if (string.IsNullOrEmpty(filename))
            {
               return GenerateFileName(System.Windows.Forms.Application.StartupPath + "\\" +  "UnKnown.txt");
            }
            string fNameOnly=System.IO.Path.GetFileNameWithoutExtension(filename);
            if (fNameOnly.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0)
            {
                int indexOfInvalid= filename.IndexOfAny(System.IO.Path.GetInvalidFileNameChars());
                string nfilename= filename.Replace(filename[indexOfInvalid], '_');
                return GenerateFileName(nfilename);
            }
            if (File.Exists(filename))
            {
                string nfilename = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filename), System.IO.Path.GetFileNameWithoutExtension(filename) + "_1" + System.IO.Path.GetExtension(filename));
                return GenerateFileName(nfilename);
            }
            else
            {
                string nFileName = filename;
                if (CheckOrCreateFolderFromFile(out nFileName, filename))
                {
                    return nFileName;
                }
                else
                {
                    return System.Windows.Forms.Application.StartupPath + "\\" + "UnKnown.txt";
                }
            }
        }

        public static Bitmap GetImageFromSymbol(ISymbol pSymbol, int width, int height)
        {
            IStyleGalleryClass styleGalleryClass = null;
            if (pSymbol is IMarkerSymbol)
            {
                styleGalleryClass = new MarkerSymbolStyleGalleryClass();
            }
            else if (pSymbol is ILineSymbol)
            {
                styleGalleryClass = new LineSymbolStyleGalleryClass();
            }
            else if (pSymbol is IFillSymbol)
            {
                styleGalleryClass = new FillSymbolStyleGalleryClassClass();
            }
            if (styleGalleryClass != null)
            {
                return GetImage(styleGalleryClass, pSymbol, width, height);
            }

            return null;
        }

        private static Bitmap GetImage(IStyleGalleryClass styleGalleryClass,
            ISymbol symbol, int width, int height)
        {
            Bitmap img = new Bitmap(width, height);
            Graphics gc = Graphics.FromImage(img);
            IntPtr hdc = gc.GetHdc();

            var rect = new tagRECT();
            rect.left = 0;
            rect.top = 0;
            rect.right = width;
            rect.bottom = height;
            styleGalleryClass.Preview(symbol, hdc.ToInt32(), ref rect);
            gc.ReleaseHdc(hdc);
            gc.Dispose();
            //ImgToBase64String(img);
            return img;
        }

        private static string ImgToBase64String(Bitmap bmp)
        {
            try
            {
                MemoryStream ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                byte[] arr = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(arr, 0, (int)ms.Length);
                ms.Close();
                String strbaser64 = Convert.ToBase64String(arr);
                return strbaser64;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("ImgToBase64String 转换失败\nException:" + ex.Message);
                return null;
            }
        }

        private static IGeometry GetSymbolGeometry(ISymbol symbol, IEnvelope envelop)
        {
            IGeometry geometry = null;

            if (symbol is IMarkerSymbol)
            {
                var area = (IArea)envelop;
                geometry = area.Centroid;
            }
            else if (symbol is ILineSymbol)
            {
                IPolyline polyline = new PolylineClass();
                var pointCollection = (IPointCollection)polyline;

                IPoint point = new PointClass();
                object before = Type.Missing;
                object after = Type.Missing;

                // 自定义一条具有三段的折线
                point.PutCoords(envelop.XMin, envelop.YMax);
                pointCollection.AddPoint(point, ref before, ref after);
                point.PutCoords((envelop.XMax - envelop.XMin) / 3, envelop.YMin);
                pointCollection.AddPoint(point, ref before, ref after);
                point.PutCoords((envelop.XMax - envelop.XMin) * 2 / 3, envelop.YMax);
                pointCollection.AddPoint(point, ref before, ref after);
                point.PutCoords((envelop.XMax - envelop.XMin), envelop.YMin);
                pointCollection.AddPoint(point, ref before, ref after);

                geometry = polyline;
            }
            else if (symbol is IFillSymbol)
            {
                geometry = envelop;
            }
            else if (symbol is ITextSymbol)
            {
                var area = (IArea)envelop;
                geometry = area.Centroid;
            }

            return geometry;
        }
        
        private static bool CheckOrCreateFolderFromFile(out string nFilePath, string inFilePath)
        {
            bool result = false;
            nFilePath = inFilePath;            
            if (inFilePath != null && inFilePath.Length > 0)
            {
                try
                {
                    string outFolder = System.IO.Path.GetDirectoryName(nFilePath);
                    if (!Directory.Exists(outFolder))
                    {
                        Directory.CreateDirectory(outFolder);
                    }
                    result = true;
                }
                catch
                {
                    result = false;
                }

            }
            return result;
        }
        
        private static bool WriteToFile(string filepath, string msg)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(filepath, true, Encoding.UTF8))
                {
                    sw.Write(msg);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }
    }
}
