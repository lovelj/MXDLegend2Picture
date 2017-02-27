using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.esriSystem;
using System.IO;
using ESRI.ArcGIS.Geometry;
using System.Drawing.Imaging;

namespace MXDLegendExport
{
    public partial class NMainForm :Form
    {

        public NMainForm()
        {
            InitializeComponent();
        }

        

        private bool Generate(string mxdfile, LegendOutPutType outputtype, string outputpath)
        {
            IMapDocument pMapDocument;
            pMapDocument = new MapDocumentClass();
            pMapDocument.Open(mxdfile);
            IPageLayout pPageLayout = pMapDocument.PageLayout;

            bool hasLegend = false;

            IGraphicsContainer pGraphicsContainer = pPageLayout as IGraphicsContainer;
            pPageLayout.Page.Units = esriUnits.esriInches;
            pGraphicsContainer.Reset();
            IElement pElement = pGraphicsContainer.Next();
            while (pElement != null)
            {
                if (pElement is IMapSurroundFrame)
                {
                    hasLegend = true;
                    IMapSurroundFrame mapSurroundFrame = pElement as IMapSurroundFrame;
                    
                    IMapSurround mapSurround = mapSurroundFrame.MapSurround;
                    ILegend pLegend = mapSurround as ILegend;
                    switch (outputtype)
                    {
                        case LegendOutPutType.JsonType:
                            GwLegendExport.GenerateJson(pLegend, outputpath);
                            break;
                        case LegendOutPutType.LabelType:
                            GwLegendExport.GenerateLabelImage(pLegend, outputpath);
                            break;
                        case LegendOutPutType.ImageType:
                            GwLegendExport.GenerateWholeImage(mapSurround, pElement.Geometry.Envelope.Width, pElement.Geometry.Envelope.Height, outputpath);
                            break;
                        default:
                            return false;
                    }
                    string mapSurroundName = mapSurround.Name;
                }
                pElement = pGraphicsContainer.Next();
            }
            if (!hasLegend)
            {
                MessageBox.Show("未检测到图例元素！", "提示");
            }
            return true;
        }


        
        private void button1_Click(object sender, EventArgs e)
        {
            Generate(@"D:\workspace\TEST\无标题.mxd", LegendOutPutType.ImageType, @"D:\workspace\TEST\A.bmp");
            //
        }
    }
}