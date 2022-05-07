using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using IOPath = System.IO.Path;

namespace TestDemo
{
    class Program
    {
        static void Main()
        {
            var document = PresentationDocument.Open("漏斗图.pptx", false);

            var slideParts = document.PresentationPart?.SlideParts;

            var currentSlidePart = slideParts?.ToArray()[0];

            var slideCommonSlideData = currentSlidePart?.Slide.CommonSlideData;

            var shapeTree = slideCommonSlideData?.ShapeTree;

            var graphicFrame = shapeTree?.GetFirstChild<AlternateContent>();

            var graphicData = graphicFrame?.Descendants<GraphicData>()?.FirstOrDefault();

            var chartRef = graphicData?.GetFirstChild<OpenXmlUnknownElement>();
            var id = chartRef.ExtendedAttributes.FirstOrDefault().Value;
            var part = (ExtendedChartPart)currentSlidePart.GetPartById(id);
            var series = part.ChartSpace.Chart.PlotArea.PlotAreaRegion.GetFirstChild<Series>();
            var layoutIdValue = series.LayoutId.Value;
        }
    }
}
