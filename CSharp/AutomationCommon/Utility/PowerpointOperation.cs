using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.IO;

namespace AutomationCommon.Utility
{
    public class PowerpointOperation
    {
        public void GenPowerPoint(Dictionary<int, List<string>> powerpointImageDic, int imageCountPerSlide, string outputFile)
        {
            const int gap = 20;
            Application app = new Application();
            Presentation presentation = app.Presentations.Add(MsoTriState.msoFalse);
            Slides slides = presentation.Slides;
            slides.AddSlide(1, presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle]);
            foreach (var images in powerpointImageDic)
            {
                int remainslide = images.Value.Count % imageCountPerSlide;
                int slideLoop = remainslide == 0 ? images.Value.Count / imageCountPerSlide : images.Value.Count / imageCountPerSlide + 1;
                int cnt = 0;
                for (int i = 0; i < slideLoop; i++)
                {
                    if (imageCountPerSlide == 1)
                    {
                        CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[7];
                        Slide slide = slides.AddSlide(slides.Count + 1, customLayout);
                        int height = 350;
                        int top = 140;
                        slide.Shapes.AddPicture(images.Value[cnt], MsoTriState.msoFalse, MsoTriState.msoTrue, 10, top, 700, height);
                        var testbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, top + height, 700, height);
                        testbox.TextFrame.TextRange.Text = images.Value[cnt];
                        testbox.TextFrame.TextRange.Font.Size = 8;
                        testbox.TextFrame.TextRange.Font.Color.RGB = Rgb(255, 255, 255);
                        testbox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                    }
                    else
                    {
                        CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[7];
                        Slide slide = slides.AddSlide(slides.Count + 1, customLayout);
                        for (int j = 0; j < imageCountPerSlide; j++)
                        {
                            int remain = cnt % imageCountPerSlide;
                            int height = (500 - (imageCountPerSlide - 1) * gap) / imageCountPerSlide;
                            int top = 10 + remain * (height + gap);
                            if (cnt < images.Value.Count)
                            {
                                slide.Shapes.AddPicture(images.Value[cnt], MsoTriState.msoFalse, MsoTriState.msoTrue, 10, top, 700, height);
                                {
                                    var testbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, top + height, 700, height);
                                    testbox.TextFrame.TextRange.Text = images.Value[cnt];
                                    testbox.TextFrame.TextRange.Font.Size = 8;
                                    testbox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                                    testbox.TextFrame.TextRange.Font.Color.RGB = Rgb(255, 255, 255);
                                }
                            }
                            cnt++;
                        }
                    }
                }
            }
            var templateFile = Path.Combine(Directory.GetCurrentDirectory(), @"Config\Template.potx");
            presentation.Slides[1].ApplyTheme(templateFile);
            presentation.SaveAs(Path.ChangeExtension(outputFile, ".pptx"));
            presentation.Close();
            app.Quit();
        }

        private int Rgb(int r, int g, int b)
        {
            return r * 65536 + g * 256 + b;
        }
    }
}
