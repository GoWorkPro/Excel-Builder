using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using SixLabors.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml.Linq;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using DocumentFormat.OpenXml.Drawing.Charts;
using PdfSharp.Drawing.Layout;

namespace PDFBuilder
{
    public class PDFBuilder
    {

        public static void Build()
        {
            GlobalFontSettings.FontResolver = new FileFontResolver();
            // Create a new PDF document
            PdfDocument document = new PdfDocument();

            // Add a page to the document
            PdfPage page = document.AddPage();
 
            // Create a drawing object for the page
            XGraphics gfx = XGraphics.FromPdfPage(page);

        
            // Create a font
            XFont titleFont = new XFont("Verdana", 16, XFontStyleEx.Bold);
            XFont regularFont = new XFont("Verdana", 12);

            // Load user image
            XImage image = XImage.FromFile("user_image.png"); // Replace with the path to the user's image

            // Draw user image
            gfx.DrawImage(image, 20, 20, 80, 80);

            // Draw user information
            DrawText(gfx, "John Doe", titleFont, 120, 20);
            DrawText(gfx, "Software Engineer", regularFont, 120, 40);
            DrawText(gfx, "john.doe@example.com | (555) 123-4567", regularFont, 120, 60);

            // Draw education section
            DrawSectionHeader(gfx, "Education", 120);
            DrawParagraph(gfx, regularFont, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.", 20, 140, 400);

            // Draw awards section
            DrawSectionHeader(gfx, "Awards", 200);
            DrawParagraph(gfx, regularFont, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.", 20, 220, 400);

            // Save the document to a file
            document.Save("abc.pdf");

            // Close the document
            document.Close();
        }

        static void DrawSectionHeader(XGraphics gfx, string header, double y)
        {
            gfx.DrawString(header, new XFont("Verdana", 16, XFontStyleEx.Bold), XBrushes.Black, new XRect(20, y, 400, 20), XStringFormats.TopLeft);
        }

        static void DrawParagraph(XGraphics gfx, XFont font, string text, double x, double y, double maxWidth)
        {
            XTextFormatter formatter = new XTextFormatter(gfx);

            // Calculate the height needed for the text
            XRect layoutRectangle = new XRect(x, y, maxWidth, 500);
            XSize textSize = gfx.MeasureString(text, font, XStringFormats.TopLeft);
            textSize.Width = 400;
            // Adjust the layout rectangle based on the measured text size
            layoutRectangle = new XRect(x, y, maxWidth, textSize.Height);

            // Draw the text
            formatter.DrawString(text, font, XBrushes.Black, layoutRectangle, XStringFormats.TopLeft);
        }

        static void DrawText(XGraphics gfx, string text, XFont font, double x, double y)
        {
            gfx.DrawString(text, font, XBrushes.Black, new XRect(x, y, 400, 20), XStringFormats.TopLeft);
        }
    }

    public class FileFontResolver : IFontResolver // FontResolverBase
    {
        public string DefaultFontName => throw new NotImplementedException();

        public byte[] GetFont(string faceName)
        {
            using (var ms = new MemoryStream())
            {
                using (var fs = File.Open(faceName, FileMode.Open))
                {
                    fs.CopyTo(ms);
                    ms.Position = 0;
                    return ms.ToArray();
                }
            }
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (familyName.Equals("Verdana", StringComparison.CurrentCultureIgnoreCase))
            {
                //if (isBold && isItalic)
                //{
                //    return new FontResolverInfo("Fonts/Verdana-BoldItalic.ttf");
                //}
                //else if (isBold)
                //{
                    return new FontResolverInfo("Fonts/verdanab.ttf");
                //}
                //else if (isItalic)
                //{
                //    return new FontResolverInfo("Fonts/Verdana-Italic.ttf");
                //}
                //else
                //{
                //    return new FontResolverInfo("Fonts/Verdana-Regular.ttf");
                //}
            }
            return null;
        }
    }
}
