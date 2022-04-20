/*
 * cd openxlsx2/inst/extdata/fontwidth
 * csc fw.cs
 * mono fw.exe
 * 
 * requires 96 dpi monitor settings
 * 
 */

using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;

namespace openxlsx2
{
    class Program
    {
        static void Main(string[] args)
        {            
            var file = @"FontWidth.csv";
            

            using (var stream = File.CreateText(file))
            {

                string row  = string.Format("FontFamilyName, FontSize, Width, dpiX, dpiY");
                
                stream.WriteLine(row);
            
                InstalledFontCollection sysFontCollection = new InstalledFontCollection();
                FontFamily[] fontFamilies = sysFontCollection.Families;
                
                for (int i =0; i < fontFamilies.Length; i++)
                {
                
                        Bitmap bitmap = new Bitmap(100, 100);
                        Graphics graphics = Graphics.FromImage(bitmap);
                        graphics.PageUnit = GraphicsUnit.Pixel;
                        StringFormat strFormat = new StringFormat(StringFormat.GenericTypographic);
                        
                        for (int fontSize = 2; fontSize <= 25; fontSize++)
                        {
                            Font font = new Font(fontFamilies[i].Name, fontSize, FontStyle.Regular);
                            float digitWidth = graphics.MeasureString("00000000", font, 1000, strFormat).Width / 8;
                            font.Dispose();
                            float dpix = graphics.DpiX;
                            float dpiy = graphics.DpiY;
                            row  = string.Format("{0}, {1}, {2}, {3}, {4}",
                                                 font.OriginalFontName,
                                                 font.SizeInPoints,
                                                 Math.Round(digitWidth),
                                                 dpix,
                                                 dpiy); // font.FontFamily.Name
                            
                            stream.WriteLine(row);
                        }
                        graphics.Dispose();
                        bitmap.Dispose();
                }
            }
        }
    }
}
