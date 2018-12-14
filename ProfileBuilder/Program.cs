using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using ProfileBuilder.ExcelData;
using System.Drawing.Text;
using System.Linq;
using CommandLine;
using ProfileBuilder.System.Data;
using Environment = System.Environment;
using System.Reflection;

namespace ProfileBuilder
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine();
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            var consoleName = Assembly.GetExecutingAssembly().ManifestModule.ToString();
            Console.WriteLine(consoleName);
            Console.WriteLine(version);

           var files = new DirectoryInfo(System.Data.Environment.InputFolder.ToString())
          .GetFiles()
          .Where(f => f.IsImage());

            foreach (var file in files)
            {
                using (var image = Image.FromFile(file.FullName))
                {
                    using (var newImage = ScaleImage(image, 450, 450))
                    {
                        try
                        {
                            var newImageName = Path.Combine(System.Data.Environment.InputFolder.ToString(), Path.GetFileNameWithoutExtension(file.Name) + "_resized" + file.Extension);
                            newImage.Save(newImageName, ImageFormat.Jpeg);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
            try
            {
                string path = Path.GetFullPath(args[0]);

                ExcelData.DataReader reader = new ExcelData.DataReader(path);
                int width = 2100;
                int height = 750;
                int width2 = 213;
                int height2 = 363;

                PrivateFontCollection pfc = new PrivateFontCollection();
                pfc.AddFontFile(System.Data.Environment.Font.ToString());


                while (reader.nextLine())
                {
                    reader.readLine();

                    using (Bitmap bmp = new Bitmap(width, height))
                    {
                        bmp.SetResolution(300, 300);
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.Clear(Color.Transparent);
                            String title = reader.JobTitle;
                            String name = reader.Name;
                            String contactInfo = reader.ContactInfo;
                            String bio = reader.Bio;

                            RectangleF name2Rect = new RectangleF();
                            using (Font useFont = new Font(pfc.Families[0], 20, FontStyle.Regular))
                            {
                                name2Rect.Location = new Point(368, 15);
                                name2Rect.Size = new Size(1200, ((int)g.MeasureString(name, useFont, 600, StringFormat.GenericTypographic).Height));
                                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                                g.DrawString(name, useFont, Brushes.Blue, name2Rect);
                            }

                            RectangleF title2Rect = new RectangleF();
                            using (Font useFont = new Font(pfc.Families[0], 12, FontStyle.Regular))
                            {
                                title2Rect.Location = new Point(375, (int)name2Rect.Bottom + 10);
                                title2Rect.Size = new Size(1200, ((int)g.MeasureString(title, useFont, 600, StringFormat.GenericTypographic).Height));
                                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                                g.DrawString(title, useFont, Brushes.Blue, title2Rect);
                            }

                            RectangleF contactInfo2Rect = new RectangleF();
                            using (Font useFont = new Font(pfc.Families[0], 12, FontStyle.Regular))
                            {
                                contactInfo2Rect.Location = new Point(375, (int)title2Rect.Bottom + 7);
                                contactInfo2Rect.Size = new Size(900, ((int)g.MeasureString(title, useFont, 900, StringFormat.GenericTypographic).Height));
                                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                                g.DrawString(contactInfo, useFont, Brushes.Blue, contactInfo2Rect);
                            }

                            RectangleF bio2Rect = new RectangleF();
                            StringFormat fmt = new StringFormat(StringFormatFlags.FitBlackBox);
                            Rectangle rc = new Rectangle(375, 250, 1700, 750);
                            using (Font useFont = new Font(pfc.Families[0], 10, FontStyle.Regular))
                            {
                                bio2Rect.Location = new Point(375, (int)contactInfo2Rect.Bottom);
                                bio2Rect.Size = new Size(1000, ((int)g.MeasureString(title, useFont, 600, StringFormat.GenericTypographic).Height));
                                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                                g.DrawString(bio, useFont, SystemBrushes.WindowText, rc, fmt);
                            }

                            
                            string image = Path.Combine(System.Data.Environment.InputFolder.ToString() + Path.GetFileNameWithoutExtension(reader.ImageFileName) + "_resized.jpg");
                            Image headshot = Image.FromFile(image);
                            Rectangle sourceRect = new Rectangle(new Point(10, 10), new Size(300, 300));

                            Rectangle crop = new Rectangle(0, 20, 305, 305);
                                g.DrawImage(headshot, sourceRect, crop, GraphicsUnit.Pixel);
                                g.DrawRectangle(
                                    pen: new Pen(
                                        color: Color.White,
                                        width: 12
                                    ),
                                    rect: new Rectangle(1, 1, 305, 305)
                                );
                           
                            bmp.Save(System.Data.Environment.OutputFolder.ToString() + name + ".png", ImageFormat.Png);

                        }
                    }

                    using (Bitmap bmp = new Bitmap(width2, height2))
                    {
                        bmp.SetResolution(300, 300);
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.Clear(Color.Transparent);
                            string name = reader.Name;
                            var nameSplit = name.Split(' ');
                            string firstName = nameSplit[0];
                            string lastName = nameSplit[1];

                            string image = Path.Combine(System.Data.Environment.InputFolder.ToString() + Path.GetFileNameWithoutExtension(reader.ImageFileName) + "_resized.jpg");
                            Image headshot = Image.FromFile(image);
                            Rectangle sourceRect = new Rectangle(new Point(0, 0), new Size(213, 213));

                            Rectangle crop = new Rectangle(0, 30, 290, 300);
                            g.DrawImage(headshot, sourceRect, crop, GraphicsUnit.Pixel);
                            
                            RectangleF firstname2Rect = new RectangleF(new Point(0, 160), new Size(213, 213));
                            using (Font useFont = new Font(pfc.Families[0], 10, FontStyle.Bold))
                            {
                                StringFormat stringFormat = new StringFormat();
                                stringFormat.Alignment = StringAlignment.Center;
                                stringFormat.LineAlignment = StringAlignment.Center;
                                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                                g.DrawString(firstName, useFont, Brushes.White, firstname2Rect, stringFormat);
                            }
                            RectangleF lastname2Rect = new RectangleF(new Point(0, 200), new Size(213, 213));
                            using (Font useFont = new Font(pfc.Families[0], 10, FontStyle.Bold))
                            {
                                StringFormat stringFormat = new StringFormat();
                                stringFormat.Alignment = StringAlignment.Center;
                                stringFormat.LineAlignment = StringAlignment.Center;
                                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                                g.DrawString(lastName, useFont, Brushes.White, lastname2Rect, stringFormat);
                            }
                            
                            bmp.Save(System.Data.Environment.OutputFolder.ToString() + name + "2.png", ImageFormat.Png);
                        }
                    }
                }
            }
            catch (Exception e)
            {

                Console.WriteLine("Error : " + e.StackTrace);
            }
        }
        
        public static Image ScaleImage(Image image, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / image.Width;
            var ratioY = (double)maxHeight / image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);

            var newImage = new Bitmap(newWidth, newHeight);

            using (var graphics = Graphics.FromImage(newImage))
                graphics.DrawImage(image, 0, 0, newWidth, newHeight);

            return newImage;
        }
    }

    public static class Extensions
    {
        public static bool IsImage(this FileInfo file)
        {
            var allowedExtensions = new[] { ".jpg", ".png", ".gif", ".jpeg" };
            return allowedExtensions.Contains(file.Extension.ToLower());
        }
    }
}
