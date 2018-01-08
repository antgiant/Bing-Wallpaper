using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace Wallpaper
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string file = "";
            uint monitor = uint.MaxValue;
			Rectangle crop = new Rectangle();
			Regex is_crop = new Regex(@"^([0-9]+,){3}([0-9]+)$", RegexOptions.IgnoreCase);
			Regex is_number = new Regex(@"^[0-9,]{1,9}$", RegexOptions.IgnoreCase);

			//Only evaluate 1 or two arguments
			if (args.Length >= 1 && args.Length <= 3)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    //If argument is a number set wallpaper on only that monitor (Assumes less than 1,000,000,000 monitors.)
                    if (is_crop.Match(args[i]).Success)
                    {
                        try
                        {
							string[] arg_split = args[i].Split(',');
							crop = new Rectangle(
												Convert.ToInt32(arg_split[0]),	//x
												Convert.ToInt32(arg_split[1]),	//y
												Convert.ToInt32(arg_split[2]),	//width
												Convert.ToInt32(arg_split[3]));	//height

                        }
                        catch
                        {
                            //If monitor number is invalid do nothing and exit.
                            Environment.Exit(0);
                        }
                    }
					//If argument is a number set wallpaper on only that monitor (Assumes less than 1,000,000,000 monitors.)
					else if (is_number.Match(args[i]).Success)
					{
						try
						{
							//Subtract one so that passed in monitor number equals what windows displays
							monitor = Convert.ToUInt32(args[i]) - 1;

							//If specified monitor number is greater than the current number of monitors do nothing and exit.
							var wallpaper = (IDesktopWallpaper)(new DesktopWallpaperClass());
							if (wallpaper.GetMonitorDevicePathCount() < monitor)
							{
								Environment.Exit(0);
							}
						}
						catch
						{
							//If monitor number is invalid do nothing and exit.
							Environment.Exit(0);
						}
					}
					//If argument is a URL download and set wallpaper to that. (Only use trusted URLs as this will download anything.)
					else if (args[i].Substring(0, 4).ToLower() == "http")
                    {
                        string path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\Wallpapers";

                        if (!Directory.Exists(path))
                            Directory.CreateDirectory(path);

                        var bd =
                            new Downloader(path)
                            {
                                UseHttps = true,
                                Size =
                                    new Rect
                                    {
                                        Right = Screen.PrimaryScreen.Bounds.Width,
                                        Bottom = Screen.PrimaryScreen.Bounds.Height
                                    }
                            };

                        file = bd.DownloadSync(args[i]);

                        //If specified URL cannot be downloaded do nothing and exit.
                        if (file == string.Empty)
                        {
                            Environment.Exit(0);
                        }
                    }
                    //Assume argument was a wallpaper image
                    else
                    {
                        //If specified file does not exist and path is not a folder do nothing and exit.
                        if (File.Exists(args[i]))
                        {
                            file = args[i];
                        }
                        else if (Directory.Exists(args[i]))
                        {
                            var files = Directory.EnumerateFiles(args[i], "*.*")
                                                .Where(s => s.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".gif", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".tif", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".tiff", StringComparison.OrdinalIgnoreCase)
                                                       );
                            if (files.Count() > 0)
                            {
                                Random rand = new Random();
                                file = files.ElementAt(rand.Next(0, files.Count()));
                            }
                            else
                            {
                                Environment.Exit(0);
                            }
                        }
                        else
                        {
                            Environment.Exit(0);
                        }
                    }
                }
            }

            //Default to Bing Wallpaper
            if (file == "")
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\Wallpapers";

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                var bd =
                    new Downloader(path)
                    {
                        UseHttps = true,
                        Size =
                            new Rect
                            {
                                Right = Screen.PrimaryScreen.Bounds.Width,
                                Bottom = Screen.PrimaryScreen.Bounds.Height
                            }
                    };

                file = bd.DownloadSync();
            }

            if (!string.IsNullOrEmpty(file))
            {
				//Crop image if crop data is provided
				if (!crop.IsEmpty)
				{
					String crop_file = Path.GetDirectoryName(file) + "\\" + Path.GetFileNameWithoutExtension(file) + "_cropped" + Path.GetExtension(file);

                    int width = crop.Width;
                    int height = crop.Height;
                    float scale = System.Math.Max((float)crop.Width / (float)Screen.PrimaryScreen.Bounds.Width, (float)crop.Height / (float)Screen.PrimaryScreen.Bounds.Height);
                    if (scale > 1)
                    {
                        width = (int)(crop.Width / scale);
                        height = (int)(crop.Height / scale);
                    }
                    try
                    {
                        Bitmap bmpImage = new Bitmap(file);
                        ResizeImage(bmpImage.Clone(crop, bmpImage.PixelFormat), width, height).Save(crop_file);
                    }
                    catch
                    {
                        Environment.Exit(0);
                    }

                    file = crop_file;
				}

                if (monitor == uint.MaxValue)
                {
                    SetAsWallPaper(file);
                }
                else
                {
                    SetAsWallPaper(file, monitor);
                }
            }
        }

        private static void SetAsWallPaper(string file)
        {
            try
            {
                var wallpaper = (IDesktopWallpaper)(new DesktopWallpaperClass());

                for (uint i = 0; i < wallpaper.GetMonitorDevicePathCount(); i++)
                {
                    wallpaper.SetWallpaper(wallpaper.GetMonitorDevicePathAt(i), file);
                    wallpaper.SetPosition(DesktopWallpaperPosition.Fill);
                }
            }
            catch
            {
                Environment.Exit(0);
            }
        }
        private static void SetAsWallPaper(string file, uint monitor)
        {
            try
            {
                var wallpaper = (IDesktopWallpaper)(new DesktopWallpaperClass());
                Rect monitorRect = wallpaper.GetMonitorRECT(wallpaper.GetMonitorDevicePathAt(monitor));
                wallpaper.SetWallpaper(wallpaper.GetMonitorDevicePathAt(monitor), file);
                wallpaper.SetPosition(DesktopWallpaperPosition.Fill);
            }
            catch
            {
                Environment.Exit(0);
            }
        }
        /// <summary>
        /// Resize the image to the specified width and height. Source https://stackoverflow.com/questions/1922040/resize-an-image-c-sharp
        /// </summary>
        /// <param name="image">The image to resize.</param>
        /// <param name="width">The width to resize to.</param>
        /// <param name="height">The height to resize to.</param>
        /// <returns>The resized image.</returns>
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
    }
}