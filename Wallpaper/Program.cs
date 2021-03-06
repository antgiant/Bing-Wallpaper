﻿using System;
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
            string url = "";
            bool download_is_temporary = false;
            uint monitor = uint.MaxValue;
			Rectangle crop = new Rectangle();
			Regex is_crop = new Regex(@"^([0-9]+,){3}([0-9]+)$", RegexOptions.IgnoreCase);
			Regex is_number = new Regex(@"^[0-9,]{1,9}$", RegexOptions.IgnoreCase);

			//Only evaluate with correct number of arguments
			if (args.Length >= 1 && args.Length <= 4)
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
                        url = args[i];
                    }
                    else if (args[i].ToLower() == "download_is_temporary") {
                        download_is_temporary = true;
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

            //Download URL, but Default to Bing Wallpaper if none provided
            if (file == "")
            {
                string path = "";
                if (download_is_temporary) {
                    path = Path.GetTempPath();
                } else {
                    path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\Wallpapers";
                }

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

                if (url == "") {
                    file = bd.DownloadSync();
                    url = "bing";
                } else {
                    file = bd.DownloadSync(url);
                }
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
                        ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);

                        // Create an Encoder object based on the GUID  
                        // for the Quality parameter category.  
                        System.Drawing.Imaging.Encoder myEncoder =  
                            System.Drawing.Imaging.Encoder.Quality;  
            
                        // Create an EncoderParameters object.  
                        // An EncoderParameters object has an array of EncoderParameter  
                        // objects. In this case, there is only one  
                        // EncoderParameter object in the array.  
                        EncoderParameters myEncoderParameters = new EncoderParameters(1);  
            
                        EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 80L);  
                        myEncoderParameters.Param[0] = myEncoderParameter;  
            
                        Bitmap bmpImage = new Bitmap(file);
                        ResizeImage(bmpImage.Clone(crop, bmpImage.PixelFormat), width, height).Save(crop_file, jpgEncoder, myEncoderParameters);
                        if (url != "") {
                            bmpImage.Dispose();
                            File.Delete(file);
                            File.Move(crop_file, file);
                        }
                    }
                    catch
                    {
                        Environment.Exit(0);
                    }
                    if (url == "") {
                        file = crop_file;
                    }
				}

                if (monitor == uint.MaxValue)
                {
                    SetAsWallPaper(file);
                }
                else
                {
                    SetAsWallPaper(file, monitor);
                }
                if (download_is_temporary) {
                    //Wait 30 Seconds for windows to actually set the wallpaper before deleting the file.
                    System.Threading.Thread.Sleep(30000);
                    File.Delete(file);
                }
            }
        }
        private static ImageCodecInfo GetEncoder(ImageFormat format)  
        {  
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();  
            foreach (ImageCodecInfo codec in codecs)  
            {  
                if (codec.FormatID == format.Guid)  
                {  
                    return codec;  
                }  
            }  
            return null;  
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