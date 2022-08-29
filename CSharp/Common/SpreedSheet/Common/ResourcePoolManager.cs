#define WPF

#if WINFORM || WPF

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using unvell.ReoGrid.Graphics;
using FontFamily = System.Drawing.FontFamily;
using FontStyle = System.Windows.FontStyle;
using SystemFonts = System.Drawing.SystemFonts;
using WFFont = System.Drawing.Font;
using WFFontStyle = System.Drawing.FontStyle;
using WFGraphics = System.Drawing.Graphics;

#if WINFORM
using RGFloat = System.Single;

using RGDashStyle = System.Drawing.Drawing2D.DashStyle;
using RGDashStyles = System.Drawing.Drawing2D.DashStyle;

using RGPen = System.Drawing.Pen;
using RGSolidBrush = System.Drawing.SolidBrush;

using HatchBrush = System.Drawing.Drawing2D.HatchBrush;
using HatchStyle = System.Drawing.Drawing2D.HatchStyle;
#elif WPF
using RGFloat = System.Double;
using RGPen = System.Windows.Media.Pen;
using RGSolidBrush = System.Windows.Media.SolidColorBrush;
using RGBrushes = System.Windows.Media.Brushes;
using RGDashStyle = System.Windows.Media.DashStyle;
using RGDashStyles = System.Windows.Media.DashStyles;
#endif // WPF

namespace unvell.Common
{
    public sealed class ResourcePoolManager : IDisposable
    {
        //private static readonly ResourcePoolManager instance = new ResourcePoolManager();
        //public static ResourcePoolManager Instance { get { return instance; } }

        internal ResourcePoolManager()
        {
            Logger.Log("resource pool", "create resource pool...");
        }

        #region Brush

#if WINFORM || WPF
        private readonly Dictionary<SolidColor, RGSolidBrush>
            cachedBrushes = new Dictionary<SolidColor, RGSolidBrush>();

        public RGSolidBrush GetBrush(SolidColor color)
        {
            if (color.A == 0) return null;

            lock (cachedBrushes)
            {
                RGSolidBrush b;
                if (cachedBrushes.TryGetValue(color, out b)) return b;

                b = new RGSolidBrush(color);
                cachedBrushes.Add(color, b);

                if (cachedBrushes.Count % 10 == 0)
                    Logger.Log("resource pool", "solid brush count: " + cachedBrushes.Count);

                return b;
            }
        }
#endif // WINFORM || WPF

#if WINFORM
		private Dictionary<HatchStyleBrushInfo, HatchBrush> hatchBrushes =
 new Dictionary<HatchStyleBrushInfo, HatchBrush>();

		public HatchBrush GetHatchBrush(HatchStyle style, SolidColor foreColor, SolidColor backColor)
		{
			HatchStyleBrushInfo info = new HatchStyleBrushInfo(style, foreColor, backColor);

			lock (this.hatchBrushes)
			{
				if (hatchBrushes.TryGetValue(info, out var hb))
				{
					return hb;
				}
				else
				{
					HatchBrush b = new HatchBrush(style, foreColor, backColor);
					hatchBrushes.Add(info, b);

					Logger.Log("resource pool", "add hatch brush, count: " + hatchBrushes.Count);
					return b;
				}
			}
		}
		private struct HatchStyleBrushInfo
		{
			internal HatchStyle style;
			internal SolidColor foreColor;
			internal SolidColor backgroundColor;

			public HatchStyleBrushInfo(HatchStyle style, SolidColor foreColor, SolidColor backgroundColor)
			{
				this.style = style;
				this.foreColor = foreColor;
				this.backgroundColor = backgroundColor;
			}

			public override bool Equals(object obj)
			{
				if (!(obj is HatchStyleBrushInfo)) return false;

				HatchStyleBrushInfo right = (HatchStyleBrushInfo)obj;
				return (this.style == right.style
					&& this.foreColor == right.foreColor
					&& this.backgroundColor == right.backgroundColor);
			}

			public static bool operator ==(HatchStyleBrushInfo left, HatchStyleBrushInfo right)
			{
				return left.Equals(right);

				// type converted from class
				//if (left == null && right == null) return true;
				//if (left == null || right == null) return false;

				//if (left == null)
				//	return right.Equals(left);
				//else
				//	return left.Equals(right);
			}

			public static bool operator !=(HatchStyleBrushInfo left, HatchStyleBrushInfo right)
			{
				return !(left == right);
			}

			public override int GetHashCode()
			{
				return (short)style * (foreColor.ToArgb() + backgroundColor.ToArgb());
			}
		}
#endif // WINFORM

        #endregion Brush

        #region Pen

        private readonly Dictionary<SolidColor, List<RGPen>> cachedPens = new Dictionary<SolidColor, List<RGPen>>();

        public RGPen GetPen(SolidColor color)
        {
            return GetPen(color, 1, RGDashStyles.Solid);
        }

        public RGPen GetPen(SolidColor color, double weight, RGDashStyle style)
        {
            if (color.A == 0) return null;

            RGPen pen;
            List<RGPen> penlist;

            lock (cachedPens)
            {
                if (!cachedPens.TryGetValue(color, out penlist))
                {
                    penlist = cachedPens[color] = new List<RGPen>();
#if WINFORM
					penlist.Add(pen = new RGPen(color, weight));
#elif WPF
                    penlist.Add(pen = new RGPen(new RGSolidBrush(color), weight));
#endif // WPF

                    pen.DashStyle = style;

                    if (cachedPens.Count % 10 == 0) Logger.Log("resource pool", "wf pen count: " + cachedPens.Count);
                }
                else
                {
                    lock (penlist)
                    {
#if WINFORM
						pen = penlist.FirstOrDefault(p => p.Width == weight && p.DashStyle == style);
#elif WPF
                        pen = penlist.FirstOrDefault(p => p.Thickness == weight && p.DashStyle == style);
#endif // WPF
                    }

                    if (pen == null)
                    {
#if WINFORM
						penlist.Add(pen = new RGPen(color, weight));
#elif WPF
                        penlist.Add(pen = new RGPen(new RGSolidBrush(color), weight));
#endif // WPF
                        pen.DashStyle = style;

                        if (cachedPens.Count % 10 == 0) Logger.Log("resource pool", "pen count: " + cachedPens.Count);
                    }
                }
            }

            return pen;
        }

        #endregion // Pen

        #region Font

        private readonly Dictionary<string, List<WFFont>> fonts = new Dictionary<string, List<WFFont>>();

#if WINFORM
		public WFFont GetFont(string familyName, float emSize, WFFontStyle wfs)
		{

#elif WPF
        public WFFont GetFont(string familyName, double emSizeD, WFFontStyle wfs)
        {
            var emSize = (float)emSizeD;
#endif // WPF

#if DEBUG
            var sw = Stopwatch.StartNew();
#endif // DEBUG

            if (string.IsNullOrEmpty(familyName)) familyName = SystemFonts.DefaultFont.FontFamily.Name;

            WFFont font = null;
            List<WFFont> fontGroup = null;
            FontFamily family = null;

            lock (fonts)
            {
                if (fonts.TryGetValue(familyName, out fontGroup))
                {
                    if (fontGroup.Count > 0) family = fontGroup[0].FontFamily;

                    lock (fontGroup)
                    {
                        font = fontGroup.FirstOrDefault(f => f.Size == emSize && f.Style == wfs);
                    }
                }
            }

            if (font != null) return font;

            if (family == null)
            {
                try
                {
                    family = new FontFamily(familyName);
                }
                catch (ArgumentException ex)
                {
                    //throw new FontNotFoundException(ex.ParamName);
                    family = SystemFonts.DefaultFont.FontFamily;
                    Logger.Log("resource pool", "font family error: " + familyName + ": " + ex.Message);
                }

                if (!family.IsStyleAvailable(wfs))
                    try
                    {
                        wfs = FindFirstAvailableFontStyle(family);
                    }
                    catch
                    {
                        return SystemFonts.DefaultFont;
                    }
            }

            lock (fonts)
            {
                if (fonts.TryGetValue(family.Name, out fontGroup))
                    lock (fontGroup)
                    {
                        font = fontGroup.FirstOrDefault(f => f.Size == emSize && f.Style == wfs);
                    }
            }

            if (font == null)
            {
                font = new WFFont(family, emSize, wfs);

                if (fontGroup == null)
                {
                    lock (fonts)
                    {
                        fonts.Add(family.Name, fontGroup = new List<WFFont> { font });
                    }

                    Logger.Log("resource pool", "font resource group added. font groups: " + fonts.Count);
                }
                else
                {
                    lock (fontGroup)
                    {
                        fontGroup.Add(font);
                    }

                    Logger.Log("resource pool", "font resource added. fonts: " + fontGroup.Count);
                }
            }

#if DEBUG
            sw.Stop();
            var ms = sw.ElapsedMilliseconds;
            if (ms > 10) Debug.WriteLine("resource pool: font scan: " + sw.ElapsedMilliseconds + " ms.");
#endif // DEBUG
            return font;
        }

        private static WFFontStyle FindFirstAvailableFontStyle(FontFamily ff)
        {
            if (ff.IsStyleAvailable(WFFontStyle.Regular)) return WFFontStyle.Regular;

            if (ff.IsStyleAvailable(WFFontStyle.Bold)) return WFFontStyle.Bold;

            if (ff.IsStyleAvailable(WFFontStyle.Italic)) return WFFontStyle.Italic;

            if (ff.IsStyleAvailable(WFFontStyle.Strikeout)) return WFFontStyle.Strikeout;

            if (ff.IsStyleAvailable(WFFontStyle.Underline)) return WFFontStyle.Underline;

            Logger.Log("resource pool", "no available font style found: " + ff.Name);
            throw new NoAvailableFontStyleException();
        }

        internal class NoAvailableFontStyleException : Exception
        {
        }

#if WPF

        private readonly Dictionary<string, System.Windows.Media.FontFamily> fontFamilies
            = new Dictionary<string, System.Windows.Media.FontFamily>();

        public System.Windows.Media.FontFamily GetFontFamily(string name)
        {
            System.Windows.Media.FontFamily ff = null;
            fontFamilies.TryGetValue(name, out ff);
            if (ff == null)
            {
                ff = new System.Windows.Media.FontFamily(name);
                fontFamilies[name] = ff;
            }

            return ff;
        }

        private readonly Dictionary<string, List<Typeface>> typefaces
            = new Dictionary<string, List<Typeface>>();

        public Typeface GetTypeface(string name)
        {
            return GetTypeface(name, FontWeights.Regular, FontStyles.Normal,
                FontStretches.Normal);
        }

        public Typeface GetTypeface(string name, FontWeight weight,
            FontStyle style, FontStretch stretch)
        {
            List<Typeface> list;

            if (!typefaces.TryGetValue(name, out list)) typefaces[name] = list = new List<Typeface>();

            var typeface = list.FirstOrDefault(t => t.Weight == weight && t.Style == style);
            if (typeface == null)
                list.Add(typeface = new Typeface(new System.Windows.Media.FontFamily(name), style, weight, stretch));

            return typeface;
        }
#endif // WPF

        #endregion // Font

        #region Image

#if WINFORM && IMAGE_POOL
		private Dictionary<Guid, ImageResource> images
			= new Dictionary<Guid, ImageResource>();
		public ImageResource GetImageResource(Guid id)
		{
			return images.Values.FirstOrDefault(i => i.ResId.Equals(id));
		}
		public ImageResource GetImage(string fullPath)
		{
			ImageResource res = images.Values.FirstOrDefault(
				i => i.FullPath != null &&
					i.FullPath.ToLower().Equals(fullPath.ToLower()));
			if (res != null)
			{
				if (res.Image != null) res.Image.Dispose();
				res.Image = Image.FromFile(fullPath);
				return res;
			}
			else
			{
				Image image;
				try
				{
					image = Image.FromFile(fullPath);
				}
				catch(Exception ex) {
					Logger.Log("resource pool", "add image file failed: " + ex.Message);
					return null;
				}

				return AddImage(Guid.NewGuid(), image, fullPath);
			}
		}
		public ImageResource AddImage(Guid id, Image image, string fullPath)
		{
			ImageResource res;

			if (!images.TryGetValue(id, out res))
			{
				images.Add(id, res = new ImageResource()
				{
					ResId = id,
					FullPath = fullPath,
				});

				Logger.Log("resource pool", "image added. count: " + images.Count);
			}

			if (res.Image != null)
			{
				res.Image.Dispose();
			}

			res.Image = image;

			return res;
		}
#endif

        #endregion

        #region Graphics

        private static Bitmap bitmapForCachedGDIGraphics;
        private static WFGraphics cachedGDIGraphics;

        public static WFGraphics CachedGDIGraphics
        {
            get
            {
                if (cachedGDIGraphics == null)
                {
                    bitmapForCachedGDIGraphics = new Bitmap(1, 1);
                    cachedGDIGraphics = WFGraphics.FromImage(bitmapForCachedGDIGraphics);
                }

                return cachedGDIGraphics;
            }
        }

        #endregion // Graphics

        #region FormattedText

        #endregion // FormattedText

        internal void ReleaseAllResources()
        {
            Logger.Log("resource pool", "release all resources...");

            var count =
                    cachedPens.Count +

#if WINFORM
				hatchBrushes.Count + fonts.Values.Sum(f => f.Count) +
#endif
                    /*images.Count +*/ cachedBrushes.Count
#if WPF
                    + typefaces.Sum(t => t.Value.Count)
#endif
                ;

            // pens
            foreach (var plist in cachedPens.Values)
            {
#if WINFORM
				foreach (var p in plist) p.Dispose();
#endif // WINFORM
                plist.Clear();
            }

            cachedPens.Clear();

#if WINFORM
			// fonts
			foreach (var fl in fonts.Values)
			{
				foreach (var f in fl)
				{
					f.FontFamily.Dispose();
					f.Dispose();
				}
				fl.Clear();
			}

			fonts.Clear();

			foreach (var hb in this.hatchBrushes.Values)
			{
				hb.Dispose();
			}

			hatchBrushes.Clear();

			foreach (var sb in this.cachedBrushes.Values)
			{
				sb.Dispose();
			}
#elif WPF
            foreach (var list in typefaces) list.Value.Clear();
#endif // WPF

            cachedBrushes.Clear();

#if WINFORM
			//if (cachedGDIGraphics != null) cachedGDIGraphics.Dispose();
			//if (bitmapForCachedGDIGraphics != null) bitmapForCachedGDIGraphics.Dispose();
#endif // WINFORM

            Logger.Log("resource pool", count + " objects released.");
        }

        public void Dispose()
        {
            ReleaseAllResources();
        }
    }
}

#endif // WINFORM || WPF