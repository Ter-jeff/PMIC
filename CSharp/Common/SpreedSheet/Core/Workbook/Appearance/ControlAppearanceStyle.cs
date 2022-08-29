using System.Collections.Generic;
using SpreedSheet.WPF;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Utility;

namespace SpreedSheet.Core.Workbook.Appearance
{
    /// <summary>
    ///     Appearance Colors
    /// </summary>
    public class ControlAppearanceStyle
    {
        /// <summary>
        ///     Construct empty control appearance
        /// </summary>
        public ControlAppearanceStyle()
        {
            SelectionBorderWidth = 2f;
        }

        internal SheetControl CurrentControl { get; set; }

        internal Dictionary<ControlAppearanceColors, SolidColor> Colors { get; set; } =
            new Dictionary<ControlAppearanceColors, SolidColor>(100);

        /// <summary>
        ///     Get or set color for appearance items
        /// </summary>
        /// <param name="colorKey"></param>
        /// <returns></returns>
        public SolidColor this[ControlAppearanceColors colorKey]
        {
            get
            {
                SolidColor color;
                if (Colors.TryGetValue(colorKey, out color))
                    return color;
                return SolidColor.Black;
            }
            set { SetColor(colorKey, value); }
        }

        /// <summary>
        ///     Get or set selection border weight
        /// </summary>
        public float SelectionBorderWidth { get; set; }

        /// <summary>
        ///     Get color for appearance item
        /// </summary>
        /// <param name="colorKey">key to get the color item</param>
        /// <param name="color">output color get by specified key</param>
        /// <returns>true if color is found by specified key</returns>
        public bool GetColor(ControlAppearanceColors colorKey, out SolidColor color)
        {
            return Colors.TryGetValue(colorKey, out color);
        }

        /// <summary>
        ///     Set color for appearance item
        /// </summary>
        /// <param name="colorKey">Key of appearance item</param>
        /// <param name="color">Color to be set</param>
        public void SetColor(ControlAppearanceColors colorKey, SolidColor color)
        {
            Colors[colorKey] = color;
            CurrentControl?.ApplyControlStyle();
        }

        /// <summary>
        ///     Try get a color item from control appearance style set
        /// </summary>
        /// <param name="key">Key used to specify a item</param>
        /// <param name="color">Output color struction</param>
        /// <returns>True if key was found and color could be returned; otherwise return false</returns>
        public bool TryGetColor(ControlAppearanceColors key, out SolidColor color)
        {
            return Colors.TryGetValue(key, out color);
        }

        internal SolidColor GetColHeadStartColor(bool isHover, bool isSelected, bool isFullSelected, bool isInvalid)
        {
            if (isFullSelected)
                return Colors[ControlAppearanceColors.ColHeadFullSelectedStart];
            if (isSelected)
                return Colors[ControlAppearanceColors.ColHeadSelectedStart];
            if (isHover)
                return Colors[ControlAppearanceColors.ColHeadHoverStart];
            if (isInvalid)
                return Colors[ControlAppearanceColors.ColHeadInvalidStart];
            return Colors[ControlAppearanceColors.ColHeadNormalStart];
        }

        internal SolidColor GetColHeadEndColor(bool isHover, bool isSelected, bool isFullSelected, bool isInvalid)
        {
            if (isFullSelected)
                return Colors[ControlAppearanceColors.ColHeadFullSelectedEnd];
            if (isSelected)
                return Colors[ControlAppearanceColors.ColHeadSelectedEnd];
            if (isHover)
                return Colors[ControlAppearanceColors.ColHeadHoverEnd];
            if (isInvalid)
                return Colors[ControlAppearanceColors.ColHeadInvalidEnd];
            return Colors[ControlAppearanceColors.ColHeadNormalEnd];
        }

        internal SolidColor GetRowHeadEndColor(bool isHover, bool isSelected, bool isFullSelected, bool isInvalid)
        {
            if (isFullSelected)
                return Colors[ControlAppearanceColors.RowHeadFullSelected];
            if (isSelected)
                return Colors[ControlAppearanceColors.RowHeadSelected];
            if (isHover)
                return Colors[ControlAppearanceColors.RowHeadHover];
            if (isInvalid)
                return Colors[ControlAppearanceColors.RowHeadInvalid];
            return Colors[ControlAppearanceColors.RowHeadNormal];
        }

        /// <summary>
        ///     Create default style for grid control.
        /// </summary>
        /// <returns>Default style created</returns>
        public static ControlAppearanceStyle CreateDefaultControlStyle()
        {
            return new ControlAppearanceStyle
            {
                Colors = new Dictionary<ControlAppearanceColors, SolidColor>
                {
                    { ControlAppearanceColors.LeadHeadNormal, SolidColor.Lavender },
                    { ControlAppearanceColors.LeadHeadSelected, SolidColor.Lavender },
                    { ControlAppearanceColors.LeadHeadIndicatorStart, SolidColor.Gainsboro },
                    { ControlAppearanceColors.LeadHeadIndicatorEnd, SolidColor.Silver },
                    { ControlAppearanceColors.ColHeadSplitter, SolidColor.LightSteelBlue },
                    { ControlAppearanceColors.ColHeadNormalStart, SolidColor.White },
                    { ControlAppearanceColors.ColHeadNormalEnd, SolidColor.Lavender },
                    { ControlAppearanceColors.ColHeadHoverStart, SolidColor.LightGoldenrodYellow },
                    { ControlAppearanceColors.ColHeadHoverEnd, SolidColor.Goldenrod },
                    { ControlAppearanceColors.ColHeadSelectedStart, SolidColor.LightGoldenrodYellow },
                    { ControlAppearanceColors.ColHeadSelectedEnd, SolidColor.Goldenrod },
                    { ControlAppearanceColors.ColHeadFullSelectedStart, SolidColor.WhiteSmoke },
                    { ControlAppearanceColors.ColHeadFullSelectedEnd, SolidColor.LemonChiffon },
                    { ControlAppearanceColors.ColHeadText, SolidColor.DarkBlue },
                    { ControlAppearanceColors.RowHeadSplitter, SolidColor.LightSteelBlue },
                    { ControlAppearanceColors.RowHeadNormal, SolidColor.AliceBlue },
                    { ControlAppearanceColors.RowHeadHover, SolidColor.LightSteelBlue },
                    { ControlAppearanceColors.RowHeadSelected, SolidColor.PaleGoldenrod },
                    { ControlAppearanceColors.RowHeadFullSelected, SolidColor.LemonChiffon },
                    { ControlAppearanceColors.RowHeadText, SolidColor.DarkBlue },
                    { ControlAppearanceColors.GridText, SolidColor.Black },
                    { ControlAppearanceColors.GridBackground, SolidColor.White },
                    { ControlAppearanceColors.GridLine, SolidColor.FromArgb(255, 208, 215, 229) },
                    {
                        ControlAppearanceColors.SelectionBorder,
                        ColorUtility.FromAlphaColor(255, StaticResources.SystemColor_Highlight)
                    },
                    {
                        ControlAppearanceColors.SelectionFill,
                        ColorUtility.FromAlphaColor(30, StaticResources.SystemColor_Highlight)
                    },
                    { ControlAppearanceColors.OutlineButtonBorder, SolidColor.Black },
                    { ControlAppearanceColors.OutlinePanelBackground, StaticResources.SystemColor_Control },
                    { ControlAppearanceColors.OutlinePanelBorder, SolidColor.Silver },
                    { ControlAppearanceColors.OutlineButtonText, StaticResources.SystemColor_WindowText },
                    { ControlAppearanceColors.SheetTabText, StaticResources.SystemColor_WindowText },
                    { ControlAppearanceColors.SheetTabBorder, StaticResources.SystemColor_Highlight },
                    { ControlAppearanceColors.SheetTabBackground, SolidColor.White },
                    { ControlAppearanceColors.SheetTabSelected, StaticResources.SystemColor_Window }
                }
            };
        }
    }
}