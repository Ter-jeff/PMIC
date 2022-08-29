﻿#define WPF

#if WINFORM
using System.Windows.Forms;
using RGFloat = System.Single;
using RGImage = System.Drawing.Image;
#else
#endif // WINFORM
using System;
using SpreedSheet.Interaction;
using unvell.ReoGrid;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.CellTypes
{
    /// <summary>
    ///     Representation for hyperlink of cell body.
    /// </summary>
    [Serializable]
    public class HyperlinkCell : CellBody
    {
        /// <summary>
        ///     Create hyperlink cell body instance.
        /// </summary>
        public HyperlinkCell()
            : this(null, true)
        {
        }

        /// <summary>
        ///     Create instane of hyperlink cell body with specified navigation url and AutoNavigate property.
        /// </summary>
        /// <param name="navigationURL">Navigation url redirected to when hyperlink clicked. (Default is emtpy)</param>
        public HyperlinkCell(string navigationURL)
            : this(navigationURL, true)
        {
        }

        /// <summary>
        ///     Create instane of hyperlink cell body with specified navigation url and AutoNavigate property.
        /// </summary>
        /// <param name="navigationURL">Navigation url redirected to when hyperlink clicked. (Default is emtpy)</param>
        /// <param name="autoNavigate">
        ///     Determine whether or not redirect to specified url
        ///     when hyperlink clicked automatically. (Default is true)
        /// </param>
        public HyperlinkCell(string navigationURL, bool autoNavigate)
        {
            ActivateColor = SolidColor.Red;
            LinkColor = SolidColor.Blue;
            VisitedColor = SolidColor.Purple;

            LinkURL = navigationURL;
            AutoNavigate = autoNavigate;
        }

        /// <summary>
        ///     Get or set color of mouse-hover status.
        /// </summary>
        public SolidColor ActivateColor { get; set; }

        /// <summary>
        ///     Get or set color of normal hyperlink.
        /// </summary>
        public SolidColor LinkColor { get; set; }

        /// <summary>
        ///     Get or set color of visited status.
        /// </summary>
        public SolidColor VisitedColor { get; set; }

        /// <summary>
        ///     Get or set the navigation url. (Redirected when AutoNavigation property is true)
        /// </summary>
        public string LinkURL { get; set; }

        /// <summary>
        ///     Determine whether or not the hyperlink is in pressed status.
        /// </summary>
        public bool IsPressed { get; set; }

        /// <summary>
        ///     Determine whether or not redirect to navigation url when hyperlink was pressed
        /// </summary>
        public bool AutoNavigate { get; set; }

        /// <summary>
        ///     Handle event when the cell of this body entered edit mode.
        /// </summary>
        /// <returns>True to allow edit; False to disallow edit.</returns>
        public override bool OnStartEdit()
        {
            return !IsPressed;
        }

        /// <summary>
        ///     Initialize cell body when set up into a cell.
        /// </summary>
        /// <param name="cell">Instance of cell to be set up.</param>
        public override void OnSetup(Cell cell)
        {
            // set cell text style
            cell.Style.TextColor = LinkColor;
            cell.Style.Underline = true;

            // set cell default value
            if (LinkURL != null)
                cell.Data = LinkURL;
            else if (cell.InnerData != null) LinkURL = Convert.ToString(cell.InnerData);
        }

        /// <summary>
        ///     Change color of hyperlink to activate-status when mouse button pressed inside cell.
        /// </summary>
        /// <param name="e">Event argument of cell body mouse-down.</param>
        /// <returns>True if event has been handled.</returns>
        public override bool OnMouseDown(CellMouseEventArgs e)
        {
            IsPressed = true;

            e.Cell.Style.TextColor = ActivateColor;

            return true;
        }

        /// <summary>
        ///     Restore color of hyperlink to normal-status or hover-status when mouse button was released from cell.
        /// </summary>
        /// <param name="e">Event argument of cell body mouse-up.</param>
        /// <returns>True if event has been handled.</returns>
        public override bool OnMouseUp(CellMouseEventArgs e)
        {
            if (IsPressed)
                if (Bounds.Contains(e.RelativePosition))
                    PerformClick();

            IsPressed = false;

            e.Cell.Style.TextColor = VisitedColor;

            return true;
        }

        /// <summary>
        ///     Change color of hyperlink to hover-status when mouse moved into the cell.
        /// </summary>
        /// <param name="e">Event argument of cell body mouse-enter.</param>
        /// <returns>True if event has been handled.</returns>
        public override bool OnMouseEnter(CellMouseEventArgs e)
        {
            e.Worksheet.controlAdapter.ChangeSelectionCursor(CursorStyle.Hand);
            return false;
        }

        /// <summary>
        ///     Restore color of hyperlink from hover-status when mouse leaved from cell.
        /// </summary>
        /// <param name="e">Argument of mouse leaving event.</param>
        /// <returns>True if this event has been handled; Otherwise return false.</returns>
        public override bool OnMouseLeave(CellMouseEventArgs e)
        {
            // change current cursor to hand
            e.Worksheet.ControlAdapter.ChangeSelectionCursor(CursorStyle.PlatformDefault);
            return false;
        }

        /// <summary>
        ///     Handle keyboard down event.
        /// </summary>
        /// <param name="keyCode">Virtual keys code that is converted from system platform.</param>
        /// <returns>True if event has been handled; Otherwise return false.</returns>
        public override bool OnKeyDown(KeyCode keyCode)
        {
            if (keyCode == KeyCode.Space)
            {
                IsPressed = true;
                Cell.Style.TextColor = ActivateColor;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Handle keyboard up event.
        /// </summary>
        /// <param name="keyCode">Virtual keys code that is converted from system platform.</param>
        /// <returns>True if event has been handled; Otherwise return false;</returns>
        public override bool OnKeyUp(KeyCode keyCode)
        {
            if (IsPressed)
            {
                IsPressed = false;

                PerformClick();

                Cell.Style.TextColor = VisitedColor;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Handle event if cell has lost focus.
        /// </summary>
        public override void OnLostFocus()
        {
            if (IsPressed) IsPressed = false;
        }

        /// <summary>
        ///     Manually fire the hyperlink click event.
        /// </summary>
        public void PerformClick()
        {
            if (AutoNavigate && !string.IsNullOrWhiteSpace(LinkURL))
                try
                {
                    RGUtility.OpenFileOrLink(LinkURL);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Error to open link: " + ex.Message);
                    Cell?.Worksheet?.NotifyExceptionHappen(ex);
                }

            Click?.Invoke(this, null);
        }

        /// <summary>
        ///     When data of cell set, read navigation url from cell data
        /// </summary>
        /// <param name="data">New data to be set</param>
        /// <returns>Replacement data if needed</returns>
        public override object OnSetData(object data)
        {
            LinkURL = Convert.ToString(data);
            return data;
        }

        /// <summary>
        ///     Event raised when hyperlink was preseed
        /// </summary>
        public event EventHandler Click;
    }
}