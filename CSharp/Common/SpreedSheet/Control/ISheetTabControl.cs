﻿using System;
using SpreedSheet.Interaction;
using RGPoint = System.Windows.Point;
using RGColor = System.Windows.Media.Color;

namespace SpreedSheet.Control
{
    /// <summary>
    ///     Mouse event arguments for sheet tab control.
    /// </summary>
    public class SheetTabMouseEventArgs : EventArgs
    {
        /// <summary>
        ///     Mouse button flags. (Left, Right or Middle)
        /// </summary>
        public MouseButtons MouseButtons { get; set; }

        /// <summary>
        ///     Mouse location related to sheet tab control.
        /// </summary>
        public RGPoint Location { get; set; }

        /// <summary>
        ///     Number of tab specified by this index to be moved.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        ///     Get or set whether the user-code handled this event.
        ///     Built-in operations will be cancelled if this property is set to true.
        /// </summary>
        public bool Handled { get; set; }
    }

    /// <summary>
    ///     Sheet moved event arguments.
    /// </summary>
    public class SheetTabMovedEventArgs : EventArgs
    {
        /// <summary>
        ///     Number of tab specified by this index to be moved.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        ///     Number of tab as position moved to.
        /// </summary>
        public int TargetIndex { get; set; }
    }

    /// <summary>
    ///     Represents the border style of tab item.
    /// </summary>
    public enum SheetTabBorderStyle
    {
        /// <summary>
        ///     Sharp Rectangle
        /// </summary>
        RectShadow,

        /// <summary>
        ///     Separated Rounded Rectangle
        /// </summary>
        SplitRouned,

        /// <summary>
        ///     No Borders (Windows 8 Style)
        /// </summary>
        NoBorder
    }

    /// <summary>
    ///     Position of tab control will be located.
    /// </summary>
    public enum SheetTabControlPosition
    {
        /// <summary>
        ///     Put at top to other controls.
        /// </summary>
        Top,

        /// <summary>
        ///     Put at bottom to other controls.
        /// </summary>
        Bottom
    }

    /// <summary>
    ///     Representes the sheet tab control interface.
    /// </summary>
    public interface ISheetTabControl
    {
        ///// <summary>
        ///// Get or set the border color.
        ///// </summary>
        //[Description("Get or set the border color")]
        //RGColor BorderColor { get; set; }

        ///// <summary>
        ///// Get or set the background color for selected tab.
        ///// </summary>
        //[Description("Get or set the background color for selected tab")]
        //RGColor SelectedBackColor { get; set; }

        ///// <summary>
        ///// Get or set the text color for selected tab.
        ///// </summary>
        //[Description("Get or set the text color for selected tab")]
        //RGColor SelectedTextColor { get; set; }

        /// <summary>
        ///     Get or set the current tab index.
        /// </summary>
        int SelectedIndex { get; set; }

        /// <summary>
        ///     Get or set the width of sheet tab control
        /// </summary>
        double ControlWidth { get; set; }

        /// <summary>
        ///     Determine whether or not allow to move tab by dragging mouse.
        /// </summary>
        bool AllowDragToMove { get; set; }

        /// <summary>
        ///     Determine whether or not to show new sheet button.
        /// </summary>
        bool NewButtonVisible { get; set; }

        /// <summary>
        ///     Event raised when tab item is moved.
        /// </summary>
        event EventHandler<SheetTabMovedEventArgs> TabMoved;

        ///// <summary>
        ///// Convert the absolute point on this sheet tab control to scrolled view point.
        ///// </summary>
        ///// <param name="p">point to be converted.</param>
        ///// <returns>converted view point.</returns>
        //RGFloat TranslateScrollPoint(int p);

        ///// <summary>
        ///// Get rectangle of specified tab item.
        ///// </summary>
        ///// <param name="index">Number of tab to get bounds.</param>
        ///// <returns>Rectangle bounds of specified tab.</returns>
        //RGRect GetItemBounds(int index);

        /// <summary>
        ///     Event raised when selected tab is changed.
        /// </summary>
        event EventHandler SelectedIndexChanged;

        /// <summary>
        ///     Event raised when splitter is moved.
        /// </summary>
        event EventHandler SplitterMoving;

        /// <summary>
        ///     Event raised when sheet list button is clicked.
        /// </summary>
        event EventHandler SheetListClick;

        /// <summary>
        ///     Event raised when new sheet butotn is clicked.
        /// </summary>
        event EventHandler NewSheetClick;

        /// <summary>
        ///     Event raised when mouse is pressed down on tab items.
        /// </summary>
        event EventHandler<SheetTabMouseEventArgs> TabMouseDown;

        ///// <summary>
        ///// Move item to specified position.
        ///// </summary>
        ///// <param name="index">number of tab to be moved.</param>
        ///// <param name="targetIndex">position of moved to.</param>
        //void MoveItem(int index, int targetIndex);

        /// <summary>
        ///     Scroll view to show tab item by specified index.
        /// </summary>
        /// <param name="index">Number of item to scrolled.</param>
        void ScrollToItem(int index);

        /// <summary>
        ///     Add tab.
        /// </summary>
        /// <param name="title">Title of tab.</param>
        void AddTab(string title);

        /// <summary>
        ///     Insert tab
        /// </summary>
        /// <param name="index">Zero-based number of tab.</param>
        /// <param name="title">Title of tab.</param>
        void InsertTab(int index, string title);

        /// <summary>
        ///     Update tab title.
        /// </summary>
        /// <param name="index">Zero-based number of tab.</param>
        /// <param name="title">Title of tab.</param>
        void UpdateTab(int index, string title, RGColor backgroundColor, RGColor foregroundColor);

        /// <summary>
        ///     Remove specified tab.
        /// </summary>
        /// <param name="index">Zero-based number of tab.</param>
        void RemoveTab(int index);

        /// <summary>
        ///     Clear all tabs.
        /// </summary>
        void ClearTabs();
    }
}