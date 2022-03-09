//------------------------------------------------------------------------------
// Copyright (C) 2018 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Bug#            Notes
//
// 2018 Feb 28 Oliver Ou                      Initial creation
//
//------------------------------------------------------------------------------ 

using System.Drawing;

namespace PmicAutomation.Utility.nWire.Component
{
    /// <summary>
    /// ProgressBarEx
    /// </summary>
    public class ProgressBarEx : System.Windows.Forms.ProgressBar
    {
        private const int WM_PAINT = 0xf;
        private const float TWO = 2f;

        private Graphics _g = default(Graphics);
        private bool _bs = default(bool);
        private string _stage = default(string);
        private SizeF _sz = default(SizeF);

        /// <summary>
        /// Stage
        /// </summary>
        public string Stage
        {
            get { return _stage; }
            set { _stage = value; }
        }

        /// <summary>
        /// Draw text
        /// </summary>
        private void DrawText()
        {
            if (!string.IsNullOrEmpty(Stage))
            {
                int m;
                int n;
                _sz = _g.MeasureString(Stage, Font, Size, StringFormat.GenericDefault, out m, out n);
                _g.DrawString(Stage, Font, Brushes.Black, (Width - _sz.Width) / TWO, (Height - _sz.Height) / TWO);
            }
        }

        /// <summary>
        /// OnHandleCreated Event Handler
        /// </summary>
        /// <param name="e"></param>
        protected override void OnHandleCreated(System.EventArgs e)
        {
            base.OnHandleCreated(e);
            _g = Graphics.FromHwnd(Handle);
            _bs = true;
        }

        /// <summary>
        /// OnHandleDestroyed Event Handler
        /// </summary>
        /// <param name="e"></param>
        protected override void OnHandleDestroyed(System.EventArgs e)
        {
            base.OnHandleDestroyed(e);
            _g.Dispose();
            _bs = false;
        }

        /// <summary>
        /// WndProc
        /// </summary>
        /// <param name="m"></param>
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            base.WndProc(ref m);

            switch (m.Msg)
            {
                case WM_PAINT:
                    if (_bs)
                    {
                        DrawText();
                    }
                    break;
                default:
                    return;
            }
        }
    }
}
