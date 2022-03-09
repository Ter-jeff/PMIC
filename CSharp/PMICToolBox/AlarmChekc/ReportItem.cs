//------------------------------------------------------------------------------
// Copyright (C) 2021 Teradyne, Inc. All rights reserved.
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
// Date        Name           Task#           Notes
// 2022-2-14  Terry Zhang     #312       Initial creation
//------------------------------------------------------------------------------ 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AlarmChekc
{
    public class ReportItem
    {
        private string _filename = string.Empty;

        public string FileName
        {
            get
            {
                return this._filename;
            }
            set
            {
                this._filename = value;
            }
        }

        private string _functionname = string.Empty;
        public string FunctionName
        {
            get
            {
                return this._functionname;
            }
            set
            {
                this._functionname = value;
            }
        }

        private string _instrument = string.Empty;
        public string Instrument
        {
            get
            {
                return this._instrument;
            }
            set
            {
                this._instrument = value;
            }
        }

        private string _alarmcategory = string.Empty;
        public string AlarmCategory
        {
            get
            {
                return this._alarmcategory;
            }
            set
            {
                this._alarmcategory = value;
            }
        }

        private string _pins = string.Empty;
        public string Pins
        {
            get
            {
                return this._pins;
            }
            set
            {
                this._pins = value;
            }
        }

        private string _alarmbehavior = string.Empty;
        public string AlarmBehavior
        {
            get
            {
                return this._alarmbehavior;
            }
            set
            {
                this._alarmbehavior = value;
            }
        }

        private string _comment = string.Empty;
        public string Comment
        {
            get
            {
                if(!string.IsNullOrEmpty(this._alarmcategory)&&!Regex.IsMatch(this._alarmcategory,"^tl.*?"))
                {
                    return this._comment + @"Alarm category is not start with tl.";
                }
                return this._comment;
            }
            set
            {
                this._comment = value;
            }
        }


        public ReportItem()
        {

        }

        public ReportItem(string p_strFileName,string p_strFunctionName,string p_strInstrument,string p_strAlarmCategory,string p_strPins,
            string p_strAlarmBehavior,string p_strComment)
        {
            this._filename = p_strFileName;
            this._functionname = p_strFunctionName;
            this._instrument = p_strInstrument;
            this._alarmcategory = p_strAlarmCategory;
            this._pins = p_strPins;
            this._alarmbehavior = p_strAlarmBehavior;
            this._comment = p_strComment;
        }

        public override bool Equals(object obj)
        {
            bool l_Rtn = true;

            ReportItem l_Item = obj as ReportItem;

            if(!this._filename.Equals(l_Item.FileName))
            {
                l_Rtn = false;
            }
            else if (!this._functionname.Equals(l_Item.FunctionName))
            {
                l_Rtn = false;
            }
            else if (!this._instrument.Equals(l_Item.Instrument))
            {
                l_Rtn = false;
            }
            else if (!this._alarmcategory.Equals(l_Item.AlarmCategory))
            {
                l_Rtn = false;
            }
            else if (!this._pins.Equals(l_Item.Pins))
            {
                l_Rtn = false;
            }
            else if (!this._alarmbehavior.Equals(l_Item.AlarmBehavior))
            {
                l_Rtn = false;
            }
            else
            {
                //do nothing
            }

            return l_Rtn;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
