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
using System.Collections.Generic;

namespace FWFrame.Model
{
    public class IgxlBinTablesEntry
    {
        public string Space = string.Empty;
        public string Name = string.Empty;
        public string ItemList = string.Empty;
        public string Op = string.Empty;
        public string Sort = string.Empty;
        public string Bin = string.Empty;
        public string Result = string.Empty;
        public List<string> ItemArgs = new List<string>(new string[80]);
        public string Comment = string.Empty;

        public override string ToString()
        {
            List<string> items = new List<string>();

            items.Add(Space);
            items.Add(Name);
            items.Add(ItemList);
            items.Add(Op);
            items.Add(Sort);
            items.Add(Bin);
            items.Add(Result);
            items.AddRange(ItemArgs);
            items.Add(Comment);

            return string.Join("\t", items);
        }

        public class IgxlBinTablesEntryComparer : IEqualityComparer<IgxlBinTablesEntry>
        {
            public bool Equals(IgxlBinTablesEntry x, IgxlBinTablesEntry y)
            {
                return x.Name.Equals(y.Name);
            }

            public int GetHashCode(IgxlBinTablesEntry obj)
            {
                return obj.Name.GetHashCode();
            }
        }
    }
}
