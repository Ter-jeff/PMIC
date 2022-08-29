using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using unvell.ReoGrid;

namespace SpreedSheet.CellTypes
{
    /// <summary>
    ///     Manage the collection of available cell types
    /// </summary>
    public static class CellTypesManager
    {
        private static Dictionary<string, Type> cellTypes;

        /// <summary>
        ///     Get the available collection of cell types
        /// </summary>
        public static Dictionary<string, Type> CellTypes
        {
            get
            {
                if (cellTypes == null)
                {
                    cellTypes = new Dictionary<string, Type>();

                    try
                    {
                        var types = Assembly.GetAssembly(typeof(Worksheet)).GetTypes();

                        foreach (var type in types.OrderBy(t => t.Name))
                            if (type != typeof(ICellBody) && type != typeof(CellBody)
                                                          && (type.IsSubclassOf(typeof(ICellBody))
                                                              || type.IsSubclassOf(typeof(CellBody)))
                                                          && type.IsPublic
                                                          && !type.IsAbstract)
                                cellTypes[type.Name] = type;
                    }
                    catch
                    {
                    }
                }

                return cellTypes;
            }
        }
    }
}