using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using SpreedSheet.Interface;
using unvell.ReoGrid;

namespace SpreedSheet.Core.Workbook
{
    /// <summary>
    ///     Collection of Worksheet
    /// </summary>
    public class WorksheetCollection : IList<Worksheet>
    {
        private readonly unvell.ReoGrid.Workbook workbook;
        private IControlAdapter adapter;

        internal WorksheetCollection(unvell.ReoGrid.Workbook workbook)
        {
            adapter = workbook.controlAdapter;
            this.workbook = workbook;
        }

        /// <summary>
        ///     Get worksheet by specified name.
        /// </summary>
        /// <param name="name">Name to find worksheet</param>
        /// <returns>Instacne of worksheet found by specified name</returns>
        public Worksheet this[string name]
        {
            get { return workbook.worksheets.FirstOrDefault(s => string.Compare(s.Name, name, true) == 0); }
        }

        /// <summary>
        ///     Add worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to be added</param>
        public void Add(Worksheet sheet)
        {
            workbook.InsertWorksheet(workbook.worksheets.Count, sheet);
        }

        /// <summary>
        ///     Insert worksheet at specified position
        /// </summary>
        /// <param name="index">Zero-based number of worksheet to insert the worksheet</param>
        /// <param name="sheet">Worksheet to be inserted</param>
        public void Insert(int index, Worksheet sheet)
        {
            workbook.InsertWorksheet(index, sheet);
        }

        /// <summary>
        ///     Clear all worksheet from this workbook
        /// </summary>
        public void Clear()
        {
            workbook.ClearWorksheets();
        }

        /// <summary>
        ///     Check whether or not specified worksheet is contained in this workbook
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public bool Contains(Worksheet sheet)
        {
            return workbook.worksheets.Contains(sheet);
        }

        /// <summary>
        ///     Get number of worksheets in this workbook
        /// </summary>
        public int Count
        {
            get { return workbook.worksheets.Count; }
        }

        /// <summary>
        ///     Check whether or not current workbook is read-only
        /// </summary>
        public bool IsReadOnly
        {
            get { return workbook.Readonly; }
        }

        /// <summary>
        ///     Remove worksheet instance
        /// </summary>
        /// <param name="sheet">Instace of worksheet to be removed</param>
        /// <returns></returns>
        public bool Remove(Worksheet sheet)
        {
            return workbook.RemoveWorksheet(sheet);
        }

        /// <summary>
        ///     Get enumerator of worksheet list
        /// </summary>
        /// <returns>Enumerator of worksheet list</returns>
        public IEnumerator<Worksheet> GetEnumerator()
        {
            return workbook.worksheets.GetEnumerator();
        }

        /// <summary>
        ///     Get enumerator of worksheet list
        /// </summary>
        /// <returns>Enumerator of worksheet list</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return workbook.worksheets.GetEnumerator();
        }

        /// <summary>
        ///     Get or set worksheet by specified position
        /// </summary>
        /// <param name="index">Zero-based number of worksheet</param>
        /// <returns>Instance of worksheet found at specified position</returns>
        public Worksheet this[int index]
        {
            get
            {
                if (index < 0 || index >= workbook.worksheets.Count) throw new ArgumentOutOfRangeException("index");

                return workbook.worksheets[index];
            }
            set
            {
                if (index < 0 || index >= workbook.worksheets.Count) throw new ArgumentOutOfRangeException("index");

                workbook.worksheets[index] = value;
            }
        }

        /// <summary>
        ///     Get the index position of specified worksheet
        /// </summary>
        /// <param name="sheet">Instace of worksheet</param>
        /// <returns>Zero-based number of worksheet</returns>
        public int IndexOf(Worksheet sheet)
        {
            return workbook.GetWorksheetIndex(sheet);
        }

        /// <summary>
        ///     Remove worksheet from specified position
        /// </summary>
        /// <param name="index">Zero-based number of worksheet to locate the worksheet to be removed</param>
        public void RemoveAt(int index)
        {
            workbook.RemoveWorksheet(index);
        }

        /// <summary>
        ///     Copy all worksheet instances into specified array
        /// </summary>
        /// <param name="array">Array used to store worksheets</param>
        /// <param name="arrayIndex">Start index to copy the worksheets</param>
        public void CopyTo(Worksheet[] array, int arrayIndex)
        {
            workbook.worksheets.CopyTo(array, arrayIndex);
        }

        /// <summary>
        ///     Create worksheet by specified name
        /// </summary>
        /// <param name="name">Unique name used to identify the worksheet</param>
        /// <returns>Instance of worksheet created by specified name</returns>
        public Worksheet Create(string name = null)
        {
            return workbook.CreateWorksheet(name);
        }
    }
}