#define WPF

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using SpreedSheet.Core;
using SpreedSheet.Core.Workbook;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Enum;
using SpreedSheet.Interaction;
using SpreedSheet.Interface;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.Common;
using unvell.ReoGrid;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.IO;
using unvell.ReoGrid.Rendering;
using DrawingContext = System.Windows.Media.DrawingContext;
using WPFPoint = System.Windows.Point;

namespace SpreedSheet.WPF
{
    public sealed class SheetControl : Canvas, IVisualWorkbook,
        IRangePickableControl, IContextMenuControl, IPersistenceWorkbook, IActionControl, IWorkbook
    {
        internal IRenderer Renderer;

        #region Initialize

        private void InitControl()
        {
            #region initialize cursors

            // normal grid selector
            builtInCellsSelectionCursor = LoadCursorFromResource(Properties.Resources.grid_select);
            internalCurrentCursor = builtInCellsSelectionCursor;

            // cell picking
            defaultPickRangeCursor = LoadCursorFromResource(Properties.Resources.pick_range);

            // full-row and full-col selector
            builtInFullColSelectCursor = LoadCursorFromResource(Properties.Resources.full_col_select);
            builtInFullRowSelectCursor = LoadCursorFromResource(Properties.Resources.full_row_select);

            builtInEntireSheetSelectCursor = builtInCellsSelectionCursor;

            builtInCrossCursor = LoadCursorFromResource(Properties.Resources.cross);

            #endregion

            ControlStyle = ControlAppearanceStyle.CreateDefaultControlStyle();
        }

        private void InitWorkbook(IControlAdapter adapter)
        {
            // create workbook
            _workbook = new Workbook(adapter);

            #region Workbook Event Attach

            _workbook.WorksheetCreated += (s, e) => { WorksheetCreated?.Invoke(this, e); };

            _workbook.WorksheetInserted += (s, e) => { WorksheetInserted?.Invoke(this, e); };

            _workbook.WorksheetRemoved += (s, e) =>
            {
                ClearActionHistoryForWorksheet(e.Worksheet);

                if (_workbook.worksheets.Count > 0)
                {
                    var index = _sheetTab.SelectedIndex;

                    if (index >= _workbook.worksheets.Count) index = _workbook.worksheets.Count - 1;

                    _sheetTab.SelectedIndex = index;
                    _activeWorksheet = _workbook.worksheets[_sheetTab.SelectedIndex];
                }
                else
                {
                    _sheetTab.SelectedIndex = -1;
                    _activeWorksheet = null;
                }

                WorksheetRemoved?.Invoke(this, e);
            };

            _workbook.WorksheetNameChanged += (s, e) => { WorksheetNameChanged?.Invoke(this, e); };

            _workbook.SettingsChanged += (s, e) =>
            {
                if (_workbook.HasSettings(WorkbookSettings.View_ShowSheetTabControl))
                    ShowSheetTabControl();
                else
                    HideSheetTabControl();

                if (_workbook.HasSettings(WorkbookSettings.View_ShowHorScroll))
                    ShowHorScrollBar();
                else
                    HideHorScrollBar();

                if (_workbook.HasSettings(WorkbookSettings.View_ShowVerScroll))
                    ShowVerScrollBar();
                else
                    HideVerScrollBar();

                SettingsChanged?.Invoke(this, null);
            };

            _workbook.ExceptionHappened += Workbook_ErrorHappened;

            #endregion // Workbook Event Attach

            // create and set default worksheet
            _workbook.AddWorksheet(_workbook.CreateWorksheet());

            //RefreshWorksheetTabs();
            ActiveWorksheet = _workbook.worksheets[0];

            _sheetTab.SelectedIndexChanged += (s, e) =>
            {
                if (_sheetTab.SelectedIndex >= 0 && _sheetTab.SelectedIndex < _workbook.worksheets.Count)
                    ActiveWorksheet = _workbook.worksheets[_sheetTab.SelectedIndex];
            };

            _workbook.WorkbookLoaded += (s, e) =>
            {
                if (_workbook.worksheets.Count <= 0)
                {
                    _activeWorksheet = null;
                }
                else
                {
                    if (_activeWorksheet != _workbook.worksheets[0])
                        ActiveWorksheet = _workbook.worksheets[0];
                    else
                        _activeWorksheet.UpdateViewportControlBounds();
                }

                WorkbookLoaded?.Invoke(s, e);
            };

            _workbook.WorkbookSaved += (s, e) => { WorkbookSaved?.Invoke(s, e); };

            actionManager.BeforePerformAction += (s, e) =>
            {
                if (BeforeActionPerform != null)
                {
                    var arg = new BeforeActionPerformEventArgs(e.Action);

                    BeforeActionPerform(this, arg);

                    e.Cancel = arg.IsCancelled;
                }
            };

            // register for monitoring reusable action
            actionManager.AfterPerformAction += (s, e) =>
            {
                if (e.Action is WorksheetReusableAction) lastReusableAction = e.Action as WorksheetReusableAction;

                ActionPerformed?.Invoke(this, new WorkbookActionEventArgs(e.Action));
            };
        }

        #endregion

        #region Memory Workbook

        /// <summary>
        ///     Create an instance of ReoGrid workbook in memory. <br />
        ///     The memory workbook is the non-GUI version of ReoGrid control, which can do almost all operations,
        ///     such as reading and saving from Excel file, RGF file, changing data, formulas, styles, borders and etc.
        /// </summary>
        /// <returns>Instance of memory workbook.</returns>
        public static IWorkbook CreateMemoryWorkbook()
        {
            var workbook = new Workbook(null);

            var defaultWorksheet = workbook.CreateWorksheet();
            workbook.AddWorksheet(defaultWorksheet);

            return workbook;
        }

        #endregion // Memory Workbook

        #region Workbook & Worksheet

        private Workbook _workbook;

        #region Save & Load

        /// <summary>
        ///     Save workbook into file
        /// </summary>
        /// <param name="path">Full file path to save workbook</param>
        /// <param name="fileFormat">Specified file format used to save workbook</param>
        public void Save(string path)
        {
            Save(path, FileFormat._Auto);
        }

        /// <summary>
        ///     Save workbook into file
        /// </summary>
        /// <param name="path">Full file path to save workbook</param>
        /// <param name="fileFormat">Specified file format used to save workbook</param>
        public void Save(string path, FileFormat fileFormat)
        {
            Save(path, fileFormat, Encoding.Default);
        }

        /// <summary>
        ///     Save workbook into file
        /// </summary>
        /// <param name="path">Full file path to save workbook</param>
        /// <param name="fileFormat">Specified file format used to save workbook</param>
        /// <param name="encoding">Encoding used to read plain-text from resource. (Optional)</param>
        public void Save(string path, FileFormat fileFormat, Encoding encoding)
        {
            _workbook.Save(path, fileFormat, encoding);
        }

        /// <summary>
        ///     Save workbook into stream with specified format
        /// </summary>
        /// <param name="stream">Stream to output data of workbook</param>
        /// <param name="fileFormat">Specified file format used to save workbook</param>
        public void Save(Stream stream, FileFormat fileFormat)
        {
            _workbook.Save(stream, fileFormat, Encoding.Default);
        }

        /// <summary>
        ///     Save workbook into stream with specified format
        /// </summary>
        /// <param name="stream">Stream to output data of workbook</param>
        /// <param name="fileFormat">Specified file format used to save workbook</param>
        /// <param name="encoding">Encoding used to read plain-text from resource. (Optional)</param>
        public void Save(Stream stream, FileFormat fileFormat, Encoding encoding)
        {
            _workbook.Save(stream, fileFormat, encoding);
        }

        /// <summary>
        ///     Load workbook from file by specified path.
        /// </summary>
        /// <param name="path">Path to open file and read data.</param>
        public void Load(string path)
        {
            Load(path, FileFormat._Auto, Encoding.Default);
        }

        /// <summary>
        ///     Load workbook from file by specified path.
        /// </summary>
        /// <param name="path">Path to open file and read data.</param>
        /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
        public void Load(string path, FileFormat fileFormat)
        {
            Load(path, fileFormat, Encoding.Default);
        }

        /// <summary>
        ///     Load workbook from file with specified format
        /// </summary>
        /// <param name="path">Path to open file and read data.</param>
        /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
        /// <param name="encoding">Encoding used to read plain-text from resource. (Optional)</param>
        public void Load(string path, FileFormat fileFormat, Encoding encoding)
        {
            _workbook.Load(path, fileFormat, encoding);
        }

        /// <summary>
        ///     Load workbook from stream with specified format.
        /// </summary>
        /// <param name="stream">Stream to read data of workbook.</param>
        /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
        /// <param name="sheetName"></param>
        public void Load(Stream stream, FileFormat fileFormat, string sheetName)
        {
            Load(stream, fileFormat, Encoding.Default, sheetName);
        }

        /// <summary>
        ///     Load workbook from stream with specified format.
        /// </summary>
        /// <param name="stream">Stream to read data of workbook.</param>
        /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
        /// <param name="encoding">Encoding used to read plain-text data from specified stream.</param>
        /// <param name="sheetName"></param>
        public void Load(Stream stream, FileFormat fileFormat, Encoding encoding, string sheetName)
        {
            _workbook.Load(stream, fileFormat, encoding, sheetName);

            //if (workbook.worksheets.Count > 0)
            //{
            //    CurrentWorksheet = workbook.worksheets[0];
            //}
        }

        #endregion // Save & Load

        /// <summary>
        ///     Event raised when workbook loaded from stream or file.
        /// </summary>
        public event EventHandler WorkbookLoaded;

        /// <summary>
        ///     Event raised when workbook saved into stream or file.
        /// </summary>
        public event EventHandler WorkbookSaved;

        #region Worksheet Management

        private Worksheet _activeWorksheet { get; set; }

        /// <summary>
        ///     Get or set the current worksheet
        /// </summary>
        public Worksheet ActiveWorksheet
        {
            get { return _activeWorksheet; }
            set
            {
                if (value == null) throw new ArgumentNullException("cannot set current worksheet to null");

                if (_activeWorksheet != value)
                {
                    if (_activeWorksheet != null && _activeWorksheet.IsEditing)
                        _activeWorksheet.EndEdit(EndEditReason.NormalFinish);

                    _activeWorksheet = value;

                    // update bounds for viewport of worksheet
                    _activeWorksheet.UpdateViewportControlBounds();

                    // update bounds for viewport of worksheet
                    var scrollableViewportController =
                        _activeWorksheet.ViewportController as IScrollableViewportController;
                    if (scrollableViewportController != null) scrollableViewportController.SynchronizeScrollBar();

                    CurrentWorksheetChanged?.Invoke(this, null);

                    _sheetTab.SelectedIndex = GetWorksheetIndex(_activeWorksheet);
                    _sheetTab.ScrollToItem(_sheetTab.SelectedIndex);

                    Adapter.Invalidate();
                }
            }
        }

        /// <summary>
        ///     Create new instance of worksheet with default available name. (e.g. Sheet1, Sheet2 ...)
        /// </summary>
        /// <returns>Instance of worksheet to be created.</returns>
        /// <remarks>
        ///     This method creates a new worksheet, but doesn't add it into the collection of worksheet.
        ///     Worksheet will only be available until adding into a workbook, by using these methods:
        ///     <code>InsertWorksheet</code>, <code>Worksheets.Add</code> or <code>Worksheets.Insert</code>
        /// </remarks>
        public Worksheet CreateWorksheet()
        {
            return CreateWorksheet(null);
        }

        /// <summary>
        ///     Create new instance of worksheet.
        /// </summary>
        /// <param name="name">
        ///     name of new worksheet to be created.
        ///     If name is null, ReoGrid will find an available name automatically. e.g. 'Sheet1', 'Sheet2'...
        /// </param>
        /// <returns>instance of worksheet to be created</returns>
        /// <remarks>
        ///     This method creates a new worksheet, but doesn't add it into the collection of worksheet.
        ///     Worksheet will only be available until adding into a workbook, by using these methods:
        ///     <code>InsertWorksheet</code>, <code>Worksheets.Add</code> or <code>Worksheets.Insert</code>
        /// </remarks>
        public Worksheet CreateWorksheet(string name)
        {
            return _workbook.CreateWorksheet(name);
        }

        /// <summary>
        ///     Add specified worksheet into this workbook
        /// </summary>
        /// <param name="sheet">worksheet to be added</param>
        public void AddWorksheet(Worksheet sheet)
        {
            _workbook.AddWorksheet(sheet);
        }

        /// <summary>
        ///     Create and append a new instance of worksheet into workbook.
        /// </summary>
        /// <param name="name">Optional name for new worksheet.</param>
        /// <returns>Instance of created new worksheet.</returns>
        public Worksheet NewWorksheet(string name = null)
        {
            var worksheet = CreateWorksheet(name);

            AddWorksheet(worksheet);

            return worksheet;
        }

        /// <summary>
        ///     Insert specified worksheet into this workbook.
        /// </summary>
        /// <param name="index">position of zero-based number of worksheet used to insert specified worksheet.</param>
        /// <param name="sheet">worksheet to be inserted.</param>
        public void InsertWorksheet(int index, Worksheet sheet)
        {
            _workbook.InsertWorksheet(index, sheet);
        }

        /// <summary>
        ///     Remove worksheet from this workbook by specified index.
        /// </summary>
        /// <param name="index">zero-based number of worksheet to be removed.</param>
        /// <returns>true if specified worksheet can be found and removed successfully.</returns>
        public bool RemoveWorksheet(int index)
        {
            return _workbook.RemoveWorksheet(index);
        }

        /// <summary>
        ///     Remove worksheet from this workbook.
        /// </summary>
        /// <param name="sheet">worksheet to be removed.</param>
        /// <returns>true if specified worksheet can be found and removed successfully.</returns>
        public bool RemoveWorksheet(Worksheet sheet)
        {
            return _workbook.RemoveWorksheet(sheet);
        }

        /// <summary>
        ///     Create a cloned worksheet and put into specified position.
        /// </summary>
        /// <param name="index">Index of source worksheet to be copied</param>
        /// <param name="newIndex">Target index used to insert the copied worksheet</param>
        /// <param name="newName">Name for new worksheet, set as null to use a default worksheet name e.g. Sheet1, Sheet2...</param>
        /// <returns>New instance of copid worksheet</returns>
        public Worksheet CopyWorksheet(int index, int newIndex, string newName = null)
        {
            return _workbook.CopyWorksheet(index, newIndex, newName);
        }

        /// <summary>
        ///     Create a cloned worksheet and put into specified position.
        /// </summary>
        /// <param name="sheet">Source worksheet to be copied, the worksheet must be already added into this workbook</param>
        /// <param name="newIndex">Target index used to insert the copied worksheet</param>
        /// <param name="newName">Name for new worksheet, set as null to use a default worksheet name e.g. Sheet1, Sheet2...</param>
        /// <returns>New instance of copid worksheet</returns>
        public Worksheet CopyWorksheet(Worksheet sheet, int newIndex, string newName = null)
        {
            return _workbook.CopyWorksheet(sheet, newIndex, newName);
        }

        /// <summary>
        ///     Move worksheet from a position to another position.
        /// </summary>
        /// <param name="index">Worksheet in this position to be moved</param>
        /// <param name="newIndex">Target position moved to</param>
        /// <returns>Instance of moved worksheet</returns>
        public Worksheet MoveWorksheet(int index, int newIndex)
        {
            return _workbook.MoveWorksheet(index, newIndex);
        }

        /// <summary>
        ///     Create a cloned worksheet and put into specified position.
        /// </summary>
        /// <param name="sheet">Instance of worksheet to be moved, the worksheet must be already added into this workbook.</param>
        /// <param name="newIndex">Zero-based target position moved to.</param>
        /// <returns>Instance of moved worksheet.</returns>
        public Worksheet MoveWorksheet(Worksheet sheet, int newIndex)
        {
            return _workbook.MoveWorksheet(sheet, newIndex);
        }

        /// <summary>
        ///     Get index of specified worksheet from the collection in this workbook
        /// </summary>
        /// <param name="sheet">Worksheet to get.</param>
        /// <returns>zero-based number of worksheet in this workbook's collection</returns>
        public int GetWorksheetIndex(Worksheet sheet)
        {
            return _workbook.GetWorksheetIndex(sheet);
        }

        /// <summary>
        ///     Get the index of specified worksheet by name from workbook.
        /// </summary>
        /// <param name="sheet">Name of worksheet to get.</param>
        /// <returns>Zero-based number of worksheet in worksheet collection of workbook. Returns -1 if not found.</returns>
        public int GetWorksheetIndex(string name)
        {
            return _workbook.GetWorksheetIndex(name);
        }

        /// <summary>
        ///     Find worksheet by specified name.
        /// </summary>
        /// <param name="name">Name to find worksheet.</param>
        /// <returns>Instance of worksheet that is found by specified name; otherwise return null.</returns>
        public Worksheet GetWorksheetByName(string name)
        {
            return _workbook.GetWorksheetByName(name);
        }

        /// <summary>
        ///     Get the collection of worksheet.
        /// </summary>
        //[System.ComponentModel.Editor(typeof(WinForm.Designer.WorkbookEditor),
        //	typeof(System.Drawing.Design.UITypeEditor))]
        public WorksheetCollection Worksheets
        {
            get { return _workbook.Worksheets; }
        }

        /// <summary>
        ///     Event raised when current worksheet is changed.
        /// </summary>
        public event EventHandler CurrentWorksheetChanged;

        /// <summary>
        ///     Event raised when worksheet is created.
        /// </summary>
        public event EventHandler<WorksheetCreatedEventArgs> WorksheetCreated;

        /// <summary>
        ///     Event raised when worksheet is inserted into this workbook.
        /// </summary>
        public event EventHandler<WorksheetInsertedEventArgs> WorksheetInserted;

        /// <summary>
        ///     Event raised when worksheet is removed from this workbook.
        /// </summary>
        public event EventHandler<WorksheetRemovedEventArgs> WorksheetRemoved;

        /// <summary>
        ///     Event raised when the name of worksheet managed by this workbook is changed.
        /// </summary>
        public event EventHandler<WorksheetNameChangingEventArgs> WorksheetNameChanged;

        /// <summary>
        ///     Event raised when background color of worksheet name is changed.
        /// </summary>
        public event EventHandler<WorksheetEventArgs> WorksheetNameBackColorChanged;

        /// <summary>
        ///     Event raised when text color of worksheet name is changed.
        /// </summary>
        public event EventHandler<WorksheetEventArgs> WorksheetNameTextColorChanged;

        #endregion // Worksheet Management

        /// <summary>
        ///     Determine whether or not this workbook is read-only (Reserved v0.8.8)
        /// </summary>
        [Description("Determine whether or not this workbook is read-only")]
        [DefaultValue(false)]
        public bool Readonly
        {
            get { return _workbook.Readonly; }
            set { _workbook.Readonly = value; }
        }

        /// <summary>
        ///     Reset control and workbook (remove all worksheets and put one new)
        /// </summary>
        public void Reset()
        {
            _workbook.Reset();

            ActiveWorksheet = _workbook.worksheets[0];
        }

        /// <summary>
        ///     Check whether or not current workbook is empty (all worksheets don't have any cells)
        /// </summary>
        public bool IsWorkbookEmpty
        {
            get { return _workbook.IsEmpty; }
        }

        #endregion // Workbook & Worksheet

        #region Actions

        internal ActionManager actionManager = new ActionManager();

        private WorksheetReusableAction lastReusableAction;

        public void DoAction(BaseWorksheetAction action)
        {
            DoAction(_activeWorksheet, action);
        }

        /// <summary>
        ///     Do specified action.
        ///     An action does the operation as well as undoes for worksheet.
        ///     Actions performed by this method will be appended to action history stack
        ///     in order to undo, redo and repeat.
        ///     There are built-in actions available for many base operations, such as:
        ///     <code>SetCellDataAction</code> - set cell data
        ///     <code>SetRangeDataAction</code> - set data into range
        ///     <code>SetRangeBorderAction</code> - set border to specified range
        ///     <code>SetRangeStyleAction</code> - set styles to specified range
        ///     ...
        ///     It is possible to make custom action by inherting BaseWorksheetAction.
        /// </summary>
        /// <example>
        ///     ReoGrid uses ActionManager, unvell lightweight undo framework,
        ///     to implement the Do/Undo/Redo/Repeat method.
        ///     To do action:
        ///     <code>
        ///    var action = new SetCellDataAction("B1", 10);
        ///    workbook.DoAction(targetSheet, action);
        ///  </code>
        ///     To undo action:
        ///     <code>
        ///    workbook.Undo();
        ///  </code>
        ///     To redo action:
        ///     <code>
        /// 		workbook.Redo();
        ///  </code>
        ///     To repeat last action:
        ///     <code>
        /// 		workbook.RepeatLastAction(targetSheet, new ReoGridRange("B1:C3"));
        ///  </code>
        ///     It is possible to do multiple actions at same time:
        ///     <code>
        ///    var action1 = new SetRangeDataAction(...);
        ///    var action2 = new SetRangeBorderAction(...);
        ///    var action3 = new SetRangeStyleAction(...);
        ///    
        /// 		var actionGroup = new WorksheetActionGroup();
        /// 		actionGroup.Actions.Add(action1);
        /// 		actionGroup.Actions.Add(action2);
        /// 		actionGroup.Actions.Add(action3);
        /// 		
        /// 		workbook.DoAction(targetSheet, actionGroup);
        ///  </code>
        ///     Actions added into action group will be performed by one time,
        ///     they will be also undone by one time.
        /// </example>
        /// <seealso cref="ActionGroup" />
        /// <seealso cref="BaseWorksheetAction" />
        /// <seealso cref="WorksheetActionGroup" />
        /// <param name="sheet">worksheet of the target container to perform specified action</param>
        /// <param name="action">action to be performed</param>
        public void DoAction(Worksheet sheet, BaseWorksheetAction action)
        {
            action.Worksheet = sheet;

            actionManager.DoAction(action);

            var reusableAction = action as WorksheetReusableAction;
            if (reusableAction != null) lastReusableAction = reusableAction;

            if (_activeWorksheet != sheet)
            {
                sheet.RequestInvalidate();
                ActiveWorksheet = sheet;
            }

            // fix #282, https://github.com/unvell/ReoGrid/issues/282
            // comment out to avoid invoke ActionPerformed event, which is already invoked by actionManager above.
            //if (ActionPerformed != null) ActionPerformed(this, new WorkbookActionEventArgs(action));
        }

        /// <summary>
        ///     Undo the last action.
        /// </summary>
        public void Undo()
        {
            if (_activeWorksheet != null)
                if (_activeWorksheet.IsEditing)
                    _activeWorksheet.EndEdit(EndEditReason.NormalFinish);

            var action = actionManager.Undo();

            if (action != null)
            {
                if (action is WorkbookAction)
                {
                    // seems nothing to do
                }
                else
                {
                    var worksheetAction = action as BaseWorksheetAction;
                    if (worksheetAction != null)
                    {
                        var sheet = worksheetAction.Worksheet;

                        var reusableAction = action as WorksheetReusableAction;
                        if (reusableAction != null)
                            if (sheet != null)
                                sheet.SelectRange(reusableAction.Range);

                        if (sheet != null)
                        {
                            sheet.RequestInvalidate();
                            ActiveWorksheet = sheet;
                        }
                    }
                }

                Undid?.Invoke(this, new WorkbookActionEventArgs(action));
            }
        }

        /// <summary>
        ///     Redo the last action.
        /// </summary>
        public void Redo()
        {
            if (_activeWorksheet != null)
                if (_activeWorksheet.IsEditing)
                    _activeWorksheet.EndEdit(EndEditReason.NormalFinish);

            var action = actionManager.Redo();

            if (action != null)
            {
                var worksheetAction = action as BaseWorksheetAction;
                if (worksheetAction != null)
                {
                    var sheet = worksheetAction.Worksheet;

                    var reusableAction = action as WorksheetReusableAction;
                    if (reusableAction != null)
                    {
                        lastReusableAction = reusableAction;

                        if (sheet != null) sheet.SelectRange(lastReusableAction.Range);
                    }

                    if (sheet != null && _activeWorksheet != sheet)
                    {
                        sheet.RequestInvalidate();
                        ActiveWorksheet = sheet;
                    }
                }

                Redid?.Invoke(this, new WorkbookActionEventArgs(action));
            }
        }

        /// <summary>
        ///     Repeat to do last action and apply to another specified range.
        /// </summary>
        /// <param name="range">The new range to be applied for the last action.</param>
        public void RepeatLastAction(RangePosition range)
        {
            RepeatLastAction(_activeWorksheet, range);
        }

        /// <summary>
        ///     Repeat to do last action and apply to another specified range and worksheet.
        /// </summary>
        /// <param name="worksheet">The target worksheet to perform the action.</param>
        /// <param name="range">The new range to be applied for the last action.</param>
        public void RepeatLastAction(Worksheet worksheet, RangePosition range)
        {
            if (_activeWorksheet != null)
                if (_activeWorksheet.IsEditing)
                    _activeWorksheet.EndEdit(EndEditReason.NormalFinish);

            if (CanRedo())
            {
                Redo();
            }
            else
            {
                if (lastReusableAction != null)
                {
                    var newAction = lastReusableAction.Clone(range);
                    newAction.Worksheet = worksheet;

                    actionManager.DoAction(newAction);

                    // fix #282, https://github.com/unvell/ReoGrid/issues/282
                    //this.ActionPerformed?.Invoke(this, new WorkbookActionEventArgs(newAction));

                    _activeWorksheet.RequestInvalidate();
                }
            }
        }

        /// <summary>
        ///     Determine whether there is any actions can be undone.
        /// </summary>
        /// <returns>True if any actions can be undone</returns>
        public bool CanUndo()
        {
            return actionManager.CanUndo();
        }

        /// <summary>
        ///     Determine whether there is any actions can be redid.
        /// </summary>
        /// <returns>True if any actions can be redid</returns>
        public bool CanRedo()
        {
            return actionManager.CanRedo();
        }

        /// <summary>
        ///     Clear all undo/redo actions from workbook action history.
        /// </summary>
        public void ClearActionHistory()
        {
            actionManager.Reset();

            lastReusableAction = null;
        }

        /// <summary>
        ///     Delete all actions that belongs to specified worksheet.
        /// </summary>
        /// <param name="sheet">Actions belongs to this worksheet will be deleted from workbook action histroy.</param>
        public void ClearActionHistoryForWorksheet(Worksheet sheet)
        {
            var undoActions = actionManager.UndoStack;
            for (var i = 0; i < undoActions.Count;)
            {
                var action = undoActions[i];

                var worksheetAction = action as BaseWorksheetAction;

                if (worksheetAction != null && worksheetAction.Worksheet == sheet)
                {
                    undoActions.RemoveAt(i);
                    continue;
                }

                i++;
            }

            var totalActions = undoActions.Count;
            var redoActions = new List<IUndoableAction>(actionManager.RedoStack);

            for (var i = 0; i < redoActions.Count;)
            {
                var action = redoActions[i];

                var worksheetAction = action as BaseWorksheetAction;

                if (worksheetAction != null && worksheetAction.Worksheet == sheet)
                {
                    redoActions.RemoveAt(i);
                    continue;
                }

                i++;
            }

            actionManager.RedoStack.Clear();

            for (var i = redoActions.Count - 1; i >= 0; i--) actionManager.RedoStack.Push(redoActions[i]);

            totalActions += redoActions.Count;

            if (totalActions <= 0) lastReusableAction = null;
        }

        /// <summary>
        ///     Event fired before action perform.
        /// </summary>
        public event EventHandler<WorkbookActionEventArgs> BeforeActionPerform;

        /// <summary>
        ///     Event fired when any action performed.
        /// </summary>
        public event EventHandler<WorkbookActionEventArgs> ActionPerformed;

        /// <summary>
        ///     Event fired when Undo operation performed by user.
        /// </summary>
        public event EventHandler<WorkbookActionEventArgs> Undid;

        /// <summary>
        ///     Event fired when Reod operation performed by user.
        /// </summary>
        public event EventHandler<WorkbookActionEventArgs> Redid;

        #endregion // Actions

        #region Settings

        /// <summary>
        ///     Set specified workbook settings
        /// </summary>
        /// <param name="settings">Settings to be set</param>
        /// <param name="value">True to enable the settings, false to disable the settings</param>
        public void SetSettings(WorkbookSettings settings, bool value)
        {
            _workbook.SetSettings(settings, value);
        }

        /// <summary>
        ///     Get current settings of workbook
        /// </summary>
        /// <returns>Workbook settings set</returns>
        public WorkbookSettings GetSettings()
        {
            return _workbook.GetSettings();
        }

        /// <summary>
        ///     Determine whether or not the specified workbook settings has been set
        /// </summary>
        /// <param name="settings">Settings to be checked</param>
        /// <returns>True if specified settings has been set</returns>
        public bool HasSettings(WorkbookSettings settings)
        {
            return _workbook.HasSettings(settings);
        }

        /// <summary>
        ///     Enable specified settings for workbook.
        /// </summary>
        /// <param name="settings">Settings to be enabled.</param>
        public void EnableSettings(WorkbookSettings settings)
        {
            _workbook.SetSettings(settings, true);
        }

        /// <summary>
        ///     Disable specified settings for workbook.
        /// </summary>
        /// <param name="settings">Settings to be disabled.</param>
        public void DisableSettings(WorkbookSettings settings)
        {
            _workbook.SetSettings(settings, false);
        }

        /// <summary>
        ///     Event raised when settings is changed
        /// </summary>
        public event EventHandler SettingsChanged;

        #endregion // Settings

        #region Script

        /// <summary>
        ///     Get or set script content
        /// </summary>
        public string Script
        {
            get { return _workbook.Script; }
            set { _workbook.Script = value; }
        }

#if EX_SCRIPT
		// TODO: srm should have only one instance 
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		[Browsable(false)]
		public unvell.ReoScript.ScriptRunningMachine Srm
		{
			get { return this.workbook.Srm; }
		}

		/// <summary>
		/// Run workbook script.
		/// </summary>
		/// <returns>Return value from script.</returns>
		public object RunScript()
		{
			return this.workbook.RunScript();
		}

		/// <summary>
		/// Run specified script by workbook.
		/// </summary>
		/// <param name="script">Script to be executed.</param>
		/// <returns>Return value from specified script.</returns>
		public object RunScript(string script = null)
		{
			return this.workbook.RunScript(script);
		}
#endif

        #endregion // Script

        #region Internal Exceptions

        /// <summary>
        ///     Event raised when exception has been happened during internal operations.
        ///     Usually the internal operations are raised by hot-keys pressed by end-user.
        /// </summary>
        public event EventHandler<ExceptionHappenEventArgs> ExceptionHappened;

        private void Workbook_ErrorHappened(object sender, ExceptionHappenEventArgs e)
        {
            ExceptionHappened?.Invoke(this, e);
        }

        /// <summary>
        ///     Notify that there are exceptions happen on any worksheet.
        ///     The event ExceptionHappened of workbook will be invoked.
        /// </summary>
        /// <param name="sheet">Worksheet where the exception happened.</param>
        /// <param name="ex">Exception to describe the details of error information.</param>
        public void NotifyExceptionHappen(Worksheet sheet, Exception ex)
        {
            if (_workbook != null) _workbook.NotifyExceptionHappen(sheet, ex);
        }

        #endregion // Internal Exceptions

        #region Cursors

#if WINFORM || WPF
        private Cursor builtInCellsSelectionCursor;
        internal Cursor builtInFullColSelectCursor;
        internal Cursor builtInFullRowSelectCursor;
        private Cursor builtInEntireSheetSelectCursor;
        internal Cursor builtInCrossCursor;

        private Cursor customCellsSelectionCursor;
        private Cursor defaultPickRangeCursor;
        internal Cursor internalCurrentCursor;

        /// <summary>
        ///     Get or set the mouse cursor on cells selection
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Cursor CellsSelectionCursor
        {
            get { return customCellsSelectionCursor ?? builtInCellsSelectionCursor; }
            set
            {
                customCellsSelectionCursor = value;
                internalCurrentCursor = value;
            }
        }

        /// <summary>
        ///     Cursor symbol displayed when moving mouse over on row headers
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Cursor FullRowSelectionCursor { get; set; }

        /// <summary>
        ///     Cursor symbol displayed when moving mouse over on column headers
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Cursor FullColumnSelectionCursor { get; set; }

        /// <summary>
        ///     Get or set the mouse cursor of lead header part
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Cursor EntireSheetSelectionCursor { get; set; }

        private static Cursor LoadCursorFromResource(byte[] res)
        {
            using (var ms = new MemoryStream(res))
            {
                return new Cursor(ms);
            }
        }
#endif // WINFORM || WPF

        #endregion Cursors

        #region Pick Range

#if WINFORM || WPF
        /// <summary>
        ///     Start to pick a range from current worksheet.
        /// </summary>
        /// <param name="onPicked">Callback function invoked after range is picked.</param>
        public void PickRange(Func<Worksheet, RangePosition, bool> onPicked)
        {
            PickRange(onPicked, defaultPickRangeCursor);
        }

        /// <summary>
        ///     Start to pick a range from current worksheet.
        /// </summary>
        /// <param name="onPicked">Callback function invoked after range is picked.</param>
        /// <param name="pickerCursor">Cursor style during picking.</param>
        public void PickRange(Func<Worksheet, RangePosition, bool> onPicked, Cursor pickerCursor)
        {
            internalCurrentCursor = pickerCursor;

            _activeWorksheet.PickRange((sheet, range) =>
            {
                var ret = onPicked(sheet, range);
                return ret;
            });
        }

        /// <summary>
        ///     Start to pick ranges and copy the styles to the picked range
        /// </summary>
        public void StartPickRangeAndCopyStyle()
        {
            _activeWorksheet.StartPickRangeAndCopyStyle();
        }

        /// <summary>
        ///     End pick range operation
        /// </summary>
        public void EndPickRange()
        {
            _activeWorksheet.EndPickRange();

            internalCurrentCursor = customCellsSelectionCursor ?? builtInCellsSelectionCursor;
        }
#endif // WINFORM || WPF

        #endregion // Pick Range

        #region Appearance

        /// <summary>
        ///     Retrieve control instance of workbook.
        /// </summary>
        public SheetControl ControlInstance
        {
            get { return null; }
        }

        internal ControlAppearanceStyle controlStyle;

        /// <summary>
        ///     Control Style Settings
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ControlAppearanceStyle ControlStyle
        {
            get { return controlStyle; }
            set
            {
                if (value == null) throw new ArgumentNullException("ControlStyle", "cannot set ControlStyle to null");

                if (controlStyle != value)
                {
                    if (controlStyle != null) controlStyle.CurrentControl = null;
                    controlStyle = value;
                }
                //workbook.SetControlStyle(value);

                ApplyControlStyle();
            }
        }

        internal void ApplyControlStyle()
        {
            controlStyle.CurrentControl = this;
            _sheetTab.Background = new SolidColorBrush(controlStyle[ControlAppearanceColors.SheetTabBackground]);
            Adapter?.Invalidate();
        }

        //private AppearanceStyle appearanceStyle = new AppearanceStyle(this);

        #endregion // Appearance

        #region Mouse

        private void OnWorksheetMouseDown(WPFPoint location, MouseButtons buttons)
        {
            var sheet = _activeWorksheet;

            if (sheet != null)
            {
                // if currently control is in editing mode, make the input fields invisible
                if (sheet.CurrentEditingCell != null)
                {
                    var editableAdapter = Adapter as IEditableControlAdapter;
                    if (editableAdapter != null) sheet.EndEdit(editableAdapter.GetEditControlText());
                }

                sheet.ViewportController?.OnMouseDown(location, buttons);
            }
        }

        private void OnWorksheetMouseMove(WPFPoint location, MouseButtons buttons)
        {
            _activeWorksheet?.ViewportController?.OnMouseMove(location, buttons);
        }

        private void OnWorksheetMouseUp(WPFPoint location, MouseButtons buttons)
        {
            _activeWorksheet?.ViewportController?.OnMouseUp(location, buttons);
        }

        #endregion // Mouse

        protected override void OnMouseLeave(MouseEventArgs e)
        {
            base.OnMouseLeave(e);

            if (_activeWorksheet != null)
            {
                Adapter.ChangeCursor(CursorStyle.PlatformDefault);
                _activeWorksheet.HoverPos = CellPosition.Empty;
            }
        }

        #region SheetTabControl

        /// <summary>
        ///     Show or hide the built-in sheet tab control.
        /// </summary>
        public bool SheetTabVisible
        {
            get { return HasSettings(WorkbookSettings.View_ShowSheetTabControl); }
            set
            {
                if (value)
                    EnableSettings(WorkbookSettings.View_ShowSheetTabControl);
                else
                    DisableSettings(WorkbookSettings.View_ShowSheetTabControl);
            }
        }

        /// <summary>
        ///     Get or set the width of sheet tab control.
        /// </summary>
        public double SheetTabWidth
        {
            get { return _sheetTab.ControlWidth; }
            set { _sheetTab.ControlWidth = value; }
        }

        /// <summary>
        ///     Determines that whether or not to display the new button on sheet tab control.
        /// </summary>
        public bool SheetTabNewButtonVisible
        {
            get { return _sheetTab.NewButtonVisible; }
            set { _sheetTab.NewButtonVisible = value; }
        }

        #endregion // SheetTabControl

        #region Scroll

        /// <summary>
        ///     Scroll current active worksheet.
        /// </summary>
        /// <param name="x">Scroll value on horizontal direction.</param>
        /// <param name="y">Scroll value on vertical direction.</param>
        public void ScrollCurrentWorksheet(double x, double y)
        {
            var svc = _activeWorksheet?.ViewportController as IScrollableViewportController;
            if (svc != null)
            {
                svc.ScrollViews(ScrollDirection.Both, x, y);

                svc.SynchronizeScrollBar();
            }
        }

        /// <summary>
        ///     Event raised when current worksheet is scrolled.
        /// </summary>
        public event EventHandler<WorksheetScrolledEventArgs> WorksheetScrolled;

        /// <summary>
        ///     Raise the event of worksheet scrolled.
        /// </summary>
        /// <param name="worksheet">Instance of scrolled worksheet.</param>
        /// <param name="x">Scroll value on horizontal direction.</param>
        /// <param name="y">Scroll value on vertical direction.</param>
        public void RaiseWorksheetScrolledEvent(Worksheet worksheet, double x, double y)
        {
            WorksheetScrolled?.Invoke(this, new WorksheetScrolledEventArgs(worksheet)
            {
                X = x,
                Y = y
            });
        }

        private bool showScrollEndSpacing = true;

        [DefaultValue(100)]
        [Browsable(true)]
        [Description("Determines whether or not show the white spacing at bottom and right of worksheet.")]
        public bool ShowScrollEndSpacing
        {
            get { return showScrollEndSpacing; }
            set
            {
                if (showScrollEndSpacing != value)
                {
                    showScrollEndSpacing = value;
                    _activeWorksheet.UpdateViewportController();
                }
            }
        }

        #endregion Scroll

        internal const int ScrollBarSize = 18;

        public SheetControlAdapter Adapter { get; set; }
        public InputTextBox EditTextBox { get; set; }
        internal readonly SheetTabControl _sheetTab;
        internal readonly ScrollBar _horScrollbar;
        internal readonly ScrollBar _verScrollbar;

        /// <summary>
        ///     Create ReoGrid spreadsheet control
        /// </summary>
        public SheetControl()
        {
            SnapsToDevicePixels = true;
            Focusable = true;
            FocusVisualStyle = null;

            BeginInit();

            _sheetTab = new SheetTabControl
            {
                ControlWidth = 800
            };

            _horScrollbar = new ScrollBar
            {
                Orientation = Orientation.Horizontal,
                Height = ScrollBarSize,
                SmallChange = Worksheet.InitDefaultColumnWidth
            };

            _verScrollbar = new ScrollBar
            {
                Orientation = Orientation.Vertical,
                Width = ScrollBarSize,
                SmallChange = Worksheet.InitDefaultRowHeight
            };

            Children.Add(_sheetTab);
            Children.Add(_horScrollbar);
            Children.Add(_verScrollbar);

            _horScrollbar.Scroll += (s, e) =>
            {
                if (_activeWorksheet.ViewportController is IScrollableViewportController)
                    ((IScrollableViewportController)_activeWorksheet.ViewportController).HorizontalScroll(e.NewValue);
            };

            _verScrollbar.Scroll += (s, e) =>
            {
                if (_activeWorksheet.ViewportController is IScrollableViewportController)
                    ((IScrollableViewportController)_activeWorksheet.ViewportController).VerticalScroll(e.NewValue);
            };

            _sheetTab.SplitterMoving += (s, e) =>
            {
                var width = Mouse.GetPosition(this).X + 3;
                if (width < 75)
                    width = 75;
                if (width > RenderSize.Width - ScrollBarSize)
                    width = RenderSize.Width - ScrollBarSize;

                SheetTabWidth = width;

                UpdateSheetTabAndScrollBarsLayout();
            };

            InitControl();

            EditTextBox = new InputTextBox
            {
                Owner = this,
                BorderThickness = new Thickness(0),
                Visibility = Visibility.Hidden,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                VerticalScrollBarVisibility = ScrollBarVisibility.Hidden,
                Padding = new Thickness(0),
                Margin = new Thickness(0)
            };

            Children.Add(EditTextBox);

            Adapter = new SheetControlAdapter(this);
            Adapter.EditTextBox = EditTextBox;

            InitWorkbook(Adapter);

            TextCompositionManager.AddPreviewTextInputHandler(this, OnTextInputStart);

            EndInit();

            Renderer = new WPFRenderer();

            Dispatcher.BeginInvoke(DispatcherPriority.Input,
                new Action(delegate
                {
                    if (!string.IsNullOrEmpty(LoadFromFile))
                    {
                        var file = new FileInfo(LoadFromFile);
                        _activeWorksheet.Load(file.FullName);
                    }
                }));
        }

        #region SheetTab & Scroll Bars Visibility

        /// <summary>
        ///     Handle event on render size changed
        /// </summary>
        /// <param name="sizeInfo">size information</param>
        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            //this.bottomGrid.ColumnDefinitions[1].Width = new GridLength(this.RenderSize.Width
            //	- this.bottomGrid.ColumnDefinitions[0].ActualWidth - ScrollBarWidth);

            //Canvas.SetTop(this.bottomGrid, this.RenderSize.Height - ScrollBarHeight);
            //this.bottomGrid.Width = this.RenderSize.Width;
            //this.bottomGrid.Height = ScrollBarHeight;

            if (Visibility == Visibility.Visible)
                if (sizeInfo.PreviousSize.Width > 0)
                {
                    SheetTabWidth += sizeInfo.NewSize.Width - sizeInfo.PreviousSize.Width;
                    if (SheetTabWidth < 0) SheetTabWidth = 0;
                }

            UpdateSheetTabAndScrollBarsLayout();

            InvalidateVisual();
        }

        private void SetHorizontalScrollBarSize()
        {
            var hsbWidth = ActualWidth;

            if (_sheetTab.Visibility == Visibility.Visible) hsbWidth -= SheetTabWidth;

            if (_verScrollbar.Visibility == Visibility.Visible) hsbWidth -= ScrollBarSize;

            if (hsbWidth < 0) hsbWidth = 0;
            _horScrollbar.Width = hsbWidth;
        }

        private void SetSheetTabSize()
        {
            double stWidth = 0;

            if (_horScrollbar.Visibility == Visibility.Visible)
                stWidth = SheetTabWidth;
            else
                stWidth = ActualWidth;

            if (_verScrollbar.Visibility == Visibility.Visible) stWidth -= ScrollBarSize;

            if (stWidth < 0) stWidth = 0;
            _horScrollbar.Width = stWidth;
        }

        private void UpdateSheetTabAndScrollBarsLayout()
        {
            SetTop(_sheetTab, ActualHeight - ScrollBarSize);
            SetTop(_horScrollbar, ActualHeight - ScrollBarSize);

            _sheetTab.Height = ScrollBarSize;
            _horScrollbar.Height = ScrollBarSize;

            SetLeft(_verScrollbar, RenderSize.Width - ScrollBarSize);

            var vsbHeight = RenderSize.Height - ScrollBarSize;
            if (vsbHeight < 0) vsbHeight = 0;
            _verScrollbar.Height = vsbHeight;

            if (_sheetTab.Visibility == Visibility.Visible
                && _horScrollbar.Visibility == Visibility.Visible)
            {
                _sheetTab.Width = SheetTabWidth;

                SetLeft(_horScrollbar, SheetTabWidth);
                SetHorizontalScrollBarSize();
            }
            else if (_sheetTab.Visibility == Visibility.Visible)
            {
                _sheetTab.Width = ActualWidth;
            }
            else if (_horScrollbar.Visibility == Visibility.Visible)
            {
                SetLeft(_horScrollbar, 0);
                SetHorizontalScrollBarSize();
            }
            else
            {
                _verScrollbar.Height = RenderSize.Height;
            }

            _activeWorksheet.UpdateViewportControlBounds();
        }

        private void ShowSheetTabControl()
        {
            if (_sheetTab.Visibility != Visibility.Visible)
            {
                _sheetTab.Visibility = Visibility.Visible;
                UpdateSheetTabAndScrollBarsLayout();
            }
        }

        private void HideSheetTabControl()
        {
            if (_sheetTab.Visibility != Visibility.Hidden)
            {
                _sheetTab.Visibility = Visibility.Hidden;
                UpdateSheetTabAndScrollBarsLayout();
            }
        }

        private void ShowHorScrollBar()
        {
            if (_horScrollbar.Visibility != Visibility.Visible)
            {
                _horScrollbar.Visibility = Visibility.Visible;
                UpdateSheetTabAndScrollBarsLayout();
            }
        }

        private void HideHorScrollBar()
        {
            if (_horScrollbar.Visibility != Visibility.Hidden)
            {
                _horScrollbar.Visibility = Visibility.Hidden;
                UpdateSheetTabAndScrollBarsLayout();
            }
        }

        private void ShowVerScrollBar()
        {
            if (_verScrollbar.Visibility != Visibility.Visible)
            {
                _verScrollbar.Visibility = Visibility.Visible;
                UpdateSheetTabAndScrollBarsLayout();
            }
        }

        private void HideVerScrollBar()
        {
            if (_verScrollbar.Visibility != Visibility.Hidden)
            {
                _verScrollbar.Visibility = Visibility.Hidden;
                UpdateSheetTabAndScrollBarsLayout();
            }
        }

        #endregion // SheetTab & Scroll Bars Visibility

        #region Render

        /// <summary>
        ///     Handle repaint event to draw component.
        /// </summary>
        /// <param name="dc">Platform independence drawing context.</param>
        protected override void OnRender(DrawingContext dc)
        {
#if DEBUG
            var watch = Stopwatch.StartNew();
#endif

            if (_activeWorksheet != null
                && _activeWorksheet.workbook != null
                && _activeWorksheet.controlAdapter != null)
            {
                SolidColorBrush bgBrush;
                SolidColor bgColor;
                if (controlStyle.TryGetColor(ControlAppearanceColors.GridBackground, out bgColor))
                    bgBrush = new SolidColorBrush(bgColor);
                else
                    bgBrush = Brushes.White;

                dc.DrawRectangle(bgBrush, null, new Rect(0, 0, RenderSize.Width, RenderSize.Height));

                Renderer.Reset();

                ((WPFRenderer)Renderer).SetPlatformGraphics(dc);

                var rgdc = new CellDrawingContext(_activeWorksheet, DrawMode.View, Renderer);
                _activeWorksheet.ViewportController.Draw(rgdc);
            }

#if DEBUG
            watch.Stop();
            var ms = watch.ElapsedMilliseconds;
            if (ms > 30) Debug.WriteLine("end draw: {0} ms.", watch.ElapsedMilliseconds);
#endif
        }

        #endregion // Render

        #region Mouse

        private bool _mouseCaptured;

        public static readonly DependencyProperty EditTextProperty = DependencyProperty.Register("EditText",
            typeof(string), typeof(SheetControl), new PropertyMetadata(default(object)));

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);

            Focus();

            var pos = e.GetPosition(this);

            var right = RenderSize.Width;
            var bottom = RenderSize.Height;

            if (_verScrollbar.Visibility == Visibility.Visible) right = GetLeft(_verScrollbar);

            if (_sheetTab.Visibility == Visibility.Visible)
                bottom = GetTop(_sheetTab);
            else if (_horScrollbar.Visibility == Visibility.Visible) bottom = GetTop(_horScrollbar);

            if (pos.X < right && pos.Y < bottom)
            {
                if (e.ClickCount == 2)
                {
                    _activeWorksheet.OnMouseDoubleClick(e.GetPosition(this), WpfUtility.ConvertToUiMouseButtons(e));
                }
                else
                {
                    OnWorksheetMouseDown(e.GetPosition(this), WpfUtility.ConvertToUiMouseButtons(e));
                    if (CaptureMouse()) _mouseCaptured = true;
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            OnWorksheetMouseMove(e.GetPosition(this), WpfUtility.ConvertToUiMouseButtons(e));
        }

        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            base.OnMouseUp(e);

            OnWorksheetMouseUp(e.GetPosition(this), WpfUtility.ConvertToUiMouseButtons(e));

            if (_mouseCaptured) ReleaseMouseCapture();
        }

        protected override void OnMouseWheel(MouseWheelEventArgs e)
        {
            base.OnMouseWheel(e);

            _activeWorksheet.OnMouseWheel(e.GetPosition(this), e.Delta, WpfUtility.ConvertToUiMouseButtons(e));
        }

        #endregion // Mouse

        #region Keyboard

        /// <summary>
        ///     Handle event when key down.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (!_activeWorksheet.IsEditing)
            {
                var wfkeys = (KeyCode)KeyInterop.VirtualKeyFromKey(e.Key);

                if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                    wfkeys |= KeyCode.Control;
                else if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
                    wfkeys |= KeyCode.Shift;
                else if ((Keyboard.Modifiers & ModifierKeys.Alt) == ModifierKeys.Alt) wfkeys |= KeyCode.Alt;

                if (wfkeys != KeyCode.Control
                    && wfkeys != KeyCode.Shift
                    && wfkeys != KeyCode.Alt)
                    if (_activeWorksheet.OnKeyDown(wfkeys))
                        e.Handled = true;
            }
        }

        protected override void OnKeyUp(KeyEventArgs e)
        {
            if (!_activeWorksheet.IsEditing)
            {
                var wfkeys = (KeyCode)KeyInterop.VirtualKeyFromKey(e.Key);

                if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                    wfkeys |= KeyCode.Control;
                else if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
                    wfkeys |= KeyCode.Shift;
                else if ((Keyboard.Modifiers & ModifierKeys.Alt) == ModifierKeys.Alt) wfkeys |= KeyCode.Alt;

                if (wfkeys != KeyCode.Control
                    && wfkeys != KeyCode.Shift
                    && wfkeys != KeyCode.Alt)
                    if (_activeWorksheet.OnKeyUp(wfkeys))
                        e.Handled = true;

                //base.OnKeyUp(e);
            }
        }

        /// <summary>
        ///     Handle event when text inputted
        /// </summary>
        /// <param name="e"></param>
        protected override void OnTextInput(TextCompositionEventArgs e)
        {
            base.OnTextInput(e);
        }

        private void OnTextInputStart(object sender, TextCompositionEventArgs args)
        {
            if (!_activeWorksheet.IsEditing)
            {
                _activeWorksheet.StartEdit();
                _activeWorksheet.CellEditText = string.Empty;
            }
        }

        #endregion // Keyboard

        #region Adapter

        #endregion // Adapter

        #region Context Menu Strips

        internal ContextMenu BaseContextMenu
        {
            get { return ContextMenu; }
            set { ContextMenu = value; }
        }

        /// <summary>
        ///     Get or set the cells context menu
        /// </summary>
        public ContextMenu CellsContextMenu { get; set; }

        /// <summary>
        ///     Get or set the row header context menu
        /// </summary>
        public ContextMenu RowHeaderContextMenu { get; set; }

        /// <summary>
        ///     Get or set the column header context menu
        /// </summary>
        public ContextMenu ColumnHeaderContextMenu { get; set; }

        /// <summary>
        ///     Get or set the lead header context menu
        /// </summary>
        public ContextMenu LeadHeaderContextMenu { get; set; }

        #endregion // Context Menu Strips

        /// <summary>
        ///     Get or set filepath of startup template file
        /// </summary>
        public string LoadFromFile { get; set; }

        public string EditText
        {
            get
            {
                if (_activeWorksheet == null)
                    if (_activeWorksheet.selectionRange.IsSingleCell)
                        return _activeWorksheet.selectionRange.ToString();
                return (string)GetValue(EditTextProperty);
            }
            set { SetValue(EditTextProperty, value); }
        }

        public void Dispose()
        {
        }
    }
}