#define WPF

#if OUTLINE
namespace unvell.ReoGrid.Actions
{
	/// <summary>
	/// Action to collapse outline
	/// </summary>
	public class ExpandOutlineAction : OutlineAction
	{
		/// <summary>
		/// Create action instance to expand outline
		/// </summary>
		/// <param name="rowOrColumn">Row or column to find specified outline</param>
		/// <param name="start">Number of line of specified outline</param>
		/// <param name="count">Number of lines of specified outline</param>
		public ExpandOutlineAction(RowOrColumn rowOrColumn, int start, int count)
			: base(rowOrColumn, start, count)
		{
		}

		/// <summary>
		/// Do this action
		/// </summary>
		public override void Do()
		{
			if (this.Worksheet != null)
			{
				this.Worksheet.ExpandOutline(this.rowOrColumn, this.start, this.count);
			}
		}

		/// <summary>
		/// Undo this action
		/// </summary>
		public override void Undo()
		{
			if (this.Worksheet != null)
			{
				this.Worksheet.CollapseOutline(this.rowOrColumn, this.start, this.count);
			}
		}

		/// <summary>
		/// Get friendly name of this action
		/// </summary>
		/// <returns>Name of action</returns>
		public override string GetName()
		{
			return string.Format("Expand {0} Outline, Start at {1}, Count: {2}",
				base.GetRowOrColumnDesc(), this.start, this.count);
		}
	}
}

#endif // OUTLINE