#define WPF

namespace unvell.ReoGrid
{
#if EX_EVENT_STYLE
	public class SGStyleTrigger
	{
		public SGStyleTrigger(SheetControl grid)
		{
			grid.CellStyleChanged += new EventHandler<ReoGridCellEventArgs>(grid_CellStyleChanged);
		}

		void grid_CellStyleChanged(object sender, ReoGridCellEventArgs e)
		{
			OnCellStyleChanged(e.Cell);
		}

		protected virtual void OnCellStyleChanged(ReoGridCell cell)
		{
			
		}
	}


#endif // EX_EVENT_STYLE

#if EX_DATA_TRIGGER
	public class RGDataTrigger
	{
		public void AttchGrid(SheetControl grid)
		{
			grid.CellDataChanged += new EventHandler<CellEventArgs>(grid_CellDataChanged);
		}

		void grid_CellDataChanged(object sender, CellEventArgs e)
		{
			OnCellDataChanged(e.Cell);
		}

		protected virtual void OnCellDataChanged(ReoGridCell cell)
		{

		}
	}

	public class RGDataTriggerActionPerformer : RGDataTrigger
	{
		public ReoGridRange TargetRange { get; set; }

		protected override void OnCellDataChanged(ReoGridCell cell)
		{
		}
	}

	public class RGDataTriggerStyleSetter : RGDataTrigger
	{
		public ReoGridPos TestCell { get; set; }
	
		public string DataContains { get; set; }
		public string ValueGreatThan { get; set; }

		public ReoGridRange StyleRange { get; set; }

		protected override void OnCellDataChanged(ReoGridCell cell)
		{
		}
	}
#endif // EX_DATA_TRIGGER
}