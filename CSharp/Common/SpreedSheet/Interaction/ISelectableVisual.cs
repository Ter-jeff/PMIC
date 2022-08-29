namespace SpreedSheet.Interaction
{
    public interface ISelectableVisual
    {
        bool IsSelected { get; set; }

        void OnSelect();

        void OnDeselect();
    }
}