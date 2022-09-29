using System;

namespace CommonLib.Utility
{
    public abstract class TimeProvider
    {
        private static TimeProvider _current = DefaultTimeProvider.Instance;

        public static TimeProvider Current
        {
            get { return _current; }
            set { _current = value; }
        }

        public abstract DateTime Now { get; }

        public static void ResetToDefault()
        {
            _current = DefaultTimeProvider.Instance;
        }
    }

    public class DefaultTimeProvider : TimeProvider
    {
        private DefaultTimeProvider()
        {
        }

        public override DateTime Now
        {
            get { return DateTime.Now; }
        }

        public static DefaultTimeProvider Instance { get; } = new DefaultTimeProvider();
    }
}