using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CommonLib.Utility
{
    public sealed class StaTaskScheduler : TaskScheduler, IDisposable
    {
        private readonly List<Thread> threads;
        private BlockingCollection<Task> tasks;

        public override int MaximumConcurrencyLevel
        {
            get { return threads.Count; }
        }

        public StaTaskScheduler(int concurrencyLevel)
        {
            if (concurrencyLevel < 1) throw new ArgumentOutOfRangeException("concurrencyLevel");

            tasks = new BlockingCollection<Task>();
            threads = Enumerable.Range(0, concurrencyLevel).Select(i =>
            {
                var thread = new Thread(() =>
                {
                    foreach (var t in tasks.GetConsumingEnumerable())
                    {
                        TryExecuteTask(t);
                    }
                });
                thread.IsBackground = true;
                thread.SetApartmentState(ApartmentState.STA);
                return thread;
            }).ToList();

            threads.ForEach(t => t.Start());
        }

        protected override void QueueTask(Task task)
        {
            tasks.Add(task);
        }
        protected override IEnumerable<Task> GetScheduledTasks()
        {
            return tasks.ToArray();
        }
        protected override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
        {
            return Thread.CurrentThread.GetApartmentState() == ApartmentState.STA && TryExecuteTask(task);
        }

        public void Dispose()
        {
            if (tasks != null)
            {
                tasks.CompleteAdding();

                foreach (var thread in threads) thread.Join();

                tasks.Dispose();
                tasks = null;
            }
        }
    }
}