using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Threading;

namespace Cat.UI
{
    /// <summary>Serial STA queue for WebBrowser SVG rasterization (never use ThreadPool).</summary>
    internal static class SvgPreviewQueue
    {
        private struct Job
        {
            public string SvgPath;
            public int Size;
            public Dispatcher UiDispatcher;
            public WeakReference PreviewTarget;
        }

        private static readonly BlockingCollection<Job> Queue = new BlockingCollection<Job>();
        private static int _workerStarted;

        public static void Enqueue(string svgPath, int size, Image preview)
        {
            if (preview == null || string.IsNullOrEmpty(svgPath)) return;
            EnsureWorker();
            Queue.Add(new Job
            {
                SvgPath = svgPath,
                Size = size,
                UiDispatcher = preview.Dispatcher,
                PreviewTarget = new WeakReference(preview)
            });
        }

        private static void EnsureWorker()
        {
            if (Interlocked.Exchange(ref _workerStarted, 1) == 1) return;
            var t = new Thread(WorkerLoop)
            {
                IsBackground = true,
                Name = "CAT.SvgPreview"
            };
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        private static void WorkerLoop()
        {
            foreach (var job in Queue.GetConsumingEnumerable())
            {
                byte[] png = null;
                try { png = SvgIconHelper.RenderPreviewPngBytes(job.SvgPath, job.Size); }
                catch { png = null; }
                if (png == null || job.UiDispatcher == null) continue;

                try
                {
                    job.UiDispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
                    {
                        try
                        {
                            if (job.PreviewTarget?.Target is Image img && img.Parent != null)
                                img.Source = SvgIconHelper.BytesToBitmapImage(png);
                        }
                        catch { }
                    }));
                }
                catch { }
            }
        }
    }
}
