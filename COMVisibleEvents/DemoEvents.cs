using System;
using System.Diagnostics;
using System.EnterpriseServices;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMVisibleEvents
{
    [ComVisible(true)]
    [Guid("8403C952-E751-4DE1-BD91-F35DEE19206E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEvents
    {
        [DispId(1)]
        void OnDownloadCompleted();

        [DispId(2)]
        void OnDownloadFailed(string message);
    }

    [ComVisible(true)]
    [Guid("2BF7DA6B-DDB3-42A5-BD65-92EE93ABB473")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDemoEvents
    {
        [DispId(1)]
        Task DownloadFileAsync(string address, string filename);
    }

    [ComVisible(true)]
    [Guid("56C41646-10CB-4188-979D-23F70E0FFDF5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEvents))]
    [ProgId("COMVisibleEvents.DemoEvents")]
    public class DemoEvents : ServicedComponent, IDemoEvents
    {
        public delegate void OnDownloadCompletedDelegate();
        public delegate void OnDownloadFailedDelegate(string message);

        public event OnDownloadCompletedDelegate OnDownloadCompleted;
        public event OnDownloadFailedDelegate OnDownloadFailed;

        private string FileNamePath(string filename) 
            => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), filename);

        public async Task DownloadFileAsync(string address, string filename)
        {
            try
            {
                using (var webClient = new WebClient())
                {
                    await webClient
                        .DownloadFileTaskAsync(new Uri(address), FileNamePath(filename))
                        .ContinueWith(t =>
                        {
                            if (t.Status == TaskStatus.Faulted)
                            {
                                var failed = OnDownloadFailed;
                                failed?.Invoke(GetExceptions(t));
                            }
                            else
                            {
                                var completed = OnDownloadCompleted;
                                completed?.Invoke();
                            }
                        }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            #region Local

            string GetExceptions(Task task)
            {
                var innerExceptions = task.Exception?.Flatten().InnerExceptions;
                if (innerExceptions == null)
                    return string.Empty;
                var builder = new StringBuilder();
                foreach (var e in innerExceptions)
                    builder.AppendLine(e.Message);
                return builder.ToString();
            }

            #endregion Local
        }
    }
}
