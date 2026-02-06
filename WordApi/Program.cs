using System;
using System.Windows.Forms;

namespace WordApiService
{
    public class WordTask
    {
        public string TaskId { get; set; } = string.Empty;
        public string InputFile { get; set; } = string.Empty;
        public string OutputDocx { get; set; } = string.Empty;
        public string OutputPdf { get; set; } = string.Empty;
    }

    class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
