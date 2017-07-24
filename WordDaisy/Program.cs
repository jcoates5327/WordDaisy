using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordDaisy {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            /*
             Console.WriteLine("Enter a source file.");
            String src = Console.ReadLine();
            Console.WriteLine("Enter a destination.");
            String dst = Console.ReadLine();

            System.IO.Directory.CreateDirectory(dst + @"\test\XML");
            Processor p = new Processor(src, dst);
            */
        }
    }
}
