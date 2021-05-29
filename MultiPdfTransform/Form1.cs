using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.IO;
using System.Drawing.Drawing2D;
using System.Threading.Tasks;
using System.Linq;
using System.Threading;

namespace MultiPdfTransform
{
    public partial class Images2PDFs : Form
    {
        private Queue<string> currentImages = new Queue<string>();
        IEnumerable<PaperSize> paper;
        
        public Images2PDFs()
        {
            InitializeComponent();
            updateText = new UpdateText(UpdateTextMet);
            initializeProgress = new InitializeProgressBar(InitializeProgress);
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            progressBar1.Visible = true;
            progressBar1.Minimum = 0;
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        public delegate void UpdateText(string msg);
        public UpdateText updateText;
        public delegate void InitializeProgressBar(int Maximum);
        public InitializeProgressBar initializeProgress;
        public void InitializeProgress(int Maximum)
        {
            progressBar1.Maximum = Maximum;
        }
        public void UpdateTextMet(string msg)
        {
            textBox1.AppendText(msg + "\r\n");
            textBox1.ScrollToCaret();
        }
        public void ThreadText(string path)
        {
            this.BeginInvoke(updateText, path);
        }
        public void ThreadInitProgress(int Maximum)
        {
            this.Invoke(initializeProgress, Maximum);
        }


        public void ExecuteIt(string path)
        {
            ThreadText("Start Processing:"+path);
            List<string> dirs = new List<string>(Directory.EnumerateDirectories(path));
            int totalNum = dirs.Count();
            Console.WriteLine(totalNum);
            ThreadInitProgress(totalNum);
            /*dirs.Sort(delegate(string a, string b) {
                int i = Directory.EnumerateFiles(a).Count();
                int j = Directory.EnumerateFiles(b).Count();
                if (i < j) return 1;
                else if (i == j) return 0;
                else return -1;
            });*/
            int passedPages = 0;
            this.Invoke(new Action(() => { label1.Text = passedPages + "/" + totalNum; }));
            Parallel.ForEach(dirs, new ParallelOptions { MaxDegreeOfParallelism = 30 } , dir =>
            {
                string dirt = dir.Insert(dir.LastIndexOf('\\'), "Out");
                //Console.WriteLine("Executing:" + Directory.EnumerateFiles(dir).Count() + " " + passedPages + "," + dir);
                passedPages++;
                this.Invoke(new Action(() => { progressBar1.PerformStep(); }));
                if (File.Exists(dirt + ".pdf"))
                {
                    ThreadText("Passed "+passedPages+"with :" + dir.Substring(dir.LastIndexOf('\\')));
                    this.Invoke(new Action(()=> { label1.Text = passedPages + "/" + totalNum; }));
                    return;
                }
                //thistimePages++;
                Queue<string> current = new Queue<string>(Directory.EnumerateFiles(dir));
                Console.WriteLine(dirt);
                ThreadText("Processing Id="+passedPages+"with "+current.Count+"pages:" + dir.Substring(dir.LastIndexOf('\\')));
                printSinglePDF(dirt + ".pdf", current);
                this.Invoke(new Action(() => { label1.Text = passedPages + "/" + totalNum; }));
                return;
            });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //this.TopMost = true;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                progressBar1.Value = 0;
                ThreadPool.QueueUserWorkItem(w =>
                {
                    ExecuteIt(folderBrowserDialog1.SelectedPath);
                    this.Invoke(new MethodInvoker(() => MessageBox.Show("Done")));
                },null);
            }
        }

        private Boolean isIllegal(string p)
        {
            return !p.EndsWith(".jpg") && !p.EndsWith(".png") && !p.EndsWith(".gif");
        }

        private void printSinglePDF(String filename, Queue<string> que)
        {
            Queue<string> pkf = que;
            try
            {
                using (PrintDocument pdo = new PrintDocument())
                {
                    pdo.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pdo.PrinterSettings.PrintToFile = true;
                    pdo.PrinterSettings.PrintFileName = filename;
                    //pdo.PrintController = new StandardPrintController();
                    paper = pdo.PrinterSettings.PaperSizes.Cast<PaperSize>();
                    pdo.PrintPage += new PrintPageEventHandler((object sender, PrintPageEventArgs e) =>
                    {
                        string p = pkf.Dequeue();
                        while (isIllegal(p) && pkf.Count > 0) p = pkf.Dequeue();
                        if (pkf.Count > 0)
                        {
                            try
                            {
                                using (Image img = Image.FromFile(p))
                                {
                                    int wid = paper.First().Width;
                                    int hei = img.Height * paper.First().Width / img.Width;
                                    int hei2 = hei > paper.First().Height ? paper.First().Height : hei;
                                    if ((float)hei2 / paper.First().Height > 0.85f) hei2 = paper.First().Height;
                                    if ((float)wid / paper.First().Width > 0.85f) wid = paper.First().Width;
                                    if (img != null)
                                    {
                                        using (Image timg = resizeImage(img, new Size(wid, hei2)))
                                        {

                                            img.Dispose();
                                            e.Graphics.DrawImage(timg, new Point(0, 0));
                                            timg.Dispose();
                                        }
                                    }
                                    else
                                    {

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Broken Image:" + p);
                            }
                        }
                        e.HasMorePages = pkf.Count != 0;
                    });
                    pdo.Print();
                    pdo.Dispose();
                }
            }catch(Exception ex)
            {
                Console.WriteLine("Error Happened in:" + filename +"\n"+ ex.StackTrace);
            }
        }

        private void pdoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            string p = currentImages.Dequeue();
            while (isIllegal(p)) p = currentImages.Dequeue();
            using (Image img = Image.FromFile(p))
            {
                int wid = paper.First().Width;
                int hei = img.Height * paper.First().Width / img.Width;
                int hei2 = hei > paper.First().Height ? paper.First().Height : hei;
                if ((float)hei2 / paper.First().Height > 0.85f) hei2 = paper.First().Height;
                if ((float)wid / paper.First().Width > 0.85f) wid = paper.First().Width;
                using (Image temp = resizeImage(img, new Size(wid, hei2)))
                {
                    e.Graphics.DrawImage(temp, new Point(0, 0));
                }
            }
            e.HasMorePages = currentImages.Count != 0;
        }

        private static System.Drawing.Image resizeImage(System.Drawing.Image imgToResize, Size size)
        {
            //Get the image current width  
            int sourceWidth = imgToResize.Width;
            //Get the image current height  
            int sourceHeight = imgToResize.Height;
            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;
            //Calulate  width with new desired size  
            nPercentW = ((float)size.Width / (float)sourceWidth);
            //Calculate height with new desired size  
            nPercentH = ((float)size.Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;
            //New Width  
            int destWidth = (int)(sourceWidth * nPercent);
            //New Height  
            int destHeight = (int)(sourceHeight * nPercent);
            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((System.Drawing.Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // Draw image with new width and height  
            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();
            return (System.Drawing.Image)b;
        }

        private void Images2PDFs_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
