using Spire.Doc;
using Spire.Pdf;
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Workbook = Spire.Xls.Workbook;

namespace PrintButton
{
    public partial class Form1 : Form
    {
        private NotifyIcon trayIcon;
        private ImageList menuImageList;
        private ContextMenuStrip trayMenu;
        private string[] _files;
        public System.Windows.Forms.ToolTip formToolTip;
        public ToolStripMenuItem titleMenuItem;
        public ToolStripMenuItem excelMenuItem;
        public ToolStripMenuItem pdfMenuItem;
        public ToolStripMenuItem wordMenuItem;
        public ToolStripMenuItem exitMenuItem;

        public Form1(string[] files = null)
        {
            InitializeComponent();
            _files = files;

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(FormDragEnter);
            this.DragDrop += new DragEventHandler(FormDragDrop);

            // Initialize and configure the ToolTip
            formToolTip = new System.Windows.Forms.ToolTip();
            formToolTip.SetToolTip(this, PrintButton.Properties.Resources.DropDocXPdfOrXlsXFilesHereToPrintThem);

            // Set icon
            this.Icon = Properties.Resources.printbutton;
            this.Cursor = Cursors.Help;
            // Set TopMost
            this.TopMost = true;

            // Set up tray icon
            Version shortVersion = Assembly.GetExecutingAssembly().GetName().Version;
            trayIcon = new NotifyIcon();
            trayIcon.Text = string.Format(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + $" {shortVersion.Major}.{shortVersion.Minor}.{shortVersion.Build}");
            trayIcon.Icon = Properties.Resources.printbutton;
            trayIcon.Visible = true;

            // Set up tray menu
            trayMenu = new ContextMenuStrip();
            titleMenuItem = new ToolStripMenuItem(string.Format(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + $" {shortVersion.Major}.{shortVersion.Minor}.{shortVersion.Build}"));
            excelMenuItem = new ToolStripMenuItem("");
            pdfMenuItem = new ToolStripMenuItem("");
            wordMenuItem = new ToolStripMenuItem("");
            exitMenuItem = new ToolStripMenuItem(Properties.Resources.Leave);

            // Load the icons into the ImageList
            menuImageList = new ImageList();
            menuImageList.Images.Add(Properties.Resources.word);
            menuImageList.Images.Add(Properties.Resources.excel);
            menuImageList.Images.Add(Properties.Resources.pdf);
            menuImageList.Images.Add(Properties.Resources.exit);
            menuImageList.Images.Add(Properties.Resources.printbutton);

            // Set ImageList for the context menu
            trayMenu.ImageList = menuImageList;

            // Set ImageIndex for menu items
            titleMenuItem.ImageIndex = 4;
            excelMenuItem.ImageIndex = 1;
            pdfMenuItem.ImageIndex = 2;
            wordMenuItem.ImageIndex = 0;
            exitMenuItem.ImageIndex = 3;

            // Add event handlers
            excelMenuItem.Click += SetExcelPrinter;
            pdfMenuItem.Click += SetPdfPrinter;
            wordMenuItem.Click += SetWordPrinter;
            exitMenuItem.Click += ExitMenuItem_Click;

            // Add menu items to the trayMenu
            trayMenu.Items.Add(titleMenuItem);
            trayMenu.Items.Add(new ToolStripSeparator()); // Optional: Add a separator after title
            trayMenu.Items.Add(excelMenuItem);
            trayMenu.Items.Add(pdfMenuItem);
            trayMenu.Items.Add(wordMenuItem);
            trayMenu.Items.Add(new ToolStripSeparator()); // Optional: Add a separator before exit
            trayMenu.Items.Add(exitMenuItem);

            // Custom renderer for centering the titleMenuItem text
            trayMenu.Renderer = new TitleTextRenderer(titleMenuItem, this);

            trayIcon.ContextMenuStrip = trayMenu;

            // Form properties
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.BackColor = Color.Gray;
            this.TransparencyKey = Color.Gray; // Make the form background fully transparent
            this.Size = new Size(128, 128); // Set form size

            // Set the form icon in the middle
            this.Paint += new PaintEventHandler(DrawIconInCenter);

            // Position the form at the bottom right of the screen with 50px margins
            this.Load += new EventHandler(Form1_Load);

            // Close the app when double-clicking the tray icon
            trayIcon.DoubleClick += (s, e) => this.Close();

            // Get the system's default printer
            PrinterSettings printerSettings = new PrinterSettings();
            string defaultPrinterName = printerSettings.PrinterName;

            // Set menu item text based on settings or default printer
            if (string.IsNullOrEmpty(Properties.Settings.Default.printWordOn))
            {
                wordMenuItem.Text = defaultPrinterName;
            }
            else
            {
                wordMenuItem.Text = Properties.Settings.Default.printWordOn;
            }

            if (string.IsNullOrEmpty(Properties.Settings.Default.printExcelOn))
            {
                excelMenuItem.Text = defaultPrinterName;
            }
            else
            {
                excelMenuItem.Text = Properties.Settings.Default.printExcelOn;
            }

            if (string.IsNullOrEmpty(Properties.Settings.Default.printPdfOn))
            {
                pdfMenuItem.Text = defaultPrinterName;
            }
            else
            {
                pdfMenuItem.Text = Properties.Settings.Default.printPdfOn;
            }
        }

        // Custom renderer for the titleMenuItem
        public class TitleTextRenderer : ToolStripProfessionalRenderer
        {
            private ToolStripMenuItem _titleMenuItem;
            private Form _form;

            public TitleTextRenderer(ToolStripMenuItem titleMenuItem, Form form)
            {
                _titleMenuItem = titleMenuItem;
                _form = form;
            }

            protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
            {
                // Check if the item being rendered is the titleMenuItem
                if (e.Item != _titleMenuItem)
                {
                    // Default behavior for other menu items
                    base.OnRenderMenuItemBackground(e);
                }
            }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                // Center the text for the titleMenuItem
                if (e.Item == _titleMenuItem)
                {
                    // Get the text size
                    var textSize = e.Graphics.MeasureString(e.Text, e.TextFont);

                    // Calculate centered position
                    float x = (e.Item.Width - textSize.Width) / 2;
                    float y = (e.Item.Height - textSize.Height) / 2;

                    // Draw the text with the default text color at the centered position
                    e.Graphics.DrawString(e.Text, e.TextFont, new SolidBrush(Color.Black), new PointF(x, y));
                }
                else
                {
                    // Use the default text rendering for other items
                    base.OnRenderItemText(e);
                }
            }
        }

        private void SetExcelPrinter(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.printExcelOn = printDialog.PrinterSettings.PrinterName;
                Properties.Settings.Default.Save();
                if (!string.IsNullOrEmpty(Properties.Settings.Default.printExcelOn))
                {
                    excelMenuItem.Text = Properties.Settings.Default.printExcelOn;
                }
            }
        }

        private void SetPdfPrinter(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.printPdfOn = printDialog.PrinterSettings.PrinterName;
                Properties.Settings.Default.Save();
                if (!string.IsNullOrEmpty(Properties.Settings.Default.printPdfOn))
                {
                    pdfMenuItem.Text = Properties.Settings.Default.printPdfOn;
                }
            }
        }

        private void SetWordPrinter(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.printWordOn = printDialog.PrinterSettings.PrinterName;
                Properties.Settings.Default.Save();
                if (!string.IsNullOrEmpty(Properties.Settings.Default.printWordOn))
                {
                    wordMenuItem.Text = Properties.Settings.Default.printWordOn;
                }
            }
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Position the form at the bottom right of the screen with 50px margins
            this.Location = new Point(Screen.PrimaryScreen.WorkingArea.Width - this.Width,
                                       Screen.PrimaryScreen.WorkingArea.Height - this.Height);
        }

        private void FormDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void FormDragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                if (Path.GetExtension(file).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    PrintPdf(file);
                }
                else if (Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    PrintWord(file);
                }
                else if (Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    PrintExcel(file);
                } else
                {
                    ShowNotificationToast("Error", Properties.Resources.DropDocXPdfOrXlsXFilesHereToPrintThem, 5);
                }
            }
        }

        private void PrintPdf(string filePath)
        {
            PdfDocument pdfDocument = new PdfDocument();
            pdfDocument.LoadFromFile(filePath);
            pdfDocument.PrintSettings.PrinterName = pdfMenuItem.Text;
            pdfDocument.DocumentInformation.Title = Path.GetFileName(filePath);
            pdfDocument.Print();
        }

        private void PrintWord(string filePath)
        {
            Document document = new Document();
            document.LoadFromFile(filePath);
            document.PrintDocument.PrinterSettings.PrinterName = wordMenuItem.Text;
            document.PrintDocument.DocumentName = Path.GetFileName(filePath);
            document.PrintDocument.Print();
        }

        private void PrintExcel(string filePath)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);
            workbook.PrintDocument.PrinterSettings.PrinterName = excelMenuItem.Text;
            workbook.PrintDocument.DocumentName = Path.GetFileName(filePath);
            workbook.PrintDocument.Print();
        }

        private void DrawIconInCenter(object sender, PaintEventArgs e)
        {
            // Draw the icon in the center of the form
            int x = (this.ClientSize.Width - this.Icon.Width) / 2;
            int y = (this.ClientSize.Height - this.Icon.Height) / 2;
            e.Graphics.DrawIcon(this.Icon, x, y);
        }
        private void ShowNotificationToast(string iconType, string message, int seconds)
        {
            ToolTipIcon tooltipIcon;

            switch (iconType.ToLower())
            {
                case "info":
                    tooltipIcon = ToolTipIcon.Info;
                    break;
                case "warning":
                    tooltipIcon = ToolTipIcon.Warning;
                    break;
                case "error":
                    tooltipIcon = ToolTipIcon.Error;
                    break;
                default:
                    tooltipIcon = ToolTipIcon.None;
                    break;
            }
            trayIcon.BalloonTipIcon = tooltipIcon;
            trayIcon.BalloonTipText = message;
            trayIcon.ShowBalloonTip(seconds * 1000);
        }
    }
}
