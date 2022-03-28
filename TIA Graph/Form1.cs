using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using DGVPrinterHelper;
using Excel = Microsoft.Office.Interop.Excel;

namespace TIA_Graph
{
    public partial class Form1 : Form
    {

        #region Var

        private XmlNamespaceManager _ns;

        private XmlDocument _document;

        public XmlDocument Document
        {
            get { return _document; }
            set { _document = value; }
        }

        private XmlNode _rootNode;

        public XmlNode RootNode
        {
            get { return _rootNode; }
            set { _rootNode = value; }
        }
        
        private string FileinfoName;

        private List<(string Key, int row, int col)> BranchRefList = new List<(string Key, int row, int col)>();

        private bool RowAligment;

        #endregion


        public Form1()
        {
            InitializeComponent();
        }


        void ProcessFile(string Filename)
        {
            Document = new XmlDocument();

            _ns = new XmlNamespaceManager(Document.NameTable);
            _ns.AddNamespace("SI", "http://www.siemens.com/automation/Openness/SW/Interface/v3");
            _ns.AddNamespace("Graph", "http://www.siemens.com/automation/Openness/SW/NetworkSource/Graph/v4");

            //Load Xml File with fileName into memory
            Document.Load(Filename);
            //get root node of xml file
            RootNode = Document.DocumentElement;


            var listOfNetworks = RootNode.SelectNodes("//SW.Blocks.CompileUnit");


            if (listOfNetworks != null)
            {

                foreach (XmlNode network in listOfNetworks)
                {

                    XmlNode Sequence = network.SelectSingleNode(".//Graph:Sequence", _ns);

                    XmlNodeList listOfConnections = network.SelectNodes(".//Graph:Connection", _ns);


                    // Init Step
                    var InitStepNumber = network.SelectSingleNode(".//Graph:Step[@Init='true']", _ns)
                        .Attributes["Number"].Value;
                    
                    BranchRefList.Clear();

                    var count = listOfConnections.Count;

                    var NextNodeNumber = "";

                    var NextNodeName = "";

                    
                    var currentIteration = 0;

                    var skipped = 0;

                    int altRow = 0;
                    int altColumn = 0;

                    int ActualAltBranchOut = 0;


                    int watchDog = 0;

                RestartIteration:

                    if (currentIteration > 0)
                    {

                        listOfConnections[currentIteration - 1].RemoveAll();
                        

                        currentIteration = 0;
                        skipped = 0;

                    }

                    foreach (XmlNode con in listOfConnections)
                    {

                        currentIteration++;

                        if (!con.HasChildNodes)
                        {
                            skipped++;

                            if (skipped == listOfConnections.Count)
                            {
                                MessageBox.Show("Hotovo");
                                goto Finish;
                            }

                            continue;
                        }


                        var ConnectionLinkType = con["LinkType"].InnerText;
                        

                        // Prvy krok + prva Trans
                        if (con.SelectSingleNode(".//Graph:StepRef[@Number='" + InitStepNumber + "']", _ns) != null && NextNodeNumber == "" && NextNodeName == "")
                        {
                            // Init Step
                            GetStep("StepRef", InitStepNumber, Sequence, ConnectionLinkType);


                            // Next Transition
                            XmlNode member = con.SelectSingleNode(".//Graph:StepRef[@Number='" + InitStepNumber + "']", _ns).ParentNode.NextSibling;

                            NextNodeName = member.SelectSingleNode("*").Name;

                            NextNodeNumber = member.SelectSingleNode("*").Attributes["Number"].Value;

                            GetStep(NextNodeName, NextNodeNumber, Sequence, ConnectionLinkType);

                            goto RestartIteration;
                        }


                        // Kazdy dalsi krok - hlavna vetva
                        if (con["NodeFrom"].SelectSingleNode(".//Graph:*", _ns).Name == NextNodeName && con["NodeFrom"].SelectSingleNode(".//Graph:*", _ns).Attributes["Number"].Value == NextNodeNumber)
                        {

                            if (con["NodeFrom"].SelectSingleNode(".//Graph:*", _ns).Name == "BranchRef")
                            {
                                if (con["NodeFrom"].SelectSingleNode(".//Graph:BranchRef", _ns).Attributes.GetNamedItem("Out") != null)
                                {
                                    if (Int32.Parse(con["NodeFrom"].SelectSingleNode(".//Graph:BranchRef", _ns).Attributes["Out"].Value) > 0)
                                        continue;
                                }
                            }
                            
                            NextNodeName = con["NodeTo"].SelectSingleNode(".//Graph:*", _ns).Name;

                            if (NextNodeName == "EndConnection")
                            {
                                NextNodeNumber = "";
                            }
                            else
                            {
                                NextNodeNumber = con["NodeTo"].SelectSingleNode(".//Graph:*", _ns).Attributes["Number"].Value;
                            }

                            GetStep(NextNodeName, NextNodeNumber, Sequence, ConnectionLinkType);

                            goto RestartIteration;

                        }
                        
                    }
                    

                    // najdi alternativnu vetvu
                    
                    currentIteration = 0;
                    skipped = 0;
                    NextNodeNumber = "";
                    NextNodeName = "";
                    altRow = 0;

                RestartIteration2:


                    if (currentIteration > 0)
                    {

                        listOfConnections[currentIteration - 1].RemoveAll();
                        currentIteration = 0;
                        skipped = 0;

                    }


                    if (NextNodeNumber =="" && NextNodeName =="")
                    {
                        foreach (XmlNode con in listOfConnections)
                        {
                            currentIteration++;

                            if (!con.HasChildNodes)
                            {
                                skipped++;

                                if (skipped == listOfConnections.Count)
                                {
                                    MessageBox.Show("Hotovo");
                                    goto Finish;
                                }

                                continue;
                            }

                            var ConnectionLinkType = con["LinkType"].InnerText;

                            if (con["NodeFrom"].SelectSingleNode(".//Graph:BranchRef", _ns) != null)
                            {


                                ActualAltBranchOut = Int32.Parse(con["NodeFrom"]
                                    .SelectSingleNode(".//Graph:BranchRef", _ns).Attributes["Out"].Value);

                                //if (con["NodeFrom"].SelectSingleNode(".//Graph:BranchRef", _ns).Attributes.GetNamedItem("Out") != null)
                                //{
                                //    if (Int32.Parse(con["NodeFrom"].SelectSingleNode(".//Graph:BranchRef", _ns).Attributes["Out"].Value) > 1)
                                //        continue;
                                //}


                                altRow = 0;

                                altColumn = AddColumn(altColumn);

                                

                                var nodefrom = con["NodeFrom"].SelectSingleNode(".//Graph:*", _ns).Name;

                                var nodefromnumber = con["NodeFrom"].SelectSingleNode(".//Graph:*", _ns).Attributes["Number"].Value;

                                GetAltStep(nodefrom, nodefromnumber, Sequence, ConnectionLinkType, ref altRow, ref altColumn);


                                NextNodeName = con["NodeTo"].SelectSingleNode(".//Graph:*", _ns).Name;

                                NextNodeNumber = con["NodeTo"].SelectSingleNode(".//Graph:*", _ns).Attributes["Number"].Value;


                                GetAltStep(NextNodeName, NextNodeNumber, Sequence, ConnectionLinkType, ref altRow, ref altColumn);


                                goto RestartIteration2;

                            }

                        }
                    }

                    currentIteration = 0;
                    skipped = 0;


                    foreach (XmlNode con in listOfConnections)
                    {

                        currentIteration++;

                        if (!con.HasChildNodes)
                        {
                            skipped++;

                            if (skipped == listOfConnections.Count)
                            {
                                MessageBox.Show("OK");
                                goto Finish;
                            }

                            continue;
                        }


                        var ConnectionLinkType = con["LinkType"].InnerText;


                        if (con["NodeFrom"].SelectSingleNode(".//Graph:*", _ns).Name == NextNodeName && con["NodeFrom"].SelectSingleNode(".//Graph:*", _ns).Attributes["Number"].Value == NextNodeNumber)
                        {
                           

                            NextNodeName = con["NodeTo"].SelectSingleNode(".//Graph:*", _ns).Name;

                            NextNodeNumber = con["NodeTo"].SelectSingleNode(".//Graph:*", _ns).Attributes["Number"].Value;


                            GetAltStep(NextNodeName, NextNodeNumber, Sequence, ConnectionLinkType, ref altRow, ref altColumn);


                            goto RestartIteration2;

                        }
                        
                    }


                    if (currentIteration == count)
                    {
                        watchDog++;
                        if (watchDog == 200)
                        {
                            MessageBox.Show("NOK");
                            goto Finish;
                        }

                        currentIteration = 0;
                        skipped = 0;
                        NextNodeNumber = "";
                        NextNodeName = "";
                        ActualAltBranchOut = 0;
                        goto RestartIteration2;
                    }
                    

                    Finish:
                    {

                    }

                    DgRefresh();

                }

            }

        }


        private void GetStep(string _NextNodeName, string _NextNodeNumber, XmlNode Sequence, string ConnectionLinkType)
        {
            var md5 = MD5.Create();

            StringBuilder text = new StringBuilder();


            if (_NextNodeName == "StepRef")
            {
                XmlNode member = Sequence.SelectSingleNode(".//Graph:Step[@Number='" + _NextNodeNumber + "']", _ns);

                var listOfOneStepActions = member.SelectNodes(".//Graph:Token", _ns);

                var StepName = Sequence.SelectSingleNode(".//Graph:Step[@Number='" + _NextNodeNumber + "']", _ns).Attributes["Name"].Value;

                int n = dataGridView1.Rows.Add();


                if (ConnectionLinkType == "Direct")
                {
                    text.AppendLine($"Step {_NextNodeNumber} : {StepName}");
                    dataGridView1.Rows[n].Cells[0].Style.BackColor = Color.LightGray;

                    foreach (XmlNode nodeOneStepAction in listOfOneStepActions)
                    {
                        var stepActionValue = nodeOneStepAction.Attributes["Text"].Value;

                        if (stepActionValue != "\n")
                        {

                            text.AppendLine($"Step Action: { stepActionValue}");
                            
                        }

                    }

                }
                else if(ConnectionLinkType == "Jump")
                {
                    text.AppendLine($" Jump to -> Step {_NextNodeNumber} : {StepName}");
                    dataGridView1.Rows[n].Cells[0].Style.BackColor = Color.Orange;
                }

                DGVwriteValue(n, 0, text.ToString());

                text.Clear();


            }
            else if (_NextNodeName == "TransitionRef")
            {
                XmlNode member = Sequence.SelectSingleNode(".//Graph:Transition[@Number='" + _NextNodeNumber + "']", _ns);

                var TansName = Sequence.SelectSingleNode(".//Graph:Transition[@Number='" + _NextNodeNumber + "']", _ns).Attributes["Name"].Value;

                int n = dataGridView1.Rows.Add();
                
                DGVwriteValue(n, 0, $"Transition {_NextNodeNumber} : {TansName}");

            }
            else if (_NextNodeName == "BranchRef")
            {
                XmlNode member = Sequence.SelectSingleNode(".//Graph:Branch[@Number='" + _NextNodeNumber + "']", _ns);

                var BranchName = Sequence.SelectSingleNode(".//Graph:Branch[@Number='" + _NextNodeNumber + "']", _ns).Attributes["Type"].Value;


                int n = dataGridView1.Rows.Add();

                DGVwriteValue(n, 0, $"Branch: {BranchName} Nr.: {_NextNodeNumber}");

                string pattern = @"Begin";
                Regex rg = new Regex(pattern);

                if (rg.IsMatch(BranchName))
                {
                    BranchRefList.Add((_NextNodeNumber, n, 0));
                }
                

                var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(BranchName + _NextNodeNumber));
                var color = Color.FromArgb(hash[0], hash[1], hash[2]);

                dataGridView1.Rows[n].Cells[0].Style.BackColor = color;


            }
            else if (_NextNodeName == "EndConnection")
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Style.BackColor = Color.Yellow;
                DGVwriteValue(n, 0, "EndConnection");
            }

        }
        

        private void GetAltStep(string _NextNodeName, string _NextNodeNumber, XmlNode Sequence, string ConnectionLinkType, ref int altRow, ref int altColumn)
        {

            var md5 = MD5.Create();

            if (dataGridView1.RowCount < altRow + 1)
            {
                //Add row
                altRow = dataGridView1.Rows.Add();
            }
            

            StringBuilder text = new StringBuilder();


            if (_NextNodeName == "StepRef")
            {
                XmlNode member = Sequence.SelectSingleNode(".//Graph:Step[@Number='" + _NextNodeNumber + "']", _ns);

                var listOfOneStepActions = member.SelectNodes(".//Graph:Token", _ns);

                var StepName = Sequence.SelectSingleNode(".//Graph:Step[@Number='" + _NextNodeNumber + "']", _ns).Attributes["Name"].Value;
                

                if (ConnectionLinkType == "Direct")
                {
                    text.AppendLine($"Step {_NextNodeNumber} : {StepName}");
                    dataGridView1.Rows[altRow].Cells[altColumn].Style.BackColor = Color.LightGray;

                    foreach (XmlNode nodeOneStepAction in listOfOneStepActions)
                    {
                        var stepActionValue = nodeOneStepAction.Attributes["Text"].Value;

                        if (stepActionValue != "\n")
                        {

                            text.AppendLine($"Step Action: { stepActionValue}");

                        }

                    }

                }
                else if (ConnectionLinkType == "Jump")
                {
                    text.AppendLine($" Jump to -> Step {_NextNodeNumber} : {StepName}");
                    dataGridView1.Rows[altRow].Cells[altColumn].Style.BackColor = Color.Orange;
                }

                DGVwriteValue(altRow, altColumn, text.ToString());
                text.Clear();


            }
            else if (_NextNodeName == "TransitionRef")
            {
                XmlNode member = Sequence.SelectSingleNode(".//Graph:Transition[@Number='" + _NextNodeNumber + "']", _ns);

                var TansName = Sequence.SelectSingleNode(".//Graph:Transition[@Number='" + _NextNodeNumber + "']", _ns).Attributes["Name"].Value;
                
                DGVwriteValue(altRow, altColumn, $"Transition {_NextNodeNumber} : {TansName}");

            }
            else if (_NextNodeName == "BranchRef")
            {
                XmlNode member = Sequence.SelectSingleNode(".//Graph:Branch[@Number='" + _NextNodeNumber + "']", _ns);

                var BranchName = Sequence.SelectSingleNode(".//Graph:Branch[@Number='" + _NextNodeNumber + "']", _ns).Attributes["Type"].Value;


                if (BranchRefList.Any(a => a.Key == _NextNodeNumber) && RowAligment)
                {
                    altRow = BranchRefList.Where(w => w.Key == _NextNodeNumber).Select(s => s.row).FirstOrDefault();
                }

                if (!BranchRefList.Any(a => a.Key == _NextNodeNumber) && RowAligment)
                {
                    string pattern = @"Begin";
                    Regex rg = new Regex(pattern);

                    if (rg.IsMatch(BranchName))
                    {
                        BranchRefList.Add((_NextNodeNumber, altRow, altColumn));
                    }
                }

                DGVwriteValue(altRow, altColumn, $"Branch: {BranchName} Nr.: {_NextNodeNumber}");

                var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(BranchName + _NextNodeNumber));
                var color = Color.FromArgb(hash[0], hash[1], hash[2]);

                dataGridView1.Rows[altRow].Cells[altColumn].Style.BackColor = color;
                
            }

            altRow++;

        }


        private int AddColumn(int altColumn)
        {
            altColumn = 2 + altColumn;

            dataGridView1.Columns.Add($"newColumnName{altColumn - 1}", "");
            dataGridView1.Columns[altColumn - 1].Width = 10;


            dataGridView1.Columns.Add($"newColumnName{altColumn}", $"Alternative {altColumn / 2}");
            dataGridView1.Columns[altColumn].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView1.AutoResizeRow(altColumn, DataGridViewAutoSizeRowMode.AllCells);

            return altColumn;

        }


        private void DgRefresh()
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            //dataGridView1.AutoResizeColumns();
            //dataGridView1.AutoResizeRows();
            dataGridView1.Refresh();
        }


        private void DGVwriteValue(int row, int column, string text)
        {
            if (dataGridView1.Rows[row].Cells[column].Value == null)
            {
                dataGridView1.Rows[row].Cells[column].Value = text;
            }
            else
            {
                MessageBox.Show($"Error by write value! \n Cell {row},{column} > have value : {dataGridView1.Rows[row].Cells[column].Value} \n string for write is : {text}");
            }
        }


        void ExportToExcel()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                Excel.Application xcelApp = new Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            xcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                            if (dataGridView1.Rows[i].Cells[j].Style.BackColor.Name != "0")
                            {
                                xcelApp.Cells[i + 2, j + 1].Interior.Color = ColorTranslator.ToWin32(dataGridView1.Rows[i].Cells[j].Style.BackColor);
                            }

                            if (dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 6) == "Branch")
                            {
                                xcelApp.Cells[i + 2, j + 1].Font.Color = ColorTranslator.ToWin32(Color.White);
                            }


                        }
                        else
                        {
                            xcelApp.Cells[i + 2, j + 1] = "";
                        }

                    }
                }
                xcelApp.Columns.AutoFit();
                xcelApp.Visible = true;
                Marshal.ReleaseComObject(xcelApp);

            }
        }


        void PrintToPdf()
        {
            DGVPrinter printer = new DGVPrinter();

            printer.Title = "Schrittkete : " + FileinfoName;

            //printer.SubTitle = "";

            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit |

                                          StringFormatFlags.NoClip;

            printer.PageNumbers = true;

            printer.PageNumberInHeader = false;

            printer.PorportionalColumns = true;

            printer.HeaderCellAlignment = StringAlignment.Near;

            //printer.Footer = "";

            printer.FooterSpacing = 7;


            printer.PrintDataGridView(dataGridView1);
        }


        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            var openFileDialog = new OpenFileDialog();

            //openFileDialog.InitialDirectory = @"D:\VS repos\TIA Graph";
            openFileDialog.InitialDirectory = @"D:\Projekty\28.Audi PPE_EA690\_temp2\AS_";
            openFileDialog.Filter = "TIA XML (*.xml)|*.xml";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                dataGridView1.Rows.Clear();

                while (dataGridView1.Columns.Count > 1)
                {
                    dataGridView1.Columns.RemoveAt(dataGridView1.Columns.Count - 1);
                }

                dataGridView1.Refresh();


                string Path = openFileDialog.FileName;
                System.IO.FileInfo info = new System.IO.FileInfo(Path);
                FileinfoName = info.Name;
                this.Text = "File : " + FileinfoName;
                
                ProcessFile(openFileDialog.FileName);
                
            }
        }


        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintToPdf();
        }


        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        
        private void autoResizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoResizeRows();
            dataGridView1.Refresh();
        }


        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

            toolStripMenuItem1.Checked = !toolStripMenuItem1.Checked;

            RowAligment = Properties.Settings.Default.RowAligment = toolStripMenuItem1.Checked;

        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
            
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            RowAligment = toolStripMenuItem1.Checked = Properties.Settings.Default.RowAligment;
        }
    }
}