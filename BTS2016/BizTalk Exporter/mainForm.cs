using Microsoft.BizTalk.ExplorerOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace BizTalk_Exporter
{
    /// <summary>
    /// This code is protected by the laws of God.
    /// Because God help you if you understand this or how it works.
    /// Also, be kind, rewind. Share the knowledge you've gained and give feedback to those
    /// that were kind enough to share this with you.
    /// </summary>
    public partial class mainForm : Form
    {
        //Global Vars
        public string outputPath = @"c:\BTSExporter\";
        public static string excelFile = "";
        XmlHelper helper = new XmlHelper();
        public mainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                statusLbl.Text = "oLading Resources...";
                System.Windows.Forms.Application.DoEvents();
                envCombo.SelectedIndex = 0;
                LoadResources();
                statusLbl.Text = "Resources loaded. Waiting...";
            }
            catch (Exception ex)
            {
                actionlbl.Text = ex.Message;
            }
            #region Rd btns
            asmApprdBtn.Select();
            #endregion
        }

        #region Applications - This little piggy went to the market
        private void addbtn_Click(object sender, EventArgs e)
        {
            if (applicationsTreeView.SelectedNode == null)
                return;
            //Adds selected Item to list if new
            if (applicationsTreeView.SelectedNode.Parent == null)
            {
                if (!appsExportListBox.Items.Contains(applicationsTreeView.SelectedNode.Text))
                    appsExportListBox.Items.Add(applicationsTreeView.SelectedNode.Text);
            }
        }
        private void removeBtn_Click(object sender, EventArgs e)
        {//Removes selected Item from list
            appsExportListBox.Items.Remove(appsExportListBox.SelectedItem);
        }
        private void clearBtn_Click(object sender, EventArgs e)
        {//Clears list
            appsExportListBox.Items.Clear();
            ControlsReset();
        }
        private void exportAppsBtn_Click(object sender, EventArgs e)
        {
            try
            {
                setOutputBtn_Click(sender, e);
                statusLbl.Text = "Export started";
                System.Windows.Forms.Application.DoEvents();
                foreach (string app in appsExportListBox.Items)
                {
                    //Powershell call to generate bindings file for selected App.
                    BtsTask(null, "BTSTask.exe", @"ExportBindings /Destination:""" + outputPath + @"""\\" + app.Replace(' ', '_') + ".BindingInfo.xml /ApplicationName:" +
                                            (app.Contains(" ") ? "\"" + app + "\"" : app), 2, "");
                }
            }
            catch (Exception ex)
            {
                statusLbl.Text = "Error";
                actionlbl.Text = ex.Message;
            }
        }

        #endregion Applications

        #region Assemblies - This little piggy stayed home
        private void addAsmBtn_Click(object sender, EventArgs e)
        {
            bool foundParent = false;
            TreeNode parentAsm = new TreeNode();
            //Adds selected Item to list if new
            //If node is null, or already in tree or parent in tree, exit
            if (assembliesTreeView.SelectedNode == null)
                return;

            #region Loop tree to find match
            foreach (TreeNode node in asmExportTreeView.Nodes)
            {   //loop child nodes
                if (new XmlHelper().FindSubNodeInNode(node, assembliesTreeView.SelectedNode.Text))
                    return;
                if (assembliesTreeView.SelectedNode.Parent != null)
                {
                    if (node.Text == assembliesTreeView.SelectedNode.Parent.Text)
                    {
                        foundParent = true;
                        parentAsm = node;
                    }
                }
            }
            #endregion

            //If parent, add all child nodes
            if (assembliesTreeView.SelectedNode.Parent == null)
            {
                parentAsm = helper.FindParent(asmExportTreeView, assembliesTreeView.SelectedNode.Text);
                if (parentAsm != null)
                {
                    foreach (TreeNode subnodeParent in assembliesTreeView.SelectedNode.Nodes)
                    {
                        if (helper.FindSubNodeInNode(parentAsm, subnodeParent.Text))
                            continue;
                        else
                            parentAsm.Nodes.Add(subnodeParent.Text);
                    }
                }
                else
                    helper.AddRootAndChild(assembliesTreeView, asmExportTreeView);
            }
            else if (assembliesTreeView.SelectedNode.Parent != null)
            {
                //if parent node is not in in tree, add it
                if (!foundParent)
                {
                    var subnode = asmExportTreeView.Nodes.Add(assembliesTreeView.SelectedNode.Parent.Text);
                    subnode.Nodes.Add(assembliesTreeView.SelectedNode.Text);
                    subnode.Expand();
                }
                else
                    parentAsm.Nodes.Add(assembliesTreeView.SelectedNode.Text);
            }
        }
        private void removeAsmBtn_Click(object sender, EventArgs e)
        {   //Removes selected Item from list
            if (asmExportTreeView.Nodes.Count > 0 && asmExportTreeView.SelectedNode != null)
                asmExportTreeView.Nodes.Remove(asmExportTreeView.SelectedNode);
        }
        private void clearAsmBtn_Click(object sender, EventArgs e)
        {   //Clears list
            asmExportTreeView.Nodes.Clear();
            ControlsReset();
        }
        private void exportAsmbtn_Click(object sender, EventArgs e)
        {
            try
            {
                setOutputBtn_Click(sender, e);
                statusLbl.Text = "Export started"; System.Windows.Forms.Application.DoEvents();

                //method to export joined by app or in separeted files.
                if (asmApprdBtn.Checked)
                {
                    //export by Application
                    foreach (TreeNode application in asmExportTreeView.Nodes)
                    {
                        BtsTask(null, "BTSTask.exe", "ExportBindings /Destination:\"" + outputPath + application.Text +
                            ".BindingInfo.xml\" /ApplicationName:" + application.Text, 2, "");
                    }
                }
                else if (asmAsmrdBtn.Checked)
                {
                    foreach (TreeNode appNode in asmExportTreeView.Nodes)
                    {
                        foreach (TreeNode asm in appNode.Nodes)
                        {
                            //Powershell call to generate bindings file for selected App.
                            BtsTask(null, "BTSTask.exe", "ExportBindings /Destination:\"" + outputPath + "\"\\" +
                                (includeAppNameCkbox.Checked ? appNode.Text + "." : "") +
                                asm.Text.Split(',')[0]
                                + ".BindingInfo.xml /AssemblyName:\"" + asm.Text + "\"", 2, "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                statusLbl.Text = "Error!";
                actionLbl2.Text = ex.Message;
            }
        }

        #endregion Assemblies

        #region Receive ports - this little piggy had roast beef
        private void addReceiveBtn_Click(object sender, EventArgs e)
        {
            //Adds selected Item to list if new
            bool foundParent = false;
            TreeNode parentRcv = new TreeNode();
            //Adds selected Item to list if new
            //If node is null, or already in tree or parent in tree, exit
            if (receiveTreeView.SelectedNode == null)
                return;
            #region Loop tree to find match
            foreach (TreeNode node in rcvExportTreeView.Nodes)
            {
                if (helper.FindSubNodeInNode(node, receiveTreeView.SelectedNode.Text))
                    return;
                if (receiveTreeView.SelectedNode.Parent != null)
                {
                    if (node.Text == receiveTreeView.SelectedNode.Parent.Text)
                    {
                        foundParent = true;
                        parentRcv = node;
                    }
                }
            }
            #endregion

            //If parent, add all child nodes
            if (receiveTreeView.SelectedNode.Parent == null)
            {
                parentRcv = helper.FindParent(rcvExportTreeView, receiveTreeView.SelectedNode.Text);
                if (parentRcv != null)
                {
                    foreach (TreeNode subnodeParent in receiveTreeView.SelectedNode.Nodes)
                    {
                        if (helper.FindSubNodeInNode(parentRcv, subnodeParent.Text))
                            continue;
                        else
                            parentRcv.Nodes.Add(subnodeParent.Text);
                    }
                }
                else
                    helper.AddRootAndChild(receiveTreeView, rcvExportTreeView);
            }
            else if (receiveTreeView.SelectedNode.Parent != null)
            {
                //if parent node is not in in tree, add it
                if (!foundParent)
                {
                    var subnode = rcvExportTreeView.Nodes.Add(receiveTreeView.SelectedNode.Parent.Text);
                    subnode.Nodes.Add(receiveTreeView.SelectedNode.Text);
                    subnode.Expand();
                }
                else
                    parentRcv.Nodes.Add(receiveTreeView.SelectedNode.Text);
            }
        }
        private void removeReceiveBtn_Click(object sender, EventArgs e)
        {    //Removes selected Item from list
            if (rcvExportTreeView.Nodes.Count > 0 && rcvExportTreeView.SelectedNode != null)
                rcvExportTreeView.Nodes.Remove(rcvExportTreeView.SelectedNode);
        }
        private void clearReceiveBtn_Click(object sender, EventArgs e)
        {//Clears list
            rcvExportTreeView.Nodes.Clear();
            ControlsReset();
        }
        private void exportReceivebtn_Click(object sender, EventArgs e)
        {
            try
            {
                #region Excel file and SaveFile dialogs check
                if (string.IsNullOrEmpty(excelFile) && !string.IsNullOrEmpty((string)envCombo.SelectedItem))
                    excelBtn_Click(sender, e);
                if (saveFileDialog.ShowDialog() == DialogResult.Cancel)
                    return;
                #endregion

                statusLbl.Text = "Export started"; System.Windows.Forms.Application.DoEvents();
                string originalPath = saveFileDialog.FileName;

                #region Get all children
                List<string> children = new List<string>();
                foreach (TreeNode node in rcvExportTreeView.Nodes)
                    helper.CollectChildren(node, ref children);
                #endregion

                if (rcvAppRbtn.Checked)
                {
                    foreach (TreeNode application in rcvExportTreeView.Nodes)
                    {
                        BtsTask(children, "BTSTask.exe", "ExportBindings /Destination:" + saveFileDialog.FileName + " /ApplicationName:"
                            + (application.Text.Contains(" ") ? "\"" + application.Text + "\"" :
                            application.Text), 0,
                            saveFileDialog.FileName);
                    }
                }
                else if (rcvPortRbtn.Checked)
                {
                    foreach (TreeNode appNode in rcvExportTreeView.Nodes)
                    {
                        #region Handle Original names and add App name
                        saveFileDialog.FileName = originalPath;
                        if (includeAppRcvCb.Checked)
                        {
                            FileInfo fi = new FileInfo(saveFileDialog.FileName);
                            //Add AppName
                            saveFileDialog.FileName = fi.Directory + "\\" + appNode.Text.Replace(" ", "_") + "." + fi.Name;
                        }
                        string appOriginal = saveFileDialog.FileName;
                        #endregion

                        //I create a file with every port in the Application
                        BtsTask(children, "BTSTask.exe", "ExportBindings /Destination:" + saveFileDialog.FileName + " /ApplicationName:"
                                + (appNode.Text.Contains(" ") ? "\"" + appNode.Text + "\"" :
                                appNode.Text), 0, saveFileDialog.FileName);

                        //Then, for each Port I'll create a new file where I remove the other ports, based on the global.
                        foreach (TreeNode rcv in appNode.Nodes)
                        {
                            //reset for each Port
                            saveFileDialog.FileName = appOriginal;
                            //split the path to an array
                            string[] pathArray = saveFileDialog.FileName.Split('\\');
                            saveFileDialog.FileName = ""; //reset the path
                            for (int i = 0; i < pathArray.Length; i++)
                            {
                                //if it's the last, add the port name to it.
                                if (i == pathArray.Length - 1)
                                    saveFileDialog.FileName += rcv.Text + "." + pathArray[i];
                                else
                                    saveFileDialog.FileName += pathArray[i] + "\\";
                            }
                            new XmlHelper().RemoveExcessBindings(appOriginal, saveFileDialog.FileName, true, new List<string>() { rcv.Text });
                        }
                    }
                }
                //If Env is QA or PROD, replace Configs read in file
                if ((string)envCombo.SelectedItem == "QA" || (string)envCombo.SelectedItem == "PROD")
                {
                    var ports = new ExcelOperations().ReadExcelFile(envCombo.SelectedItem.ToString(), excelFile);
                    helper.ReplaceEnvironmentBindings(ports, saveFileDialog.FileName, "Receive");
                }
            }
            catch (Exception ex)
            {
                statusLbl.Text = "Error";
                actionLbl3.Text = ex.Message;
            }
        }
        #endregion

        #region Send Ports - This little piggy had none.
        private void addSendBtn_Click(object sender, EventArgs e)
        {//Adds selected Item to list if new
            bool foundParent = false;
            TreeNode parentSend = new TreeNode();
            //Adds selected Item to list if new //If node is null, or already in tree or parent in tree, exit
            if (sendTreeView.SelectedNode == null)
                return;
            #region Loop tree to find match
            foreach (TreeNode node in sendExportTreeView.Nodes)
            {
                if (helper.FindSubNodeInNode(node, sendTreeView.SelectedNode.Text))
                    return;
                if (sendTreeView.SelectedNode.Parent != null)
                {
                    if (node.Text == sendTreeView.SelectedNode.Parent.Text)
                    {
                        foundParent = true;
                        parentSend = node;
                    }
                }
            }
            #endregion
            //If parent and doesnt exist, add all child nodes
            if (sendTreeView.SelectedNode.Parent == null)
            {
                parentSend = helper.FindParent(sendExportTreeView, sendTreeView.SelectedNode.Text);
                if (parentSend != null)
                {
                    foreach (TreeNode subnodeParent in sendTreeView.SelectedNode.Nodes)
                    {
                        if (helper.FindSubNodeInNode(parentSend, subnodeParent.Text))
                            continue;
                        else
                            parentSend.Nodes.Add(subnodeParent.Text);
                    }
                }
                else
                    helper.AddRootAndChild(sendTreeView, sendExportTreeView);
            }
            else if (sendTreeView.SelectedNode.Parent != null)
            {
                //if parent node is not in in tree, add it
                if (!foundParent)
                {
                    var subnode = sendExportTreeView.Nodes.Add(sendTreeView.SelectedNode.Parent.Text);
                    subnode.Nodes.Add(sendTreeView.SelectedNode.Text);
                    subnode.Expand();
                }
                else
                    parentSend.Nodes.Add(sendTreeView.SelectedNode.Text);
            }
        }

        private void removeSendBtn_Click(object sender, EventArgs e)
        {    //Removes selected Item from list
            if (sendExportTreeView.Nodes.Count > 0 && sendExportTreeView.SelectedNode != null)
                sendExportTreeView.Nodes.Remove(sendExportTreeView.SelectedNode);
        }

        private void clearSendBtn_Click(object sender, EventArgs e)
        {//Clears list
            sendExportTreeView.Nodes.Clear();
            ControlsReset();
        }

        private void exportSendbtn_Click(object sender, EventArgs e)
        {
            try
            {
                #region Excel file and SaveFile dialogs check
                if (string.IsNullOrEmpty(excelFile) && !string.IsNullOrEmpty((string)envCombo.SelectedItem))
                    excelBtn_Click(sender, e);
                if (saveFileDialog.ShowDialog() == DialogResult.Cancel)
                    return;
                #endregion

                statusLbl.Text = "Export started"; System.Windows.Forms.Application.DoEvents();
                string originalPath = saveFileDialog.FileName;
                #region Get all children
                List<string> children = new List<string>();
                foreach (TreeNode node in sendExportTreeView.Nodes)
                    helper.CollectChildren(node, ref children);
                #endregion

                if (sendAppRbtn.Checked)
                {
                    foreach (TreeNode application in sendExportTreeView.Nodes)
                    {
                        //I create a file with every port in the Application
                        BtsTask(children, "BTSTask.exe", "ExportBindings /Destination:\"" + saveFileDialog.FileName + "\" /ApplicationName:"
                      + (application.Text.Contains(" ") ? "\"" + application.Text + "\"" :
                      application.Text), 1, saveFileDialog.FileName);
                    }
                }
                else if (sendPortRbtn.Checked)
                {
                    foreach (TreeNode appNode in sendExportTreeView.Nodes)
                    {
                        #region Handle Original names and add App name
                        saveFileDialog.FileName = originalPath;
                        if (includeAppSendCb.Checked)
                        {
                            FileInfo fi = new FileInfo(saveFileDialog.FileName);
                            //Add AppName
                            saveFileDialog.FileName = fi.Directory + "\\" + appNode.Text.Replace(" ", "_") + "." + fi.Name;
                        }
                        string appOriginal = saveFileDialog.FileName;
                        #endregion

                        BtsTask(children, "BTSTask.exe", "ExportBindings /Destination:" + saveFileDialog.FileName + " /ApplicationName:"
                      + (appNode.Text.Contains(" ") ? "\"" + appNode.Text + "\"" :
                      appNode.Text), 1, saveFileDialog.FileName);
                        #region  Then, for each Port I'll create a new file where I remove the other ports, based on the global.
                        foreach (TreeNode send in appNode.Nodes)
                        {
                            //reset for each Port
                            saveFileDialog.FileName = appOriginal;
                            //split the path to an array
                            string[] pathArray = saveFileDialog.FileName.Split('\\');
                            saveFileDialog.FileName = ""; //reset the path
                            for (int i = 0; i < pathArray.Length; i++)
                            {
                                //if it's the last, add the port name to it.
                                if (i == pathArray.Length - 1)
                                    saveFileDialog.FileName += send.Text + "." + pathArray[i];
                                else
                                    saveFileDialog.FileName += pathArray[i] + "\\";
                            }
                            new XmlHelper().RemoveExcessBindings(appOriginal, saveFileDialog.FileName, true, new List<string>() { send.Text });
                        }
                        #endregion
                    }
                }

                #region handle Environments
                //If Env is QA or PROD, replace Configs read in file
                if ((string)envCombo.SelectedItem == "QA" || (string)envCombo.SelectedItem == "PROD")
                {
                    var ports = new ExcelOperations().ReadExcelFile(envCombo.SelectedItem.ToString(), excelFile);
                    helper.ReplaceEnvironmentBindings(ports, saveFileDialog.FileName, "Send");
                }
                #endregion
            }
            catch (Exception ex)
            {
                actionLbl4.Text = ex.Message;
                statusLbl.Text = "Error";
            }
        }
        #endregion

        #region Controls lifecycle
        /// <summary>
        /// Sets the Output directory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        #region RadioButtons
        private void asmApprdBtn_CheckedChanged(object sender, EventArgs e)
        {
            if (asmAsmrdBtn.Checked)
            {
                asmApprdBtn.Checked = false;
                includeAppNameCkbox.Enabled = true;
            }
        }
        private void asmAsmrdBtn_CheckedChanged(object sender, EventArgs e)
        {
            if (asmApprdBtn.Checked)
            {
                asmAsmrdBtn.Checked = false;
                asmApprdBtn.Checked = true;
                includeAppNameCkbox.Enabled = false;
            }
        }
        private void rcvAppRbtn_CheckedChanged(object sender, EventArgs e)
        {
            if (rcvPortRbtn.Checked)
            {
                rcvAppRbtn.Checked = false;
                includeAppRcvCb.Enabled = true;
            }
        }
        private void rcvPortRbtn_CheckedChanged(object sender, EventArgs e)
        {
            if (rcvAppRbtn.Checked)
            {
                rcvPortRbtn.Checked = false;
                rcvAppRbtn.Checked = true;
                includeAppRcvCb.Enabled = false;
            }
        }
        private void sendAppRbtn_CheckedChanged(object sender, EventArgs e)
        {
            if (sendPortRbtn.Checked)
            {
                sendAppRbtn.Checked = false;
                includeAppSendCb.Enabled = true;
            }
        }
        private void sendPortRbtn_CheckedChanged(object sender, EventArgs e)
        {
            if (sendAppRbtn.Checked)
            {
                sendPortRbtn.Checked = false;
                sendAppRbtn.Checked = true;
                includeAppSendCb.Enabled = false;
            }
        }
        #endregion
        private void loadResourcesBtn_Click(object sender, EventArgs e)
        {
            LoadResources();
        }
        private void connStrBtn_Click(object sender, EventArgs e)
        {
            connectionForm f = new connectionForm();
            f.ShowDialog();
        }
        private void excelBtn_Click(object sender, EventArgs e)
        {
            openFileDialog.ShowDialog();
            excelFile = openFileDialog.FileName; //catches the returned file and path
        }
        private void setOutputBtn_Click(object sender, EventArgs e)
        {//Initial path
            folderBrowserDialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
            folderBrowserDialog.ShowDialog(); //shows dialog without possibility to click outside
            outputPath = folderBrowserDialog.SelectedPath; // sets the global var
        }
        private void openOutputBtn_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(outputPath))
                Process.Start(outputPath); //Opens Explorer with the path.
            else
                MessageBox.Show("Output folder doesn't exist yet.", "Warning!", MessageBoxButtons.OK);
        }
        private void ControlsReset()
        {
            actionlbl.ResetText();
            actionLbl2.ResetText();
            actionLbl3.ResetText();
            actionLbl4.ResetText();
            statusLbl.Text = "Waiting";
        }
        private void loadResourcesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadResources();
        }
        #endregion

        #region Helper Methods
        private int BtsTask(List<string> childNodes, string command, string args, int receiveSend, string path)
        {
            try
            {
                var proc = new Process()
                {
                    StartInfo = {
                        FileName = command,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        Arguments = args,
                        CreateNoWindow = false
                    }
                };

                proc.Start();
                string stdout = proc.StandardOutput.ReadToEnd();
                proc.WaitForExit();
                var exitCode = proc.ExitCode;
                proc.Close();
                if (exitCode != 0)
                    actionlbl.Text = actionLbl2.Text = actionLbl3.Text = actionLbl4.Text = stdout;
                else
                {
                    actionlbl.Text = actionLbl2.Text = actionLbl3.Text = actionLbl4.Text = "Export successful";

                    if (receiveSend == 0)//receive
                        helper.RemoveExcessBindings(path, true, childNodes);
                    else if (receiveSend == 1)//send
                        helper.RemoveExcessBindings(path, false, childNodes);

                    statusLbl.Text = "Operation Completed";
                }
                return exitCode;
            }
            catch (Exception ex)
            { throw ex; }
        }
        private void LoadResources()
        {
            using (var expl = new BtsCatalogExplorer() { ConnectionString = ConfigurationManager.AppSettings["connString"].ToString() })
            {
                ApplicationCollection orqs = expl.Applications;
                #region Load TreeViews
                #region  Clear trees
                applicationsTreeView.Nodes.Clear();
                assembliesTreeView.Nodes.Clear();
                receiveTreeView.Nodes.Clear();
                sendTreeView.Nodes.Clear();
                #endregion

                foreach (Microsoft.BizTalk.ExplorerOM.Application app in orqs)
                {
                    #region  Creates a root node for each App in the corresponding Tree
                    var appNode = applicationsTreeView.Nodes.Add(app.Name);
                    var assembliesNode = assembliesTreeView.Nodes.Add(app.Name);
                    var receiveNode = receiveTreeView.Nodes.Add(app.Name);
                    var sendNode = sendTreeView.Nodes.Add(app.Name);
                    #endregion

                    #region Assemblies
                    if (app.Assemblies.Count > 0)
                    {
                        var subNode = appNode.Nodes.Add("Assemblies");
                        foreach (BtsAssembly assembly in app.Assemblies)
                        {
                            subNode.Nodes.Add(assembly.Name, assembly.DisplayName);
                            assembliesNode.Nodes.Add(assembly.Name, assembly.DisplayName);
                        }
                    }
                    #endregion
                    #region ReceivePorts
                    if (app.ReceivePorts.Count > 0)
                    {
                        var subNode = appNode.Nodes.Add("Receive Ports");
                        foreach (ReceivePort port in app.ReceivePorts)
                        {
                            subNode.Nodes.Add(port.Name);
                            receiveNode.Nodes.Add(port.Name);
                        }
                    }
                    #endregion
                    #region SendPorts
                    if (app.SendPorts.Count > 0)
                    {
                        var subNode = appNode.Nodes.Add("Send Ports");
                        foreach (SendPort port in app.SendPorts)
                        {
                            subNode.Nodes.Add(port.Name);
                            sendNode.Nodes.Add(port.Name);
                        }
                    }
                    #endregion
                }
                #endregion
            }
        }

        #endregion

        private void exitBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
