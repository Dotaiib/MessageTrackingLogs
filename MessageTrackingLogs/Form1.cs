using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace MessageTrackingLogs
{
    public partial class Form1 : Form
    {
        //SqlConnection cn = new SqlConnection(@"data source = wb025 ; initial catalog = WBExchange ; integrated security = false; User ID = Paie; Password = P@ie@mdesk0823;");
        SqlConnection cn = new SqlConnection(@"data source = wb025 ; initial catalog = WBExchange ; integrated security = true;");


        public Form1()
        {
            InitializeComponent(); TextboxScript();
        }

        private void TextboxScript()
        {
            DateTime dateTimeNow = DateTime.Now;
            DateTime datePartOnly = dateTimeNow.Date;
            TimeSpan TSDay = new TimeSpan(1, 0, 0, 0);
            TimeSpan TSminute = new TimeSpan(0, 1, 0);
            DateTime DTStart = datePartOnly.Subtract(TSDay);
            DateTime DTEnd = datePartOnly.Subtract(TSminute);
            //label1.Text = datePartOnly.ToString("MM/dd/yyyy HH:mm:ss");
            String TxtScript = "Get-MessageTrackingLog -Start \"" + DTStart.ToString("MM/dd/yyyy HH:mm:ss") + "\" -End \"" + DTEnd.ToString("MM/dd/yyyy HH:mm:ss") + "\" -resultsize unlimited";
            //String TxtScript = "Get-User -resultsize unlimited";
            textBox1.Text = TxtScript;
        }


        void ExecutePowerShellUsingRemotimg()
        {
            string userName = textBox2.Text;
            string password = textBox3.Text;
            var securestring = new SecureString();
            foreach (Char c in password)
            {
                securestring.AppendChar(c);
            }

            PSCredential creds = new PSCredential(userName, securestring);

            System.Uri uri = new Uri("http://Ex2016/powershell");

            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            Pipeline pipeline = runspace.CreatePipeline();


            string serverFqdn = "EX2016.warnecke.int";
            pipeline.Commands.AddScript("Set-ExecutionPolicy RemoteSigned");
            pipeline.Commands.AddScript(string.Format("$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "+ uri +" -Authentication Kerberos", serverFqdn));
            pipeline.Commands.AddScript("Import-PSSession $Session");
            pipeline.Commands.AddScript("Get-User");
            //pipeline.Commands.Add("Out-String");
            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();

            StringBuilder sb = new StringBuilder();

            if (pipeline.Error != null && pipeline.Error.Count > 0)
            {
                // Read errors
                //succeeded = false;
                Collection<object> errors = pipeline.Error.ReadToEnd();
                foreach (object error in errors)
                    sb.Append(error.ToString());
            }
            else
            {
                // Read output
                /*foreach (PSObject obj in results)
                    sb.Append(obj.ToString());*/

                ArrayList objects = new ArrayList();
                objects.AddRange(results);

                dataGridView1.DataSource = objects;
            }

            runspace.Dispose();
            pipeline.Dispose();
        }

        private void CreateExchange()
        {
            string userName = textBox2.Text;
            string password = textBox3.Text;
            var securestring = new SecureString();
            foreach (Char c in password)
            {
                securestring.AppendChar(c);
            }

            PSCredential creds = new PSCredential(userName, securestring);

            //System.Uri uri = new Uri("http://Ex2016/powershell?serializationLevel=Full");
            //System.Uri uri = new Uri("http://Ex2016/powershell");

            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            //Pipeline pip = runspace.CreatePipeline();
            PowerShell pip = PowerShell.Create();


            pip.Commands.AddScript(string.Format("$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://Ex2016/powershell/ -Credential " + creds + ""));
            //pip.Commands.AddScript(string.Format("$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://Ex2016/powershell/ -Authentication Kerberos"));
            pip.Commands.AddScript("Import-PSSession $Session");
            pip.Commands.AddScript("Get-User");
            //pip.Commands.Add("Out-String");
            Collection<PSSession> result = pip.Invoke<PSSession>();
            runspace.Close();


            ArrayList objects = new ArrayList();
            objects.AddRange(result);

            dataGridView1.DataSource = objects;
        }

        private void myttst()
        {
            string userName = textBox2.Text;
            string password = textBox3.Text;
            var securestring = new SecureString();
            foreach (Char c in password)
            {
                securestring.AppendChar(c);
            }

            PSCredential creds = new PSCredential(userName, securestring);

            //System.Uri uri = new Uri("http://Ex2016/powershell?serializationLevel=Full");
            System.Uri uri = new Uri("http://Ex2016/powershell");

            Runspace runspace = RunspaceFactory.CreateRunspace();

            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();
            command.AddCommand("New-PSSession");
            command.AddParameter("ConfigurationName", "Microsoft.Exchange");
            command.AddParameter("ConnectionUri", uri);
            command.AddParameter("Credential", creds);
            command.AddParameter("Authentication", "Default");
            powershell.Commands = command;
            runspace.Open(); 
            powershell.Runspace = runspace;
            Collection<PSSession> result = powershell.Invoke<PSSession>();

            // fade
            powershell = PowerShell.Create();
            command = new PSCommand();
            command.AddCommand("Set-variable");
            command.AddParameter("Name", "ra");
            command.AddParameter("Value", result[0]);
            powershell.Commands = command;
            powershell.Runspace = runspace;
            powershell.Invoke();

            powershell = PowerShell.Create();
            command = new PSCommand();
            command.AddScript("Import-PSSession -Session $ra");
            powershell.Commands = command;
            powershell.Runspace = runspace;
            powershell.Invoke();

            powershell = PowerShell.Create();
            command = new PSCommand();
            command.AddScript("Get-User");
            powershell.Commands = command;
            powershell.Runspace = runspace;
            Collection<PSObject> results = powershell.Invoke();
            //fade

            ArrayList objects = new ArrayList();
            objects.AddRange(results);

            dataGridView1.DataSource = objects;

        }

        private void CreateForm()
        {
            // Create a PowerShell object. Creating this object takes care of 
            // building all of the other data structures needed to run the command.
            using (PowerShell powershell = PowerShell.Create())
            {
                powershell.AddCommand(textBox1.Text).AddCommand("sort-object");
                if (Runspace.DefaultRunspace == null)
                {
                    Runspace.DefaultRunspace = powershell.Runspace;
                }

                Collection<PSObject> results = powershell.Invoke();

                // The generic collection needs to be re-wrapped in an ArrayList
                // for data-binding to work.
                ArrayList objects = new ArrayList();
                objects.AddRange(results);

                // The DataGridView will use the PSObjectTypeDescriptor type
                // to retrieve the properties.
                dataGridView1.DataSource = objects;
            }
        }

        private void RT()
        {
            string userName = textBox2.Text;
            string password = textBox3.Text;
            var securestring = new SecureString();
            foreach (Char c in password)
            {
                securestring.AppendChar(c);
            }
            System.Uri uri = new Uri("http://Ex2016/powershell");

            PSCredential credential = new PSCredential(userName, securestring);    
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(uri,"http://EX2016/powershell/Microsoft.Exchange",credential);

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                runspace.Open();
                using (PowerShell powershell = PowerShell.Create())
                {
                    powershell.Runspace = runspace;
                    //Create the command and add a parameter
                    powershell.AddScript("Set-ExecutionPolicy RemoteSigned");
                    powershell.AddCommand("Get-Mailbox");
                    powershell.AddParameter("RecipientTypeDetails", "UserMailbox");
                    //Invoke the command and store the results in a PSObject collection
                    Collection<PSObject> results = powershell.Invoke();
                    //Iterate through the results and write the DisplayName and PrimarySMTP
                    //address for each mailbox
                    ArrayList objects = new ArrayList();
                    objects.AddRange(results);
                    // The DataGridView will use the PSObjectTypeDescriptor type
                    // to retrieve the properties.
                    dataGridView1.DataSource = objects;
                }
            }
        }


        private void CallingPSCmdlet()
        {
            string userName = textBox2.Text;
            string password = textBox3.Text;
            var securestring = new SecureString();
            foreach (Char c in password)
            {
                securestring.AppendChar(c);
            }

            PSCredential creds = new PSCredential(userName, securestring);
            System.Uri uri = new Uri("http://Ex2016/powershell?serializationLevel=Full");

            Runspace runspace = RunspaceFactory.CreateRunspace();

            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();
            command.AddCommand("New-PSSession");
            command.AddParameter("ConfigurationName", "Microsoft.Exchange");
            command.AddParameter("ConnectionUri", uri);
            command.AddParameter("Credential", creds);
            command.AddParameter("Authentication", "Default");
            PSSessionOption sessionOption = new PSSessionOption();
            sessionOption.SkipCACheck = true;
            sessionOption.SkipCNCheck = true;
            sessionOption.SkipRevocationCheck = true;
            command.AddParameter("SessionOption", sessionOption);

            powershell.Commands = command;

            try
            {
                // open the remote runspace
                runspace.Open();

                // associate the runspace with powershell
                powershell.Runspace = runspace;

                // invoke the powershell to obtain the results
                Collection<PSSession> result = powershell.Invoke<PSSession>();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    MessageBox.Show("Exception: " + current.Exception.ToString());
                    MessageBox.Show("Inner Exception: " + current.Exception.InnerException);
                    //Console.WriteLine("Exception: " + current.Exception.ToString());
                    //Console.WriteLine("Inner Exception: " + current.Exception.InnerException);
                }

                if (result.Count != 1)
                    throw new Exception("Unexpected number of Remote Runspace connections returned.");

                // Set the runspace as a local variable on the runspace
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "ra");
                command.AddParameter("Value", result[0]);
                powershell.Commands = command;
                powershell.Runspace = runspace;

                powershell.Invoke();


                // First import the cmdlets in the current runspace (using Import-PSSession)
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Import-PSSession -Session $ra");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();


                // Now run Exchange PowerShell
                powershell = PowerShell.Create();
                command = new PSCommand();
                //command.AddScript("Get-ExchangeServer | where-object{$_.Name -like \"*MBX\"}");
                command.AddScript(""+textBox1.Text+"");
                powershell.Commands = command;
                powershell.Runspace = runspace;

                Collection<PSObject> results = new Collection<PSObject>();
                results = powershell.Invoke();

                foreach (PSObject PSresult in results)
                {
                    ArrayList objects = new ArrayList();
                    objects.AddRange(results);

                    dataGridView1.DataSource = objects;

                    //Console.WriteLine(PSresult.Properties["Name"].Value.ToString());
                }

            }

            finally
            {
                // dispose the runspace and enable garbage collection
                runspace.Dispose();
                runspace = null;

                // Finally dispose the powershell and set all variables to null to free
                // up any resources.
                powershell.Dispose();
                powershell = null;
            }


        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            CallingPSCmdlet();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView1.Rows.Count - 0; i++)
            {
                SqlCommand cmd = new SqlCommand("insert into WB_Exchange_MessageTrackingLog (MTLog_Timestamp,MTLog_Sender,MTLog_Recipients,MTLog_MsgSubject,MTLog_Source,MTLog_EventID,MTLog_SrcCntxt,MTLog_MsgId,MTLog_Save_Time) values (@MTLog_Timestamp,@MTLog_Sender,@MTLog_Recipients,@MTLog_MsgSubject,@MTLog_Source,@MTLog_EventID,@MTLog_SrcCntxt,@MTLog_MsgId,@MTLog_Save_Time) ", cn);
                cmd.Parameters.AddWithValue("@MTLog_Timestamp", dataGridView1.Rows[i].Cells[4].Value);
                cmd.Parameters.AddWithValue("@MTLog_Sender", dataGridView1.Rows[i].Cells[23].Value.ToString());
                cmd.Parameters.AddWithValue("@MTLog_Recipients", dataGridView1.Rows[i].Cells[16].Value.ToString());
                cmd.Parameters.AddWithValue("@MTLog_MsgSubject", dataGridView1.Rows[i].Cells[22].Value.ToString());
                cmd.Parameters.AddWithValue("@MTLog_Source", dataGridView1.Rows[i].Cells[11].Value.ToString());
                cmd.Parameters.AddWithValue("@MTLog_EventID", dataGridView1.Rows[i].Cells[12].Value.ToString());
                cmd.Parameters.AddWithValue("@MTLog_SrcCntxt", dataGridView1.Rows[i].Cells[9].Value.ToString());
                cmd.Parameters.AddWithValue("@MTLog_MsgId", dataGridView1.Rows[i].Cells[14].Value.ToString());
                cmd.Parameters.AddWithValue("@MTLog_Save_Time", DateTime.Now);
                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();
            }


            MessageBox.Show("The add was successful!", "Saving", MessageBoxButtons.OK);
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            ArrayList dt = new ArrayList();
            //DataTable dt = new DataTable();
            dataGridView1.DataSource = dt;
            dt.Clear();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Extension.ExportTOExcel(dataGridView1);
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            /*if(e.KeyCode == Keys.Enter)
            {
                CallingPSCmdlet();
            }*/
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Extension.WaitNSeconds(10);
            button1.PerformClick();

            if(dataGridView1.DataSource != null)
            {
                Extension.WaitNSeconds(10);
                button2.PerformClick();
            }
        }
    }
}
