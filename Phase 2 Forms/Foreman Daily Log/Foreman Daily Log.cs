using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Pechkin;

// Above is extra code written by microsoft to speed up development

namespace Phase_2_Forms
{
    public partial class Daily_Log : Form
    {
        public string HTMLString { get; private set; }

        /*
         * This constructor Initializes everything, adds information to tables, and loads
         * the data from the saved file.
         */
        public Daily_Log()
        {
            InitializeComponent();

            manpower.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            weeklySchedule.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            //Code to add info in the rows of weeklySchedule as I could not figure out a cleaner way to do so
            weeklySchedule.Rows.Add("Task & Schedule Start/Finish by Trade");
            weeklySchedule.Rows.Add("Inspections");
            weeklySchedule.Rows.Add("Quality Control");
            weeklySchedule.Rows.Add("Projected Footage");
            weeklySchedule.Rows.Add("Manpower");
            weeklySchedule.Rows.Add("Obstacles");
            weeklySchedule.Rows.Add("GC & Trades");
            weeklySchedule.Rows.Add("Materials Needed");
            weeklySchedule.Rows.Add("Equipment");
            weeklySchedule.Rows.Add("STS & Fasteners");
            weeklySchedule.Rows.Add("P2 Coordination");

            // Code to add the version number to the main page (top-right)
            versionNumber.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            // Code to fill the log with what is written in the StoredData.txt file
            LoadSaveFile();
        }

        /*
         * Stops the initial close, saves the data, and then closes.
         */
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Code to save the log into the StoredData.txt file
            StoreSaveFile();
            // Closes Daily Log
            base.OnFormClosing(e);
        }

        /*
         * Creates outlook application and then returns the newly created application
         */
        private Outlook._Application CreateOutlookApp()
        {
            // Finds the registry key that outlook is in
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE");
            // Finds the executable path for outlook
            string path = (string)key.GetValue("Path");
            // Tests to see if path exists (outlook is installed)
            if (path != null)
                // Runs outlook
                System.Diagnostics.Process.Start("OUTLOOK.EXE");
            else
                // Throws an error if no outlook path is found (outlook is not installed)
                MessageBox.Show("There is no Outlook in this computer!", "SystemError", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            // Sets a variable to outlook application
            Outlook._Application app = new Outlook.Application();
            // Returns the application for other code to use
            return app;
        }

        /*
         * Does sanity checks, opens outlook, create pdf, sends email.
         */
        private void SendButton_Click(object sender, EventArgs e)
        {
            // for the if statement to check if jobNumberTextBox contains "\" there has to be two \ 
            // since it is an escape character meaning remove special meaning in code.
            // Thus you have to us \ to remove the special meaning behind \ or use \ to remove special
            // meaning behind ".

            // If statement checks to see if job number contains anything that windows file system
            // does not allow since it is used to create the pdf name.
            if (jobNumberTextBox.Text.Contains("~")
                || jobNumberTextBox.Text.Contains("\"")
                || jobNumberTextBox.Text.Contains("#")
                || jobNumberTextBox.Text.Contains("%")
                || jobNumberTextBox.Text.Contains("&")
                || jobNumberTextBox.Text.Contains("*")
                || jobNumberTextBox.Text.Contains(":")
                || jobNumberTextBox.Text.Contains(">")
                || jobNumberTextBox.Text.Contains("<")
                || jobNumberTextBox.Text.Contains("?")
                || jobNumberTextBox.Text.Contains("/")
                || jobNumberTextBox.Text.Contains("\\")
                || jobNumberTextBox.Text.Contains("{")
                || jobNumberTextBox.Text.Contains("|")
                || jobNumberTextBox.Text.Contains("}"))
            {
                MessageBox.Show(jobNumberLabel.Text.ToString() + " cannot contain ~ \" # % & * : < > ? / \\ { | }");
            }
            // All these else if's make sure that data has been entered
            else if (jobNumberTextBox.Text.ToString() == "")
                MessageBox.Show(jobNumberLabel.Text.ToString() + " cannot be blank");
            else if (foremanOnSiteTextBox.Text.ToString() == "")
                MessageBox.Show(foremanOnSiteLabel.Text.ToString() + " cannot be blank");
            else if (siteForemanAssistantTextBox.Text.ToString() == "")
                MessageBox.Show(siteForemanAssistentLabel.Text.ToString() + " cannot be blank");
            else if (projectNameTextBox.Text.ToString() == "")
                MessageBox.Show(projectNameLabel.Text.ToString() + " cannot be blank");
            else if (manPowerAssessmentTextBox.Text.ToString() == "")
                MessageBox.Show(manPowerAssessmentLabel.Text.ToString() + " cannot be blank");
            else if (safetyConcernsTextBox.Text.ToString() == "")
                MessageBox.Show(safetyConcernsLabel.Text.ToString() + " cannot be blank");
            else if (ahaReviewedTextBox.Text.ToString() == "")
                MessageBox.Show(ahaReviewedLabel.Text.ToString() + " cannot be blank");
            else if (scheduleConcernsTextBox.Text.ToString() == "")
                MessageBox.Show(scheduleConcernsLabel.Text.ToString() + " cannot be blank");
            else if (budgetConcernsTextBox.Text.ToString() == "")
                MessageBox.Show(budgetConcernsLabel.Text.ToString() + " cannot be blank");
            else if (deliveriesReceivedTextBox.Text.ToString() == "")
                MessageBox.Show(deliveriesReceivedLabel.Text.ToString() + " cannot be blank");
            else if (deliveriesNeededTextBox.Text.ToString() == "")
                MessageBox.Show(deliveriesNeededLabel.Text.ToString() + " cannot be blank");
            else if (newWorkAuthorizationsTextBox.Text.ToString() == "")
                MessageBox.Show(newWorkAuthorizationsLabel.Text.ToString() + " cannot be blank");
            else if (qcInspectionTextBox.Text.ToString() == "")
                MessageBox.Show(qcInspectionLabel.Text.ToString() + " cannot be blank");
            else if (notesCorrespondenceTextBox.Text.ToString() == "")
                MessageBox.Show(notesCorrespondenceLabel.Text.ToString() + " cannot be blank");
            else if (actionItemsTextBox.Text.ToString() == "")
                MessageBox.Show(actionItemsLabel.Text.ToString() + " cannot be blank");
            else if (commentsAboutShopTextBox.Text.ToString() == "")
                MessageBox.Show(commentsAboutShopLabel.Text.ToString() + " cannot be blank");
            else if (toTextBox.Text.ToString() == "")
                MessageBox.Show(toLabel.Text.ToString() + " cannot be blank");
            else if (toTextBox.Text.ToString().Contains("/n"))
                MessageBox.Show(toLabel.Text.ToString() + " cannot have 'enters' or page breaks");
            else
            {
                // try block catches and boiles up the errors that are found. Especially since this is
                // an error prone area.
                try
                {
                    string fileLocation;
                    string fileName;


                    // Convert current date and time to string (words)
                    String sDate = DateTime.Now.ToString();
                    DateTime datevalue = (Convert.ToDateTime(sDate.ToString()));

                    // Extract day, month, and year to be used in the email and pdf name.
                    String dy = datevalue.Day.ToString();
                    String mn = datevalue.Month.ToString();
                    String yy = datevalue.Year.ToString();

                    String date = yy + "-" + mn + "-" + dy;

                    // Compile the beginning part of the Daily Log as html for email.
                    // Used *Label.Text to make the email reflect the daily log labels text for
                    // both consistency, and easier editing.
                    string body = "<html><body><b>" + jobNumberLabel.Text + "</b> " + jobNumberTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                                + "<b>" + foremanOnSiteLabel.Text + "</b> " + foremanOnSiteTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                                + "<b>" + siteForemanAssistentLabel.Text + "</b> " + siteForemanAssistantTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                                + "<b>" + projectNameLabel.Text + "</b> " + projectNameTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                                + "<b>" + "Date:" + "</b> " + DateTime.Now + "<br><br>"
                                + "<b>" + "To:" + "</b> " + toTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                                // Start on table in html for email
                                + "<table class=fixed border=1 style=width:100% cellpadding='1'>"
                                + "<caption>" + manpowerLabel.Text.ToString() + "</caption>";
                    // loop through Daily Log table to construct html table
                    // this for loop goes through the rows
                    for (int row = 0; row < manpower.RowCount; row++)
                    {
                        // Bad sloppy code that does not account for number of column headers. Works though...
                        if (row == 0)
                        {
                            body += "<tr><th><b>" + manpower.Columns[0].HeaderText.ToString() + "</b></th>";
                            body += "<th><b>" + manpower.Columns[1].HeaderText.ToString() + "</b></th>";
                            body += "<th><b>" + manpower.Columns[2].HeaderText.ToString() + "</b></th></tr>";
                        }
                        body += "<tr>";
                        // this for loop goes through the columns
                        for (int column = 0; column < manpower.ColumnCount; column++)
                        {
                            body += "<td>";
                            body += manpower.Rows[row].Cells[column].Value;
                            body += "</td>";
                        }
                        body += "</tr>";
                    }
                    body += "</table><br>"
                        // Start on table in html for email
                        + "<table class=fixed border=1 style=width:100% cellpadding='1'>"
                         + "<caption>" + weeklyScheduleLabel.Text.ToString() + "</caption>";
                    // this for loop goes through the rows
                    for (int row = 0; row < weeklySchedule.RowCount; row++)
                    {
                        // Bad sloppy code that does not account for number of column headers. Works though...
                        if (row == 0)
                        {
                            body += "<tr><th><b>" + weeklySchedule.Columns[0].HeaderText.ToString() + "</b></th>";
                            body += "<th><b>" + weeklySchedule.Columns[1].HeaderText.ToString() + "</b></th>";
                            body += "<th><b>" + weeklySchedule.Columns[2].HeaderText.ToString() + "</b></th>";
                            body += "<th><b>" + weeklySchedule.Columns[3].HeaderText.ToString() + "</b></th>";
                            body += "<th><b>" + weeklySchedule.Columns[4].HeaderText.ToString() + "</b></th>";
                            body += "<th><b>" + weeklySchedule.Columns[5].HeaderText.ToString() + "</b></th></tr>";
                        }
                        body += "<tr>";
                        // this for loop goes through the columns
                        for (int column = 0; column < weeklySchedule.ColumnCount; column++)
                        {
                            body += "<td>";
                            if (column == 0)
                                body += "<b>" + weeklySchedule.Rows[row].Cells[column].Value + "</b>";
                            else
                                body += weeklySchedule.Rows[row].Cells[column].Value;
                            body += "</td>";
                        }
                        body += "</tr>";
                    }
                    body += "</table>"
                    // Compiles all the ending information into html for table use
                    + "<br>"
                    + "<b>" + manPowerAssessmentLabel.Text + "</b><br>" + manPowerAssessmentTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + safetyConcernsLabel.Text + "</b><br>" + safetyConcernsTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + ahaReviewedLabel.Text + "</b><br>" + ahaReviewedTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + safetyConcernsLabel.Text + "</b><br>" + scheduleConcernsTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + budgetConcernsLabel.Text + "</b><br>" + budgetConcernsTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + deliveriesReceivedLabel.Text + "</b><br>" + deliveriesReceivedTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + deliveriesNeededLabel.Text + "</b><br>" + deliveriesNeededTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + newWorkAuthorizationsLabel.Text + "</b><br>" + newWorkAuthorizationsTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + qcInspectionLabel.Text + "</b><br>" + qcInspectionTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + notesCorrespondenceLabel.Text + "</b><br>" + notesCorrespondenceTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + actionItemsLabel.Text + "</b><br>" + actionItemsTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "<b>" + commentsAboutShopLabel.Text + "</b><br>" + commentsAboutShopTextBox.Text.Replace("\n", "<br>") + "<br><br>"
                    + "</body></html>";

                    // Creates outlook application by calling constructor
                    Outlook._Application app = CreateOutlookApp();

                    // Creates new email from the application
                    Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem);

                    // Adds the "To" box into the email
                    mail.To = toTextBox.Text.ToString();

                    // Adds the cc text into the email
                    mail.CC = ccLabelBox.Text;

                    // Compiles the subject
                    mail.Subject = "Foreman Daily Log " + jobNumberTextBox.Text.ToString() + "_" + date;

                    // Adds the html made in above code before the creation of an outlook app
                    mail.HTMLBody = body;

                    // Sets the importance to normal
                    mail.Importance = Outlook.OlImportance.olImportanceNormal;
                    
                    // Determines the name of the folder to store all the pdf's
                    fileLocation = "Saved Foreman Daily Logs";

                    // Creates the name for the pdf's
                    fileName = jobNumberTextBox.Text.ToString() + "_" + date + ".pdf";

                    // This line of code reads best right to left. It creates the pdf getting
                    // the html info, the object location, and then the name of the object.
                    // Then attaches the newly made file as a pdf.
                    mail.Attachments.Add(CreatePDF(body, fileLocation, fileName));

                    // Sends the email
                    ((Outlook._MailItem)mail).Send();
                    
                    // Display a sent message
                    MessageBox.Show("Your message has been successfully sent!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                catch (Exception ex)
                {
                    // Catch if error is outlook does not recognize all emails
                    if (ex.GetHashCode() == 44307222)
                        MessageBox.Show("Outlook does not recognize all the emails entered in the 'To' box. Make sure that they are all correct, and that there is a ';' in between each email.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // If other error messages then show them
                    else
                        MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }

        /*
         * Opens StoredData.txt and reads each line looking for a specific word.
         * If that word corrolates with a label then it skips a line, and then filles in the next line.
         */
        private void LoadSaveFile()
        {
            // Stores info found on line
            String line;
            try
            {
                // Opens a file reader to read lines
                StreamReader sr = new StreamReader("StoredData.txt");

                // Read a line (prime the file reader since it starts on line 0 which is
                // automatically null)
                line = sr.ReadLine();

                // As long as there is info on line keep going. This works since there has to be
                // something on every line from the first line to the last
                while (line != null)
                {
                    // if the line reads the same as the text then add the data 2 lines down
                    if (line == jobNumberLabel.Text.ToString())
                    {
                        line = sr.ReadLine();
                        line = sr.ReadLine();
                        jobNumberTextBox.Text = line;
                    }
                    // if the line reads the same as the text then add the data 2 lines down
                    else if (line == foremanOnSiteLabel.Text.ToString())
                    {
                        line = sr.ReadLine();
                        line = sr.ReadLine();
                        foremanOnSiteTextBox.Text = line;
                    }
                    // if the line reads the same as the text then add the data 2 lines down
                    else if (line == siteForemanAssistentLabel.Text.ToString())
                    {
                        line = sr.ReadLine();
                        line = sr.ReadLine();
                        siteForemanAssistantTextBox.Text = line;
                    }
                    // if the line reads the same as the text then add the data 2 lines down
                    else if (line == projectNameLabel.Text.ToString())
                    {
                        line = sr.ReadLine();
                        line = sr.ReadLine();
                        projectNameTextBox.Text = line;
                    }
                    // if the line reads the same as the text then add the data 2 lines down
                    else if (line == toLabel.Text.ToString())
                    {
                        line = sr.ReadLine();
                        line = sr.ReadLine();
                        toTextBox.Text = line;
                    }
                    // if the line reads the same as the text then then add all lines in the loop
                    // Depending on the table it will look like in the text:
                    //  row
                    //  column
                    //  column
                    //  column
                    //  column
                    //  row
                    //  column
                    //  column
                    //  column
                    //  column
                    //  row
                    //  column
                    //  column
                    //  column
                    //  column
                    // ...
                    else if (line == weeklyScheduleLabel.Text.ToString())
                    {
                        for (int row = 0; row < weeklySchedule.RowCount; row++)
                        {
                            for (int column = 0; column < weeklySchedule.ColumnCount; column++)
                            {
                                line = sr.ReadLine();
                                weeklySchedule.Rows[row].Cells[column].Value = line;
                            }
                        }
                    }
                    // load next line (in case there is more after this) or if the current line did
                    // not match a label
                    line = sr.ReadLine();
                }

                // Once all lines are read through close file so other programs (including the user) can
                // use it
                sr.Close();
            }
            catch (Exception ex)
            {
                // Push error if it occurs most likely the error is a file not found
                MessageBox.Show("Exception: " + ex.Message + " might need to make a file named 'StoredData.txt' in the same folder as the Daily Log exe.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /*
         * Errases all in the StoredData.txt and then fills it with data that is needed to be saved 
         */
        private void StoreSaveFile()
        {
            try
            {
                // Writes all lines and data into ""
                File.WriteAllText("StoredData.txt", "");

                // Creates a stream writer to write to file
                StreamWriter sw = new StreamWriter("StoredData.txt");

                // Write data that is wanted to be saved to file
                sw.WriteLine(jobNumberLabel.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(jobNumberTextBox.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(foremanOnSiteLabel.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(foremanOnSiteTextBox.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(siteForemanAssistentLabel.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(siteForemanAssistantTextBox.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(projectNameLabel.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(projectNameTextBox.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(toLabel.Text.ToString());
                sw.WriteLine();
                sw.WriteLine(toTextBox.Text.ToString());
                sw.WriteLine();
                // Step through all data to be saved in file
                // Depending on the table it will look like in the text:
                //  row
                //  column
                //  column
                //  column
                //  column
                //  row
                //  column
                //  column
                //  column
                //  column
                //  row
                //  column
                //  column
                //  column
                //  column
                // ...
                sw.WriteLine(weeklyScheduleLabel.Text.ToString());
                for (int row = 0; row < weeklySchedule.RowCount; row++)
                {
                    for (int column = 0; column < weeklySchedule.ColumnCount; column++)
                    {
                            sw.WriteLine(weeklySchedule.Rows[row].Cells[column].Value);
                    }
                }

                //Close the file
                sw.Close();
            }
            catch (Exception e)
            {
                // Show error if it occurs most likely the error is a file not found
                MessageBox.Show("Exception: " + e.Message + " might need to make a file named 'StoredData.txt' in the same folder as the Daily Log exe.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /*
          Asks for HTML Text, File Location, and File Name then creates a pdf containing the HTML text
          at the given location with the given name.
          Returns created file location.
        */
        private string CreatePDF(string pdfText, string fileLocation, string fileName)
        {
            // Creates file if none exists where the Daily Log exe is, if there is a file then it
            // does nothing
            System.IO.Directory.CreateDirectory(fileLocation);

            byte[] pdfBuffer = new SimplePechkin(new GlobalConfig()).Convert(pdfText);

            if (ByteArrayToFile(fileLocation + "\\" + fileName, pdfBuffer))
            {
                Console.WriteLine("PDF Succesfully created");
            }
            else
            {
                Console.WriteLine("Cannot create PDF");
            }

            // Returns where the pdf was created
            return Path.GetFullPath(fileLocation + "\\" + fileName);
        }
        public bool ByteArrayToFile(string _FileName, byte[] _ByteArray)
        {
            try
            {
                // Open file for reading
                FileStream _FileStream = new FileStream(_FileName, FileMode.Create, FileAccess.Write);
                // Writes a block of bytes to this stream using data from  a byte array.
                _FileStream.Write(_ByteArray, 0, _ByteArray.Length);

                // Close file stream
                _FileStream.Close();

                return true;
            }
            catch (Exception _Exception)
            {
                Console.WriteLine("Exception caught in process while trying to save : {0}", _Exception.ToString());
            }

            return false;
        }
    }
}

