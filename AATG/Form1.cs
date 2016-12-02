using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace AATG
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public List<Applicant> applicants = new List<Applicant>();
        public string date = "December 10, 2016";
        public string time = "8:45-11:15";
        public string applicantsListFile = "";
        public string ticketsListFile = "";
        
        private bool proceedWithGeneration = true;

        private void readApplicantList()
        {

            if (applicantsListFile.Length > 0)
            {

                label6.Text = "Status: Reading data from Applicant List File";

                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                Document document = word.Documents.Open(applicantsListFile);

                for (int i = 1; i < document.Paragraphs.Count; i++)
                {
                    if (document.Paragraphs[i + 1].Range.Text.ToString().IndexOf(" –") != -1)
                    {
                        Applicant a = new Applicant();
                        a.Name = document.Paragraphs[i + 1].Range.Text.ToString().Split(new string[] { " –" }, StringSplitOptions.None)[0];
                        a.ID = document.Paragraphs[i + 1].Range.Text.ToString().Substring(document.Paragraphs[i + 1].Range.Text.ToString().Length - 4, 3);
                        applicants.Add(a);
                    }
                    else if (document.Paragraphs[i + 1].Range.Text.ToString().IndexOf(", ") != -1)
                    {
                        Applicant a = new Applicant();
                        a.Name = document.Paragraphs[i + 1].Range.Text.ToString().Split(new string[] { "–" }, StringSplitOptions.None)[0];
                        a.ID = document.Paragraphs[i + 1].Range.Text.ToString().Substring(document.Paragraphs[i + 1].Range.Text.ToString().Length - 4, 3);
                        applicants.Add(a);
                    }
                }

                document.Close();
                word.NormalTemplate.Saved = true;
                word.Quit();
            }
            else
            {
                MessageBox.Show("Please select the Applicants List File.");
            }

        }

        private void generateAdmissionTickets()
        {
            if (ticketsListFile.Length > 0)
            {

                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                var document = word.Documents.Add();

                for (int i = 0; i < applicants.Count; i++)
                {

                    label6.Text = "Status: Writing Ticket " + (i + 1) + " of " + applicants.Count;

                    StringBuilder wordDocBuilder = new StringBuilder();
                    wordDocBuilder.Append("Middlesex County Academy	Admissions Test Ticket\n");

                    var paragraph = document.Paragraphs.Add();
                    paragraph.Range.Font.Name = "Calibri";
                    paragraph.Range.Font.Size = 11;
                    paragraph.LineSpacing = 1F;
                    paragraph.Range.Text = wordDocBuilder.ToString();

                    wordDocBuilder = new StringBuilder();

                    wordDocBuilder.Append(date + "		" + time + "\n\n");

                    paragraph.Range.Bold = 1;

                    paragraph = document.Paragraphs.Add();
                    paragraph.Range.Font.Name = "Calibri";
                    paragraph.Range.Font.Size = 11;
                    paragraph.LineSpacing = 1F;
                    paragraph.Range.Text = wordDocBuilder.ToString();

                    wordDocBuilder = new StringBuilder();

                    wordDocBuilder.Append("Student Name/ID: " + applicants[i].Name + " / " + applicants[i].ID + "\n\n");

                    paragraph.Range.Bold = 0;

                    paragraph = document.Paragraphs.Add();
                    paragraph.Range.Font.Name = "Calibri";
                    paragraph.Range.Font.Size = 11;
                    paragraph.LineSpacing = 1F;
                    paragraph.Range.Text = wordDocBuilder.ToString();

                    wordDocBuilder = new StringBuilder();

                    wordDocBuilder.Append("Room Number: " + applicants[i].Room + "\n\n");

                    paragraph.Range.Bold = 0;

                    paragraph = document.Paragraphs.Add();
                    paragraph.Range.Font.Name = "Calibri";
                    paragraph.Range.Font.Size = 11;
                    paragraph.LineSpacing = 1F;
                    paragraph.Range.Text = wordDocBuilder.ToString();

                    wordDocBuilder = new StringBuilder();

                    wordDocBuilder.Append("*This ticket is required for entry to the Admissions Test\n\n\n\n");

                    paragraph.Range.Bold = 1;

                    paragraph = document.Paragraphs.Add();
                    paragraph.Range.Font.Name = "Calibri";
                    paragraph.Range.Font.Size = 11;
                    paragraph.LineSpacing = 1F;
                    paragraph.Range.Text = wordDocBuilder.ToString();

                    paragraph.Range.Bold = 0;

                }

                label6.Text = "Status: Tickets List Generated";

                word.ActiveDocument.Paragraphs.SpaceAfter = 0;
                word.ActiveDocument.SaveAs(ticketsListFile, WdSaveFormat.wdFormatDocumentDefault);
                document.Close();
                word.NormalTemplate.Saved = true;
                word.Quit();
            }
            else
            {
                MessageBox.Show("Please select the Tickets List File.");
            }

        }

        private void assignRoomNumbers()
        {

            int numApplicantsOffset = 0;

            for (int i = 1; i <= 7; i++)
            {
                if (Controls.Find("roomTextBox" + i, true).Length > 0 && ((TextBox)(Controls.Find("roomTextBox" + i, true)[0])).Text.Length > 0 && Controls.Find("numApplicantsTextBox" + i, true).Length > 0 && ((TextBox)(Controls.Find("numApplicantsTextBox" + i, true)[0])).Text.Length > 0)
                {
                    int numApplicantsInRoom = 0;

                    if (Int32.TryParse(((TextBox)(Controls.Find("numApplicantsTextBox" + i, true)[0])).Text, out numApplicantsInRoom))
                    {
                        for (int j = 0; j < numApplicantsInRoom; j++)
                        {
                            if (j + numApplicantsOffset < applicants.Count)
                            {
                                applicants[j + numApplicantsOffset].Room = ((TextBox)(Controls.Find("roomTextBox" + i, true)[0])).Text;
                            }
                        }

                        numApplicantsOffset += numApplicantsInRoom;

                    }
                }
            }

            if (numApplicantsOffset < applicants.Count)
            {
                proceedWithGeneration = false;
                MessageBox.Show("Total number of applicants in all rooms should meet or exceed " + applicants.Count);
            }

        }

        private void selectApplicantListFile()
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Word Documents (*.docx)|*.docx";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                applicantsListFile = ofd.FileName;
                textBox1.Text = applicantsListFile;
            }

        }

        private void selectTicketsListFile()
        {

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word Documents (*.docx)|*.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ticketsListFile = sfd.FileName;
                textBox2.Text = ticketsListFile;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            readApplicantList();
            assignRoomNumbers();
            if (proceedWithGeneration)
            {
                generateAdmissionTickets();
            }
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            System.Diagnostics.Process.Start(ticketsListFile);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("How to Use\n1. Select Applicant List File.\n2. Select where to save the Ticket List File.\n3. Enter the Room Name(s)/Number(s) and the Number of Applicants in the Room.\n4. Click generate and wait for generation of the Tickets List to complete.");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            selectApplicantListFile();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            selectTicketsListFile();
        }

    }

    public class Applicant
    {
        public string Name;
        public string ID;
        public string Room;
    }

}
