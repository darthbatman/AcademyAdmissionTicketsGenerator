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
    public class Applicant
    {
        public string Name;
        public string ID;
    }

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<Applicant> applicants = new List<Applicant>();
        string date = "December 10, 2016";
        string time = "8:45-11:15";
        string roomNumber = "100";

        private void readApplicantList()
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Word Documents (*.docx)|*.docx";

            if (ofd.ShowDialog() == DialogResult.OK)
            {

                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                Document document = word.Documents.Open(ofd.FileName);

                for (int i = 1; i < document.Paragraphs.Count; i++)
                {
                    if (document.Paragraphs[i + 1].Range.Text.ToString().IndexOf(" –") != -1)
                    {
                        Applicant a = new Applicant();
                        a.Name = document.Paragraphs[i + 1].Range.Text.ToString().Split(new string[] { " –" }, StringSplitOptions.None)[0];
                        a.ID = document.Paragraphs[i + 1].Range.Text.ToString().Substring(document.Paragraphs[i + 1].Range.Text.ToString().Length - 4, 3);
                        applicants.Add(a);
                    }
                }

                document.Close();
                word.NormalTemplate.Saved = true;
                word.Quit();
            }
        }

        private void generateAdmissionTickets()
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            var document = word.Documents.Add();

            for (int i = 0; i < applicants.Count; i++)
            {
                StringBuilder wordDocBuilder = new StringBuilder();
                Console.WriteLine(i.ToString());
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

                wordDocBuilder.Append("Room Number: " + roomNumber + "\n\n");

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

            string fileName = Directory.GetCurrentDirectory() + "/test.doc";

            word.ActiveDocument.Paragraphs.SpaceAfter = 0;
            word.ActiveDocument.SaveAs(fileName, WdSaveFormat.wdFormatDocument);
            document.Close();
            word.NormalTemplate.Saved = true;
            word.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            readApplicantList();
            generateAdmissionTickets();
        }
    }
}
