using System;
using System.Windows;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Practica3
{
    public class APIInteraction
    {
        private string fullname;
        private bool ContainsExtraChars(string input)
        {
            return !Regex.IsMatch(input, @"^[А-Яа-я\s]+$");
        }
        public string GetFullName()
        {
            string URL = "http://localhost:4444/TransferSimulator/fullName";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
            request.Method = "GET";
            request.Proxy.Credentials = new NetworkCredential("student", "student");
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string text = reader.ReadToEnd();
            JObject jobject = JObject.Parse(text);
            string value = (string)jobject["value"];
            fullname = value;
            return fullname;
        }
        public string FillDocument()
        {
            string result = "";
            if (fullname == null)
            {
                MessageBox.Show("<<!Данные не были получены!>>");
                return result;
            }
            bool isValidFullname = ContainsExtraChars(fullname);
            result = isValidFullname ? "ФИО содержит запрещенные символы" : "ФИО не содержит запрещенные символы";
            string[] rowData = { $"Введены данные \n{fullname}", result, "Успешно" };
            AddToWordTable(rowData);
            return result;
        }
        public void AddToWordTable(string[] rowData)
        {
            var openFileDlg = new System.Windows.Forms.OpenFileDialog();
            string filePath = "";
            if (openFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                filePath = openFileDlg.FileName;
            else return;
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(filePath);
                wordApp.Visible = false;
                Word.Table table = doc.Tables[1];
                Word.Row row = table.Rows.Add();
                for (int i = 0; i < rowData.Length; i++)
                {
                    row.Cells[i + 1].Range.Text = rowData[i];
                }

                doc.Save();

                MessageBox.Show("");
            }
            catch
            {
                MessageBox.Show("");
            } 
            finally
            {
                if (doc != null)
                {
                    doc.Close(Word.WdSaveOptions.wdSaveChanges);
                }
                wordApp.Quit(Word.WdSaveOptions.wdSaveChanges);
            }
            
        }
    }
}
