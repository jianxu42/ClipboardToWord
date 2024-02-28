using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ClipboardToWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            // Configure your form and add controls here
            this.Text = "Seamless Clipboard to Word Transfer Tool";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            var triggerButton = new Button
            {
                Text = "Clipboard to Word",
                Size = new System.Drawing.Size(400, 50)
            };
            triggerButton.Location = new System.Drawing.Point(
                (this.ClientSize.Width - triggerButton.Width) / 2,
                (this.ClientSize.Height - triggerButton.Height) / 2
            );
            triggerButton.Click += button1_Click;
            Controls.Add(triggerButton);
        }

        private void button1_Click(object? sender, EventArgs e)
        {
            try
            {
                if (Clipboard.ContainsText())
                {
                    // Initialize Word application
                    Word.Application wordApp = new Word.Application();
                    wordApp.Visible = false;

                    // Create a new document
                    Word.Document newDoc = wordApp.Documents.Add();

                    // Paste the content into the new document
                    newDoc.Content.Paste();

                    // Now, copy the content of the new document to the clipboard
                    newDoc.Content.Copy();

                    // Close the documents
                    newDoc.Close(false);
                    wordApp.Quit();

                    // Cleanup
                    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(newDoc);
                    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

                    MessageBox.Show("Content copied to clipboard.");
                }
                else
                {
                    MessageBox.Show("The current content on the clipboard is not text.");
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }
    }
}
