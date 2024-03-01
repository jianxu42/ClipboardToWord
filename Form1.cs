using Word = Microsoft.Office.Interop.Word;

namespace ClipboardToWord
{
    public partial class Form1 : Form
    {
        Button transferButton; // Declare it at the class level to access it in both methods

        public Form1()
        {
            this.Text = "Seamless Clipboard to Word Transfer Tool";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            var validateRTFButton = new Button
            {
                Text = "Validate RTF in Clipboard",
                Size = new System.Drawing.Size(400, 50),
                Location = new System.Drawing.Point(
                    (this.ClientSize.Width - 400) / 2,
                    (this.ClientSize.Height - 50) / 2 - 60 // Position appropriately
                )
            };
            validateRTFButton.Click += ValidateRTFButton_Click;
            Controls.Add(validateRTFButton);

            transferButton = new Button
            {
                Text = "Clipboard to Word",
                Size = new System.Drawing.Size(400, 50),
                Location = new System.Drawing.Point(
                    (this.ClientSize.Width - 400) / 2,
                    (this.ClientSize.Height - 50) / 2 + 10 // Position appropriately
                ),
                Visible = false // Initially hidden
            };
            transferButton.Click += TransferButton_Click;
            Controls.Add(transferButton);
        }

        private void ValidateRTFButton_Click(object? sender, EventArgs e)
        {
            if (!Clipboard.ContainsData(DataFormats.Rtf))
            {
                transferButton.Visible = true; // Show the transfer button only if clipboard does not contain RTF
            }
            else
            {
                MessageBox.Show("The clipboard already contains RTF content.");
            }
        }

        private void TransferButton_Click(object? sender, EventArgs e)
        {
            try
            {
                if (Clipboard.ContainsText())
                {
                    if (!Clipboard.ContainsData(DataFormats.Rtf))
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
                        MessageBox.Show("The clipboard already contains RTF content; there is no need to paste it into Word.");
                    }
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
