using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Trilogen
{
    public partial class StatusForm : Form
    {
        // public delegates
        public delegate void CancelCallback();

        // events
        public event CancelCallback CancelClicked;

        // private delegates
        private delegate void SetProgressValueCallback(int value);
        private delegate void SetStatusTextCallback(string text);
        private delegate void SetTitleCallback(string text);
        private delegate void CloseWindowCallback();

        public StatusForm()
        {
            InitializeComponent();

            // set cancel clicked event
            CancelClicked += StatusForm_CancelClicked;
        }

        void StatusForm_CancelClicked()
        {
            // close status form window
            CloseWindow();
        }

        // thread safe
        public void SetProgressValue(int value)
        {
            if (pbStatus.InvokeRequired)
            {
                SetProgressValueCallback progressValueCallback = new SetProgressValueCallback(SetProgressValue);
                Invoke(progressValueCallback, new object[] { value });
            }
            else
            {
                pbStatus.Value = value;
            }
        }

        // thread safe - change status text
        public void SetStatusText(string text)
        {
            if (lblStatus.InvokeRequired)
            {
                SetStatusTextCallback statusTextCallback = new SetStatusTextCallback(SetStatusText);
                Invoke(statusTextCallback, new object[] { text });
            }
            else
            {
                lblStatus.Text = text;
            }
        }

        // thread safe - change title
        public void SetTitle(string text)
        {
            if (InvokeRequired)
            {
                SetTitleCallback titleCallback = new SetTitleCallback(SetTitle);
                Invoke(titleCallback, new object[] { text });
            }
            else
            {
                Text = text;
            }
        }

        // thread safe - close window
        public void CloseWindow()
        {
            if (InvokeRequired)
            {
                CloseWindowCallback closeWindowCallback = new CloseWindowCallback(CloseWindow);
                Invoke(closeWindowCallback);
            }
            else
            {
                // close and dispose of form window
                Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // confirm cancel
            DialogResult result = MessageBox.Show("Are you sure you want to cancel?", "Confirmation required", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            // user hit yes
            bool confirmCancel = (result == System.Windows.Forms.DialogResult.Yes);

            // cancel
            if (confirmCancel)
            {
                // cancel clicked event
                CancelClicked();
            } 
        }
    }
}
