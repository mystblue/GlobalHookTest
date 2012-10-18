using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;

namespace GlobalHookTest
{
    // ボタンをクリックすると、レジストリ（LocalMachine）に書き込むテスト
    class MyForm2 : Form
    {
        public MyForm2()
        {
            this.Text = "Registry Test";

            // 位置
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(200, 100);

            // サイズ
            this.Size = new System.Drawing.Size(300, 300);

            AddRegistButton();
            AddUnRegistButton();
        }

        private void AddRegistButton()
        {
            Button registButton = new Button();
            registButton.Text = "レジストリ登録";
            registButton.Parent = this;
            registButton.Location = new System.Drawing.Point(50, 50);
            registButton.Size = new System.Drawing.Size(200, 50);
            registButton.Click += new EventHandler(registButtonClicked);
            registButton.DialogResult = DialogResult.OK;
        }

        private void AddUnRegistButton()
        {
            Button registButton = new Button();
            registButton.Text = "レジストリ解除";
            registButton.Parent = this;
            registButton.Location = new System.Drawing.Point(50, 150);
            registButton.Size = new System.Drawing.Size(200, 50);
            registButton.Click += new EventHandler(unregistButtonClicked);
            registButton.DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// レジストリ登録ボタンが押された時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void registButtonClicked(object sender, EventArgs e)
        {
            ProcessStartInfo psInfo = new ProcessStartInfo();
            psInfo.FileName = "MSOfficeKeyRegister.exe";
            //psInfo.Arguments = "user_id factroy_id 1008";
            Process exec = Process.Start(psInfo);
            exec.WaitForExit();
        }

        /// <summary>
        /// OKボタンが押された時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void unregistButtonClicked(object sender, EventArgs e)
        {
            try{
                string regKeyName = @"SOFTWARE\Classes\Excel.Sheet.8";
                RegistryKey regKey = Registry.LocalMachine.CreateSubKey(regKeyName);
                regKey.SetValue("BrowserFlags", 8, Microsoft.Win32.RegistryValueKind.DWord);
                regKey.Close();
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("レジストリにアクセスする権限がありません。", "権限なし");
            }
        }
    }
}
