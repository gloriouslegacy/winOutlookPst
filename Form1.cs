using System;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pstFinder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnPatch_Click(object sender, EventArgs e)
        {
            textBox1.Text = "패치 중...";

           //Outlook.exe 프로세스 검색
           Process[] processes = Process.GetProcessesByName("Outlook");

            if (processes.Length > 0)
            {
                foreach (Process process in processes)
                {
                    process.Kill();
                }
           
            }
            
            //패치
            string registryPath = @"HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\office\16.0\outlook";
            string registryName = "disablepst";
            int registryValue = 0;

            try
            {
                Registry.SetValue(registryPath, registryName, registryValue, RegistryValueKind.DWord);
                MessageBox.Show("패치 완료");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"레지스트리 값 변경 중 오류가 발생했습니다: {ex.Message}");
            }

            textBox1.Text = "패치 완료";
        }
        
        //
        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}
