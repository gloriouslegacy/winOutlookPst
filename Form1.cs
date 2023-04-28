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
                //MessageBox.Show("Outlook.exe 프로세스가 종료되었습니다.");
            }
            //else
            //{
            //    MessageBox.Show("Outlook.exe 프로세스가 실행 중이 아닙니다.");
            //}

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

        //private void btnSearch_Click(object sender, EventArgs e)
        //{
        //    // 검색할 파일 확장자
        //    string searchPattern = "*.pst";

        //    // 검색할 드라이브 목록
        //    string[] drives = new string[] { "C:", "D:" };

        //    // 검색 결과 출력할 텍스트 박스 등
        //    // ...

        //    // 각 드라이브에서 파일 검색
        //    foreach (string drive in drives)
        //    {
        //        try
        //        {
        //            DirectoryInfo di = new DirectoryInfo(drive);
        //            FileInfo[] files = di.GetFiles(searchPattern, SearchOption.AllDirectories);
        //            foreach (FileInfo file in files)
        //            {
        //                richTextBox1.AppendText(file.FullName + Environment.NewLine);
        //                // 검색 결과 출력할 텍스트 박스 등에 결과 추가
        //                // ...
        //            }
        //        }
        //        catch (UnauthorizedAccessException)
        //        {
        //            //MessageBox.Show($"파일 검색 중 오류가 발생했습니다: {ex.Message}");
        //        }
        //    }

        //    //string fileExtension = ".png";
        //    //string[] drives = Environment.GetLogicalDrives();
        //    //foreach (string drive in drives)
        //    //{
        //    //    try
        //    //    {
        //    //        DirectoryInfo dir = new DirectoryInfo(drive);
        //    //        foreach (FileInfo file in dir.GetFiles($"*{fileExtension}", SearchOption.AllDirectories))
        //    //        {
        //    //            richTextBox1.AppendText(file.FullName + Environment.NewLine);
        //    //        }
        //    //    }
        //    //    catch (UnauthorizedAccessException)
        //    //    {
        //    //        // 권한이 없는 경우 생략
        //    //    }
        //    //}

        //}

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}
