using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace grbloxy.Properties
{
    public class Class_Save
    {
        private void ShowSaveDialog()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.FileName = "默认文件";
                saveFileDialog.Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";

                // 此时调试到这一行就能正常弹出对话框了
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string path = saveFileDialog.FileName;
                    MessageBox.Show($"选中路径：{path}");
                }
            }
        }
        public void BtnSelectSavePath_Click(object sender, EventArgs e)
        {

            // 1. 创建保存文件对话框实例
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                // 2. 配置对话框参数（关键）
                // 设置默认文件名
                saveFileDialog.FileName = "默认文件名";
                // 设置文件类型筛选（格式："显示文本|扩展名"，多个类型用|分隔）
                saveFileDialog.Filter = "文本文件 (*.txt)|*.txt|Excel文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                // 设置默认选中的文件类型（对应Filter的索引，从0开始）
                saveFileDialog.FilterIndex = 0;
                // 设置对话框标题
                saveFileDialog.Title = "选择文件保存位置";
                // 是否覆盖已有文件前提示
                saveFileDialog.OverwritePrompt = true;
                // 是否检查路径是否存在（默认true）
                saveFileDialog.CheckPathExists = true;
                // 是否允许创建新文件夹（Win10+支持，需设置）
                saveFileDialog.CreatePrompt = false; // 若设为true，会提示创建新文件（非文件夹）

                // 3. 显示对话框并判断用户是否点击"保存"
                ShowSaveDialog();
            }

        }
    }

}

