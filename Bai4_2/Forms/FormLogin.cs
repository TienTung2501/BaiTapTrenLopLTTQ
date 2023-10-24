using Bai4_2.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Bai4_2.Data;
namespace Bai4_2.Forms
{
    public partial class FormLogin : Form
    {
        DataBaseProcesser dbData=new DataBaseProcesser();
        public FormLogin()
        {
            InitializeComponent();
        }

        private void FormLogin_Load(object sender, EventArgs e)
        {
            txtUserName.Text = "UserName";
            txtUserName.ForeColor = Color.Gray;
            txtPassWord.Text = "Password";
            txtPassWord.ForeColor = Color.Gray;
            txtUserName.Select();
        }

        private void txtUserName_Enter(object sender, EventArgs e)
        {
            if (txtUserName.Text == "UserName")
            {
                txtUserName.Text = "";
                txtUserName.ForeColor = Color.Black;
            }
        }

        private void txtUserName_Leave(object sender, EventArgs e)
        {
            if (txtUserName.Text.Trim() == "")
            {
                txtUserName.Text = "UserName";
                txtUserName.ForeColor = Color.Gray;
            }    
        }
        private void txtPassWord_Leave(object sender, EventArgs e)
        {
            if (txtPassWord.Text.Trim() == "")
            {
                txtPassWord.Text = "Password";
                txtPassWord.ForeColor = Color.Gray;
            }
        }

        private void txtPassWord_Enter(object sender, EventArgs e)
        {
            if (txtPassWord.Text == "Password")
            {
                txtPassWord.Text = "";
                txtPassWord.ForeColor = Color.Black;
            }
        }
        private void btnDangNhap_Click(object sender, EventArgs e)
        {
            if (txtUserName.Text.Trim() == "" || txtUserName.Text == "UserName")
                MessageBox.Show("Bạn cần nhập UserName");
            else if (txtPassWord.Text == "Password")
            {
                MessageBox.Show("Bạn cần nhập Password");
            }
            else
            {
                string userName = txtUserName.Text;
                string passWord = txtPassWord.Text;

                string sql = "SELECT * FROM tblUser WHERE userName = '" + userName + "' AND passWordd = '" + passWord + "'";

                if (dbData.ReadData(sql).Rows.Count > 0)
                {
                    staticVariable.userName = userName;
                    FormHangHoa formHangHoa= new FormHangHoa();
                    this.Hide();
                    formHangHoa.Show();
                }
                else
                {
                    MessageBox.Show("Sai Username or Password! Vui lòng nhập lại!!");
                    txtUserName.Clear();
                    txtPassWord.Clear(); 
                    txtUserName.Focus();
                }
            }

        }

       
    }
}
