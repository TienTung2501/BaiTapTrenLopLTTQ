using Bai4_2.Data;
using Bai4_2.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bai4_2
{
    public partial class FormHangHoa : Form
    {
        public string fileAnh;
        DataBaseProcesser dataBase = new DataBaseProcesser();
        CommonFunction cmmf = new CommonFunction();
        public FormHangHoa()
        {
            InitializeComponent();
            DataTable dtChatLieu = dataBase.ReadData("Select * from tblChatLieu");
            cmmf.fillComboBox(cbChatLieu, dtChatLieu, "TenChatLieu", "MaChatLieu");
        }


        private void FormHangHoa_Load(object sender, EventArgs e)
        {
            lbXinChao.Text += staticVariable.userName;
            loadData();

        }
        void loadData()
        {

            string sqlhang = "Select * from tblHang";
            DataTable dt = dataBase.ReadData(sqlhang);
            dataGrid.DataSource = dt;
            dataGrid.Columns[0].HeaderText = "Mã hàng";
            dataGrid.Columns[0].HeaderText = "Tên hàng";
            dataGrid.Columns[0].HeaderText = "Chất liệu hàng";
            dataGrid.Columns[0].HeaderText = "Số lương ";
            dataGrid.Columns[0].HeaderText = "Đơn giá nhập";
            dataGrid.Columns[0].HeaderText = "Đơn giá bán";
            dataGrid.Columns[0].HeaderText = "Đơn giá nhập";
            dataGrid.Columns[0].HeaderText = "Ảnh";
            dataGrid.Columns[0].HeaderText = "Ghi chú";

        }

        private void dataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaHang.Text = dataGrid.CurrentRow.Cells[0].Value.ToString();
            txtTenHang.Text = dataGrid.CurrentRow.Cells[1].Value.ToString();
            cbChatLieu.SelectedValue = dataGrid.CurrentRow.Cells[2].Value.ToString();
            txtSoLuong.Text = dataGrid.CurrentRow.Cells[3].Value.ToString();
            txtDonGiaNhap.Text = dataGrid.CurrentRow.Cells[4].Value.ToString();
            txtDonGiaBan.Text = dataGrid.CurrentRow.Cells[5].Value.ToString();
            txtGhiChu.Text = dataGrid.CurrentRow.Cells[7].Value.ToString();
            fileAnh = dataGrid.CurrentRow.Cells[6].Value.ToString();
            picAnh.Image = Image.FromFile(Application.StartupPath + "\\images\\products\\" + fileAnh);
            btnThem.Enabled = false;
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string[] imagePath;
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "JPEG Images|*.jpg|PNG Images|*png|All file|*.*";
            openfile.FilterIndex = 1;
            openfile.InitialDirectory = Application.StartupPath;
            if (openfile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    picAnh.Image = Image.FromFile(openfile.FileName);
                }
                catch
                {
                    picAnh.Image = null;
                    MessageBox.Show("vui lòng nhập ảnh");
                }
                imagePath = openfile.FileName.ToString().Split('\\');
                fileAnh = imagePath[imagePath.Length - 1];
                MessageBox.Show(fileAnh);


            }
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            // validate dữ liệu:
            if (txtMaHang.Text.Trim() == "" || txtTenHang.Text.Trim() == "" || txtDonGiaNhap.Text.Trim() == "" || txtDonGiaBan.Text.Trim() == "" || txtSoLuong.Text.Trim() == "" || txtGhiChu.Text.Trim() == "" || cbChatLieu.SelectedItem.ToString().Trim() == "")
                MessageBox.Show("Bạn cần nhập đủ thông tin");
            else
            {
                try
                {
                    float.Parse(txtDonGiaBan.Text);
                    float.Parse(txtDonGiaNhap.Text);
                    int.Parse(txtSoLuong.Text);
                }
                catch
                {
                    MessageBox.Show("Cần nhập đúng định dạng số của các trường Đơn giá nhập or Đơn giá bán or Số lượng");
                }
                DataTable dtCheck = dataBase.ReadData("Select * from tblHang Where MaHang='" + txtMaHang.Text + "' ");
                if (dtCheck.Rows.Count > 0)
                {
                    MessageBox.Show("Mã hàng đã có mời nhập mã khác");
                    txtMaHang.Focus();
                    return;
                }
                // thêm mới hàng:
                string insertTblHang = "INSERT INTO tblHang values('" + txtMaHang.Text + "', N'" + txtTenHang.Text + "', N'" + cbChatLieu.SelectedValue + "', " + int.Parse(txtSoLuong.Text) + ", " + float.Parse(txtDonGiaNhap.Text) + ", " + float.Parse(txtDonGiaBan.Text) + ", '" + fileAnh + "', N'" + txtGhiChu.Text + "')";


                /* string insertTblHang = "insert into tblHang values('" + txtMaHang.Text + "',N'" + txtTenHang.Text + "',N'" + cbChatLieu.SelectedValue + "'," + int.Parse(txtSoLuong.Text) + ","+float.Parse(txtDonGiaNhap.Text)+","+float.Parse(txtDonGiaBan.Text)+","+fileAnh+",N'"+txtGhiChu.Text+"')";*/
                dataBase.ChangeData(insertTblHang);
                loadData();
                MessageBox.Show("Thêm thành công");
                ressetValue();
            }
        }
        void ressetValue()
        {
            txtMaHang.Text = "";
            txtTenHang.Text = "";
            txtDonGiaBan.Text = "";
            cbChatLieu.Text = "";
            txtSoLuong.Text = "";
            txtGhiChu.Text = "";
            txtDonGiaNhap.Text = "";
            picAnh.Image = null;
            fileAnh = "";
            txtMaHang.Focus();
            btnThem.Enabled = true;
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn xóa không?", "Có hoặc không", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    dataBase.ChangeData("Delete tblHang where MaHang='" + txtMaHang.Text + "'");
                    loadData();
                    ressetValue();
                }
                catch
                {
                    MessageBox.Show("Bạn không được xóa dữ liệu vì nó có liên quan đến các thông tin quan trọng khác.");
                }
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (txtMaHang.Text.Trim() == "" || txtTenHang.Text.Trim() == "" || txtDonGiaNhap.Text.Trim() == "" || txtDonGiaBan.Text.Trim() == "" || txtSoLuong.Text.Trim() == "" || txtGhiChu.Text.Trim() == "" || cbChatLieu.SelectedItem.ToString().Trim() == "")
                MessageBox.Show("Bạn cần nhập đủ thông tin");
            else
            {
                try
                {
                    float.Parse(txtDonGiaBan.Text);
                    float.Parse(txtDonGiaNhap.Text);
                    int.Parse(txtSoLuong.Text);
                }
                catch
                {
                    MessageBox.Show("Cần nhập đúng định dạng số của các trường Đơn giá nhập or Đơn giá bán or Số lượng");
                }
                dataBase.ChangeData("update tblHang set TenHang=N'" + txtTenHang.Text + "',ChatLieu='" + cbChatLieu.SelectedValue + "',SoLuong=" + int.Parse(txtSoLuong.Text) + ",DonGiaNhap=" + float.Parse(txtDonGiaNhap.Text) + ",DonGiaBan=" + float.Parse(txtDonGiaBan.Text) + ",Anh='" + fileAnh + "',GhiChu='" + txtGhiChu.Text + "' where MaHang='"+txtMaHang.Text+"'");
                loadData();
                ressetValue();
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        
        }
    }
}
