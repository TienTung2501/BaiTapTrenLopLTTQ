using Bai4_2.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using Excel = Microsoft.Office.Interop.Excel;
namespace Bai4_2.Forms
{
    public partial class FormHoaDon : Form
    {
        DataBaseProcesser dataBase = new DataBaseProcesser();
        CommonFunction cmmf = new CommonFunction();
        DataTable dataTable = new DataTable();
        public int tongtien = 0;
        public Boolean chualuu = true;
        public FormHoaDon()
        {
            InitializeComponent();
            DataTable dtHang = dataBase.ReadData("Select MaHang from tblhang");
            cbMaHang.DisplayMember = "MaHang";
            cbMaHang.DataSource = dtHang;
            DataTable dtKhach = dataBase.ReadData("Select MaKhachHang from tblKhachHang");
            cbMaKhach.DisplayMember = "MaKhachHang";
            cbMaKhach.DataSource = dtKhach;
            DataTable dtNhanVien = dataBase.ReadData("Select MaNhanVien from tblNhanVien");
            cbMaNV.DisplayMember = "MaNhanVien";
            cbMaNV.DataSource = dtNhanVien;
            DataTable dtHoaDon = dataBase.ReadData("Select MaHDBan from tblHDBan");
            cbMaHoaDon.DisplayMember = "MaHDBan";// hiển thị các giá trị trong cbbox theo thuộc tính của mã hóa đơn
            cbMaHoaDon.DataSource = dtHoaDon;// truyền dữ liệu của datatable vào cbbox;
        }

        private void FormHoaDon_Load(object sender, EventArgs e)
        {
            cbMaHoaDon.Text = "";
            reset();
            txtDonGia.Enabled = false;
            txtTenHang.Enabled = false;
            txtThanhTien.Enabled = false;
            txtTenKhach.Enabled = false;
            txtTenNV.Enabled = false;
            RtxtDiaChi.Enabled = false;
            txtDienThoai.Enabled = false;
            txtMaHD.Enabled = false;
            lbTongTien.Text = "Tong Tien Thanh Toan: 0";
            /*cbMaNV.Text = string.Empty;
            cbMaKhach.Text = string.Empty;
            cbMaHang.Text = string.Empty;*/
            // Tạo một DataTable chứa dữ liệu
           

            // Thêm các cột vào DataTable
            dataTable.Columns.Add("Mã Hàng", typeof(string));
            dataTable.Columns.Add("Tên hàng", typeof(string));
            dataTable.Columns.Add("Số Lượng", typeof(int));
            dataTable.Columns.Add("Đơn Giá", typeof(int));
            dataTable.Columns.Add("Giảm Giá %", typeof(int));
            dataTable.Columns.Add("Thành Tiền", typeof(int));


        }
        public void reset()
        {
            DataTable dtHoaDon = dataBase.ReadData("Select MaHDBan from tblHDBan");
            cbMaHoaDon.DisplayMember = "MaHDBan";// hiển thị các giá trị trong cbbox theo thuộc tính của mã hóa đơn
            cbMaHoaDon.DataSource = dtHoaDon;
            tongtien = 0;
            lbTongTien.Text = "Tổng tiền thanh toán:0";
            cbMaHoaDon.Text ="";
            cbMaNV.Text= "";
            cbMaKhach.Text="";
            cbMaHang.Text="";
            dataTable.Clear();
            txtMaHD.Text = "";
            txtDonGia.Text ="0";
            txtSoLuong.Text = string.Empty;
            txtGiamGia.Text = string.Empty;
            txtTenKhach.Text = string.Empty;    
            txtTenNV.Text = string.Empty;
            txtDienThoai.Text = string.Empty;
            txtTenHang.Text = string.Empty;
            RtxtDiaChi.Text = string.Empty;
        }
        private void btnThem_Click(object sender, EventArgs e)
        {
           
            
                string maHD = txtMaHD.Text;
                string maNV = cbMaNV.Text;
                string maKhach = cbMaKhach.Text;
                DateTime dt = txtTime.Value;
                string formattedDate = dt.ToString("yyyy-MM-dd HH:mm:ss"); // Chuyển đổi ngày tháng
                string timHD = "SELECT * FROM tblHDBan WHERE MaHDBan = '" + txtMaHD.Text + "'";

                if (dataBase.ReadData(timHD).Rows.Count > 0)
                {
                    string upateHoaDon = "Update tblHDBan set MaNhanVien='" + maNV + "', NgayBan='" + dt + "',MaKhachHang='" + maKhach + "',TongTien=" + tongtien + " where MaHDBan='" + txtMaHD.Text + "' ";
                    dataBase.ChangeData(upateHoaDon);
                }
                else
                {
                    TaoMaHD();
                    string insertHD = "Insert into tblHDBan values( '" + txtMaHD.Text + "','" + maNV + "','" + dt + "','" + maKhach + "'," + tongtien + ")";
                    dataBase.ChangeData(insertHD);
                }

                foreach (DataRow row in dataTable.Rows)
                {
                    string maHang = row["Mã Hàng"].ToString();
                    int soLuong = int.Parse(row["Số Lượng"].ToString());
                    int giamGia = int.Parse(row["Giảm Giá %"].ToString());
                    int thanhTien = int.Parse(row["Thành Tiền"].ToString());
                    string findHang = "Select MaHang from tblChiTietHDBan where MaHang='" + maHang + "' and MaHDBan='" + txtMaHD.Text + "'";
                    if (dataBase.ReadData(findHang).Rows.Count > 0)
                    {
                        string updateHang = "Update tblChiTietHDBan set SoLuong='" + soLuong + "',GiamGia=" + giamGia + ",ThanhTien=" + thanhTien + " where MaHang='" + maHang + "'";
                        dataBase.ChangeData(updateHang);
                    }
                    else
                    {
                        string insertChiTietHoaDon = "INSERT INTO tblChiTietHDBan VALUES ('" + txtMaHD.Text + "', '" + maHang + "', '" + soLuong + "', '" + giamGia + "', '" + thanhTien + "')";
                        dataBase.ChangeData(insertChiTietHoaDon);
                    }
                }
                MessageBox.Show("Thêm thành công");
                reset();
            
            


        }
        public void TaoMaHD()
        {
                string maHoaDon = "HDB";
                DateTime date = DateTime.Now;
                // Tách ngày thành ngày, tháng và năm
                int day = date.Day;
                int month = date.Month;
                int year = date.Year;
                int hour = date.Hour;
                int minute = date.Minute;
                int second = date.Second;
                maHoaDon += day.ToString() + month.ToString() + year.ToString() + "_" + hour.ToString() + minute.ToString() + second.ToString();
                txtMaHD.Text = maHoaDon;
            
        }

        private void cbMaNV_SelectedIndexChanged(object sender, EventArgs e)
        {
                try
                {
                    /*string manv = cbMaNV.Text;*/
                    string manv = cbMaNV.Text;
                    DataTable dtNV = dataBase.ReadData("Select * from tblNhanVien where MaNhanVien='" + manv + "'");
                    if (dtNV.Rows.Count > 0)
                    {
                        txtTenNV.Text = dtNV.Rows[0]["TenNhanVien"].ToString();
                    }
                    else
                    {
                        MessageBox.Show("Khong thay nhan vien");
                    }
                }
                catch
                {
                    MessageBox.Show("Co loi xay ra");
                }
           
        }
        private void cbMaKhach_SelectedIndexChanged(object sender, EventArgs e)
        {
                try
                {
                    string maKhach = cbMaKhach.Text;
                    DataTable dtKhach = dataBase.ReadData("Select * from tblKhachHang where MaKhachHang='" + maKhach + "'");
                    if (dtKhach.Rows.Count > 0)
                    {
                        txtTenKhach.Text = dtKhach.Rows[0]["TenKhachHang"].ToString();
                        txtDienThoai.Text = dtKhach.Rows[0]["DienThoai"].ToString();
                        RtxtDiaChi.Text = dtKhach.Rows[0]["DiaChi"].ToString();
                    }
                   /* else
                    {
                        MessageBox.Show("Khong thay khach hang");
                    }*/
                }
                catch
                {
                    MessageBox.Show("Co loi xay ra");
                }
        }

        private void cbMaHang_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string maHang = cbMaHang.Text;
                DataTable dtHang = dataBase.ReadData("Select * from tblHang where MaHang='" + maHang + "'");
                if (dtHang.Rows.Count > 0)
                {
                    txtTenHang.Text = dtHang.Rows[0]["TenHang"].ToString();
                    txtDonGia.Text = dtHang.Rows[0]["DonGiaBan"].ToString();
                    txtSoLuong.Text = "";
                    txtGiamGia.Text = "";
                }
            }
            catch
            {
                MessageBox.Show("Co loi xay ra");
            }

        }
        private void UpdateThanhTien()
        {
            int DonGia = 0;
            int SoLuong = 0;
            int GiamGia = 0;
            if (txtSoLuong.Text.Trim() == "")
                txtSoLuong.Text = "0";
            if (txtGiamGia.Text.Trim() == "")
                txtGiamGia.Text = "0";
            if (!int.TryParse(txtSoLuong.Text, out int soluong) || !int.TryParse(txtGiamGia.Text, out int giamgia) )
            {
                MessageBox.Show("Bạn cần nhập đúng dạng thông tin");
                return;
            }
            DonGia = int.Parse(txtDonGia.Text);
            SoLuong=int.Parse(txtSoLuong.Text);
            GiamGia = int.Parse(txtGiamGia.Text);
            double ThanhTien = ((DonGia * (100 - GiamGia) / 100) * SoLuong);
            txtThanhTien.Text = ThanhTien.ToString();
        }

        private void txtSoLuong_TextChanged(object sender, EventArgs e)
        {
            UpdateThanhTien();
        }

        private void txtGiamGia_TextChanged(object sender, EventArgs e)
        {
            UpdateThanhTien();
        }
        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (
               cbMaNV.Text == "" || cbMaKhach.Text == "" ||
               cbMaHang.Text == "" || txtSoLuong.Text == "" || txtGiamGia.Text == ""
               )
            {
                MessageBox.Show("Bạn cần nhập đầy đủ thông tin");
                cbMaNV.Select();
            }
            else
            {
                int soLuongDaChon = 0;
                int slTonKho = 0;
                DataTable dtHang = dataBase.ReadData("Select SoLuong from tblHang where MaHang='" + cbMaHang.Text + "'");
                if (dtHang.Rows.Count > 0)
                {
                    slTonKho = Convert.ToInt32(dtHang.Rows[0]["SoLuong"]);
                }
                try
                {
                    string maHang = cbMaHang.Text;
                    DataRow foundRow = null;

                    foreach (DataRow row in dataTable.Rows)
                    {
                        if (row["Mã Hàng"].ToString() == maHang)
                        {
                            foundRow = row;
                            soLuongDaChon = int.Parse(row["Số Lượng"].ToString());
                            break;
                        }
                    }
                    int slTonKhoCu = slTonKho + soLuongDaChon;
                    slTonKho = slTonKho + soLuongDaChon - int.Parse(txtSoLuong.Text);
                    if (foundRow != null && slTonKho < 0)
                    {
                        MessageBox.Show("Mặt hàng hiện tại chỉ còn '" + slTonKhoCu + "'");
                    }
                    else if (foundRow != null && slTonKho > 0)
                    {
                        // Dòng đã tồn tại, cập nhật các thuộc tính của nó
                        foundRow["Số Lượng"] = int.Parse(txtSoLuong.Text);
                        foundRow["Giảm Giá %"] = int.Parse(txtGiamGia.Text);
                        foundRow["Thành Tiền"] = int.Parse(txtThanhTien.Text);
                        // ...
                    }
                    else
                    {
                        // Tạo một dòng mới và thêm nó vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["Mã Hàng"] = cbMaHang.Text;
                        newRow["Tên Hàng"] = txtTenHang.Text;
                        newRow["Số Lượng"] = int.Parse(txtSoLuong.Text);
                        newRow["Đơn Giá"] = int.Parse(txtDonGia.Text);
                        newRow["Giảm Giá %"] = int.Parse(txtGiamGia.Text);
                        newRow["Thành Tiền"] = int.Parse(txtThanhTien.Text);
                        dataTable.Rows.Add(newRow);
                    }
                    dataBase.ChangeData("update tblHang set SoLuong=" + slTonKho + " where MaHang='" + cbMaHang.Text + "'");
                    dataGr.DataSource = dataTable;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
                TongTienThanhToan();
            }
            

        }
        private void TongTienThanhToan()
        {
            tongtien = 0;
            lbTongTien.Text = "Tổng Tiền Thanh Toán:";
            foreach (DataGridViewRow row in dataGr.Rows)
            {
                // Kiểm tra xem hàng hiện tại có dữ liệu và không phải là hàng header
                if (!row.IsNewRow)
                {
                    // Truy cập dữ liệu trong từng ô cụ thể
                    string thanhTienValue = row.Cells["Thành Tiền"].Value.ToString();
                    int thanhTien;

                    if (int.TryParse(thanhTienValue, out thanhTien))
                    {
                        tongtien += thanhTien;
                    }
                    else
                    {
                        // Xử lý trường hợp không thể chuyển đổi
                        MessageBox.Show("Lỗi: Dữ liệu không hợp lệ ở ô Thành Tiền");
                    }
                }
            }
            lbTongTien.Text += tongtien.ToString();
        }

        private void dataGr_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // Lấy chỉ số dòng được chọn
            int rowIndex = e.RowIndex;
            int slTonKho = 0;
            DataTable dtHang = dataBase.ReadData("Select SoLuong from tblHang where MaHang='" + cbMaHang.Text + "'");
            if (dtHang.Rows.Count > 0)
            {
                slTonKho = Convert.ToInt32(dtHang.Rows[0]["SoLuong"]);
            }
            if (rowIndex >= 0)
            {
                // Xóa dòng từ DataTable
                DataRow selectedRow = ((DataRowView) dataGr.Rows[rowIndex].DataBoundItem).Row;
                slTonKho += int.Parse(selectedRow["Số Lượng"].ToString());
                dataBase.ChangeData("update tblHang set SoLuong=" + slTonKho + " where MaHang='" + cbMaHang.Text + "'");
                dataBase.ChangeData("delete from tblChiTietHDBan where MaHang='" + cbMaHang.Text+"'and MaHDBan='"+txtMaHD.Text+"'");
                dataTable.Rows.Remove(selectedRow);
                // Cập nhật lại DataGridView
                dataGr.DataSource = dataTable;

            }
            TongTienThanhToan();
        }
        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            cbMaHang.Text = "";
            txtTenHang.Text ="";
            txtDonGia.Text ="0";
            txtGiamGia.Text ="";
            txtSoLuong.Text ="";

            string Mahoadon = cbMaHoaDon.Text;
            txtMaHD.Text = Mahoadon;
            DataTable dt = dataBase.ReadData("Select * from tblHDBan where MaHDBan='" + Mahoadon + "'");
            if (dt.Rows.Count > 0) // Kiểm tra xem có dữ liệu trả về từ câu truy vấn hay không
            {
                cbMaNV.ValueMember = "MaNhanVien";
                cbMaNV.DisplayMember = "MaNhanVien";
                cbMaNV.DataSource = dt;

                txtTime.Value = Convert.ToDateTime(dt.Rows[0]["NgayBan"]);

                cbMaKhach.ValueMember = "MaKhachHang";
                cbMaKhach.DisplayMember = "MaKhachHang";
                cbMaKhach.DataSource = dt;
            }
            else
            {
                MessageBox.Show("Hóa đơn không hợp lệ");
            }

            /*tongtien = int.Parse(dt.Rows[0]["TongTien"].ToString());*/
            updateDataGr(Mahoadon);
            
        }
        public void updateDataGr(string Mahoadon)
        {
            dataTable.Rows.Clear();
            DataTable dtTableGr = dataBase.ReadData("Select tblChiTietHDBan.MaHang,tblHang.TenHang,tblChiTietHDBAn.SoLuong,tblHang.DonGiaBan,tblChiTietHDBan.GiamGia,tblChiTietHDBan.ThanhTien from tblChiTietHDBan inner join tblHang on tblHang.MaHang=tblChiTietHDBan.MaHang where MaHDBan='" + Mahoadon + "'");
            // Vòng lặp để sao chép từng dòng từ dtTableGr vào dataTable
            foreach (DataRow dr in dtTableGr.Rows)
            {
                DataRow newRow = dataTable.NewRow();
                newRow.ItemArray = dr.ItemArray;
                dataTable.Rows.Add(newRow);
            }
            // Cuối cùng, cập nhật lại DataGridView và tổng tiền
            dataGr.DataSource = dataTable;
            TongTienThanhToan();
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            string maHoaDon = txtMaHD.Text;
            string deleteChiTietHoaDon = "delete from tblChiTietHDBan where MaHDBan='" + maHoaDon + "'";
            string deleteHoaDon = "delete from tblHDBan where MaHDBan='" + maHoaDon + "'";
            if (txtMaHD.Text == "")
            {

                if (dataGr.Rows.Count == 0)
                {
                    MessageBox.Show("Bạn chưa thêm vật phẩm vào hóa đơn nên không cần hủy");
                }
                else
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        int sltonkho = 0;
                        string maHang = row["Mã Hàng"].ToString();

                        DataTable dt = dataBase.ReadData("SELECT SoLuong FROM tblHang WHERE MaHang='" + maHang + "'");

                        if (dt.Rows.Count > 0)
                        {
                            sltonkho = int.Parse(dt.Rows[0]["SoLuong"].ToString());
                            int soLuongMua = int.Parse(row["Số Lượng"].ToString());
                            sltonkho += soLuongMua;

                            // Cập nhật số lượng hàng trong cơ sở dữ liệu
                            dataBase.ChangeData("UPDATE tblHang SET SoLuong=" + sltonkho + " WHERE MaHang='" + maHang + "'");
                        }
                    }

                    // Xóa tất cả dữ liệu khỏi DataGridView và DataTable
                    dataGr.DataSource = null;
                    dataTable.Rows.Clear();
                    MessageBox.Show("Hủy thành công");






                }
            }
            else
            {
                if (dataGr.Rows.Count == 0)
                {
                    MessageBox.Show("Bạn chưa thêm vật phẩm vào hóa đơn nên không cần hủy");
                }
                else
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        int sltonkho = 0;
                        string maHang = row["Mã Hàng"].ToString();

                        DataTable dt = dataBase.ReadData("SELECT SoLuong FROM tblHang WHERE MaHang='" + maHang + "'");

                        if (dt.Rows.Count > 0)
                        {
                            sltonkho = int.Parse(dt.Rows[0]["SoLuong"].ToString());
                            int soLuongMua = int.Parse(row["Số Lượng"].ToString());
                            sltonkho += soLuongMua;

                            // Cập nhật số lượng hàng trong cơ sở dữ liệu
                            dataBase.ChangeData("UPDATE tblHang SET SoLuong=" + sltonkho + " WHERE MaHang='" + maHang + "'");
                        }
                    }

                    // Xóa tất cả dữ liệu khỏi DataGridView và DataTable
                    dataGr.DataSource = null;
                    dataTable.Rows.Clear();
                }
                dataBase.ChangeData(deleteChiTietHoaDon);
            
                dataBase.ChangeData(deleteHoaDon);
                updateDataGr(maHoaDon);
                MessageBox.Show("Hủy hóa đơn thành công");
            }
                reset();
            
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            Excel.Application exApp=new Excel.Application();// ứng dụng
            Excel.Workbook exWorkbook=exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);// file excel
            Excel.Worksheet exSheet= (Excel.Worksheet)exWorkbook.Worksheets[1];//1 trang tính
            Excel.Range exRange = (Excel.Range)exSheet.Cells[1, 1];// đưa con trỏ vào ô a[1,1]//1 ô
            exRange.Font.Size = 15;
            exRange.Font.Bold = true;
            exRange.Font.Color= Color.Blue;
            exRange.Value = "Trung tâm thương mại Hà Thành";
            Excel.Range diachi = (Excel.Range)exSheet.Cells[2, 1];// đưa con trỏ vào ô a[1,1]
            diachi.Font.Size = 15;
            diachi.Font.Bold = true;
            diachi.Font.Color = Color.Blue;
            diachi.Value = "Hoàng Cầu Đống Đa Hà Nội";

            // in hóa đơn bán:
            exSheet.Range["D4"].Font.Size=20;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["D4"].Font.Size = 13;
            exSheet.Range["D4"].Font.Bold = true;
            exSheet.Range["D4"].Font.Color = Color.Red;
            exSheet.Range["D4"].Value = "HÓA ĐƠN BÁN HÀNG";
            exSheet.Range["D4"].ColumnWidth =20;
            exSheet.Range["A5:A8"].Font.Size = 12;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["A5"].Value= "Mã Hóa Đơn: "+txtMaHD.Text;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["A6"].Value= "Khách Hàng: "+cbMaKhach.Text+"-"+txtTenKhach.Text;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["A7"].Value= "Địa Chỉ: "+ RtxtDiaChi.Text;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["A8"].Value= "Điện thoại: "+ txtDienThoai.Text;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["A10:G10"].Font.Size=12;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["A10:G10"].Font.Bold=true;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["C10"].ColumnWidth = 25;
            exSheet.Range["G10"].ColumnWidth = 25;
            exSheet.Range["E10"].ColumnWidth = 20;
            exSheet.Range["A10"].Value="STT";// đưa con trỏ vào ô a[1,1]
            exSheet.Range["B10"].Value="Mã Hàng";// đưa con trỏ vào ô a[1,1]
            exSheet.Range["C10"].Value="Tên Hàng";// đưa con trỏ vào ô a[1,1]
            exSheet.Range["D10"].Value="Số Lượng";// đưa con trỏ vào ô a[1,1]
            exSheet.Range["E10"].Value="Đơn Giá Bán";// đưa con trỏ vào ô a[1,1]
            exSheet.Range["F10"].Value="Giảm Giá";// đưa con trỏ vào ô a[1,1]
            exSheet.Range["F10"].ColumnWidth=40;// đưa con trỏ vào ô a[1,1]
            exSheet.Range["G10"].Value="Thành Tiền";// đưa con trỏ vào ô a[1,1]
            int dong = 11;
            //in danh sách các chi tiết sản phẩm 
            for(var i = 0; i < dataGr.Rows.Count - 1; i++)
            {
                exSheet.Range["A"+(dong+i).ToString()].Value=(i+1).ToString();// bắt đầu in data vào excel từ dòng thứ dòng +i và số thứ tự vì i chạy từ 0 nên phải cộng thêm 1
                exSheet.Range["B" + (dong + i).ToString()].Value = dataGr.Rows[i].Cells[0].Value.ToString();           
                exSheet.Range["C" + (dong + i).ToString()].Value = dataGr.Rows[i].Cells[1].Value.ToString();            
                exSheet.Range["D" + (dong + i).ToString()].Value = dataGr.Rows[i].Cells[2].Value.ToString();            
                exSheet.Range["E" + (dong + i).ToString()].Value = dataGr.Rows[i].Cells[3].Value.ToString();            
                exSheet.Range["F" + (dong + i).ToString()].Value = dataGr.Rows[i].Cells[4].Value.ToString()+" %";            
                exSheet.Range["G" + (dong + i).ToString()].Value = dataGr.Rows[i].Cells[5].Value.ToString()+" Đồng";            
            }
            dong = dong + dataGr.Rows.Count;
            exSheet.Range["F" + dong.ToString()].Value = lbTongTien.Text + " Đồng";
            exSheet.Name = txtMaHD.Text;
            exWorkbook.Activate();// kích hoạt cho file excel hoạt động
            // luu file
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel 365 .xls|*.xls|Excel 365 .xlsx|*.xlsx|All Files|*.*";
            saveFileDialog.FilterIndex = 2;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                exWorkbook.SaveAs(saveFileDialog.FileName.ToLower());// save file 
            MessageBox.Show("In thành công");
            }
            exApp.Quit();
           

            



        }
    }
}
