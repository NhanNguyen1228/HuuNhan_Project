using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;


namespace APRIORI_SIEUTHI
{
    public partial class FormMain : MetroFramework.Forms.MetroForm
    {
        public FormMain()
        {
            InitializeComponent();
        }

       

        private void FormMain_Load(object sender, EventArgs e)
        {
           
        }

        private void btnChonFile_Click(object sender, EventArgs e)
        {
            //Xử lý load file excel.
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Files (.xls)|*.*";
            OleDbConnection conn;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string fileName;
                fileName = dialog.FileName; //Lấy tên file.
                txtDuongDan.Text = fileName;
                string path = fileName;
                if (Path.GetExtension(path) == ".xls") //Tạo kết nối dạng Microsoft.Jet.OLEDB
                {
                    conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"");
                }
                else
                {
                    conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");
                }

                OleDbCommand cmd;
                OleDbDataAdapter da = new OleDbDataAdapter(); //Tạo cầu nối đến file excel.
                DataSet ds = new DataSet(); //Lưu dữ liệu từ file excel xuống Dataset.
                DataSet ds1 = new DataSet();
                string query = "";
                query = "select * from [HDBH$]"; //Tên sheet excel cần lấy dữ liệu hóa đơn.
                string queryn = "select * from [DLSP$]"; //Tên sheet excel cần lấy dữ liệu nhóm.
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open(); //Cập nhật vào file.
                }

                try
                {
                    cmd = new OleDbCommand(query, conn);
                    da = new OleDbDataAdapter(cmd);
                    da.Fill(ds);
                    dgvDanhSachHoaDon.DataSource = ds.Tables[0]; //Hiển thị dữ liêu excel lên DataGridView.
                    //Đọc nhóm dữ liệu
                    cmd = new OleDbCommand(queryn, conn);
                    da = new OleDbDataAdapter(cmd);
                    da.Fill(ds1);
                    dgvDanhMucSP.DataSource = ds1.Tables[0];
                }
                catch (Exception loi)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Lỗi: " + loi.Message);

                }
                finally
                {
                    da.Dispose();
                    conn.Close();
                }
            }
        }

        private void btnImportDuLieu_Click(object sender, EventArgs e)
        {
            //Lấy dữ liệu từ Grid.
            DataTable dt = (DataTable)dgvDanhSachHoaDon.DataSource;
            DataTable dt1 = (DataTable)dgvDanhMucSP.DataSource;

            if (dt == null && dt1 == null)
                return;

            string mahd = "";
            string ngaylap = "";
            string masp = "";
            string tensp = "";
            string soluong = "";
            string nhom = "";
         
            int dem = 0;
            for (int k = 0; k < dt1.Columns.Count; k = k + 2) //Duyệt qua từng cột của dữ liệu.
            {
                for (int i = 0; i < dt1.Rows.Count; i++) //Duyệt qua từng dòng của dữ liệu.
                {
                    masp = dt1.Rows[i][k + 1].ToString().Trim();
                    tensp = dt1.Rows[i][k].ToString().Replace('\'', ' ').Trim();

                    if (masp != "" && tensp != "")
                    {
                        string[] anpha = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };
                        nhom = anpha[dem];
                        //Lưu trữ dữ liệu sản phẩm vào Database.
                        string lenh = "insert into tblSanPham(masp, tensp, Manhom) values('" + masp + "',N'" + tensp + "','" + nhom + "')";
                        Connect_Database.Ghi_Bang(lenh);
                        if (i == dt1.Rows.Count - 1)
                            dem++;
                    }
                    else if (masp == "" && tensp == "")
                    {
                        i = dt1.Rows.Count;
                        dem++;
                    }
                }
            }

            for (int i = 3; i < dt.Rows.Count; i++)
            {
                mahd = dt.Rows[i][0].ToString();
                ngaylap = dt.Rows[i][4].ToString();
                masp = dt.Rows[i][2].ToString();
                tensp = dt.Rows[i][1].ToString().Replace('\'', ' ');
                soluong = dt.Rows[i][3].ToString();
                string lenh = "insert into tblHoaDon(mahd, ngaylap) values('" + mahd + "','" + ngaylap + "')"; //lưu trữ dữ liệu hóa đơn vào Database
                DataTable dtkthd = Connect_Database.Doc_Bang("select * from tblHoaDon where mahd='" + mahd + "'");
                if (dtkthd.Rows.Count < 1)
                    Connect_Database.Ghi_Bang(lenh); //Lưu trữ chi tiết hóa đơn vào database.
                lenh = "insert into tblCTHoaDon(mahd, ngaylap, masp, tensp, soluong) values('" + mahd + "','" + ngaylap + "','" + masp + "',N'" + tensp + "','" + soluong + "')";
                Connect_Database.Ghi_Bang(lenh);
            }

            MetroFramework.MetroMessageBox.Show(this, "Đã lưu dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            dt = Connect_Database.Doc_Bang("select * from tblCTHoaDon"); //Đọc bảng tblCTHoaDon.
            dgvDanhSachHoaDon.DataSource = dt;
            dt1 = Connect_Database.Doc_Bang("select * from tblSanPham"); //Đọc bảng tblSanPham.
            dgvDanhMucSP.DataSource = dt1;
        }

        private void btnXoaDuLieu_Click(object sender, EventArgs e)
        {
            Connect_Database.Ghi_Bang("Delete from tblHoaDon");
            Connect_Database.Ghi_Bang("Delete from tblSanPham");
            Connect_Database.Ghi_Bang("Delete from tblCTHoaDon");

            MetroFramework.MetroMessageBox.Show(this, "Đã xoá dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            DataTable dt = Connect_Database.Doc_Bang("select * from tblCTHoaDon"); //Đọc bảng tblCTHoaDon.
            DataTable dt1 = Connect_Database.Doc_Bang("select * from tblSanPham"); //Đọc bảng tblSanPham.

            dgvDanhSachHoaDon.DataSource = dt;
            dgvDanhMucSP.DataSource = dt1;
        }

        private void btnLoadDulieu_Click(object sender, EventArgs e)
        {
            DataTable dt = Connect_Database.Doc_Bang("select * from tblCTHoaDon"); //Đọc bảng tblCTHoaDon.
            dgvLoadDSHoaDon.DataSource = dt;
            lblTongDuLieu.Text = "" + dt.Rows.Count;   
            
            // Đặt tên cột.         
            dgvLoadDSHoaDon.Columns[0].HeaderText = "Mã hoá đơn";
            dgvLoadDSHoaDon.Columns[1].HeaderText = "Tên sản phẩm";
            dgvLoadDSHoaDon.Columns[2].HeaderText = "Mã sản phẩm";
            dgvLoadDSHoaDon.Columns[3].HeaderText = "Số lượng";
            dgvLoadDSHoaDon.Columns[4].HeaderText = "Ngày lập hoá đơn";

            DataTable dt1 = Connect_Database.Doc_Bang("Select masp, tensp, manhom from tblSanPham"); //Đọc bảng tblSanPham.
            dgvLoadDanhMucSP.DataSource = dt1;

            // Đặt tên cột.
            dgvLoadDanhMucSP.Columns[0].HeaderText = "Mã sản phẩm";
            dgvLoadDanhMucSP.Columns[1].HeaderText = "Tên sản phẩm";
            dgvLoadDanhMucSP.Columns[2].HeaderText = "Mã nhóm";
        }

        List<string> mang_luu;

        private void btnApriori_Click(object sender, EventArgs e)
        {
            mlvTapPhoBien.Clear();
            mlvTapLuatKH.Clear();
            mlvDanhSachMH.Clear();

            mang_luu = new List<string>();
            if ((DataTable)dgvLoadDSHoaDon.DataSource == null)
            {
                MetroFramework.MetroMessageBox.Show(this, "Vui lòng import dữ liệu.");
            }
            else
            {
                DataTable dt = ((DataTable)dgvLoadDSHoaDon.DataSource).Clone();
                dt.Columns.Add("tam", typeof(string));
             
                float minsup = 0;
                float minconf = 0;

                if (txtMinSupp.Text.Trim() == "")
                {
                    MetroFramework.MetroMessageBox.Show(this, "Vui lòng nhập độ phổ biến.");
                    return;
                }

                minsup = (float)(float.Parse(txtMinSupp.Text) / 100.0);

                if (txtMinConf.Text.Trim() == "")
                {
                    MetroFramework.MetroMessageBox.Show(this, "Vui lòng nhập độ hổ trợ.");
                    return;
                }

                minconf = (float)(float.Parse(txtMinConf.Text) / 100.0);

                if (float.Parse(txtMinSupp.Text) > float.Parse(txtMinConf.Text))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Không có luật kết hợp thỏa điều kiện.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    lblTongTLKH.Text = "0";
                    lblTongTPB.Text = "0";
                    return;
                }

                //Bắt đầu thuật toán Apriori.
                List<string> item = new List<string>();
                DataTable dtdata = Connect_Database.Doc_Bang("Select Distinct manhom From tblSanPham");
                for (int i = 0; i < dtdata.Rows.Count; i++)
                {
                    item.Add(dtdata.Rows[i]["manhom"].ToString().Trim());
                }

                //xử lý kết nối các Transaction
                List<string> trans = new List<string>();
                DataTable dtnguon = Connect_Database.Doc_Bang("select Distinct mahd from tblCTHoaDon");
                for (int i = 0; i < dtnguon.Rows.Count; i++)
                {
                    string mahd = dtnguon.Rows[i]["mahd"].ToString().Trim();
                    DataTable dthd = Connect_Database.Doc_Bang("select tblCTHoaDon.mahd, tblCTHoaDon.masp, tblSanPham.manhom from tblCTHoaDon, tblSanPham where tblSanPham.masp=tblCTHoaDon.masp and tblCTHoaDon.mahd='" + mahd + "'");
                    string str = "";
                    for (int j = 0; j < dthd.Rows.Count; j++)
                    {
                        if (!str.Contains(dthd.Rows[j]["manhom"].ToString().Trim())) //kiêm tra nhóm
                            str = str + dthd.Rows[j]["manhom"].ToString().Trim();
                    }
                    trans.Add(str);
                }

                //xứ lý thuật toán Aprirori
                ItemsetCollection db = new ItemsetCollection();
                for (int i = 0; i < trans.Count(); i++)
                {
                    string aa = trans.ElementAt(i).Trim();
                    //string[] mang =new string[aa.Length];
                    Itemset itemtraining = new Itemset();
                    for (int j = 0; j < aa.Length; j++)
                        itemtraining.Add(aa[j].ToString());
                    // mang[j] = aa[j].ToString();
                    db.Add(itemtraining);
                }
                mlvTapPhoBien.View = View.Details;
                mlvTapPhoBien.GridLines = true;
                mlvTapPhoBien.FullRowSelect = true;

                //Thêm tiêu đề cho cột.
                mlvTapPhoBien.Columns.Add("Tập phổ biến", 205);
                mlvTapPhoBien.Columns.Add("Min_supp", 90);
                Itemset uniqueItems = db.GetUniqueItems();
                ItemsetCollection L = AprioriMining.DoApriori(db, minsup);
                Console.Write("\n" + L.Count + " itemsets in L\n");
                foreach (Itemset i in L)
                {
                    string s = i.ToLuat();
                    string[] a = s.Split(':');
                    if (a[1].Trim() != "" && float.Parse(a[1]) >= minsup*100)
                    {
                        ListViewItem lstItem1 = new ListViewItem();
                        string ss = Xu_ly_chuoi(a[0].Trim());
                       
                        lstItem1.SubItems[0].Text = ss;
                        lstItem1.SubItems.Add(a[1]+"%");
                        mlvTapPhoBien.Items.Add(lstItem1);
                        
                    }

                    //Đếm tổng số tập phổ biến trong danh mục sản phẩm.
                    lblTongTPB.Text = "" + mlvTapPhoBien.Items.Count;
                }

                mlvTapLuatKH.View = View.Details;
                mlvTapLuatKH.GridLines = true;
                mlvTapLuatKH.FullRowSelect = true;

                //Thêm tiêu đề cho cột
                mlvTapLuatKH.Columns.Add("Tập luật kết hợp", 300);
                mlvTapLuatKH.Columns.Add("Min_conf", 70);
                //test mining
                List<AssociationRule> allRules = AprioriMining.Mine(db, L, minconf);
                Console.Write("\n" + allRules.Count + " rules\n");

                bool kq_kt = false;

                foreach (AssociationRule rule in allRules)
                {
                    string s = rule.ToString();
                    string[] a = s.Split(':');
                    if (a[2] != "" && float.Parse(a[1])>=minsup*100 && float.Parse(a[2]) >= minconf*100)
                    {
                        string ss = "{";
                        a[0] = a[0].Replace('{', ' ');
                        a[0] = a[0].Replace('}', ' ');
                        a[0] = a[0].Trim();
                        string[] aaa = a[0].Split('-');
                        string sss1 = Xu_ly_chuoi(aaa[0]);
                        string sss2 = Xu_ly_chuoi(aaa[1]);
                        ss = sss1 + " => " + sss2;
                        mang_luu.Add(a[0]);
                        ListViewItem lstItem1 = new ListViewItem();
                        lstItem1.SubItems[0].Text = ss;
                        lstItem1.SubItems.Add(a[2] + "%");
                        mlvTapLuatKH.Items.Add(lstItem1);

                        //Đếm tổng số luật kết hợp trong danh mục sản phẩm.
                        lblTongTLKH.Text = "" + mlvTapLuatKH.Items.Count;

                        kq_kt = true;
                    }

                  
                    //Console.Write(rule + "\n");
                }

                if (!kq_kt)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Không có luật kết hợp thỏa điều kiện.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                }

            }
        }

        private string Xu_ly_chuoi(string chuoi)
        {
            string ss = "{";
            for (int k = 0; k < chuoi.Length; k++)
            {
                string sss = Doi_Sang_Ten_Nhom(chuoi[k].ToString()).Trim();
                if (sss != "")
                    ss = ss + sss + ", ";
            }
            ss = ss.Substring(0, ss.Length - 2) + "}";
            return ss;
        }

        private void Luat_Ket_Hop_Tung_Mat_Hang(float minsup, float minconf, string chuoi)
        {
            mlvDanhSachMH.Items.Clear();
            chuoi = chuoi.Trim();
            //Tách từng nhóm.
            string where = "'" + chuoi[0].ToString();
            for (int i = 1; i < chuoi.Length; i++)
            {
                if(chuoi[i]!=' ' && chuoi[i]!='-')
                    where = where + "','" + chuoi[i].ToString();
            }
            where = where + "'";

            string lenh = "Select Distinct tblCTHoaDon.mahd from tblSanPham, tblCTHoaDon where tblCTHoaDon.masp = tblSanPham.masp and tblSanPham.manhom in(" + where + ")";
            DataTable dtHD = Connect_Database.Doc_Bang(lenh);

            lenh = "Select Distinct tblSanPham.masp from tblSanPham,tblCTHoaDon where tblCTHoaDon.masp=tblSanPham.masp and tblSanPham.manhom in(" + where + ")";
            DataTable dtMasp = Connect_Database.Doc_Bang(lenh);

            dtMasp.Columns.Add("tam", typeof(string));
            string[] anpha = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"};
            for (int i = 0; i < dtMasp.Rows.Count; i++)
            {
                dtMasp.Rows[i]["tam"] = anpha[i%anpha.Length];
            }

           
            //Xử lý kết nối các Transaction.
            List<string> trans = new List<string>();
            for (int i = 0; i < dtHD.Rows.Count; i++)
            {
                string mahd = dtHD.Rows[i]["mahd"].ToString().Trim();
                DataTable dthd = Connect_Database.Doc_Bang("select tblCTHoaDon.mahd, tblCTHoaDon.masp from tblCTHoaDon, tblSanPham where tblSanPham.masp = tblCTHoaDon.masp and tblCTHoaDon.mahd = '" + mahd + "'");
                string str = "";
                for (int j = 0; j < dthd.Rows.Count; j++)
                {
                    for (int k = 0; k < dtMasp.Rows.Count; k++)
                        if (dtMasp.Rows[k]["masp"].ToString().Trim() == dthd.Rows[j]["masp"].ToString().Trim())
                            if (!str.Contains(dtMasp.Rows[k]["tam"].ToString().Trim()))
                            {
                                str = str + dtMasp.Rows[k]["tam"].ToString().Trim();
                                k = dtMasp.Rows.Count;
                            }
                }
                trans.Add(str);
            }

            //xứ lý thuật toán Aprirori
            ItemsetCollection db = new ItemsetCollection();
            for (int i = 0; i < trans.Count(); i++)
            {
                string aa = trans.ElementAt(i).Trim();
                //string[] mang =new string[aa.Length];
                Itemset itemtraining = new Itemset();
                for (int j = 0; j < aa.Length; j++)
                    itemtraining.Add(aa[j].ToString());
                // mang[j] = aa[j].ToString();
                db.Add(itemtraining);
            }

            
            Itemset uniqueItems = db.GetUniqueItems();
            ItemsetCollection L = AprioriMining.DoApriori(db, minsup);
            
            mlvDanhSachMH.View = View.Details;
            mlvDanhSachMH.GridLines = true;
            mlvDanhSachMH.FullRowSelect = true;

            //Thêm tiêu đề cho cột.
            mlvDanhSachMH.Columns.Add("Tập luật kết hợp", 450);
            mlvDanhSachMH.Columns.Add("Min_conf", 70);
            //test mining
            List<AssociationRule> allRules = AprioriMining.Mine(db, L, minconf);
            Console.Write("\n" + allRules.Count + " rules\n");

            bool kq_kt = false;

            foreach (AssociationRule rule in allRules)
            {
                string s = rule.ToString();
                string[] a = s.Split(':');
                if (a[2] != "" && float.Parse(a[1]) >= minsup && float.Parse(a[2]) >= minconf)
                {
                    string ss = "{";
                    a[0] = a[0].Replace('{', ' ');
                    a[0] = a[0].Replace('}', ' ');
                    a[0] = a[0].Trim();
                    string[] aaa = a[0].Split('-');
                    string sss1 = Doi_Sang_Masp(dtMasp,aaa[0]);
                    string sss2 =Doi_Sang_Masp(dtMasp, aaa[1]);
                    if (sss1.Trim() != "" && sss2.Trim() != "")
                    {
                        ss = sss1 + " => " + sss2;

                        ListViewItem lstItem1 = new ListViewItem();
                        lstItem1.SubItems[0].Text = ss;
                        lstItem1.SubItems.Add(a[2] + "%");
                        mlvDanhSachMH.Items.Add(lstItem1);
                    }

                    //Đếm tổng số luật các mặt hàng gần nhau.
                    lblTongDSMH.Text = "" + mlvDanhSachMH.Items.Count;

                    kq_kt = true;

                }
                //Console.Write(rule + "\n");
            }

            if (!kq_kt)
            {
                MetroFramework.MetroMessageBox.Show(this, "Không có luật kết hợp thỏa điều kiện", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                lblTongDSMH.Text = "0";                
            }
        }

        private DataTable Luat_Ket_Hop_Tung_Mat_Hang_Luu_Chart(float minsup, float minconf, string chuoi)
        {
            //mlvDanhSachMH.Items.Clear();
            chuoi = chuoi.Trim();
            //Tách từng nhóm.
            string where = "'" + chuoi[0].ToString();
            for (int i = 1; i < chuoi.Length; i++)
            {
                if (chuoi[i] != ' ' && chuoi[i] != '-')
                    where = where + "','" + chuoi[i].ToString();
            }
            where = where + "'";

            string lenh = "Select Distinct tblCTHoaDon.mahd from tblSanPham, tblCTHoaDon where tblCTHoaDon.masp = tblSanPham.masp and tblSanPham.manhom in(" + where + ")";
            DataTable dtHD = Connect_Database.Doc_Bang(lenh);
            lenh = "Select Distinct tblSanPham.masp from tblSanPham,tblCTHoaDon where tblCTHoaDon.masp=tblSanPham.masp and tblSanPham.manhom in(" + where + ")";
            DataTable dtMasp = Connect_Database.Doc_Bang(lenh);
            dtMasp.Columns.Add("tam", typeof(string));
            string[] anpha = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            for (int i = 0; i < dtMasp.Rows.Count; i++)
            {
                dtMasp.Rows[i]["tam"] = anpha[i];
            }

            //Xử lý kết nối các Transaction.
            List<string> trans = new List<string>();
            for (int i = 0; i < dtHD.Rows.Count; i++)
            {
                string mahd = dtHD.Rows[i]["mahd"].ToString().Trim();
                DataTable dthd = Connect_Database.Doc_Bang("select tblCTHoaDon.mahd, tblCTHoaDon.masp from tblCTHoaDon, tblSanPham where tblSanPham.masp = tblCTHoaDon.masp and tblCTHoaDon.mahd = '" + mahd + "'");
                string str = "";
                for (int j = 0; j < dthd.Rows.Count; j++)
                {
                    for (int k = 0; k < dtMasp.Rows.Count; k++)
                        if (dtMasp.Rows[k]["masp"].ToString().Trim() == dthd.Rows[j]["masp"].ToString().Trim())
                            if (!str.Contains(dtMasp.Rows[k]["tam"].ToString().Trim()))
                            {
                                str = str + dtMasp.Rows[k]["tam"].ToString().Trim();
                                k = dtMasp.Rows.Count;
                            }
                }
                trans.Add(str);
            }

            //xứ lý thuật toán Apriori
            ItemsetCollection db = new ItemsetCollection();
            for (int i = 0; i < trans.Count(); i++)
            {
                string aa = trans.ElementAt(i).Trim();
                //string[] mang =new string[aa.Length];
                Itemset itemtraining = new Itemset();
                for (int j = 0; j < aa.Length; j++)
                    itemtraining.Add(aa[j].ToString());
                // mang[j] = aa[j].ToString();
                db.Add(itemtraining);
            }


            Itemset uniqueItems = db.GetUniqueItems();
            ItemsetCollection L = AprioriMining.DoApriori(db, minsup);

            //mlvDanhSachMH.View = View.Details;
            //mlvDanhSachMH.GridLines = true;
            //mlvDanhSachMH.FullRowSelect = true;
            DataTable dtkq = new DataTable();
            dtkq.Columns.Add("Luat", typeof(string));
            dtkq.Columns.Add("Confidence", typeof(string));
            ////Thêm tiêu đề cho cột.
            //mlvDanhSachMH.Columns.Add("Tập luật kết hợp", 450);
            //mlvDanhSachMH.Columns.Add("Min_conf", 70);
            //test mining
            List<AssociationRule> allRules = AprioriMining.Mine(db, L, minconf);
            Console.Write("\n" + allRules.Count + " rules\n");
            foreach (AssociationRule rule in allRules)
            {
                string s = rule.ToString();
                string[] a = s.Split(':');
                if (a[2] != "" && float.Parse(a[1]) >= minsup && float.Parse(a[2]) >= minconf)
                {
                    string ss = "{";
                    a[0] = a[0].Replace('{', ' ');
                    a[0] = a[0].Replace('}', ' ');
                    a[0] = a[0].Trim();
                    string[] aaa = a[0].Split('-');
                    string sss1 = Doi_Sang_Masp(dtMasp, aaa[0]);
                    string sss2 = Doi_Sang_Masp(dtMasp, aaa[1]);
                    if (sss1.Trim() != "" && sss2.Trim() != "")
                    {
                        ss = sss1 + " => " + sss2;
                        DataRow dr = dtkq.NewRow();
                        dr[0] = ss;
                        dr[1] =float.Parse(a[2]);
                        dtkq.Rows.Add(dr);
                        //ListViewItem lstItem1 = new ListViewItem();
                        //lstItem1.SubItems[0].Text = ss;
                        //lstItem1.SubItems.Add(a[2] + "%");
                        //mlvDanhSachMH.Items.Add(lstItem1);
                    }

                }
                //Console.Write(rule + "\n");
            }

            return dtkq;

        }

        private string Doi_Sang_Ten_Nhom(string chuoi)
        {
            string kq = "";
            if (chuoi.Trim() != "," && chuoi.Trim()!="")
            {
                DataTable dt = Connect_Database.Doc_Bang("Select tennhom from tblNhom where manhom='" + chuoi + "'");
                kq = dt.Rows[0]["tennhom"].ToString().Trim();
            }
            return kq;
        }

        private string Doi_Sang_Masp(DataTable dt, string chuoi)
        {
            string kq = "";
            for (int i = 0; i < dt.Rows.Count; i++)
                if (dt.Rows[i]["tam"].ToString().Trim() == chuoi.Trim())
                {
                    kq = dt.Rows[i]["masp"].ToString().Trim();
                    DataTable dtdata = Connect_Database.Doc_Bang("Select tensp from tblCTHoaDon where masp='" + kq + "'");
                    kq = dtdata.Rows[0]["tensp"].ToString();
                    i = dt.Rows.Count;
                }
            return kq;
        }

        private void mlvTapLuatKH_SelectedIndexChanged(object sender, EventArgs e)
        {
            mlvDanhSachMH.Clear();

            for (int i = 0; i < mlvTapLuatKH.SelectedItems.Count; i++)
            {
                int vitri = mlvTapLuatKH.SelectedIndices[i];

                float minsup = float.Parse(txtMinSupp.Text); //minsup tối thiểu
                // minsup = 5;

                float minconf = float.Parse(txtMinConf.Text);
                // minconf = 3;

                // float minconf = (float)(0.5); //minconf tối thiểu

                //if (txtMinConfDS.Text.Trim() == "")
                //{
                //    MetroFramework.MetroMessageBox.Show(this, "Vui lòng nhập độ tin cậy.");
                //    return;
                //}

                try
                {
                    Luat_Ket_Hop_Tung_Mat_Hang(minsup, minconf, mang_luu[vitri]);
                }
                catch (Exception)
                {

                }
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            //Tạo file text LuatKH_Nhom.txt
            FileStream fs = new FileStream("LuatKH_Nhom.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);

            foreach (ListViewItem itemset in mlvTapLuatKH.Items)
            {
                sw.WriteLine(itemset.SubItems[0].Text + "\tConf: " + itemset.SubItems[1].Text);
            }

            sw.Close();
            fs.Close();

            try
            {          
                for (int k = 0; k < mlvTapLuatKH.SelectedItems.Count; k++)
                {                    
                    FileStream fs1 = new FileStream("LuatKH_SanPham_Nhom" + mang_luu[mlvTapLuatKH.SelectedIndices[k]] + ".txt", FileMode.Create);
                    StreamWriter sw1 = new StreamWriter(fs1);

                    foreach (ListViewItem itemset1 in mlvDanhSachMH.Items)
                    {
                        sw1.WriteLine(itemset1.SubItems[0].Text + "\tConf: " + itemset1.SubItems[1].Text);         
                    }

                    sw1.Close();
                    fs1.Close();

                    MetroFramework.MetroMessageBox.Show(this, "Lưu luật kết hợp thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
              
                Ghi_data_Chart();
            }
            catch (Exception err)
            {
                MetroFramework.MetroMessageBox.Show(this, "Lưu không thành công, vui lòng kiểm tra lại.");
            }
        }

        private void Ghi_data_Chart()
        {
            DataTable dtchart = new DataTable();
            dtchart.Columns.Add("Luat", typeof(string));
            dtchart.Columns.Add("Confidence", typeof(float));
            
            
            for (int i = 0; i < mlvTapLuatKH.Items.Count; i++)
            {
                int vitri = i;//mlvTapLuatKH.GetItemAt(i,0);//.SelectedIndices[i];
                float minsup = float.Parse(txtMinSupp.Text); //minsup tối thiểu
                // float minconf = (float)(0.5); //minconf tối thiểu

                //if (txtMinConfDS.Text.Trim() == "")
                //{
                //    MetroFramework.MetroMessageBox.Show(this, "Vui lòng nhập độ tin cậy.");
                //    return;
                //}

                float minconf = float.Parse(txtMinConf.Text);
                // minsup = 8;

                try
                {
                  
                  DataTable dtkq = Luat_Ket_Hop_Tung_Mat_Hang_Luu_Chart(minsup, minconf, mang_luu[vitri]);
                  for (int v = 0; v < dtkq.Rows.Count; v++)
                  {
                      DataRow dr = dtchart.NewRow();
                      dr[0] = dtkq.Rows[v]["Luat"];
                      dr[1] = dtkq.Rows[v]["Confidence"];
                      dtchart.Rows.Add(dr);
                  }
                  dtkq.TableName = "Charts" + i;
                  dtkq.WriteXml("dtcharts" + i + ".xml");
                }
                catch (Exception)
                {
                    
                }
               
            }
            if (dtchart != null && dtchart.Rows.Count>0)
            {
                dtchart.TableName = "Chart";
                dtchart.WriteXml("dtcharts.xml");
            }
        }
        private void metroTabControl1_Click(object sender, EventArgs e)
        {
            
            if (metroTabControl1.SelectedIndex == 2)
            {
                listBox1.Items.Clear();
                FileStream fs1 = new FileStream("LuatKH_Nhom.txt", FileMode.Open);
                StreamReader sr = new StreamReader(fs1);
                string line = "";
                while ((line = sr.ReadLine()) != null)
                {
                    listBox1.Items.Add(line);
                }
                sr.Close();
                fs1.Close();
                listBox1.Items.Add("All");
                chart1.Series.Clear();
                chart1.Series.Add("Confidences");
                chart1.Visible = true;
                try
                {
                    DataSet ds = new DataSet();
                    ds.ReadXml("dtcharts.xml");
                    chart1.DataSource = ds.Tables[0];
                    chart1.Series["Confidences"].Points.DataBindXY(ds.Tables[0].DefaultView, "Luat", ds.Tables[0].DefaultView, "Confidence");
                    chart1.DataBind();
                    chart1.Series["Confidences"].IsValueShownAsLabel = true;
                }
                catch(Exception loi)
                {

                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {         
            if(listBox1.SelectedItem.ToString().Trim()=="All")
            {
                chart1.Series.Clear();
                chart1.Series.Add("Confidences");
                chart1.Visible = true;
                DataSet ds = new DataSet();
                ds.ReadXml("dtcharts.xml");
                chart1.DataSource = ds.Tables[0];
                //chart1.Series["Series1"].YValueMembers = "Confidence";
                //chart1.Series["Series1"].XValueMember = "Tensp";
                chart1.Series["Confidences"].Points.DataBindXY(ds.Tables[0].DefaultView, "Luat", ds.Tables[0].DefaultView, "Confidence");

                chart1.DataBind();
                chart1.Series["Confidences"].IsValueShownAsLabel = true;
            }
            else
            {
                int vitri = listBox1.SelectedIndex;
                chart1.Series.Clear();
                chart1.Series.Add("Confidences");
                chart1.Visible = true;
                try
                {
                    DataSet ds = new DataSet();
                    ds.ReadXml("dtcharts" + vitri + ".xml");
                    chart1.DataSource = ds.Tables[0];
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        ds.Tables[0].Rows[i][1] = ds.Tables[0].Rows[i][1].ToString().Replace(',', '.');
                    }
                    chart1.Series["Confidences"].Points.DataBindXY(ds.Tables[0].DefaultView, "Luat", ds.Tables[0].DefaultView, "Confidence");

                    chart1.DataBind();
                    chart1.Series["Confidences"].IsValueShownAsLabel = true;
                }
                catch(Exception loi)
                {

                }
            }
        }
    }
}
