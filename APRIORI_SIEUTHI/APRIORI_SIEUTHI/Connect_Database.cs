using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace APRIORI_SIEUTHI
{
    class Connect_Database
    {
        private static SqlConnection ket_noi;

        static Connect_Database()
        {
            string chuoi_ket_noi = @"Data Source=.\SQLEXPRESS;Initial Catalog=DBSIEUTHI_new;Integrated Security=True;Pooling=False";

            try
            {
                ket_noi = new SqlConnection(chuoi_ket_noi);
            }
            catch (Exception loi)
            {
                MessageBox.Show("Error: " + loi.Message);
            }

            ket_noi.Close();
        }

        public static DataTable Doc_Bang(string lenh)
        {

            DataTable dt = new DataTable();

            try
            {
                if (ket_noi.State == ConnectionState.Closed)
                    ket_noi.Open();
                SqlDataAdapter sda = new SqlDataAdapter(lenh, ket_noi);
                sda.FillSchema(dt, SchemaType.Source);
                sda.Fill(dt);

            }
            catch (Exception loi)
            {
                MessageBox.Show("Error: " + loi.Message);
            }

            ket_noi.Close();
            return dt;

        }

        public static bool Ghi_Bang(string lenh)
        {
            bool kq = false;

            try
            {
                if (ket_noi.State == ConnectionState.Closed)
                    ket_noi.Open();
                SqlCommand cmd = new SqlCommand(lenh, ket_noi);
                cmd.ExecuteNonQuery();
                kq = true;
            }
            catch (Exception loi)
            {
                MessageBox.Show("Loi: " + loi.Message + lenh);

            }

            ket_noi.Close();
            return kq;
        }
    }
}
