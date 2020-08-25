using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Guna;

namespace certificateSystem
{
    class CRUD
    {
        ///   AppConfigurationManager.ConnectionStrings["UniversityConnectionString"].ConnectionString;
           public static String Path = ConfigurationSettings.AppSettings["connectionString_Local"];
        //ConfigurationSettings.AppSettings["connectionString_Local"];
        //@"Data Source = DESKTOP-SHF8FLP\SQLEXPRESS;Initial Catalog = certificate; Integrated Security = True";
        //  public static String Path = @"Data Source=DESKTOP-SHF8FLP\SQLEXPRESS;Initial Catalog=certificate;Integrated Security=True";
        SqlConnection con = new SqlConnection(Path);

  
        
            public void CloseDB()
        {

            con.Close();
        }

        public void OpenBb()
        {

            con.Open();
        }
        public Boolean getDataGrid(string mysql, DataGridView Grid)
        {
              con.Open();
            using (SqlDataAdapter adapter = new SqlDataAdapter(mysql, con))
            {

                DataTable table = new DataTable();
                adapter.Fill(table);
                Grid.DataSource = table;
                Grid.Show();
               
                 con.Close();

                return true;




            }





        }

        public SqlDataReader getDataFromDataDase(string mysql)
        {
            SqlDataReader dr;
            using (SqlCommand cmd = new SqlCommand(mysql, con))
            {
                con.Open();
                dr = cmd.ExecuteReader();
             
                return dr;
            }

        }



        public int InsertUpdateDelete(string mysql)
        {
            int rtn = 0;
            using (SqlCommand cmd = new SqlCommand(mysql, con))
            {

                //using (con)
                //{
                    con.Open();
                    rtn = cmd.ExecuteNonQuery();  // -1    > = 1 
                  con.Close();
                //}

            }
            return rtn;
        }



        public int Insert(string mysql, Dictionary<string, object> myPara)
        {
            int rtn = 0;
            using (SqlCommand cmd = new SqlCommand(mysql, con))
            {



                foreach (KeyValuePair<string, object> p in myPara)
                {
                    cmd.Parameters.AddWithValue(p.Key, p.Value);
                }
                using (con)
                {
                    con.Open();
                    rtn = cmd.ExecuteNonQuery();
                    con.Close();
                }
            }
            return rtn;

        }

      

        public int NumberData(String mysql)
        {
            SqlDataReader dr;
           
            using (SqlCommand cmd = new SqlCommand(mysql, con))
            {
                con.Open();

                dr = cmd.ExecuteReader();
                dr.Read();
                int Count = (int)dr[0];
                con.Close();

                return Count;
            }

        }


    }
}
    