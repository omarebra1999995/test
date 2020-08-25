using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace certificateSystem
{


    public partial class Form3 : Form
    {

        public Form3()
        {
            CRUD DB = new CRUD();




            InitializeComponent();

            String Q = "SELECT [name_form]FROM [name_form]";
            SqlDataReader red = DB.getDataFromDataDase(Q);
            while (red.Read())
            {
                Un.Items.Add(red["name_form"]);

            }
            red.Close();
            DB.CloseDB();
            
            numberID.Items.Add(0);

            int number = 20;

            for (int i = 1; i <= number; i++)
            {
                numberID.Items.Add(i);
            }

            numberID.SelectedItem = 0;
            Un.SelectedIndex = 6;
        }


        private void Un_SelectedIndexChanged(object sender, EventArgs e)
        {
            CRUD Db = new CRUD();




            if (Un.SelectedIndex == 6)//select
            {

                listBox1Courses.Items.Clear();


            }
            else
            {
                int index = Un.SelectedIndex;
                index++;
                String Q = @"SELECT [unid],[nameForm]FROM[Forms] where unid ='" + index + "' ";
                listBox1Courses.Items.Clear();

                SqlDataReader d = Db.getDataFromDataDase(Q);
                bool check = false;
                if (d.HasRows) {
                    check = true;
                }
                while (d.Read())
                {

                    listBox1Courses.Items.Add(d["nameForm"]);


                }

                d.Close();
                Db.CloseDB();

                if (check == true) { listBox1Courses.SelectedIndex = 0; }
             




            }

        }



        private void noID_SelectedIndexChanged(object sender, EventArgs e)
        {

            TextBox[] Var = { d1, d2, d3, d4, d5, d6, d7 , d8, d9, d10,
                d11, d12, d13, d14, d15, d16, d17, d18, d19, d20};
            TextBox[] ID = { id1, id2, id3, id4, id5, id6, id7, id8, id9, id10,
                id11, id12, id13, id14, id15, id16, id17, id18, id19, id20 };

            if (numberID.SelectedIndex != null)
            {

                int no = Convert.ToInt32("20");
                for (int i = 0; i < no; i++)
                {



                    label3.Visible = false;
                    label4.Visible = false;
                    Var[i].Visible = false;
                    ID[i].Visible = false;
                    datatype.Visible = false;
                    idd.Visible = false;

                }
                int no1 = Convert.ToInt32(numberID.SelectedItem);
                for (int it = 0; it < no1; it++)
                {


                    if (no1 > 10)
                    {
                        label3.Visible = true;
                        label4.Visible = true;
                    }
                    Var[it].Visible = true;
                    ID[it].Visible = true;
                    datatype.Visible = true;
                    idd.Visible = true;
                    idd.Visible = true;


                }

            }
            else
            {
                int no = Convert.ToInt32("20");
                for (int i = 0; i < no; i++)
                {

                    label3.Visible = false;
                    label4.Visible = false;


                    Var[i].Visible = false;
                    ID[i].Visible = false;
                    datatype.Visible = false;
                    idd.Visible = false;


                }

            }
        }


        private void btnBackPage_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Form1 form = new Form1();
            form.Show();
        }

       

        private void Form3_Load(object sender, EventArgs e)
        {
            numberID.SelectedItem = 0;
            Un.SelectedIndex = 6;
        }

        private void Uploadfile(object sender, EventArgs e)
        {
            //To where your opendialog box get starting location. My initial directory location is desktop.
            open.InitialDirectory = "C://Desktop";
            //Your opendialog box title name.
            open.Title = "Select file to be upload.";
            //which type file format you want to upload in database. just add them.
            open.Filter = "Select Valid Document(*.pdf;)|*.pdf;";
            //FilterIndex property represents the index of the filter currently selected in the file dialog box.
            open.FilterIndex = 1;
            try
            {
                if (open.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (open.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(open.FileName);
                        FilePath.Text = path;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void CreateForm(object sender, EventArgs e)
        {

            if (FilePath.Text != "")
            {
                String name_file = "";
                try
                {
                    string filename = System.IO.Path.GetFileName(open.FileName);
                    name_file = filename;
                    if (filename == null)
                    {
                        MessageBox.Show("Please select a valid document.");
                    }
                    else
                    {
                        //we already define our connection globaly. We are just calling the object of connection.

                        string path = Application.StartupPath.Substring(0, (Application.StartupPath.Length - 10));
                        System.IO.File.Copy(open.FileName, path + "\\Templte\\" + filename);

                        MessageBox.Show("Document uploaded.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



                if (Un.SelectedIndex != 6)
                {
                    CRUD DB = new CRUD();
                    CRUD DB_1 = new CRUD();
                    //id int primary key  ,


                    String q = "Create table " + nametable.Text.ToString() + "( id_file int references intern(refNo)";

                    TextBox[] Var = { d1, d2, d3, d4, d5, d6, d7 , d8, d9, d10,
                d11, d12, d13, d14, d15, d16, d17, d18, d19, d20};
                    TextBox[] ID = { id1, id2, id3, id4, id5, id6, id7, id8, id9, id10,
                id11, id12, id13, id14, id15, id16, id17, id18, id19, id20 };

                    int no = Convert.ToInt32(numberID.SelectedItem);
                    for (int i = 0; i < no; i++)
                    {

                        q += "," + Var[i].Text.ToString() + " " + ID[i].Text.ToString();



                    }
                    q += ");";
                    MessageBox.Show(q);
                    DB.InsertUpdateDelete(q);
                    int index = Un.SelectedIndex;
                    index++;
                    DB_1.InsertUpdateDelete(@"insert into Forms(unid,nameForm,nametable,namefile)Values('" + index + "','" + nametable.Text.ToString() + "','" + nametable.Text.ToString() + "','" + name_file + "')");

                    this.Hide();
                    Form1 form = new Form1();
                    form.Show();
                }
                else
                {

                    MessageBox.Show("الرجاء اختيار ");
                }
            }
            else
            {
                MessageBox.Show("Please Upload document pdf.");
            }
        }

        private void DeleteForm(object sender, EventArgs e)
        {
            CRUD DB = new CRUD();
            CRUD DB_1 = new CRUD();
            CRUD DB_2 = new CRUD();


            if (Un.SelectedIndex != 6)
            {

                int index = Un.SelectedIndex;
                index++;

                SqlDataReader Read_Data_1 = DB_2.getDataFromDataDase("SELECT [namefile]   FROM  [Forms] where [unid] = '" + index + "' and nametable = '" + listBox1Courses.SelectedItem + "'");
                Read_Data_1.Read();
                
                string filename = Read_Data_1["namefile"].ToString();
                Read_Data_1.Close();
             //   ConfigurationSettings.AppSettings["Path_Templte"];
            
                File.Delete(ConfigurationSettings.AppSettings["Path_Templte"]+"\\" + filename);
                DB.InsertUpdateDelete(" delete[Forms] where [unid] = '" + index + "' and nametable = '" + listBox1Courses.SelectedItem + "'");
                DB_1.InsertUpdateDelete(@"drop table " + listBox1Courses.SelectedItem + "");
                MessageBox.Show("Delete successful ");
               // File.Delete(@"C:\Users\HP\Desktop\certificateSystem\certificateSystem\Templte\"+ filename );



                Un.SelectedIndex = 6;


            }


            ///////
            ///


        }

       

    }
}

