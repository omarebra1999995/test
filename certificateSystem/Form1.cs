using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace certificateSystem
{

    public partial class Form1 : Form
    {
        //  Temp b = new Temp();
        //class all funcation Query
        CRUD DB = new CRUD();

        ///pathFile 
        String Path_Files = ConfigurationSettings.AppSettings["Path_Files"];
        //pathTemplte
        String Path_Templte = ConfigurationSettings.AppSettings["Path_Templte"];

        //path language
        String Path_language = ConfigurationSettings.AppSettings["Path_language"];
        //class all funcation Query
        // CRUD DB = new CRUD();
        // Locatiton new file after full Templte
        string newFile = "";
        //Read and open Templte
        PdfReader pdfReader;
        //Write on Templte new data
        PdfStamper pdfStamper;
        //allow to record data on Templte
        AcroFields pdfFormFields;
        //read of data
        // SqlDataReader Read_Data_1;
        SqlDataReader d;
        //Name Templte
        //Count number No Send Email
        int cheaked_erorr_send = 0;
        //Count Number Send
        int count = 0;
        string pdfTemplate;
        //supject Emil
        String Supject = "";
        EmailManger EmailSend = new EmailManger();

        public int SearchAndDelete(String id, int typefun) {


            int indix = Un.SelectedIndex;
            indix++;

            SendToSupervisor.Enabled = true;
            String Q = @"SELECT [unid],[nameForm] ,[nametable] FROM[Forms] where unid ='" + indix + "'   and  [nameForm]='" + listBox1Courses.SelectedItem + "'";
            
            int correct = 0; ;
            SqlDataReader S = DB.getDataFromDataDase(Q);

            S.Read();

            String nameTable = S["nametable"].ToString();

            S.Close();
          DB.CloseDB();
            if (typefun == 1) { 
             
            if (id == null) {
                DB.getDataGrid(@"Select * from " + nameTable + " ", guna2DataGridView1);
                
            }
            else {
                DB.getDataGrid(@"Select * from " + nameTable + " where id_file='" + id + "'", guna2DataGridView1);
            } }else{
               correct= DB.InsertUpdateDelete(@"Delete " + nameTable + " where id_file='" + id + "'");
             
                
            }
            addheadercheckbox();
            headercheckbox.MouseClick += new MouseEventHandler(headerMousclick);
            return correct;
        }

        public void Templet(int NumberButten, String NameTable, String NameTemplet, BaseColor FontColor, String FontType , String Fontsize)
        {
            int count_doc = 0;
            String Date_tody = DateTime.Now.ToString("d/M/yyyy");
            count = 0;

            cheaked_erorr_send = 0;


            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
            {
             

                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();


                    //    it.Cells[1].Value.ToString()

                String Query = "Select [refNo],[fName],[lName],[mi],[email] ,[supervisorName] FROM[intern]where[refNo] = '" + it.Cells[1].Value.ToString() + "'";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);
                    //path
                    pdfTemplate = Path_Templte + "\\" + NameTemplet;
                    // +".pdf"


                    Read_Data_1.Read();
                    //string name = Read_Data_1["arabicName"].ToString();
                    //get first name of intern
                    //  String[] cutFristName = name.Split(' ');
                    if (NumberButten == 1)
                    {
                        //" + Read_Data_1["refNo"].ToString() + "
                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1["refNo"].ToString() + "_" + NameTemplet + ".pdf";
                    }
                    else
                    {
                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1["refNo"].ToString() + "" + NameTemplet + ".pdf";
                        // newFile = Path_Files + "\\" + Read_Data_1["refNo"].ToString() + "_" + Read_Data_1["refNo"].ToString() + "_" + NameTemplet + ".pdf";
                    }
                    Read_Data_1.Close();
                    DB.CloseDB();
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;
                    //MAJALLA.TTF
                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\" + FontType + ".TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdfFormFields.AddSubstitutionFont(arialBaseFont);


                    SqlDataReader dbm = DB.getDataFromDataDase(@"Select * from " + NameTable + "");

                    //for (int i = 0; i < Var.Length; i++)
                    dbm.Read();
                    int i = 0;
                    float FontSize = float.Parse(Fontsize);
                    while (i < dbm.FieldCount)//number column
                    {



                        //try                     name               omar
                        //{                      name column      
                        pdfFormFields.SetFieldProperty(dbm.GetName(i), "textsize", FontSize, null);
                        pdfFormFields.SetFieldProperty(dbm.GetName(i), "textcolor", FontColor, null);
                        pdfFormFields.SetField(dbm.GetName(i), dbm[i].ToString());


                        pdfFormFields.SetFieldProperty(dbm.GetName(i), "setfflags", PdfFormField.FF_READ_ONLY, null);
                    
                        i++;
                        
                       


                    }



                    pdfStamper.FormFlattening = false;
                    
                  
                    pdfStamper.Close();



                    if (NumberButten != 1)
                    {
                        Boolean cheakEmail = false;
                        if (NumberButten != 1)
                        {
                            if (NumberButten == 2)
                            {
                                String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                                Supject = @"File# " + Read_Data_1["refNo"].ToString() + "_Distance Training 2020 for " + Read_Data_1["fName"].ToString() + Read_Data_1["lName"].ToString();
                                cheakEmail = EmailSend.SendEmaile(Read_Data_1["email"].ToString(), newFile, Supject, BodyEmail);
                            }
                            if (NumberButten == 3)
                            {
                                String MessgEmail = @"Dear Supervaisor:
this is your student information";
                                Supject = "ID# " + Read_Data_1["collegeId"].ToString() + "Name#" + Read_Data_1["arabicName"].ToString();
                                cheakEmail = EmailSend.SendEmaile(Read_Data_1["supervisorName"].ToString(), newFile, Supject, MessgEmail);
                            }
                            if (cheakEmail == false)
                            {
                                cheaked_erorr_send++;
                                MessageBox.Show("Send Email Failure ((" + Read_Data_1["collegeId"].ToString() + ")" + Read_Data_1["arabicName"].ToString() + "");
                            }
                            else
                            {
                                File.Delete(newFile);
                            }
                        }
                    }
                    dbm.Close();

                }



            }


            if (count > 0)
            {
                if (cheaked_erorr_send != count)
                {
                    if (NumberButten == 1)
                    {
                        MessageBox.Show("Save successful.....");
                    }
                    else
                    {
                        MessageBox.Show("Send Email successful.....");
                    }
                }
            }
            else
            {
                MessageBox.Show("Please Chooes Intern");
            }
            pdfReader.Close();
            DB.CloseDB();

        }
       
        public void Execution(int NumberButten, BaseColor FontColor, String FontType , String fontSize)
        {
           

            int index = Un.SelectedIndex;
            index++;
            String Q = @"SELECT [nameForm] , [namefile] FROM[Forms] where unid ='" + index + "' ";

            SqlDataReader dd = DB.getDataFromDataDase(Q);
            while (dd.Read())
            {
                if (listBox1Courses.SelectedItem.Equals(dd["nameForm"].ToString()))
                {



                    String Name_form = dd["namefile"].ToString();
                    String NAmeTable = dd["nameForm"].ToString();
                    dd.Close();
                    Templet(NumberButten, NAmeTable, Name_form, FontColor, FontType, fontSize);

                    return;
                }



            }

            DB.CloseDB();

        }

        public Form1()
        {
            InitializeComponent();

            String Q = "SELECT [name_form]FROM [name_form]";
            SqlDataReader red = DB.getDataFromDataDase(Q);
            while (red.Read())
            {
                Un.Items.Add(red["name_form"]);

            }
            red.Close();
            DB.CloseDB();
           
            for (int i = 5; i <= 100; i+=5) {
                sizefont.Items.Add(i);

            }
           sizefont.SelectedIndex = 0;
            fonttype.SelectedIndex = 0;

            //create Cheak Box in header (Done) and mouseClick
            addheadercheckbox();
            headercheckbox.MouseClick += new MouseEventHandler(headerMousclick);
         
        }

        CheckBox headercheckbox = null;

        bool isheadercheckboxClicked = false;

        private void addheadercheckbox() ///create Cheak box Header (Done)
        {
            headercheckbox = new CheckBox();
            headercheckbox.Size = new Size(15, 15);

            this.guna2DataGridView1.Controls.Add(headercheckbox);
            headercheckbox.Checked = true;
            headerCheckBoxClick(headercheckbox);
            Point headerCellLocation = this.guna2DataGridView1.GetCellDisplayRectangle(0, -1, true).Location;
            headercheckbox.Location = new Point(headerCellLocation.X + 55, headerCellLocation.Y + 4);

        }

        private void headerCheckBoxClick(CheckBox hceckbox)
        {
            isheadercheckboxClicked = true;
            foreach (DataGridViewRow ROW in guna2DataGridView1.Rows)
                ((DataGridViewCheckBoxCell)ROW.Cells[0]).Value = hceckbox.Checked;
            guna2DataGridView1.RefreshEdit();
            isheadercheckboxClicked = false;

        }

        private void headerMousclick(object sender, MouseEventArgs s)
        {

            headerCheckBoxClick((CheckBox)sender);


        }

        private void button1_Click(object sender, EventArgs e)
        {// cancel button to end close the app
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            headercheckbox.MouseClick += new MouseEventHandler(headerMousclick);

            sizefont.SelectedIndex = 0;
            fonttype.SelectedIndex = 0;
            Un.SelectedIndex = 6;

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            if (txtSearch.Text != "")
            {
                SearchAndDelete(txtSearch.Text.ToString() , 1);

            }
            else
            {
                MessageBox.Show("Please Enter ID");
            }


        }

        private void btnSaveCert_Click(object sender, EventArgs e) //Butten Save File in Folder on Computer.
        {


            if (txtUploadFile.Text != "")
            {
                try
                {
                    BaseColor FontColor = new BaseColor(colorDialog1.Color);
                  
                    Execution(1, FontColor, fonttype.SelectedItem.ToString(), sizefont.SelectedItem.ToString());

                }



                catch (Exception ex)
                {
                    MessageBox.Show("There is an open File!");
                }
            }

            else
            {
                MessageBox.Show("Please choose Folder!");
            }

        }

        private void btnSendEmail_Click(object sender, EventArgs e)
            //create file and send email to intern
        {
            // call send email method

            // Execution(2);
        }
        private void btnShowAll_Click(object sender, EventArgs e)
        {


            SearchAndDelete(null ,1);
          

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            // call delete method
            int count = 0;
            if (MessageBox.Show("Do you  want really to delete user", "Message", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (DataGridViewRow it in guna2DataGridView1.Rows)
                {
                    if (bool.Parse(it.Cells[0].Value.ToString()))
                    {

                    

                        count = SearchAndDelete(it.Cells[1].Value.ToString(), 0);


                    }
                }

                
                    if (count > 0)
                        {
                            MessageBox.Show(" Delete successful.....");



                        }
                        else
                        {
                            MessageBox.Show("Failure.....");
                        }
                btnShowAll_Click(sender, e);
            }


                


           
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            SearchAndDelete(null,1);

        int NumberRow=  DB.NumberData("  Select Count(*) From  "+listBox1Courses.Text.ToString()+"");

          //  MessageBox.Show(Convert.ToString(NumberRow));
            NoRow.Text = Convert.ToString(NumberRow);


          



           
        }

        private void btnUploadFile_Click(object sender, EventArgs e)
        {
            // call upload file method

            FolderBrowserDialog Folder = new FolderBrowserDialog();
            Folder.Description = "chooes Folder";
            Folder.ShowNewFolderButton = false;
            if (Folder.ShowDialog() == DialogResult.OK)
            {

                txtUploadFile.Text = Folder.SelectedPath;
            }
        }

        private void Search(object sender, KeyPressEventArgs e)
        {
            char che = e.KeyChar;

            if (!char.IsDigit(che) && che != 8 && che != 13)
            {

                e.Handled = true;
                MessageBox.Show("Please Enter only Number");
            }
            if (txtSearch.Text != "")
            {
                if (!char.IsDigit(che) && che != 8 && che == 13)
                {

                    SearchAndDelete(txtSearch.Text.ToString(), 1);

                }

            }
        }

        private void SendSupervaisor_Click(object sender, EventArgs e) //Buten create file and send Emil to SendSupervaisor
        {
            // call send email method


            // Execution(3);
        }

        private void Un_SelectedIndexChanged(object sender, EventArgs e)
        {

            int index = Un.SelectedIndex;

            if (index == 6)//select
            {

                listBox1Courses.Items.Clear();
                btnSendAllSup.Enabled = false;
                btnSendAllToInterns.Enabled = false;

            }
            else
            {
            index++;
     
                String Q = @"SELECT [unid],[nameForm] FROM[Forms] where unid ='" + index + "' ";
                listBox1Courses.Items.Clear();

                d = DB.getDataFromDataDase(Q);
                bool cecked = false;
                if (d.HasRows) {

                    cecked = true;
                }

                while (d.Read())
                {

                    listBox1Courses.Items.Add(d["nameForm"]);


                }

                d.Close();
                DB.CloseDB();
                if (cecked == true) { listBox1Courses.SelectedIndex = 0; }
               
                addheadercheckbox();
                headercheckbox.MouseClick += new MouseEventHandler(headerMousclick);

            }


        }

        private void btnSendAllSup_Click(object sender, EventArgs e)
        {
            //string Query = "";
            //int count = 0;
            //String Datee = DateTime.Now.ToString("d/M/yyyy");
            //    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //    foreach (DataGridViewRow it in guna2DataGridView1.Rows)
            //    {
            //        if (bool.Parse(it.Cells[0].Value.ToString()) == true)
            //        {
            //            DB.CloseDB();

            //            Query = "select * from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

            //            count++;
            //            SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

            //            pdfTemplate = Path_Templte + "\\Nourh-AS.pdf";
            //            Read_Data_1.Read();
            //            string name = Read_Data_1[12].ToString();
            //            //get first name of intern
            //            String[] cutFristName = name.Split(' ');
            //            //////////////////////////////////// ---- AS
            //            newFile = Path_Files + "\\" + Read_Data_1[11].ToString() + "-" + Read_Data_1[12].ToString() + "-Attendance Sheet.pdf";

            //            pdfReader = new PdfReader(pdfTemplate);
            //            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
            //            pdfFormFields = pdfStamper.AcroFields;

            //            //                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //            pdfFormFields.AddSubstitutionFont(arialBaseFont);


            //            pdfFormFields.SetField("NameStudent", Read_Data_1[12].ToString());
            //            pdfFormFields.SetField("IDNumber", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("Major", Read_Data_1[22].ToString());
            //            for (int i = 1; i <= 25; i++)
            //            {
            //                pdfFormFields.SetField("checkin" + i, cutFristName[0]);
            //                pdfFormFields.SetField("checkout" + i, cutFristName[0]);
            //            }



            //            pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);


            //            for (int i = 1; i <= 25; i++)
            //            {
            //                pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
            //                pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            }
            //            pdfStamper.FormFlattening = false;

            //            pdfStamper.Close();
            //            /////////////////////////////////////////////////// ER1 /////////
            //            string pdfTemplateER1 = Path_Templte + "\\Nourh-ER1.pdf";

            //            newFile = Path_Files + "\\" + Read_Data_1[11].ToString() + "-" + Read_Data_1[12].ToString() + "-Evaluation(1).pdf";

            //            pdfReader = new PdfReader(pdfTemplateER1);
            //            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
            //            pdfFormFields = pdfStamper.AcroFields;

            //            pdfFormFields.AddSubstitutionFont(arialBaseFont);
            //            String Name_En = Read_Data_1[13].ToString() + " " + Read_Data_1[14].ToString() + " " + Read_Data_1[16].ToString();

            //            String[] alltextfild = { "InternName", "InternID", "internshipTitle", "DatesIntership"
            //                        , "NumberWeekly","Semester","NameStudent","IDNumber", "Date"};
            //            pdfFormFields.SetField("InternName", Name_En);
            //            pdfFormFields.SetField("InternID", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("internshipTitle", Read_Data_1[38].ToString());
            //            pdfFormFields.SetField("DatesIntership", Read_Data_1[39].ToString());
            //            pdfFormFields.SetField("NumberWeekly", Read_Data_1[37].ToString());
            //            pdfFormFields.SetField("Semester", Read_Data_1[30].ToString());
            //            pdfFormFields.SetField("NameStudent", Name_En);
            //            pdfFormFields.SetField("IDNumber", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("Date", Datee);

            //            for (int i = 0; i < alltextfild.Length; i++)
            //            {
            //                pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            }


            //            pdfStamper.FormFlattening = false;

            //            pdfStamper.Close();
            //            ////////////////////////////  ER2-------   ///
            //            string pdfTemplateER2 = Path_Templte + "\\Nourh-ER2.pdf";
            //            {
            //                newFile = Path_Files + "\\" + Read_Data_1[11].ToString() + "-" + Read_Data_1[12].ToString() + "-Evaluation(2).pdf";
            //            }
            //            arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //            pdfReader = new PdfReader(pdfTemplateER2);
            //            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
            //            pdfFormFields = pdfStamper.AcroFields;


            //            pdfFormFields.AddSubstitutionFont(arialBaseFont);



            //            pdfFormFields.SetField("NameStudent", Read_Data_1[12].ToString());
            //            pdfFormFields.SetField("IDNumber", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("Major", Read_Data_1[22].ToString());


            //            pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);




            //            pdfStamper.FormFlattening = false;

            //            pdfStamper.Close();
            //            ///////////////////// send email
            //            Boolean cheakEmail = false;
            //            String MessgEmail = @"Dear Supervaisor:
            //        this is your student information";
            //            Supject = "ID# " + Read_Data_1[11].ToString() + "Name#" + Read_Data_1[12].ToString();
            //            cheakEmail = EmailSend.sendAllFilesEmail(Read_Data_1[35].ToString(), Supject, MessgEmail);
            //            if (cheakEmail == false)
            //            {
            //                MessageBox.Show("Send Email Failure ((" + Read_Data_1[12].ToString() + ")" + Read_Data_1[11].ToString() + "");
            //            }

            //            else
            //            {
            //                // Delete all files in a directory    
            //                string[] files = Directory.GetFiles(Path_Files);
            //                foreach (string file in files)
            //                {
            //                    File.Delete(file);

            //                }
            //            }
            //            Read_Data_1.Close();

            //        }

            //    }
            //    if (count > 0)
            //    {

            //        MessageBox.Show("Send Email successful.....");

            //    }
            //    else
            //    {
            //        MessageBox.Show("Please Chooes Intern");
            //    }

        }

        private void btnSendAllToInterns_Click(object sender, EventArgs e)
        {
            //    string Query = "";
            //    // for attendance sheet
            //    int count = 0;
            //    String Datee = DateTime.Now.ToString("d/M/yyyy");
            //    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //    foreach (DataGridViewRow it in guna2DataGridView1.Rows)
            //    {
            //        if (bool.Parse(it.Cells[0].Value.ToString()) == true)
            //        {

            //            DB.CloseDB();

            //            Query = "select * from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

            //            count++;
            //            SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

            //            string pdfTemplate = Path_Templte + "\\Nourh-AS.pdf";
            //            Read_Data_1.Read();
            //            string name = Read_Data_1[12].ToString();
            //            //get first name of intern
            //            String[] cutFristName = name.Split(' ');
            //            // ---- AS
            //            newFile = Path_Files + "\\" + Read_Data_1[11].ToString() + "-" + Read_Data_1[12].ToString() + "-Attendance Sheet.pdf";

            //            pdfReader = new PdfReader(pdfTemplate);
            //            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
            //            pdfFormFields = pdfStamper.AcroFields;

            //            //                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //            pdfFormFields.AddSubstitutionFont(arialBaseFont);


            //            pdfFormFields.SetField("NameStudent", Read_Data_1[12].ToString());
            //            pdfFormFields.SetField("IDNumber", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("Major", Read_Data_1[22].ToString());
            //            for (int i = 1; i <= 25; i++)
            //            {
            //                pdfFormFields.SetField("checkin" + i, cutFristName[0]);
            //                pdfFormFields.SetField("checkout" + i, cutFristName[0]);
            //            }



            //            pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);


            //            for (int i = 1; i <= 25; i++)
            //            {
            //                pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
            //                pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            }
            //            pdfStamper.FormFlattening = false;

            //            pdfStamper.Close();
            //            /////////////////////////////////////////////////// ER1 /////////
            //            string pdfTemplateER1 = Path_Templte + "\\Nourh-ER1.pdf";

            //            newFile = Path_Files + "\\" + Read_Data_1[11].ToString() + "-" + Read_Data_1[12].ToString() + "-Evaluation(1).pdf";

            //            pdfReader = new PdfReader(pdfTemplateER1);
            //            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
            //            pdfFormFields = pdfStamper.AcroFields;

            //            pdfFormFields.AddSubstitutionFont(arialBaseFont);
            //            String Name_En = Read_Data_1[13].ToString() + " " + Read_Data_1[14].ToString() + " " + Read_Data_1[16].ToString();

            //            String[] alltextfild = { "InternName", "InternID", "internshipTitle", "DatesIntership"
            //                        , "NumberWeekly","Semester","NameStudent","IDNumber", "Date"};
            //            pdfFormFields.SetField("InternName", Name_En);
            //            pdfFormFields.SetField("InternID", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("internshipTitle", Read_Data_1[38].ToString());
            //            pdfFormFields.SetField("DatesIntership", Read_Data_1[39].ToString());
            //            pdfFormFields.SetField("NumberWeekly", Read_Data_1[37].ToString());
            //            pdfFormFields.SetField("Semester", Read_Data_1[30].ToString());
            //            pdfFormFields.SetField("NameStudent", Name_En);
            //            pdfFormFields.SetField("IDNumber", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("Date", Datee);

            //            for (int i = 0; i < alltextfild.Length; i++)
            //            {
            //                pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            }


            //            pdfStamper.FormFlattening = false;

            //            pdfStamper.Close();
            //            ////////////////////////////  ER2-------   ///
            //            string pdfTemplateER2 = Path_Templte + "\\Nourh-ER2.pdf";
            //            {
            //                newFile = Path_Files + "\\" + Read_Data_1[11].ToString() + "-" + Read_Data_1[12].ToString() + "-Evaluation(2).pdf";
            //            }
            //            arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //            pdfReader = new PdfReader(pdfTemplateER2);
            //            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
            //            pdfFormFields = pdfStamper.AcroFields;


            //            pdfFormFields.AddSubstitutionFont(arialBaseFont);



            //            pdfFormFields.SetField("NameStudent", Read_Data_1[12].ToString());
            //            pdfFormFields.SetField("IDNumber", Read_Data_1[11].ToString());
            //            pdfFormFields.SetField("Major", Read_Data_1[22].ToString());


            //            pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
            //            pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);

            //            pdfStamper.FormFlattening = false;

            //            pdfStamper.Close();

            //            Boolean cheakEmail = false;

            //            String BodyEmail = @"Dear student:
            //            thank you for end of internship this is your certificate and University Forms";
            //            Supject = @"File# " + Read_Data_1[2].ToString() + "- Distance Training 2020 for " + Read_Data_1[13].ToString() + Read_Data_1[16].ToString();
            //            cheakEmail = EmailSend.sendAllFilesEmail(Read_Data_1[24].ToString(), Supject, BodyEmail);
            //            if (cheakEmail == false)
            //            {
            //                MessageBox.Show("Send Email Failure ((" + Read_Data_1[3].ToString() + ")" + Read_Data_1[0].ToString() + "");
            //            }
            //            else
            //            {
            //                // Delete all files in a directory    
            //                string[] files = Directory.GetFiles(Path_Files);
            //                foreach (string file in files)
            //                {
            //                    File.Delete(file);

            //                }
            //                //  File.Delete(newFile);
            //            }
            //            Read_Data_1.Close();
            //        }

            //    }
            //    if (count > 0)
            //    {

            //        MessageBox.Show("Send Email successful.....");

            //    }
            //    else
            //    {
            //        MessageBox.Show("Please Chooes Intern");
            //    }

            //}
        }

        private void createform_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form3 form = new Form3();
            form.Show();
        }

        private void FontColor(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                //t1.Text =Convert.ToString( colorDialog1.Color);
                BaseColor FontColor = new BaseColor(colorDialog1.Color);
                // color1.Text = Convert.ToString(b);
                color1.BackColor = colorDialog1.Color;
            }
        }

        private void mappingExcel(object sender, EventArgs e)
        {
            if (bx1.Text.ToString()!="") {

                try
                {

                    string sexcelconnectionstring = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source=" + bx1.Text.ToString() + ";Extended Properties='Excel 12.0;IMEX=1;'";

                    string ssqlconnectionstring = ConfigurationSettings.AppSettings["connectionString_Local"];
                    // @"Data Source = DESKTOP-SHF8FLP\SQLEXPRESS;Initial Catalog = certificate; Integrated Security = True";
                    //    ConfigurationSettings.AppSettings["connectionString_Local"];

                    string sclearsql = "delete from " + listBox1Courses.SelectedItem.ToString();
                    SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                    SqlCommand sqlcmd = new SqlCommand(sclearsql, sqlconn);
                    sqlconn.Open();
                    sqlcmd.ExecuteNonQuery();
                    sqlconn.Close();
                    //series of commands to bulk copy data from the excel file into our sql table   
                    OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
                    OleDbCommand oledbcmd = new OleDbCommand(@"select * from[s1$]", oledbconn);
                    oledbconn.Open();
                    OleDbDataReader dr = oledbcmd.ExecuteReader();
                    SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                    bulkcopy.DestinationTableName = listBox1Courses.SelectedItem.ToString();
                    bulkcopy.WriteToServer(dr);
                    while (dr.Read())
                    {
                        bulkcopy.WriteToServer(dr);
                    }
                    dr.Close();
                    oledbconn.Close();

                    listBox1_SelectedIndexChanged(sender, e);


                    MessageBox.Show("operation accomplished successfully");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred !! check the number of columns");
                }
            }
            else {
                MessageBox.Show("Please Upload document Excel.");
            }

    } 

        private void uploadExcel(object sender, EventArgs e)
        {
           //To where your opendialog box get starting location. My initial directory location is desktop.
            openFileDialog1.InitialDirectory = "C://Desktop";
            //Your opendialog box title name.
            openFileDialog1.Title = "Select file to be upload.";
            //which type file format you want to upload in database. just add them.
            openFileDialog1.Filter = "Select Valid Document(*.xlsx;)| *.xlsx;";
            //"Select Valid Document(*.pdf; *.doc; *.xlsx; *.html)|*.pdf; *.docx; *.xlsx; *.html";
            //FilterIndex property represents the index of the filter currently selected in the file dialog box.
            openFileDialog1.FilterIndex = 1;
            
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog1.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(openFileDialog1.FileName);
                        bx1.Text = path;
                    }
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
  
    }
    }
    
