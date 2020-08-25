using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace certificateSystem
{
    class Temp
    {

        ///pathFile 
        String Path_Files = "omarebr";
        //ConfigurationSettings.AppSettings["Path_Files"];
        //pathTemplte
        String Path_Templte = "vv";
        //ConfigurationSettings.AppSettings["Path_Templte"];

        //path language
        String Path_language = @"C:\Users\HP\Desktop\certificateSystem\certificateSystem\language";
            //ConfigurationSettings.AppSettings["Path_language"];
        //class all funcation Query
        CRUD DB = new CRUD();
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
        //Name Templte
        //Count number No Send Email
        int cheaked_erorr_send = 0;
        //Count Number Send
        int count = 0;
        string pdfTemplate;
        //supject Emil
        String Supject = "";
        EmailManger EmailSend = new EmailManger();
       
        Form1 form1 = new Form1();
        String Query_Certificates = "SELECT [id]  ,[NameAr]  ,[nationalNum] As NationalID  ,[email]  As Email ,[NameEn]  FROM [studentInfo]";
        String Query = @"SELECT  [internId] ,[accepted] ,[nationalityId],[arabicName] ,[major],[email] ,[supervisorName],[supervisorCell],[supervisorEmail]FROM[intern]";

         

        public static string ConvertToEasternArabicNumerals(string input)
        {
            System.Text.UTF8Encoding utf8Encoder = new UTF8Encoding();
            System.Text.Decoder utf8Decoder = utf8Encoder.GetDecoder();
            System.Text.StringBuilder convertedChars = new System.Text.StringBuilder();
            char[] convertedChar = new char[1];
            byte[] bytes = new byte[] { 217, 160 };
            char[] inputCharArray = input.ToCharArray();
            foreach (char c in inputCharArray)
            {
                if (char.IsDigit(c))
                {
                    bytes[1] = Convert.ToByte(160 + char.GetNumericValue(c));
                    utf8Decoder.GetChars(bytes, 0, 2, convertedChar, 0);
                    convertedChars.Append(convertedChar[0]);
                }
                else
                {
                    convertedChars.Append(c);
                }
            }
            return convertedChars.ToString();
        }
        //Form Nourh_AS
        public void Nourh_AS(int NumberButten)
        {
            count = 0;

            cheaked_erorr_send = 0;
            count = 0;

            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "SELECT [arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName]FROM[intern]where [internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);
                    //path
                    pdfTemplate = Path_Templte + "\\Nourh-AS.pdf";



                    Read_Data_1.Read();
                    string name = Read_Data_1[0].ToString();
                    //get first name of intern
                    String[] cutFristName = name.Split(' ');
                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_AS.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Attendance Sheet.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

                    pdfFormFields.SetFieldProperty("IDNumber", "textsize", 10f, null);

                    pdfFormFields.SetField("NameStudent", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("IDNumber", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Major", Read_Data_1[2].ToString());
                    pdfFormFields.SetField("CourseCode", Read_Data_1[3].ToString());
                    for (int i = 1; i <= 25; i++)
                    {
                        pdfFormFields.SetField("checkin" + i, cutFristName[0]);
                        pdfFormFields.SetField("checkout" + i, cutFristName[0]);
                    }



                    pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);


                    for (int i = 1; i <= 25; i++)
                    {
                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();
                    //   pdfReader.Close();


                    if (NumberButten != 1)
                    {
                        //EmailSend.SendEmail(NumberButten, Read_Data_1, newFile);
                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();

        }
        //Form Suad_ER1
        public void Suad_ER1(int NumberButten)
        {
            cheaked_erorr_send = 0;

            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()))
                {

                    DB.CloseDB();



                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);


                    pdfTemplate = Path_Templte + "\\Suad_ER1.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation1.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "-Evaluation1.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));

                    pdfFormFields = pdfStamper.AcroFields;


                    pdfFormFields.AddSubstitutionFont(arialBaseFont);






                    String[] alltextfild = { "studentId", "Name", "startDate", "endDate" };


                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("endDate", "textsize", 10f, null);



                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("startDate", "07\\06\\2020");
                    pdfFormFields.SetField("endDate", "20\\07\\2020");


                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();
                    //pdfReader.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();

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
            DB.CloseDB();

        }
        //Form Suad_ER2
        public void Suad_ER2(int NumberButten)
        {
            cheaked_erorr_send = 0;

            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    string pdfTemplate = Path_Templte + "\\Suad_ER2.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation2.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation2.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





                    String[] alltextfild = { "Name", "studentId", "startDate", "institution" };



                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);


                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("startDate", "07\\06\\2020");
                    pdfFormFields.SetField("institution", "مدينة الملك فهد الطبيه");


                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[0].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();

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
            DB.CloseDB();

        }
        //Form Suad_AS
        public void Suad_AS(int NumberButten)
        {

            cheaked_erorr_send = 0;
            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    string pdfTemplate = Path_Templte + "\\Suad_AS.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_AS.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Attendance Sheet.pdf";
                    } // Evaluation(1)
                    // Attendance Sheet
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);


                    string name = Read_Data_1[0].ToString();
                    //get first name of intern
                    String[] cutFristName = name.Split(' ');
                    String Name_En = Read_Data_1[7].ToString() + " " + Read_Data_1[9].ToString() + " " + Read_Data_1[8].ToString();

                    String[] alltextfild = { "Name", "studentId", "startDate", "institution", };

                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);

                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());



                    for (int i = 1; i <= 40; i++)
                    {
                        pdfFormFields.SetField("checkin" + i, cutFristName[0]);
                        pdfFormFields.SetField("checkout" + i, cutFristName[0]);

                    }
                    for (int i = 1; i <= 40; i++)
                    {
                        pdfFormFields.SetFieldProperty("checkout" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }

                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();

        }
        //Form Suad_JF
        public void Suad_JF(int NumberButten)
        {

            cheaked_erorr_send = 0;
            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();


                    Query = "select [collegeId],[arabicName],[majorAr],[cell],[email],[supervisorName],[supervisorEmail],[refNo] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";




                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    string pdfTemplate = Path_Templte + "\\Saud-JF.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Effective Date Form.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Effective Date Form.pdf";
                    } // Evaluation(1)
                    // Attendance Sheet
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);







                    String[] alltextfild = { "Name", "collegeId", "StPhone", "StPhone2", "StEmail", "Institution", "Address", "Department", "Dep2", "SupName", "Position", "SupPhoneH", "SupPhone", "SupEmail", "StartDate", "Rd1", "Rd2" };


                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "textsize", 10f, null);

                    }
                    pdfFormFields.SetFieldProperty("Name", "textsize", 12f, null);


                    pdfFormFields.SetField("Name", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("collegeId", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("StPhone", Read_Data_1[3].ToString());
                    pdfFormFields.SetField("StPhone2", Read_Data_1[3].ToString());
                    pdfFormFields.SetField("StEmail", Read_Data_1[4].ToString());
                    pdfFormFields.SetField("Institution", "King Fahad Medical City");
                    pdfFormFields.SetField("Address", "Sulimaniyah, Riyadh 12231");
                    pdfFormFields.SetField("Department", "Executive Administration of");
                    pdfFormFields.SetField("Dep2", "Information Technology/Application Training Department");
                    pdfFormFields.SetField("SupName", "Ali Mohamed Hamidaddin");
                    pdfFormFields.SetField("Position", "IT Consultant");
                    pdfFormFields.SetField("SupPhoneH", "+96611288999");
                    pdfFormFields.SetField("SupPhone", "0538692448");
                    pdfFormFields.SetField("SupEmail", Read_Data_1[6].ToString());
                    pdfFormFields.SetField("StartDate", "07\\06\\2020");


                    string major = Read_Data_1[2].ToString();
                    string majorTrack = Read_Data_1[2].ToString();

                    if (major == "علوم الحاسب")
                    {
                        pdfFormFields.SetField("Rd1", "2");
                    }
                    else if (major == "تقنية المعلومات")
                    {
                        pdfFormFields.SetField("Rd1", "3");
                        if (majorTrack == "Network & Security")
                        {
                            pdfFormFields.SetField("Rd2", "5");
                        }
                        else if (majorTrack == "Data Management")
                        {
                            pdfFormFields.SetField("Rd2", "4");
                        }

                        else if (majorTrack == "Web Technologies & Multimedia")
                        {
                            pdfFormFields.SetField("Rd2", "3");
                        }

                        else if (majorTrack == "Cyber Security")
                        {
                            pdfFormFields.SetField("Rd2", "2");
                        }
                        else if (majorTrack == "Data Science")
                        {
                            pdfFormFields.SetField("Rd2", "1");
                        }
                        else if (majorTrack == "Networks & Internet of Things")
                        {
                            pdfFormFields.SetField("Rd2", "0");
                        }


                    }

                    else if (major == "SWE")
                    {
                        pdfFormFields.SetField("Rd1", "1");
                    }
                    else if (major == "IS")
                    {
                        pdfFormFields.SetField("Rd1", "0");
                    }



                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }



                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[0].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[0].ToString() + "Name#" + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[6].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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

            DB.CloseDB();
        }
        //Form  TVTC_JF
        public void TVTC_JF(int NumberButten)

        {
            cheaked_erorr_send = 0;

            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()))
                {

                    DB.CloseDB();



                    Query = "SELECT [arabicName], [collegeId],[major],[supervisorEmail],[email],[refNo] from [intern] where [internId]= " + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);


                    pdfTemplate = Path_Templte + "\\TVTC_JF.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_JF.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_JF.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));

                    pdfFormFields = pdfStamper.AcroFields;


                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





                    String[] alltextfild = { "numStudent", "idStudent", "department", "ddlMajor",
                        "date", "numManager", "jobManager", "telNo", "fax", "email", "mobile" };

                    for (int i = 0; i < alltextfild.Length; i++)
                    {

                        pdfFormFields.SetFieldProperty(alltextfild[i], "textsize", 10f, null);
                    }





                    pdfFormFields.SetField("numStudent", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("idStudent", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("department", "IT");
                    pdfFormFields.SetField("ddlMajor", Read_Data_1[2].ToString());
                    pdfFormFields.SetField("date", "07\\06\\2020");
                    pdfFormFields.SetField("numManager", "Ali Mohamed Hamidaddin");
                    pdfFormFields.SetField("jobManager", "IT consultant");
                    pdfFormFields.SetField("telNo", "(+966) 11 288 9999");
                    pdfFormFields.SetField("fax", "19100");
                    pdfFormFields.SetField("email", "Ahameed@kfmc.med.sa");
                    pdfFormFields.SetField("mobile", "(+966) 538692448");



                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();
                    //pdfReader.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[5].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[3].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();

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

            DB.CloseDB();
        }

        //Form Shaqra_ER
        public void Shaqra_ER(int NumberButten)
        {
            cheaked_erorr_send = 0;

            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()))
                {

                    DB.CloseDB();



                    Query = "select [arabicName],[majorAr],[registrationDate],[UploadedDate],[supervisorEmail],[email],[collegeId] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);


                    pdfTemplate = Path_Templte + "\\Shaqra_ER.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[6].ToString() + Read_Data_1[0].ToString() + "_Evaluation.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[6].ToString() + "-" + Read_Data_1[1].ToString() + "_Evaluation.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));

                    pdfFormFields = pdfStamper.AcroFields;


                    pdfFormFields.AddSubstitutionFont(arialBaseFont);






                    String[] alltextfild = { "name", "Side", "mejor", "Nmejor", "job", "date", "Edate" };


                    pdfFormFields.SetFieldProperty("name", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("Side", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("mejor", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("Nmejor", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("job", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("date", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("Edate", "textsize", 10f, null);

                    DateTime dateEnd_format = Convert.ToDateTime(Read_Data_1[3].ToString());
                    string dateeEnd = string.Format("{0:dd/MM/yyyy}", dateEnd_format);
                    DateTime dateStart_format = Convert.ToDateTime(Read_Data_1[2].ToString());
                    string dateStart = string.Format("{0:dd/MM/yyyy}", dateStart_format);


                    pdfFormFields.SetField("name", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("Side", "مدينة ملك فهد الطبية ");
                    pdfFormFields.SetField("mejor", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Nmejor", "مهندس علي حميدالدين");
                    pdfFormFields.SetField("job", "مستشار تقنية معلومات");
                    pdfFormFields.SetField("date", dateStart);
                    pdfFormFields.SetField("Edate", dateeEnd);


                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();
                    //pdfReader.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[1].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[6].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[3].ToString() + ")" + Read_Data_1[0].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();

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

            DB.CloseDB();
        }


        //Shaqra_JF
        public void Shaqra_JF(int NumberButten)
        {
            cheaked_erorr_send = 0;

            int count = 0;
            //   MessageBox.Show("processing......");
            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "select [collegeId],[arabicName],[majorAr],[cell],[email],[supervisorEmail],[refNo]  from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";
                    count++;

                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    pdfTemplate = Path_Templte + "\\Shaqra_JF.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_JF.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_JF.pdf";
                    }
                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;


                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

                    pdfFormFields.SetFieldProperty("id", "textsize", 11f, null);
                    //pdfFormFields.SetFieldProperty("NameStudent", "textsize", 14f, null);




                    pdfFormFields.SetField("id", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("name", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("track", Read_Data_1[2].ToString());
                    pdfFormFields.SetField("mobile", Read_Data_1[3].ToString());
                    pdfFormFields.SetField("email", Read_Data_1[4].ToString());
                    pdfFormFields.SetField("date", "07\\06\\2020");


                    pdfFormFields.SetFieldProperty("id", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("name", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("track", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("mobile", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("email", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("date", "setfflags", PdfFormField.FF_READ_ONLY, null);







                    pdfStamper.Close();

                    Boolean cheakEmail = false;

                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[0].ToString() + "Name#" + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[0].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();

        }
        //Form Sattam_ER
        public void Sattam_ER(int NumberButten)
        {

            cheaked_erorr_send = 0;
            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "select [collegeId],[arabicName],[fName],[lName],[email],[supervisorEmail],[refNo] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    pdfTemplate = Path_Templte + "\\Sattam_ER.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Evaluation.pdf";
                    }
                    else
                    {
                        // Evaluation(1)
                        // Attendance Sheet
                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Evaluation.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





                    String[] alltextfild = { "Name", "OrgName", "Date", "SName", "Sign" };






                    pdfFormFields.SetField("Name", Read_Data_1[1].ToString());

                    pdfFormFields.SetField("OrgName", "مدينة الملك فهد الطبيه");
                    pdfFormFields.SetField("SName", "علي محمد حميد الدين");
                    pdfFormFields.SetField("Sign", "علي ");
                    pdfFormFields.SetField("Date", DateTime.Now.ToString("dd/MM/yyyy"));
                    // pdfFormFields.SetFieldProperty("studentId", "setfflags", new , null);




                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[2].ToString() + Read_Data_1[3].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[0].ToString() + "Name#" + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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

            DB.CloseDB();
        }

        //Form Sattam_JF
        public void Sattam_JF(int NumberButten)
        {

            cheaked_erorr_send = 0;
            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "select [arabicName],[collegeId],[trainingYearId], [majorAr],[supervisorName],[supervisorEmail],[email] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    string pdfTemplate = Path_Templte + "\\Sattam_JF.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "Sattam_JF_Gr.pdf"; //////
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "-" + Read_Data_1[1].ToString() + Read_Data_1[2].ToString() + Read_Data_1[3].ToString() + Read_Data_1[4].ToString() + Read_Data_1[5].ToString() + "Sattam_JF_Gr.pdf"; /////
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);




                    String[] alltextfild = { "Name", "studentId", "Dep", "Major", "Admin", "EAdmin", "job", "phone", "organization", "UMajor", "NAdmin" };

                    pdfFormFields.SetFieldProperty("Name", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("Dep", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("Major", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("Admin", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("EAdmin", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("job", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("phone", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("organization", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("UMajor", "textsize", 10f, null);



                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Dep", Read_Data_1[2].ToString());
                    pdfFormFields.SetField("Major", Read_Data_1[3].ToString());


                    pdfFormFields.SetField("Admin", "مهندس علي حميد الدين");
                    pdfFormFields.SetField("EAdmin", "ahameed@kfmc.med.sa");
                    pdfFormFields.SetField("job", "مستشار تقنية معلومات");
                    pdfFormFields.SetField("phone", "0538692448");

                    pdfFormFields.SetField("organization", "مدينة الملك فهد الطبيه");
                    pdfFormFields.SetField("UMajor", "تدريب الانظمة");
                    pdfFormFields.SetField("NAdmin", Read_Data_1[4].ToString());

                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();
                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[1].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[6].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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

            DB.CloseDB();
        }

        //Form IMAM_ER1
        public void IMAM_ER1(int NumberButten)
        {

            cheaked_erorr_send = 0;
            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "SELECT[arabicName],[collegeId] ,[majorAr]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[cell] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    string pdfTemplate = Path_Templte + "\\Imam_ER1.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation(1).pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





                    String[] alltextfild = { "Name", "studentId", "Dep", "phone", "startDate", "endDate", "organization" };

                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("endDate", "textsize", 10f, null);
                    pdfFormFields.SetFieldProperty("phone", "textsize", 10f, null);


                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("startDate", "07\\06\\2020");
                    pdfFormFields.SetField("endDate", "20\\07\\2020");
                    pdfFormFields.SetField("Dep", Read_Data_1[2].ToString());
                    pdfFormFields.SetField("phone", Read_Data_1[9].ToString());
                    pdfFormFields.SetField("organization", "مدينة الملك فهد الطبيه");



                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[6].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();

        }
        //Form IMAM_AS
        public void IMAM_AS(int NumberButten)
        {

            cheaked_erorr_send = 0;
            int count = 0;



            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    pdfTemplate = Path_Templte + "\\Imam_AS.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_AS.pdf";
                    }
                    else
                    {
                        // Evaluation(1)
                        // Attendance Sheet
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Attendance Sheet.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

                    string name = Read_Data_1[0].ToString();
                    //get first name of intern
                    String[] cutFristName = name.Split(' ');



                    String[] alltextfild = { "Name", "studentId", "Major", "institution" };


                    pdfFormFields.SetFieldProperty("studentId", "textsize", 11f, null);




                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());

                    pdfFormFields.SetField("Major", Read_Data_1[2].ToString());

                    pdfFormFields.SetField("institution", "مدينة الملك فهد الطبيه");
                    // pdfFormFields.SetFieldProperty("studentId", "setfflags", new , null);


                    for (int i = 1; i <= 40; i++)
                    {
                        pdfFormFields.SetField("checkin" + i, cutFristName[0]);

                    }
                    for (int i = 1; i <= 40; i++)
                    {
                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);

                    }
                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();

        }
        //Form Nourh_ER2
        public void Nourh_ER2(int NumberButten)
        {
            cheaked_erorr_send = 0;

            int count = 0;
            //   MessageBox.Show("processing......");
            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "SELECT[arabicName],[collegeId] ,[majorAr] ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";
                    count++;

                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    pdfTemplate = Path_Templte + "\\Nourh-ER2.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-ER(2).pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Evaluation(2).pdf";
                    }
                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;


                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

                    pdfFormFields.SetFieldProperty("IDNumber", "textsize", 11f, null);
                    pdfFormFields.SetFieldProperty("NameStudent", "textsize", 14f, null);




                    pdfFormFields.SetField("NameStudent", Read_Data_1[0].ToString());
                    pdfFormFields.SetField("IDNumber", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Major", Read_Data_1[2].ToString());


                    pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);






                    pdfStamper.Close();

                    Boolean cheakEmail = false;

                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();
        }
        //Form Nourh_ER1
        public void Nourh_ER1(int NumberButten)
        {

            cheaked_erorr_send = 0;
            int count = 0;
            // MessageBox.Show("processing......");
            String Datee = DateTime.Now.ToString("d/M/yyyy");

            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();


                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi],[numberWeeklyInternship],[internshipTitle],[datesInternship],[trainingYearId]FROM[intern]where[internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

                    string pdfTemplate = Path_Templte + "\\Nourh-ER1.pdf";



                    Read_Data_1.Read();

                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1["arabicName"].ToString() + "-" + Read_Data_1["collegeId"].ToString() + "-Evaluation(1).pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1["arabicName"].ToString() + "-" + Read_Data_1["collegeId"].ToString() + "-Evaluation(1).pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.AddSubstitutionFont(arialBaseFont);



                    String Name_En = Read_Data_1[7].ToString() + " " + Read_Data_1[9].ToString() + " " + Read_Data_1[8].ToString();

                    String[] alltextfild = { "InternName", "InternID", "internshipTitle", "DatesIntership"
                                , "NumberWeekly","Semester","NameStudent","IDNumber", "Date"};

                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "textsize", 10f, null);

                    }

                    pdfFormFields.SetField("InternName", Name_En);
                    pdfFormFields.SetField("InternID", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("internshipTitle", Read_Data_1[11].ToString());
                    pdfFormFields.SetField("DatesIntership", Read_Data_1[12].ToString());
                    pdfFormFields.SetField("NumberWeekly", Read_Data_1[10].ToString());
                    pdfFormFields.SetField("Semester", Read_Data_1[13].ToString());
                    pdfFormFields.SetField("NameStudent", Name_En);
                    pdfFormFields.SetField("IDNumber", Read_Data_1[1].ToString());
                    pdfFormFields.SetField("Date", Datee);

                    for (int i = 0; i < alltextfild.Length; i++)
                    {
                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();

                    Boolean cheakEmail = false;
                    if (NumberButten != 1)
                    {
                        if (NumberButten == 2)
                        {
                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
                        }
                        if (NumberButten == 3)
                        {
                            String MessgEmail = @"Dear Supervaisor:
this is your student information";
                            Supject = "ID# " + Read_Data_1[11].ToString() + "Name#" + Read_Data_1[12].ToString();
                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
                        }
                        if (cheakEmail == false)
                        {
                            cheaked_erorr_send++;
                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[0].ToString() + "");
                        }
                        else
                        {
                            File.Delete(newFile);
                        }


                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();

        }
        //Certificate
        public void Certificate(int NumberButten)
        {


            cheaked_erorr_send = 0;
            //for language Arbic
            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            int count = 0;
            String Gender = "/المتدرب";
            Form1 ff = new Form1();
            
            if (form1.listBox1Courses.SelectedItem == "co-op Training Certificate")
            {
                foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
                {


                    if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                    {

                        DB.CloseDB();

                        count++;

                        Query = "select * from studentInfo where id=" + it.Cells[1].Value.ToString() + "";



                        SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);
                        //////
                        pdfTemplate = Path_Templte + "\\certificate.pdf";



                        Read_Data_1.Read();

                        DateTime dateFrom = DateTime.Parse(Read_Data_1[6].ToString());
                        DateTime dateeEnd = DateTime.Parse(Read_Data_1[7].ToString());

                        String day = dateFrom.Day.ToString();
                        String month = dateFrom.Month.ToString();
                        string year = dateFrom.Year.ToString();

                        if (dateFrom.Day < 10)
                        {
                            day = "0" + dateFrom.Day;
                        }
                        if (dateFrom.Month < 10)
                        {
                            month = "0" + dateFrom.Month;
                        }

                        String DateFromEn = day + "\\" + month + "\\" + year;

                        day = dateeEnd.Day.ToString();
                        month = dateeEnd.Month.ToString();
                        year = dateeEnd.Year.ToString();

                        if (dateFrom.Day < 10)
                        {
                            day = "0" + dateeEnd.Day;
                        }
                        if (dateFrom.Month < 10)
                        {
                            month = "0" + dateeEnd.Month;
                        }
                        String DateEndEn = day + "\\" + month + "\\" + year;

                        DateTime DateFrom_Ar = Convert.ToDateTime(Read_Data_1[3].ToString());
                        DateTime DateEnd_Ar = Convert.ToDateTime(Read_Data_1[4].ToString());

                        //  CultureInfo ci = new CultureInfo("ar-SA");
                        String DateFromAr = ConvertToEasternArabicNumerals(DateFrom_Ar.ToString("dd/MM/yyyy"));
                        String DateEndAr = ConvertToEasternArabicNumerals(DateEnd_Ar.ToString("dd/MM/yyyy"));
                        if (NumberButten == 1)
                        {
                            newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[5].ToString() + "_certificate.pdf";
                        }
                        else
                        {
                            newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + ".pdf";
                        }
                        pdfReader = new PdfReader(pdfTemplate);
                        pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                        pdfFormFields = pdfStamper.AcroFields;
                        var arialBaseFont_BoldArbic = BaseFont.CreateFont(Path_language + "\\BoldArbic.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        pdfFormFields.AddSubstitutionFont(arialBaseFont_BoldArbic);


                        String[] textFild_TemplateextFild ={"internGender","txtNameAr", "txtNumberAr"
                               ,"datefrom", "datetoAr","txtNameEn","NumberEn","datetoEn","txtdateto"};

                        String ID_ar = ConvertToEasternArabicNumerals(Read_Data_1[2].ToString());
                        /// Change Color Text in pdf 
                        for (int i = 0; i < textFild_TemplateextFild.Length; i++)
                        {
                            pdfFormFields.SetFieldProperty(textFild_TemplateextFild[i], "textcolor", new BaseColor(51, 104, 5), null);

                        }



                        if (Read_Data_1[10].Equals(2))
                        {

                            Gender = "/المتدربة"; ;
                        }
                        //fill tet field in pdf templte
                        //(Id TextFild)     (Value)
                        pdfFormFields.SetField("internGender", Gender);
                        pdfFormFields.SetField("txtNameAr", Read_Data_1[1].ToString());
                        pdfFormFields.SetField("txtNumberAr", ID_ar);
                        pdfFormFields.SetField("datefrom", DateFromAr);
                        pdfFormFields.SetField("datetoAr", DateEndAr);
                        pdfFormFields.SetField("txtNameEn", Read_Data_1[9].ToString());
                        pdfFormFields.SetField("NumberEn", Read_Data_1[5].ToString());
                        pdfFormFields.SetField("datetoEn", DateFromEn);
                        pdfFormFields.SetField("txtdateto", DateEndEn);

                        //change field Properties to be read only

                        for (int i = 0; i < textFild_TemplateextFild.Length; i++)
                        {
                            pdfFormFields.SetFieldProperty(textFild_TemplateextFild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);

                        }

                        pdfStamper.FormFlattening = true;

                        pdfStamper.Close();


                        if (NumberButten == 2)
                        {

                            String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
                            Supject += "File#" + Read_Data_1[0].ToString() + "-KFMC Trining " + Read_Data_1[9].ToString();
                            //send Email 
                            Boolean cheakEmail = EmailSend.SendEmaile(Read_Data_1[8].ToString(), newFile, Supject, BodyEmail);
                            if (cheakEmail == false)
                            {
                                cheaked_erorr_send++;
                                MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[5].ToString() + "");
                            }
                            else
                            {
                                File.Delete(newFile);
                            }

                        }
                        Read_Data_1.Close();
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
                    else
                    {
                        MessageBox.Show("Please Chooes Intern");
                    }

                }

            }
            DB.CloseDB();
        }

        public void Execution(int NumberUnv, int NumberButten)
        {
            if (form1.listBox1Courses.SelectedItem == "co-op Training Certificate")//chooes of listBox1Courses
            {
                Certificate(NumberButten);


            }

            else if (form1.listBox1Courses.SelectedIndex == 0)///AS
            {

                if (form1.Un.SelectedIndex == 3)//Suad
                {
                    Suad_AS(NumberButten);
                }
                if (form1.Un.SelectedIndex == 2)
                {
                    Nourh_AS(NumberButten);
                }
                if (form1.Un.SelectedIndex == 1)
                {
                    IMAM_AS(NumberButten);

                }
                if (form1.Un.SelectedIndex == 3)
                {
                }
                if (form1.Un.SelectedIndex == 4)
                {

                }
                if (form1.Un.SelectedIndex == 6)//TVC
                {
                    TVTC_JF(NumberButten);
                }
            }
            else if (form1.listBox1Courses.SelectedIndex == 1)///ER1
            {


                if (form1.Un.SelectedIndex == 3)
                {
                    Suad_ER1(NumberButten);
                }
                if (form1.Un.SelectedIndex == 2)
                {
                    Nourh_ER1(NumberButten);
                }
                if (form1.Un.SelectedIndex == 1)
                {
                    IMAM_ER1(NumberButten);

                }
                if (form1.Un.SelectedIndex == 4)//Sq
                {
                    Shaqra_ER(NumberButten);

                }
                if (form1.Un.SelectedIndex == 5)
                {
                    Sattam_ER(NumberButten);

                }



            }

            else if (form1.listBox1Courses.SelectedIndex == 2)//ER2
            {
                if (form1.Un.SelectedIndex == 3)
                {
                    Suad_ER2(NumberButten);
                }
                if (form1.Un.SelectedIndex == 2)
                {
                    Nourh_ER2(NumberButten);
                }
                if (form1.Un.SelectedIndex == 1)
                {


                }
                if (form1.Un.SelectedIndex == 5)
                {
                    Sattam_JF(NumberButten);
                }
                if (form1.Un.SelectedIndex == 4)
                {
                    Shaqra_JF(NumberButten);
                }
            }

            else if (form1.listBox1Courses.SelectedIndex == 3)//JF
            {
                if (form1.Un.SelectedIndex == 3)
                {
                    Suad_JF(NumberButten);
                }

            }
        }

        ////////////////////////////////////////////////////////////////
        ///

        public void Templet(int NumberButten, String []Id, String[] Var,String NameTemble)
        {
            count = 0;

            cheaked_erorr_send = 0;
            count = 0;

            foreach (DataGridViewRow it in form1.guna2DataGridView1.Rows)
            {


                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
                {

                    DB.CloseDB();



                    Query = "SELECT [arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName]FROM[intern]where [internId]=" + it.Cells[1].Value.ToString() + "";

                    count++;
                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);
                    //path
                    pdfTemplate = Path_Templte + "\\"+ NameTemble + ".pdf";



                    Read_Data_1.Read();
                    string name = Read_Data_1[Var[0]].ToString();
                    //get first name of intern
                    String[] cutFristName = name.Split(' ');
                    if (NumberButten == 1)
                    {
                        newFile = form1.txtUploadFile.Text.ToString() + "\\" + Read_Data_1[Var[0]].ToString() + "_AS.pdf";
                    }
                    else
                    {
                        newFile = Path_Files + "\\" + Read_Data_1[Var[0]].ToString() + "-" + Read_Data_1[Var[1]].ToString() + "-Attendance Sheet.pdf";
                    }
                    pdfReader = new PdfReader(pdfTemplate);
                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    pdfFormFields = pdfStamper.AcroFields;

                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

                    //  pdfFormFields.SetFieldProperty("IDNumber", "textsize", 10f, null);

                    //pdfFormFields.SetField("NameStudent", Read_Data_1[0].ToString());
                    //pdfFormFields.SetField("IDNumber", Read_Data_1[1].ToString());
                    //pdfFormFields.SetField("Major", Read_Data_1[2].ToString());
                    //pdfFormFields.SetField("CourseCode", Read_Data_1[3].ToString());
                    pdfFormFields.SetField(Id[0], Read_Data_1[Var[0]].ToString());
                    pdfFormFields.SetField(Id[1], Read_Data_1[Var[1]].ToString());
                    pdfFormFields.SetField(Id[2], Read_Data_1[Var[2]].ToString());
                    pdfFormFields.SetField(Id[3], Read_Data_1[Var[3]].ToString());
                    for (int i = 1; i <= 25; i++)
                    {
                        pdfFormFields.SetField(Id[4] + i, cutFristName[0]);
                        pdfFormFields.SetField(Id[5] + i, cutFristName[0]);
                    }

                    for (int i =0;i<Id.Length;i++) {
                        pdfFormFields.SetFieldProperty(Id[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }

                    //pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    //pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
                    //pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);


                    for (int i = 1; i <= 25; i++)
                    {
                        pdfFormFields.SetFieldProperty(Id[4] + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
                        pdfFormFields.SetFieldProperty(Id[5] + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }


                    pdfStamper.FormFlattening = false;

                    pdfStamper.Close();
                    //   pdfReader.Close();


                    if (NumberButten != 1)
                    {
                        //EmailSend.SendEmail(NumberButten, Read_Data_1, newFile);
                    }
                    Read_Data_1.Close();
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
            DB.CloseDB();

        }






        //String[] ID =   { "studentName", "studentId", "major", "companyName",
        //                "companyAddress", "supervisorName", "", "supervisorJob", "supervisorDepartment", "supevisorEmail", "supervisorPhone" ,
        //                "task1" , "task2","task3","task4","task5","houraPerWeek","hoursTotal","startDate",
        //                "supervisorName", "", "supervisorJob",  "supervisorSign","studentName", "studentId","studentSign"};


        //String[] Varr = { "arabicName", "collegeId", "majorAr", "CourseCode" };
        //String name = "SEU_Training_RF";


        //        //Form Nourh_ER1
        //        public void Nourh_ER1(int NumberButten)
        //        {

        //            cheaked_erorr_send = 0;
        //            int count = 0;
        //            // MessageBox.Show("processing......");
        //            String Datee = DateTime.Now.ToString("d/M/yyyy");

        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();


        //                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi],[numberWeeklyInternship],[internshipTitle],[datesInternship],[trainingYearId]FROM[intern]where[internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    string pdfTemplate = Path_Templte + "\\Nourh-ER1.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1["arabicName"].ToString() + "-" + Read_Data_1["collegeId"].ToString() + "-Evaluation(1).pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1["arabicName"].ToString() + "-" + Read_Data_1["collegeId"].ToString() + "-Evaluation(1).pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);



        //                    String Name_En = Read_Data_1[7].ToString() + " " + Read_Data_1[9].ToString() + " " + Read_Data_1[8].ToString();

        //                    String[] alltextfild = { "InternName", "InternID", "internshipTitle", "DatesIntership"
        //                                , "NumberWeekly","Semester","NameStudent","IDNumber", "Date"};

        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "textsize", 10f, null);

        //                    }

        //                    pdfFormFields.SetField("InternName", Name_En);
        //                    pdfFormFields.SetField("InternID", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("internshipTitle", Read_Data_1[11].ToString());
        //                    pdfFormFields.SetField("DatesIntership", Read_Data_1[12].ToString());
        //                    pdfFormFields.SetField("NumberWeekly", Read_Data_1[10].ToString());
        //                    pdfFormFields.SetField("Semester", Read_Data_1[13].ToString());
        //                    pdfFormFields.SetField("NameStudent", Name_En);
        //                    pdfFormFields.SetField("IDNumber", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Date", Datee);

        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[11].ToString() + "Name#" + Read_Data_1[12].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[0].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }



        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }

        //        public void Nourh_AS(int NumberButten)
        //        {
        //            count = 0;

        //            cheaked_erorr_send = 0;
        //            count = 0;

        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)///
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)//
        //                {

        //                    DB.CloseDB();//



        //                    Query = "SELECT [arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName]FROM[intern]where [internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;//
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);//
        //                    //path
        //                    pdfTemplate = Path_Templte + "\\Nourh-AS.pdf";//



        //                    Read_Data_1.Read();
        //                    string name = Read_Data_1[0].ToString();
        //                    //get first name of intern
        //                    String[] cutFristName = name.Split(' ');
        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_AS.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Attendance Sheet.pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

        //                    pdfFormFields.SetFieldProperty("IDNumber", "textsize", 10f, null);

        //                    pdfFormFields.SetField("NameStudent", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("IDNumber", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Major", Read_Data_1[2].ToString());
        //                    pdfFormFields.SetField("CourseCode", Read_Data_1[3].ToString());
        //                    for (int i = 1; i <= 25; i++)
        //                    {
        //                        pdfFormFields.SetField("checkin" + i, cutFristName[0]);
        //                        pdfFormFields.SetField("checkout" + i, cutFristName[0]);
        //                    }



        //                    pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);


        //                    for (int i = 1; i <= 25; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();
        //                    //   pdfReader.Close();


        //                    if (NumberButten != 1)
        //                    {
        //                        //EmailSend.SendEmail(NumberButten, Read_Data_1, newFile);
        //                    }
        //                    Read_Data_1.Close();
        //                }




        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }
        //        //Form Suad_ER1
        //        public void Suad_ER1(int NumberButten)
        //        {
        //            cheaked_erorr_send = 0;

        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()))
        //                {

        //                    DB.CloseDB();



        //                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);


        //                    pdfTemplate = Path_Templte + "\\Suad_ER1.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation1.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "-Evaluation1.pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));

        //                    pdfFormFields = pdfStamper.AcroFields;


        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);






        //                    String[] alltextfild = { "studentId", "Name", "startDate", "endDate" };


        //                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("endDate", "textsize", 10f, null);



        //                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("startDate", "07\\06\\2020");
        //                    pdfFormFields.SetField("endDate", "20\\07\\2020");


        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();
        //                    //pdfReader.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();

        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }
        //        //Form Suad_ER2
        //        public void Suad_ER2(int NumberButten)
        //        {
        //            cheaked_erorr_send = 0;

        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    string pdfTemplate = Path_Templte + "\\Suad_ER2.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation2.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation2.pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





        //                    String[] alltextfild = { "Name", "studentId", "startDate", "institution" };



        //                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);


        //                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("startDate", "07\\06\\2020");
        //                    pdfFormFields.SetField("institution", "مدينة الملك فهد الطبيه");


        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[0].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();

        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }
        //        //Form Suad_AS
        //        public void Suad_AS(int NumberButten)
        //        {

        //            cheaked_erorr_send = 0;
        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[mi] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    string pdfTemplate = Path_Templte + "\\Suad_AS.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_AS.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Attendance Sheet.pdf";
        //                    } // Evaluation(1)
        //                    // Attendance Sheet
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);


        //                    string name = Read_Data_1[0].ToString();
        //                    //get first name of intern
        //                    String[] cutFristName = name.Split(' ');
        //                    String Name_En = Read_Data_1[7].ToString() + " " + Read_Data_1[9].ToString() + " " + Read_Data_1[8].ToString();

        //                    String[] alltextfild = { "Name", "studentId", "startDate", "institution" };

        //                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);

        //                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());



        //                    for (int i = 1; i <= 40; i++)
        //                    {
        //                        pdfFormFields.SetField("checkin" + i, cutFristName[0]);
        //                        pdfFormFields.SetField("checkout" + i, cutFristName[0]);

        //                    }
        //                    for (int i = 1; i <= 40; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty("checkout" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }

        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }
        //        //Form Suad_JF
        //        public void Suad_JF(int NumberButten)
        //        {

        //            cheaked_erorr_send = 0;
        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();


        //                    Query = "select [collegeId],[arabicName],[majorAr],[cell],[email],[supervisorName],[supervisorEmail],[refNo] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";




        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    string pdfTemplate = Path_Templte + "\\Saud-JF.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Effective Date Form.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Effective Date Form.pdf";
        //                    } // Evaluation(1)
        //                    // Attendance Sheet
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);







        //                    String[] alltextfild = { "Name", "collegeId", "StPhone", "StPhone2", "StEmail", "Institution", "Address", "Department", "Dep2", "SupName", "Position", "SupPhoneH", "SupPhone", "SupEmail", "StartDate", "Rd1", "Rd2" };


        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "textsize", 10f, null);

        //                    }
        //                    pdfFormFields.SetFieldProperty("Name", "textsize", 12f, null);


        //                    pdfFormFields.SetField("Name", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("collegeId", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("StPhone", Read_Data_1[3].ToString());
        //                    pdfFormFields.SetField("StPhone2", Read_Data_1[3].ToString());
        //                    pdfFormFields.SetField("StEmail", Read_Data_1[4].ToString());
        //                    pdfFormFields.SetField("Institution", "King Fahad Medical City");
        //                    pdfFormFields.SetField("Address", "Sulimaniyah, Riyadh 12231");
        //                    pdfFormFields.SetField("Department", "Executive Administration of");
        //                    pdfFormFields.SetField("Dep2", "Information Technology/Application Training Department");
        //                    pdfFormFields.SetField("SupName", "Ali Mohamed Hamidaddin");
        //                    pdfFormFields.SetField("Position", "IT Consultant");
        //                    pdfFormFields.SetField("SupPhoneH", "+96611288999");
        //                    pdfFormFields.SetField("SupPhone", "0538692448");
        //                    pdfFormFields.SetField("SupEmail", Read_Data_1[6].ToString());
        //                    pdfFormFields.SetField("StartDate", "07\\06\\2020");


        //                    string major = Read_Data_1[2].ToString();
        //                    string majorTrack = Read_Data_1[2].ToString();

        //                    if (major == "علوم الحاسب")
        //                    {
        //                        pdfFormFields.SetField("Rd1", "2");
        //                    }
        //                    else if (major == "تقنية المعلومات")
        //                    {
        //                        pdfFormFields.SetField("Rd1", "3");
        //                        if (majorTrack == "Network & Security")
        //                        {
        //                            pdfFormFields.SetField("Rd2", "5");
        //                        }
        //                        else if (majorTrack == "Data Management")
        //                        {
        //                            pdfFormFields.SetField("Rd2", "4");
        //                        }

        //                        else if (majorTrack == "Web Technologies & Multimedia")
        //                        {
        //                            pdfFormFields.SetField("Rd2", "3");
        //                        }

        //                        else if (majorTrack == "Cyber Security")
        //                        {
        //                            pdfFormFields.SetField("Rd2", "2");
        //                        }
        //                        else if (majorTrack == "Data Science")
        //                        {
        //                            pdfFormFields.SetField("Rd2", "1");
        //                        }
        //                        else if (majorTrack == "Networks & Internet of Things")
        //                        {
        //                            pdfFormFields.SetField("Rd2", "0");
        //                        }


        //                    }

        //                    else if (major == "SWE")
        //                    {
        //                        pdfFormFields.SetField("Rd1", "1");
        //                    }
        //                    else if (major == "IS")
        //                    {
        //                        pdfFormFields.SetField("Rd1", "0");
        //                    }



        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }



        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[0].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[0].ToString() + "Name#" + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[6].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }

        //            DB.CloseDB();
        //        }
        //        //Form  TVTC_JF
        //        public void TVTC_JF(int NumberButten)

        //        {
        //            cheaked_erorr_send = 0;

        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()))
        //                {

        //                    DB.CloseDB();



        //                    Query = "SELECT [arabicName], [collegeId],[major],[supervisorEmail],[email],[refNo] from [intern] where [internId]= " + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);


        //                    pdfTemplate = Path_Templte + "\\TVTC_JF.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_JF.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_JF.pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));

        //                    pdfFormFields = pdfStamper.AcroFields;


        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





        //                    String[] alltextfild = { "numStudent", "idStudent", "department", "ddlMajor",
        //                        "date", "numManager", "jobManager", "telNo", "fax", "email", "mobile" };

        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {

        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "textsize", 10f, null);
        //                    }





        //                    pdfFormFields.SetField("numStudent", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("idStudent", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("department", "IT");
        //                    pdfFormFields.SetField("ddlMajor", Read_Data_1[2].ToString());
        //                    pdfFormFields.SetField("date", "07\\06\\2020");
        //                    pdfFormFields.SetField("numManager", "Ali Mohamed Hamidaddin");
        //                    pdfFormFields.SetField("jobManager", "IT consultant");
        //                    pdfFormFields.SetField("telNo", "(+966) 11 288 9999");
        //                    pdfFormFields.SetField("fax", "19100");
        //                    pdfFormFields.SetField("email", "Ahameed@kfmc.med.sa");
        //                    pdfFormFields.SetField("mobile", "(+966) 538692448");



        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();
        //                    //pdfReader.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[5].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[3].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();

        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }

        //            DB.CloseDB();
        //        }

        //        //Form Shaqra_ER
        //        public void Shaqra_ER(int NumberButten)
        //        {
        //            cheaked_erorr_send = 0;

        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()))
        //                {

        //                    DB.CloseDB();



        //                    Query = "select [arabicName],[majorAr],[registrationDate],[UploadedDate],[supervisorEmail],[email],[collegeId] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);


        //                    pdfTemplate = Path_Templte + "\\Shaqra_ER.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[6].ToString() + Read_Data_1[0].ToString() + "_Evaluation.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[6].ToString() + "-" + Read_Data_1[1].ToString() + "_Evaluation.pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));

        //                    pdfFormFields = pdfStamper.AcroFields;


        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);






        //                    String[] alltextfild = { "name", "Side", "mejor", "Nmejor", "job", "date", "Edate" };

        //                    // "name","id","track","mobile","email","date"
        //                    pdfFormFields.SetFieldProperty("name", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("Side", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("mejor", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("Nmejor", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("job", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("date", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("Edate", "textsize", 10f, null);

        //                    DateTime dateEnd_format = Convert.ToDateTime(Read_Data_1[3].ToString());
        //                    string dateeEnd = string.Format("{0:dd/MM/yyyy}", dateEnd_format);
        //                    DateTime dateStart_format = Convert.ToDateTime(Read_Data_1[2].ToString());
        //                    string dateStart = string.Format("{0:dd/MM/yyyy}", dateStart_format);


        //                    pdfFormFields.SetField("name", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("Side", "مدينة ملك فهد الطبية ");
        //                    pdfFormFields.SetField("mejor", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Nmejor", "مهندس علي حميدالدين");
        //                    pdfFormFields.SetField("job", "مستشار تقنية معلومات");
        //                    pdfFormFields.SetField("date", dateStart);
        //                    pdfFormFields.SetField("Edate", dateeEnd);


        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();
        //                    //pdfReader.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[1].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[6].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[3].ToString() + ")" + Read_Data_1[0].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();

        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }

        //            DB.CloseDB();
        //        }


        //        //Shaqra_JF
        //        public void Shaqra_JF(int NumberButten)
        //        {
        //            cheaked_erorr_send = 0;

        //            int count = 0;
        //            //   MessageBox.Show("processing......");
        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "select [collegeId],[arabicName],[majorAr],[cell],[email],[supervisorEmail],[refNo]  from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";
        //                    count++;

        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    pdfTemplate = Path_Templte + "\\Shaqra_JF.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_JF.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_JF.pdf";
        //                    }
        //                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;


        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

        //                    pdfFormFields.SetFieldProperty("id", "textsize", 11f, null);
        //                    //pdfFormFields.SetFieldProperty("NameStudent", "textsize", 14f, null);




        //                    pdfFormFields.SetField("id", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("name", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("track", Read_Data_1[2].ToString());
        //                    pdfFormFields.SetField("mobile", Read_Data_1[3].ToString());
        //                    pdfFormFields.SetField("email", Read_Data_1[4].ToString());
        //                    pdfFormFields.SetField("date", "07\\06\\2020");


        //                    pdfFormFields.SetFieldProperty("id", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("name", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("track", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("mobile", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("email", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("date", "setfflags", PdfFormField.FF_READ_ONLY, null);







        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;

        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[0].ToString() + "Name#" + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[0].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }

        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }
        //        //Form Sattam_ER
        //        public void Sattam_ER(int NumberButten)
        //        {

        //            cheaked_erorr_send = 0;
        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "select [collegeId],[arabicName],[fName],[lName],[email],[supervisorEmail],[refNo] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    pdfTemplate = Path_Templte + "\\Sattam_ER.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Evaluation.pdf";
        //                    }
        //                    else
        //                    {
        //                        // Evaluation(1)
        //                        // Attendance Sheet
        //                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "_" + Read_Data_1[1].ToString() + "_Evaluation.pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





        //                    String[] alltextfild = { "Name", "OrgName", "Date", "SName", "Sign" };






        //                    pdfFormFields.SetField("Name", Read_Data_1[1].ToString());

        //                    pdfFormFields.SetField("OrgName", "مدينة الملك فهد الطبيه");
        //                    pdfFormFields.SetField("SName", "علي محمد حميد الدين");
        //                    pdfFormFields.SetField("Sign", "علي ");
        //                    pdfFormFields.SetField("Date", DateTime.Now.ToString("dd/MM/yyyy"));
        //                    // pdfFormFields.SetFieldProperty("studentId", "setfflags", new , null);




        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[2].ToString() + Read_Data_1[3].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[0].ToString() + "Name#" + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }

        //            DB.CloseDB();
        //        }

        //        //Form Sattam_JF
        //        public void Sattam_JF(int NumberButten)
        //        {

        //            cheaked_erorr_send = 0;
        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "select [arabicName],[collegeId],[trainingYearId], [majorAr],[supervisorName],[supervisorEmail],[email] from [intern] where [internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    string pdfTemplate = Path_Templte + "\\Sattam_JF.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[0].ToString() + "Sattam_JF_Gr.pdf"; //////
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[0].ToString() + "-" + Read_Data_1[1].ToString() + Read_Data_1[2].ToString() + Read_Data_1[3].ToString() + Read_Data_1[4].ToString() + Read_Data_1[5].ToString() + "Sattam_JF_Gr.pdf"; /////
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);




        //                    String[] alltextfild = { "Name", "studentId", "Dep", "Major", "Admin", "EAdmin", "job", "phone", "organization", "UMajor", "NAdmin" };

        //                    pdfFormFields.SetFieldProperty("Name", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("Dep", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("Major", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("Admin", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("EAdmin", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("job", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("phone", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("organization", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("UMajor", "textsize", 10f, null);



        //                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Dep", Read_Data_1[2].ToString());
        //                    pdfFormFields.SetField("Major", Read_Data_1[3].ToString());


        //                    pdfFormFields.SetField("Admin", "مهندس علي حميد الدين");
        //                    pdfFormFields.SetField("EAdmin", "ahameed@kfmc.med.sa");
        //                    pdfFormFields.SetField("job", "مستشار تقنية معلومات");
        //                    pdfFormFields.SetField("phone", "0538692448");

        //                    pdfFormFields.SetField("organization", "مدينة الملك فهد الطبيه");
        //                    pdfFormFields.SetField("UMajor", "تدريب الانظمة");
        //                    pdfFormFields.SetField("NAdmin", Read_Data_1[4].ToString());

        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();
        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[1].ToString() + "- Distance Training 2020 for " + Read_Data_1[0].ToString() + Read_Data_1[1].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[6].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }

        //            DB.CloseDB();
        //        }

        //        //Form IMAM_ER1
        //        public void IMAM_ER1(int NumberButten)
        //        {

        //            cheaked_erorr_send = 0;
        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "SELECT[arabicName],[collegeId] ,[majorAr]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName],[cell] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    string pdfTemplate = Path_Templte + "\\Imam_ER1.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation.pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "_" + Read_Data_1[0].ToString() + "_Evaluation(1).pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);





        //                    String[] alltextfild = { "Name", "studentId", "Dep", "phone", "startDate", "endDate", "organization" };

        //                    pdfFormFields.SetFieldProperty("studentId", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("startDate", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("endDate", "textsize", 10f, null);
        //                    pdfFormFields.SetFieldProperty("phone", "textsize", 10f, null);


        //                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("startDate", "07\\06\\2020");
        //                    pdfFormFields.SetField("endDate", "20\\07\\2020");
        //                    pdfFormFields.SetField("Dep", Read_Data_1[2].ToString());
        //                    pdfFormFields.SetField("phone", Read_Data_1[9].ToString());
        //                    pdfFormFields.SetField("organization", "مدينة الملك فهد الطبيه");



        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[6].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }
        //        //Form IMAM_AS
        //        public void IMAM_AS(int NumberButten)
        //        {

        //            cheaked_erorr_send = 0;
        //            int count = 0;



        //            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "SELECT[arabicName],[collegeId] ,[major]  ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";

        //                    count++;
        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    pdfTemplate = Path_Templte + "\\Imam_AS.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "_AS.pdf";
        //                    }
        //                    else
        //                    {
        //                        // Evaluation(1)
        //                        // Attendance Sheet
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Attendance Sheet.pdf";
        //                    }
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;

        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

        //                    string name = Read_Data_1[0].ToString();
        //                    //get first name of intern
        //                    String[] cutFristName = name.Split(' ');



        //                    String[] alltextfild = { "Name", "studentId", "Major", "institution" };


        //                    pdfFormFields.SetFieldProperty("studentId", "textsize", 11f, null);




        //                    pdfFormFields.SetField("studentId", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Name", Read_Data_1[0].ToString());

        //                    pdfFormFields.SetField("Major", Read_Data_1[2].ToString());

        //                    pdfFormFields.SetField("institution", "مدينة الملك فهد الطبيه");
        //                    // pdfFormFields.SetFieldProperty("studentId", "setfflags", new , null);


        //                    for (int i = 1; i <= 40; i++)
        //                    {
        //                        pdfFormFields.SetField("checkin" + i, cutFristName[0]);

        //                    }
        //                    for (int i = 1; i <= 40; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty("checkin" + i, "setfflags", PdfFormField.FF_READ_ONLY, null);

        //                    }
        //                    for (int i = 0; i < alltextfild.Length; i++)
        //                    {
        //                        pdfFormFields.SetFieldProperty(alltextfild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    }


        //                    pdfStamper.FormFlattening = false;

        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;
        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }


        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();

        //        }
        //        //Form Nourh_ER2
        //        public void Nourh_ER2(int NumberButten)
        //        {
        //            cheaked_erorr_send = 0;

        //            int count = 0;
        //            //   MessageBox.Show("processing......");
        //            foreach (DataGridViewRow it in guna2DataGridView1.Rows)
        //            {


        //                if (bool.Parse(it.Cells[0].Value.ToString()) == true)
        //                {

        //                    DB.CloseDB();



        //                    Query = "SELECT[arabicName],[collegeId] ,[majorAr] ,[CourseCode],[email],[supervisorEmail],[refNo],[fName],[lName] FROM[intern] where[internId]=" + it.Cells[1].Value.ToString() + "";
        //                    count++;

        //                    SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);

        //                    pdfTemplate = Path_Templte + "\\Nourh-ER2.pdf";



        //                    Read_Data_1.Read();

        //                    if (NumberButten == 1)
        //                    {
        //                        newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-ER(2).pdf";
        //                    }
        //                    else
        //                    {
        //                        newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + "-" + Read_Data_1[0].ToString() + "-Evaluation(2).pdf";
        //                    }
        //                    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        //                    pdfReader = new PdfReader(pdfTemplate);
        //                    pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
        //                    pdfFormFields = pdfStamper.AcroFields;


        //                    pdfFormFields.AddSubstitutionFont(arialBaseFont);

        //                    pdfFormFields.SetFieldProperty("IDNumber", "textsize", 11f, null);
        //                    pdfFormFields.SetFieldProperty("NameStudent", "textsize", 14f, null);




        //                    pdfFormFields.SetField("NameStudent", Read_Data_1[0].ToString());
        //                    pdfFormFields.SetField("IDNumber", Read_Data_1[1].ToString());
        //                    pdfFormFields.SetField("Major", Read_Data_1[2].ToString());


        //                    pdfFormFields.SetFieldProperty("NameStudent", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("IDNumber", "setfflags", PdfFormField.FF_READ_ONLY, null);
        //                    pdfFormFields.SetFieldProperty("Major", "setfflags", PdfFormField.FF_READ_ONLY, null);






        //                    pdfStamper.Close();

        //                    Boolean cheakEmail = false;

        //                    if (NumberButten != 1)
        //                    {
        //                        if (NumberButten == 2)
        //                        {
        //                            String BodyEmail = @"Dear student:
        //thank you for end of internship this is your certificate";
        //                            Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
        //                        }
        //                        if (NumberButten == 3)
        //                        {
        //                            String MessgEmail = @"Dear Supervaisor:
        //this is your student information";
        //                            Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
        //                            cheakEmail = EmailSend.SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
        //                        }
        //                        if (cheakEmail == false)
        //                        {
        //                            cheaked_erorr_send++;
        //                            MessageBox.Show("Send Email Failure ((" + Read_Data_1[0].ToString() + ")" + Read_Data_1[1].ToString() + "");
        //                        }
        //                        else
        //                        {
        //                            File.Delete(newFile);
        //                        }


        //                    }
        //                    Read_Data_1.Close();
        //                }
        //            }

        //            if (count > 0)
        //            {
        //                if (cheaked_erorr_send != count)
        //                {
        //                    if (NumberButten == 1)
        //                    {
        //                        MessageBox.Show("Save successful.....");
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Send Email successful.....");
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Please Chooes Intern");
        //            }
        //            DB.CloseDB();
        //        }


        /////////////////////////////////////////
        //else if (listBox1Courses.SelectedIndex == 0)///AS
        //{

        //    if (Un.SelectedIndex == 8) //SEU_JF
        //    {
        //        //"task1" , "task2","task3","task4","task5"
        //        String[] alltextfild = { "studentName", "studentId", "major", "houraPerWeek", "hoursTotal", "startDate" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "numberWeeklyInternship", "trainingHours", "datesInternship" };
        //        String Name_form = "SEU_JF";

        //        Templet(NumberButten, alltextfild, Var, Name_form, 0);
        //    }


        //    if (Un.SelectedIndex == 3)//Suad
        //    {
        //        String[] id = { "Name", "studentId", "checkin", "checkout" };
        //        String[] Var = { "arabicName", "collegeId" };
        //        String name = "Suad_AS";

        //        Templet(NumberButten, id, Var, name, 1);

        //    }
        //    if (Un.SelectedIndex == 2)//Nourh
        //    {
        //        String[] id = { "NameStudent", "IDNumber", "Major", "CourseCode", "checkin", "checkout" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "CourseCode" };
        //        String name = "Nourh_AS";

        //        Templet(NumberButten, id, Var, name, 1);
        //    }
        //    if (Un.SelectedIndex == 1)//Imam
        //    {
        //        String[] id = { "Name", "studentId", "Major", "institution", "checkin", "checkout" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "majorAr" };
        //        String name = "Imam_AS";

        //        Templet(NumberButten, id, Var, name, 1);


        //    }
        //    if (Un.SelectedIndex == 3)
        //    {
        //    }
        //    if (Un.SelectedIndex == 4)
        //    {

        //    }
        //    if (Un.SelectedIndex == 6)//TVC_jf
        //    {
        //        String[] id = { "numStudent", "idStudent", "department", "ddlMajor",
        //            "date", "numManager", "jobManager", "telNo", "fax", "email", "mobile"};
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "majorAr" };
        //        String name = "TVTC_JF";

        //        Templet(NumberButten, id, Var, name, 0);

        //    }
        //}
        //else if (listBox1Courses.SelectedIndex == 1)///ER1
        //{


        //    if (Un.SelectedIndex == 3)
        //    {
        //        String[] id = { "Name", "studentId", "startDate", "endDate" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "majorAr" };
        //        String name = "Suad_ER1";
        //        Templet(NumberButten, id, Var, name, 0);
        //    }
        //    if (Un.SelectedIndex == 2)
        //    {
        //        String[] id = {"InternName", "InternID", "internshipTitle", "DatesIntership"
        //                    , "NumberWeekly","Semester","Date","NameStudent","IDNumber" };
        //        String[] Var = { "collegeId", "collegeId", "arabicName", "arabicName", "majorAr", "majorAr", "Date" };
        //        String name = "Nourh_ER1";
        //        Templet(NumberButten, id, Var, name, 0);

        //    }
        //    if (Un.SelectedIndex == 1)//IMAM_ER1
        //    {
        //        String[] id = { "Name", "studentId", "Dep", "phone", "startDate", "endDate", "organization" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "cell" };
        //        String name = "Imam_ER1";
        //        Templet(NumberButten, id, Var, name, 0);


        //    }
        //    if (Un.SelectedIndex == 4)//Sq ER
        //    {

        //        String[] id = { "name", "Side", "mejor", "Nmejor", "job", "date", "Edate" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "cell" };
        //        String name = "Shaqra_ER";
        //        Templet(NumberButten, id, Var, name, 0);


        //    }
        //    if (Un.SelectedIndex == 5)//Sattam ER
        //    {
        //        String[] id = { "Name", "OrgName", "Date", "SName", "Sign" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "cell" };
        //        String name = "Sattam_ER";
        //        Templet(NumberButten, id, Var, name, 0);

        //    }



        //}

        //else if (listBox1Courses.SelectedIndex == 2)//ER2
        //{
        //    if (Un.SelectedIndex == 3)// Suad_ER2
        //    {
        //        String[] id = { "Name", "studentId", "startDate", "institution" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "cell" };
        //        String name = "Suad_ER2";
        //        Templet(NumberButten, id, Var, name, 0);

        //    }
        //    if (Un.SelectedIndex == 2)//Nourh_ER2
        //    {
        //        String[] id = { "NameStudent", "IDNumber", "Major" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr" };
        //        String name = "Nourh-ER2";
        //        Templet(NumberButten, id, Var, name, 0);

        //    }
        //    if (Un.SelectedIndex == 1)
        //    {


        //    }
        //    if (Un.SelectedIndex == 5)//Sattam_JF
        //    {

        //        String[] id = { "Name", "studentId", "Dep", "Major", "Admin", "EAdmin", "job", "phone", "organization", "UMajor", "NAdmin" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "majorAr", "majorAr", "majorAr", "cell" };
        //        String name = "Sattam_JF";
        //        Templet(NumberButten, id, Var, name, 0);


        //    }
        //    if (Un.SelectedIndex == 4)//Shaqra_JF
        //    {
        //        String[] id = { "name", "id", "track", "mobile", "email", "date" };
        //        String[] Var = { "arabicName", "collegeId", "majorAr", "cell", "email", "majorAr" };
        //        String name = "Shaqra_JF";
        //        Templet(NumberButten, id, Var, name, 0);

        //    }
        //}

        //else if (listBox1Courses.SelectedIndex == 3)//JF
        //{
        //    if (Un.SelectedIndex == 3)
        //    {

        //        String[] id = { "Name", "collegeId", "StPhone", "StPhone2", "StEmail", "Institution", "Address", "Department", "Dep2", "SupName", "Position", "SupPhoneH", "SupPhone", "SupEmail", "StartDate", "Rd1", "Rd2" };
        //        String[] Var = { "arabicName", "collegeId", "cell", "cell", "email", "majorAr", "cell", "majorAr" };
        //        String name = "Saud-JF";
        //        Templet(NumberButten, id, Var, name, 0);


        //    }

        //}
    }
}
//public void Templet(int NumberButten, String[] IdTextFiled, String[] Var, String NameTemplet, int cheack)
//{
//    String Date_tody = DateTime.Now.ToString("d/M/yyyy");
//    count = 0;

//    cheaked_erorr_send = 0;


//    foreach (DataGridViewRow it in guna2DataGridView1.Rows)
//    {


//        if (bool.Parse(it.Cells[0].Value.ToString()) == true)
//        {

//            DB.CloseDB();

//            Query = "SELECT ";
//            for (int i = 0; i < Var.Length; i++)
//            {

//                if (Var[i] != "Date")
//                {
//                    Query += Var[i] + ",";
//                }

//            }



//            Query += " [refNo],[fName],[lName],[mi],[email] ,[supervisorName] FROM[intern]where[internId] = " + it.Cells[1].Value.ToString() + "";

//            count++;
//            SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);
//            //path
//            pdfTemplate = Path_Templte + "\\" + NameTemplet + ".pdf";



//            Read_Data_1.Read();
//            string name = Read_Data_1["arabicName"].ToString();
//            //get first name of intern
//            String[] cutFristName = name.Split(' ');
//            if (NumberButten == 1)
//            {
//                newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1["arabicName"].ToString() + "_" + NameTemplet + ".pdf";
//            }
//            else
//            {
//                newFile = Path_Files + "\\" + Read_Data_1["arabicName"].ToString() + "_" + Read_Data_1["collegeId"].ToString() + "_" + NameTemplet + ".pdf";
//            }
//            pdfReader = new PdfReader(pdfTemplate);
//            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
//            pdfFormFields = pdfStamper.AcroFields;

//            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
//            pdfFormFields.AddSubstitutionFont(arialBaseFont);


//            for (int i = 0; i < Var.Length; i++)
//            {
//                try
//                {
//                    pdfFormFields.SetField(IdTextFiled[i], Read_Data_1[Var[i]].ToString());
//                }
//                catch (Exception ex)
//                {
//                    pdfFormFields.SetField(IdTextFiled[i], Date_tody);
//                }
//            }

//            if (cheack == 1)
//            {

//                for (int i = 1; i <= 25; i++)
//                {
//                    pdfFormFields.SetField(IdTextFiled[IdTextFiled.Length - 2] + i, cutFristName[0]);
//                    pdfFormFields.SetField(IdTextFiled[IdTextFiled.Length - 1] + i, cutFristName[0]);
//                    pdfFormFields.SetFieldProperty(IdTextFiled[IdTextFiled.Length - 2] + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
//                    pdfFormFields.SetFieldProperty(IdTextFiled[IdTextFiled.Length - 1] + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
//                }
//            }

//            for (int i = 0; i < IdTextFiled.Length; i++)
//            {
//                pdfFormFields.SetFieldProperty(IdTextFiled[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
//            }





//            pdfStamper.FormFlattening = false;

//            pdfStamper.Close();



//            if (NumberButten != 1)
//            {
//                Boolean cheakEmail = false;
//                if (NumberButten != 1)
//                {
//                    if (NumberButten == 2)
//                    {
//                        String BodyEmail = @"Dear student:
//thank you for end of internship this is your certificate";
//                        Supject = @"File# " + Read_Data_1["refNo"].ToString() + "_Distance Training 2020 for " + Read_Data_1["fName"].ToString() + Read_Data_1["lName"].ToString();
//                        cheakEmail = EmailSend.SendEmaile(Read_Data_1["email"].ToString(), newFile, Supject, BodyEmail);
//                    }
//                    if (NumberButten == 3)
//                    {
//                        String MessgEmail = @"Dear Supervaisor:
//this is your student information";
//                        Supject = "ID# " + Read_Data_1["collegeId"].ToString() + "Name#" + Read_Data_1["arabicName"].ToString();
//                        cheakEmail = EmailSend.SendEmaile(Read_Data_1["supervisorName"].ToString(), newFile, Supject, MessgEmail);
//                    }
//                    if (cheakEmail == false)
//                    {
//                        cheaked_erorr_send++;
//                        MessageBox.Show("Send Email Failure ((" + Read_Data_1["collegeId"].ToString() + ")" + Read_Data_1["arabicName"].ToString() + "");
//                    }
//                    else
//                    {
//                        File.Delete(newFile);
//                    }
//                }
//            }
//            Read_Data_1.Close();

//        }



//    }


//    if (count > 0)
//    {
//        if (cheaked_erorr_send != count)
//        {
//            if (NumberButten == 1)
//            {
//                MessageBox.Show("Save successful.....");
//            }
//            else
//            {
//                MessageBox.Show("Send Email successful.....");
//            }
//        }
//    }
//    else
//    {
//        MessageBox.Show("Please Chooes Intern");
//    }
//    // DB.CloseDB();

//}////////////////////////////////////


//public void ImportDataFromExcel(string excelFilePath)
//{
//    //declare variables - edit these based on your particular situation   
//    string ssqltable = "Table1";
//    // make sure your sheet name is correct, here sheet name is sheet1,
//    so you can change your sheet name if have    different
//    string myexceldataquery = "select student,rollno,course from [Sheet1$]";
//    try
//    {
//        //create our connection strings   
//        string sexcelconnectionstring = @"provider=microsoft.jet.oledb.4.0;data source=" + excelFilePath +
//        ";extended properties=" + "\"excel 8.0;hdr=yes;\"";
//        string ssqlconnectionstring = "Data Source=SAYYED;Initial Catalog=SyncDB;Integrated Security=True";
//        //execute a query to erase any previous data from our destination table   
//        string sclearsql = "delete from " + ssqltable;
//        SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
//        SqlCommand sqlcmd = new SqlCommand(sclearsql, sqlconn);
//        sqlconn.Open();
//        sqlcmd.ExecuteNonQuery();
//        sqlconn.Close();
//        //series of commands to bulk copy data from the excel file into our sql table   
//        OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
//        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
//        oledbconn.Open();
//        OleDbDataReader dr = oledbcmd.ExecuteReader();
//        SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
//        bulkcopy.DestinationTableName = ssqltable;
//        while (dr.Read())
//        {
//            bulkcopy.WriteToServer(dr);
//        }
//        dr.Close();
//        oledbconn.Close();
//        Label1.Text = "File imported into sql server successfully.";
//    }
//    catch (Exception ex)
//    {
//        //handle exception   
//    }
//}///////////////////////////////////////////////////////////bbbbbbbbbbbbbbbbbb////
//public String FormatDate(String Date)
//{
//    DateTime date = DateTime.Parse(Date);
//    String day = date.Day.ToString();
//    String month = date.Month.ToString();
//    string year = date.Year.ToString();

//    if (date.Day < 10)
//    {
//        day = "0" + date.Day;
//    }
//    if (date.Month < 10)
//    {
//        month = "0" + date.Month;
//    }

//    return day + "\\" + month + "\\" + year;
//}
//String Query_Certificates = "SELECT [id]  ,[NameAr]  ,[nationalNum] As NationalID  ,[email]  As Email ,[NameEn]  FROM [studentInfo]";
//String Query = @"SELECT  [internId] ,[accepted] ,[nationalityId],[arabicName] ,[major],[email] ,[supervisorName],[supervisorCell],[supervisorEmail]FROM[intern]";


//public static string ConvertToEasternArabicNumerals(string input)
//{
//    System.Text.UTF8Encoding utf8Encoder = new UTF8Encoding();
//    System.Text.Decoder utf8Decoder = utf8Encoder.GetDecoder();
//    System.Text.StringBuilder convertedChars = new System.Text.StringBuilder();
//    char[] convertedChar = new char[1];
//    byte[] bytes = new byte[] { 217, 160 };
//    char[] inputCharArray = input.ToCharArray();
//    foreach (char c in inputCharArray)
//    {
//        if (char.IsDigit(c))
//        {
//            bytes[1] = Convert.ToByte(160 + char.GetNumericValue(c));
//            utf8Decoder.GetChars(bytes, 0, 2, convertedChar, 0);
//            convertedChars.Append(convertedChar[0]);
//        }
//        else
//        {
//            convertedChars.Append(c);
//        }
//    }
//    return convertedChars.ToString();
//}



//public void Templet(int NumberButten, String[] IdTextFiled, String[] Var, String NameTemplet, int cheack)
//{
//    String Date_tody = DateTime.Now.ToString("d/M/yyyy");
//    count = 0;

//    cheaked_erorr_send = 0;


//    foreach (DataGridViewRow it in guna2DataGridView1.Rows)
//    {


//        if (bool.Parse(it.Cells[0].Value.ToString()) == true)
//        {

//            DB.CloseDB();

//            Query = "SELECT ";
//            for (int i = 0; i < Var.Length; i++)
//            {

//                if (Var[i] != "Date")
//                {
//                    Query += Var[i] + ",";
//                }

//            }



//            Query += " [refNo],[fName],[lName],[mi],[email] ,[supervisorName] FROM[intern]where[internId] = " + it.Cells[1].Value.ToString() + "";

//            count++;
//            SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);
//            //path
//            pdfTemplate = Path_Templte + "\\" + NameTemplet + ".pdf";



//            Read_Data_1.Read();
//            string name = Read_Data_1["arabicName"].ToString();
//            //get first name of intern
//            String[] cutFristName = name.Split(' ');
//            if (NumberButten == 1)
//            {
//                newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1["arabicName"].ToString() + "_" + NameTemplet + ".pdf";
//            }
//            else
//            {
//                newFile = Path_Files + "\\" + Read_Data_1["arabicName"].ToString() + "_" + Read_Data_1["collegeId"].ToString() + "_" + NameTemplet + ".pdf";
//            }
//            pdfReader = new PdfReader(pdfTemplate);
//            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
//            pdfFormFields = pdfStamper.AcroFields;

//            var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
//            pdfFormFields.AddSubstitutionFont(arialBaseFont);


//            for (int i = 0; i < Var.Length; i++)
//            {
//                try
//                {
//                    pdfFormFields.SetField(IdTextFiled[i], Read_Data_1[Var[i]].ToString());
//                }
//                catch (Exception ex)
//                {
//                    pdfFormFields.SetField(IdTextFiled[i], Date_tody);
//                }
//            }

//            if (cheack == 1)
//            {

//                for (int i = 1; i <= 25; i++)
//                {
//                    pdfFormFields.SetField(IdTextFiled[IdTextFiled.Length - 2] + i, cutFristName[0]);
//                    pdfFormFields.SetField(IdTextFiled[IdTextFiled.Length - 1] + i, cutFristName[0]);
//                    pdfFormFields.SetFieldProperty(IdTextFiled[IdTextFiled.Length - 2] + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
//                    pdfFormFields.SetFieldProperty(IdTextFiled[IdTextFiled.Length - 1] + i, "setfflags", PdfFormField.FF_READ_ONLY, null);
//                }
//            }

//            for (int i = 0; i < IdTextFiled.Length; i++)
//            {
//                pdfFormFields.SetFieldProperty(IdTextFiled[i], "setfflags", PdfFormField.FF_READ_ONLY, null);
//            }





//            pdfStamper.FormFlattening = false;

//            pdfStamper.Close();



//            if (NumberButten != 1)
//            {
//                Boolean cheakEmail = false;
//                if (NumberButten != 1)
//                {
//                    if (NumberButten == 2)
//                    {
//                        String BodyEmail = @"Dear student:
//thank you for end of internship this is your certificate";
//                        Supject = @"File# " + Read_Data_1["refNo"].ToString() + "_Distance Training 2020 for " + Read_Data_1["fName"].ToString() + Read_Data_1["lName"].ToString();
//                        cheakEmail = EmailSend.SendEmaile(Read_Data_1["email"].ToString(), newFile, Supject, BodyEmail);
//                    }
//                    if (NumberButten == 3)
//                    {
//                        String MessgEmail = @"Dear Supervaisor:
//this is your student information";
//                        Supject = "ID# " + Read_Data_1["collegeId"].ToString() + "Name#" + Read_Data_1["arabicName"].ToString();
//                        cheakEmail = EmailSend.SendEmaile(Read_Data_1["supervisorName"].ToString(), newFile, Supject, MessgEmail);
//                    }
//                    if (cheakEmail == false)
//                    {
//                        cheaked_erorr_send++;
//                        MessageBox.Show("Send Email Failure ((" + Read_Data_1["collegeId"].ToString() + ")" + Read_Data_1["arabicName"].ToString() + "");
//                    }
//                    else
//                    {
//                        File.Delete(newFile);
//                    }
//                }
//            }
//            Read_Data_1.Close();

//        }



//    }


//    if (count > 0)
//    {
//        if (cheaked_erorr_send != count)
//        {
//            if (NumberButten == 1)
//            {
//                MessageBox.Show("Save successful.....");
//            }
//            else
//            {
//                MessageBox.Show("Send Email successful.....");
//            }
//        }
//    }
//    else
//    {
//        MessageBox.Show("Please Chooes Intern");
//    }
//    // DB.CloseDB();

//}

//public void Certificate(int NumberButten)
//{


//    cheaked_erorr_send = 0;
//    //for language Arbic
//    var arialBaseFont = BaseFont.CreateFont(Path_language + "\\MAJALLA.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
//    int count = 0;
//    String Gender = "/المتدرب";


//    foreach (DataGridViewRow it in guna2DataGridView1.Rows)
//    {


//        if (bool.Parse(it.Cells[0].Value.ToString()) == true)
//        {

//            DB.CloseDB();

//            count++;

//            Query = "select * from studentInfo where id=" + it.Cells[1].Value.ToString() + "";



//            SqlDataReader Read_Data_1 = DB.getDataFromDataDase(Query);
//            //////
//            pdfTemplate = Path_Templte + "\\certificate.pdf";



//            Read_Data_1.Read();



//            String DateFromEn = FormatDate(Read_Data_1[6].ToString());
//            String DateEndEn = FormatDate(Read_Data_1[7].ToString());




//            DateTime DateFrom_Ar = Convert.ToDateTime(Read_Data_1[3].ToString());
//            DateTime DateEnd_Ar = Convert.ToDateTime(Read_Data_1[4].ToString());

//            //  CultureInfo ci = new CultureInfo("ar-SA");
//            String DateFromAr = ConvertToEasternArabicNumerals(DateFrom_Ar.ToString("dd/MM/yyyy"));
//            String DateEndAr = ConvertToEasternArabicNumerals(DateEnd_Ar.ToString("dd/MM/yyyy"));
//            if (NumberButten == 1)
//            {
//                newFile = txtUploadFile.Text.ToString() + "\\" + Read_Data_1[5].ToString() + "_certificate.pdf";
//            }
//            else
//            {
//                newFile = Path_Files + "\\" + Read_Data_1[1].ToString() + ".pdf";
//            }
//            pdfReader = new PdfReader(pdfTemplate);
//            pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
//            pdfFormFields = pdfStamper.AcroFields;
//            var arialBaseFont_BoldArbic = BaseFont.CreateFont(Path_language + "\\BoldArbic.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
//            pdfFormFields.AddSubstitutionFont(arialBaseFont_BoldArbic);


//            String[] textFild_TemplateextFild ={"internGender","txtNameAr", "txtNumberAr"
//                               ,"datefrom", "datetoAr","txtNameEn","NumberEn","datetoEn","txtdateto"};

//            String ID_ar = ConvertToEasternArabicNumerals(Read_Data_1[2].ToString());
//            /// Change Color Text in pdf 
//            for (int i = 0; i < textFild_TemplateextFild.Length; i++)
//            {
//                pdfFormFields.SetFieldProperty(textFild_TemplateextFild[i], "textcolor", new BaseColor(51, 104, 5), null);

//            }



//            //if (Read_Data_1[10].Equals(2))
//            //{

//            //    Gender = "/المتدربة"; ;
//            //}
//            //fill tet field in pdf templte
//            //(Id TextFild)     (Value)
//            pdfFormFields.SetField("internGender", Gender);
//            pdfFormFields.SetField("txtNameAr", Read_Data_1[1].ToString());
//            pdfFormFields.SetField("txtNumberAr", ID_ar);
//            pdfFormFields.SetField("datefrom", DateFromAr);
//            pdfFormFields.SetField("datetoAr", DateEndAr);
//            pdfFormFields.SetField("txtNameEn", Read_Data_1[9].ToString());
//            pdfFormFields.SetField("NumberEn", Read_Data_1[5].ToString());
//            pdfFormFields.SetField("datetoEn", DateFromEn);
//            pdfFormFields.SetField("txtdateto", DateEndEn);

//            //change field Properties to be read only

//            for (int i = 0; i < textFild_TemplateextFild.Length; i++)
//            {
//                pdfFormFields.SetFieldProperty(textFild_TemplateextFild[i], "setfflags", PdfFormField.FF_READ_ONLY, null);

//            }

//            pdfStamper.FormFlattening = true;

//            pdfStamper.Close();


//            if (NumberButten == 2)
//            {

//                String BodyEmail = @"Dear student:
//thank you for end of internship this is your certificate";
//                Supject += "File#" + Read_Data_1[0].ToString() + "-KFMC Trining " + Read_Data_1[9].ToString();
//                //send Email 
//                Boolean cheakEmail = EmailSend.SendEmaile(Read_Data_1[8].ToString(), newFile, Supject, BodyEmail);
//                if (cheakEmail == false)
//                {
//                    cheaked_erorr_send++;
//                    MessageBox.Show("Send Email Failure ((" + Read_Data_1[1].ToString() + ")" + Read_Data_1[5].ToString() + "");
//                }
//                else
//                {
//                    File.Delete(newFile);
//                }

//            }
//            Read_Data_1.Close();
//        }

//    }


//    if (count > 0)
//    {
//        if (cheaked_erorr_send != count)
//        {
//            if (NumberButten == 1)
//            {
//                MessageBox.Show("Save successful.....");
//            }
//            else
//            {
//                MessageBox.Show("Send Email successful.....");
//            }
//        }
//        else
//        {
//            MessageBox.Show("Please Chooes Intern");
//        }

//    }

//    //}
//    DB.CloseDB();
//}

//String Query = @"SELECT  [internId] ,[accepted] ,[nationalityId],[arabicName] ,[major],[email] ,[supervisorName],[supervisorCell],[supervisorEmail]FROM[intern]";

// if (Un.SelectedIndex == 6) { 
//{
//    SendToSupervisor.Enabled = false;


//    // String Query = "SELECT [id]  ,[NameAr]  ,[nationalNum] As National ID  ,[email] AS Email,[NameEn] AS Name FROM[certificate].[dbo].[studentInfo]";
//    Query = "SELECT [id]  ,[NameAr]  ,[nationalNum] As NationalID  ,[email]  As Email ,[NameEn]  FROM [studentInfo]";

//    DB.getDataGrid(Query, guna2DataGridView1);
//    addheadercheckbox();
//    headercheckbox.MouseClick += new MouseEventHandler(headerMousclick);

//  }
//else
//{
///////////////excel
///
//Label1.Text = "File imported into sql server successfully.";

//  OpenFileDialog op = new OpenFileDialog();
//  op.Filter = "ALL Files |*.*| Excel Files |*.XLSX";
//if (op.ShowDialog() == DialogResult.OK)
//{
//string sexcelconnectionstring = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=1;'";
//OleDbConnection oeledbconn = new OleDbConnection(@"Provider =Microsoft.ACE.OLEDB.12.0;Data Source=" + bx1.Text.ToString() + ";Extended Properties='Excel 12.0;IMEX=1;'");
//OleDbDataAdapter adp = new OleDbDataAdapter(@"select * from[s1$]", oledbconn);
//DataTable da = new DataTable();
//adp.Fill(da);
//  listBox1Courses.SelectedIndex = listBox1Courses.SelectedIndex;






//String path = @"C:\Users\HP\Desktop\certificateSystem\certificateSystem\Files\ex11.xlsx";
////   //  @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';"
////   string sexcelconnectionstring=@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = "+ path + "; Extended Properties = Excel 12.0 Xml;HDR=YES";
//////   string sexcelconnectionstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12HDR=yes;\",";

//OpenFileDialog op = new OpenFileDialog();
//op.Filter = "ALL Files |*.*| Excel Files |*.XLSX";
//if (op.ShowDialog() == DialogResult.OK) {
//    //string sexcelconnectionstring = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=1;'";
//    OleDbConnection oledbconn = new OleDbConnection(@"Provider =Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=1;'");
//    OleDbDataAdapter adp = new OleDbDataAdapter(@"select * from[s1$]", oledbconn);
//    DataTable da = new DataTable();
//    adp.Fill(da);


//    guna2DataGridView1.DataSource = da;



//  OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
//OleDbDataAdapter adp = new OleDbDataAdapter(@"select * from[Sheet1$]", oledbconn);
//DataTable da = new DataTable();
////DataSet execl = new DataSet();
//adp.Fill(da);

////adp.Fill(da);
//guna2DataGridView1.DataSource = da;