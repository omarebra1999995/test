using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace certificateSystem
{
    class EmailManger
    {
        String Email_Sent = ConfigurationSettings.AppSettings["emailUserName"];
        String PassWord = ConfigurationSettings.AppSettings["emailPassword"];
        String Path_Files = ConfigurationSettings.AppSettings["Path_Files"];
        public Boolean SendEmaile(String reseiver, String NameFile,String Supject,String  MessegEmil)
        {


            try
            {
                SmtpClient Client = new SmtpClient("smtp.gmail.com", 587);
                Client.Credentials = new NetworkCredential(Email_Sent, PassWord);

                MailMessage Massage = new MailMessage(Email_Sent, reseiver, Supject, MessegEmil);
             
                Attachment att = new Attachment(NameFile);

                Massage.Attachments.Add(att);
                Massage.IsBodyHtml = false;
                Client.EnableSsl = true;
             
                Client.Send(Massage);
               
                Massage.Dispose();
                Client.Dispose();
                att.Dispose();


                return true;
            }
            catch (Exception ex)
            {
                return false;

            }
        }

        public int SendEmail(int NumberButten , SqlDataReader Read_Data_1,String newFile)
        {
            Boolean cheakEmail = false;
           int  cheaked_erorr_send = 0;
            if (NumberButten == 2)
            {
                String BodyEmail = @"Dear student:
thank you for end of internship this is your certificate";
               String Supject = @"File# " + Read_Data_1[6].ToString() + "- Distance Training 2020 for " + Read_Data_1[7].ToString() + Read_Data_1[8].ToString();
                cheakEmail = SendEmaile(Read_Data_1[4].ToString(), newFile, Supject, BodyEmail);
            }
            else if (NumberButten == 3)
            {
                String MessgEmail = @"Dear Supervaisor:
this is your student information";
               String Supject = "ID# " + Read_Data_1[1].ToString() + "Name#" + Read_Data_1[0].ToString();
                cheakEmail = SendEmaile(Read_Data_1[5].ToString(), newFile, Supject, MessgEmail);
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
            return cheaked_erorr_send;
        }
        public Boolean sendAllFilesEmail(String reseiver, String Supject, String MessegEmil)
        {
            try
            {
                SmtpClient Client = new SmtpClient("smtp.gmail.com", 587);
                Client.Credentials = new NetworkCredential(Email_Sent, PassWord);
                MailMessage Massage = new MailMessage(Email_Sent, reseiver, Supject, MessegEmil);
                string[] filepathhs = Directory.GetFiles(Path_Files, "*pdf");
                foreach (var filepath in filepathhs)
                {
                    var attachment = new Attachment(filepath);   // here you can attach a file as a mail attachment  
                    Massage.Attachments.Add(attachment);
                }

                Massage.IsBodyHtml = false;
                Client.EnableSsl = true;

                Client.Send(Massage);

                Massage.Dispose();
                Client.Dispose();


                return true;
            }
            catch (Exception ex)
            {
                return false;

            }
        }


    }



}


    

