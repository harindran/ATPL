﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ATPL
{
    public class STDFunc
    {

        public  string ObjtoStr(object Value)
        {
            string Returnstring = "";

            if (Value == null)
                return Returnstring;

            Returnstring = Convert.ToString(Value);
            return Returnstring;
        }

        public int Ctoint(object Pstring)
        {
            int Ctoint = 0;
            try
            {
                int LdblResult;
                if (Pstring==null)
                    return Ctoint;
                string Lstr = System.Convert.ToString(Pstring);

                if (int.TryParse(Lstr, out LdblResult))
                    Ctoint = LdblResult;
                return Ctoint;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public decimal CtoD(object Pstring)
        {
            decimal CtoD = 0;
            try
            {
                decimal LdblResult;
                if (Pstring==null)
                    return CtoD;
                string Lstr = System.Convert.ToString(Pstring);

                if (decimal.TryParse(Lstr, out LdblResult))
                    CtoD = LdblResult;
                return CtoD;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public decimal CtoD(object Pstring,int round)
        {
            decimal CtoD = 0;
            try
            {
                decimal LdblResult;
                if (Pstring == null)
                    return CtoD;
                string Lstr = System.Convert.ToString(Pstring);

                if (decimal.TryParse(Lstr, out LdblResult))
                    CtoD = LdblResult;

                if (round!=0)
                {
                    CtoD = Math.Round(CtoD, round);
                }
                return CtoD;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public DateTime GetDate(object str)
        {
            DateTime startDate;
           
            string strdate =ObjtoStr(str) ;
            string[] formatall = CultureInfo.CurrentCulture.DateTimeFormat.GetAllDateTimePatterns();
            List<string> list = new List<string>(formatall);
            list.Add("yyyyMMdd");

            formatall = list.ToArray();
            if (DateTime.TryParseExact(strdate, formatall, CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate))
            {
               
            }
            return startDate;
        }

       

        public string GetdateString(object Inv_Doc_Date, string returnformat)
        {

            string date = ObjtoStr(Inv_Doc_Date);
            if (string.IsNullOrEmpty(date))
                return "";
            

            string dtformate = GetDate(date).ToString(returnformat, CultureInfo.InvariantCulture);
            return dtformate;
        }

        public void WriteErrorLog(string Str)
        {
            try
            {

                string Foldername;
                Foldername = @"Log";
                if (Directory.Exists(Foldername))
                {
                }
                else
                {
                    Directory.CreateDirectory(Foldername);
                }

                FileStream fs;
                string chatlog = Foldername + @"\Log_" + DateTime.Now.ToString("ddMMyy") + ".txt";
                if (File.Exists(chatlog))
                {
                }
                else
                {
                    fs = new FileStream(chatlog, FileMode.Create, FileAccess.Write);
                    fs.Close();
                }
                string sdate;
                sdate = Convert.ToString(DateTime.Now);
                if (File.Exists(chatlog) == true)
                {
                    var objWriter = new StreamWriter(chatlog, true);
                    objWriter.WriteLine(sdate + " : " + Str);
                    objWriter.Close();
                }
                else
                {
                    var objWriter = new StreamWriter(chatlog, false);
                }
            }
            catch (Exception)
            {


            }
        }
        
        public  bool validMail(string Mailid)
        {
            bool validMail = true;

            validMail = Regex.IsMatch(Mailid, @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase);

            return validMail;
        }

        public  string validMail(string Mailid, char Delimeter)
        {

            string validMail = "";
            string[] split;

            split = Mailid.Split(Delimeter);
            foreach (string item in split)
            {
                if (string.IsNullOrEmpty(item)) continue;
                bool isvalidMail = Regex.IsMatch(item, @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase);
                if (!isvalidMail) { validMail += item + Delimeter; }
            }

            return validMail;
        }


    }
}