using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Common;
using LSEXT;
using DAL;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using Print_General;

namespace PrintAliquot
{

    [ComVisible(true)]
    [ProgId("PrintAliquot.PrintAliquotCls")]
    public class PrintAliquot : IWorkflowExtension
    {
        INautilusServiceProvider sp;
        private const string Type = "3";
        private int _port = 9100;
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        private IDataLayer dal;





        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {

                #region param

                string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];
                var workstationId = Parameters["WORKSTATION_ID"];
                var resultName = rs.Fields["NAME"].Value;
                var resultId = (long)rs.Fields["RESULT_ID"].Value;

                #endregion

                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);


                dal = new DataLayer();
                dal.Connect();
                Workstation ws = dal.getWorkStaitionById(workstationId);
                ReportStation reportStation = dal.getReportStationByWorksAndType(ws.NAME, Type);
                //            string ip = GetIp(printerName);
                string goodIp = ""; //removeBadChar(ip);
                string printerName = "";
                if (reportStation != null)
                {


                    if (reportStation.Destination != null)
                    {
                        //מביא את הIP של המדפסת להדפסה הזאת
                        goodIp = reportStation.Destination.ManualIP;
                    }

                    if (reportStation.Destination != null && reportStation.Destination.RawTcpipPort != null)
                    {
                        //מביא את הפורט רק במקרה שהוא דיפולט
                        _port = (int)reportStation.Destination.RawTcpipPort;
                    }
                    Result result = dal.GetResultById(resultId);
                    Aliquot aliquot = result.Test.Aliquot;// dal.GetAliquotById(resultId);

                    var sampleName = aliquot.Sample.Name;
                    // var sampleID = aliq.Sample.Name;//TODO : לבדוק אם זה הערך הנכון
                    var testcode = "";
                    //   testcode = getTestCode(aliquot, testcode);

                    Print(aliquot.Name, aliquot.AliquotId.ToString(), aliquot.ShortName, "", goodIp);
                }
                else
                    MessageBox.Show("לא הוגדרה מדפסת עבור תחנה זו.");


            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת מדבקה");
                Logger.WriteLogFile(ex);
            }
        }
        private void Print(string name, string ID, string testcode, string mihol, string ip)
        {
            string ipAddress = ip;


            // ZPL Command(s)
            string ntxt = name;
            string tctxt = testcode;
            string mtxt = mihol;
            string itxt = ID;


            string ZPLString =
                 "^XA" +
                 "^LH20,10" +
                 "^FO20,10" +
                 "^A@N20,20" +
                string.Format("^FD{0}^FS", ntxt) + //שם
                 "^FO10,60" +
                 "^A@N20,20" +

                 string.Format("^FD{0}^FS", mtxt) +
                "^FO100,60" +
                 "^A@N20,20" +

                 string.Format("^FD{0}^FS", tctxt) +
                "^FO320,0" + "^BQN,4,3" +
                //string.Format("^FD   {0}^FS", itxt) + //ברקוד
                    string.Format("^FDLA,{0}^FS", itxt) + //ברקוד
                "^XZ";
            try
            {
                // Open connection
                var client = new System.Net.Sockets.TcpClient();
                client.Connect(ipAddress, _port);
                // Write ZPL String to connection
                var writer = new StreamWriter(client.GetStream());
                writer.Write(ZPLString.Trim());
                writer.Flush();
                // Close Connection
                writer.Close();
                client.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        //     string tableName = Parameters["TABLE_NAME"];
        //     sp = Parameters["SERVICE_PROVIDER"];
        //     var rs = Parameters["RECORDS"];


        //     var workstationId = Parameters["WORKSTATION_ID"];

        //     ////////////יוצר קונקשן//////////////////////////
        //     var ntlCon = Utils.GetNtlsCon(sp);
        //     Utils.CreateConstring(ntlCon);
        //     /////////////////////////////           
        //     var dal = new DataLayer();
        //     dal.Connect();
        //     var pg = new PrintOperation(dal, workstationId, Type);
        //     var aliquotId = "";
        //     var aliquotName = "";
        //     var aliquotDescription = "";
        //     if (tableName == "ALIQUOT")
        //     {
        //         Aliquot aliquot = dal.GetAliquotById(Convert.ToInt32(rs.Fields["ALIQUOT_ID"].Value));
        //         aliquotId = aliquot.AliquotId.ToString();
        //         aliquotName = aliquot.Name;
        //         aliquotDescription = aliquot.Description;
        //     }

        //     pg.ManipulateHebrew(aliquotDescription);

        //     string zpl =
        // "^XA" +
        //     "^CI28^FT86,68^A@N,62,61,TT0003M_^FD^FS^PQ1" +//שורה שמטפלת בעברית
        // "^LH0,0" +
        // "^FO20,10" +
        // "^A@N20,20" +
        //string.Format("^FD{0}^FS", aliquotName) + //שם
        // "^FO20,60" +
        // "^A@N20,20" +

        // string.Format("^FD{0}^FS", aliquotDescription) +
        //"^FO150,60" +
        // "^A@N20,20" +

        // string.Format("^FD{0}^FS", aliquotId) +
        //"^FO260,0" + "^BQN,4,3" +
        //         //string.Format("^FD   {0}^FS", itxt) + //ברקוד
        //    string.Format("^FDLA,{0}^FS", aliquotId) + //ברקוד
        //"^XZ";
        //     pg.Print(zpl);
        // }
        // catch (Exception ex)
        // {
        //     Logger.WriteLogFile(ex);
        //     MessageBox.Show("נכשלה הדפסת מדבקה");
        // }
    }




}
