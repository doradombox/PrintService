using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Web;
using Trace.Common.UnicodeToZPL;

namespace PrintService
{
    /// <summary>
    /// Handler1 的摘要说明
    /// </summary>
    public class Handler1 : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/html";
            context.Response.AddHeader("Access-Control-Allow-Origin", "*");
            context.Response.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
            context.Response.AddHeader("Access-Control-Allow-Headers", "Content-Type");


            string materialNo = context.Request.QueryString["materialNo"];
            string branchNo = context.Request.QueryString["branchNo"];
            string serialNo = context.Request.QueryString["serialNo"];
            //模板类型1、2、3，指的是打印单排、双排、三排。每一种的纸张尺寸是不一样，这里纸张单排的尺寸为60*40*1,双排为40*20*2,三排为30*20*3
            string templateType = context.Request.QueryString["templateType"];
            string separator = context.Request.QueryString["separator"];

            string printername = ConfigurationManager.AppSettings["PrinterName"].ToString();
            if (!CheckPrinter(printername))
            {
                context.Response.Write("电脑中找不到名称为：" + printername + "的打印机");
                return;
            }
            if (!CheckPrinterInLine(printername))
            {
                context.Response.Write(printername + "打印机未连接");
                return;
            }

            ZebraPrintHelper.PrinterProgrammingLanguage = ProgrammingLanguage.ZPL;
            ZebraPrintHelper.PrinterName = printername;
            ZebraPrintHelper.PrinterType = DeviceType.DRV;
            StringBuilder sb = new StringBuilder();
            switch (templateType)
            {
                case "1":
                    print1Col(sb, materialNo, branchNo, serialNo, separator);
                    break;
                case "2":
                    print2Col(sb, materialNo, branchNo, serialNo,separator);
                    break;
                default:
                    print3Col(sb, materialNo, branchNo, serialNo, separator);
                    break;
            }
            ZebraPrintHelper.PrintCommand(sb.ToString());
            Thread.Sleep(200);
        }

        /// <summary>
        /// 打印单排的zpl指令,单排的整体偏下了，不知道为什么
        /// </summary>
        /// <param name="sb">完整的指令</param>
        /// <param name="materialNo">物料编号</param>
        /// <param name="branchNo">批次号</param>
        /// <param name="serialNo">流水号</param>
        private void print1Col(StringBuilder sb, string materialNo, string branchNo, string serialNo, string separator)
        {
            Font myfont = new Font("黑体", 16, FontStyle.Bold);
            string materialNoCN = UnicodeToZPL.UnCompressZPL("物料编号:" + materialNo, "materialNo", myfont, TextDirection.零度);
            string branchNoCN = UnicodeToZPL.UnCompressZPL("批次号:" + branchNo, "branchNo", myfont, TextDirection.零度);
            string serialNoCN = UnicodeToZPL.UnCompressZPL("流水号:" + serialNo, "serialNo", myfont, TextDirection.零度);
            string qrCode = materialNo + separator + branchNo + separator + serialNo;
            sb.Append("^XA");
            sb.Append("^FO200,30^BQ,2,5^FDQA," + qrCode + "^FS");
            sb.Append("^FO350,40^A0,26,26^FDQA," + materialNoCN + "^FS");
            sb.Append("^FO350,75^A0,26,26^FDQA," + branchNoCN + "^FS");
            sb.Append("^FO350,105^A0,26,26^FDQA," + serialNoCN + "^FS");
            sb.Append("^XZ");
        }

        /// <summary>
        /// 打印双排的zpl指令
        /// </summary>
        /// <param name="sb">完整的指令</param>
        /// <param name="materialNo">物料编号</param>
        /// <param name="branchNo">批次号</param>
        /// <param name="serialNo">流水号</param>
        private void print2Col(StringBuilder sb, string materialNo, string branchNo, string serialNo, string separator)
        {
            Font myfont = new Font("黑体", 16, FontStyle.Bold);
            string materialNoCN = UnicodeToZPL.UnCompressZPL("物料编号:" + materialNo, "materialNo", myfont, TextDirection.零度);
            string branchNoCN = UnicodeToZPL.UnCompressZPL("批次号:" + branchNo, "branchNo", myfont, TextDirection.零度);
            string serialNoCN = UnicodeToZPL.UnCompressZPL("流水号:" + serialNo, "serialNo", myfont, TextDirection.零度);
            string qrCode = materialNo + separator + branchNo + separator + serialNo;
            sb.Append("^XA");
            sb.Append("^FO80,20^BQ,2,5^FDQA," + qrCode + "^FS");
            sb.Append("^FO185,30^A0,26,26^FD," + materialNoCN + "^FS");
            sb.Append("^FO185,70^A0,26,26^FD," + branchNoCN + "^FS");
            sb.Append("^FO185,110^A0,26,26^FD," + serialNoCN + "^FS");
            sb.Append("^FO420,20^BQ,2,5^FDQA," + qrCode + "^FS");
            sb.Append("^FO530,30^A0,26,25^FD," + materialNoCN + "^FS");
            sb.Append("^FO530,70^A0,26,25^FD," + branchNoCN + "^FS");
            sb.Append("^FO530,110^A0,26,25^FD," + serialNoCN + "^FS");
            sb.Append("^XZ");
        }

        /// <summary>
        /// 打印三排的zpl指令
        /// </summary>
        /// <param name="sb">完整的指令</param>
        /// <param name="materialNo">物料编号</param>
        /// <param name="branchNo">批次号</param>
        /// <param name="serialNo">流水号</param>
        private void print3Col(StringBuilder sb, string materialNo, string branchNo, string serialNo, string separator)
        {
            Font myfont = new Font("黑体", 12, FontStyle.Bold);
            string materialNoCN = UnicodeToZPL.UnCompressZPL("物料编号:" + materialNo, "materialNo", myfont, TextDirection.零度);
            string branchNoCN = UnicodeToZPL.UnCompressZPL("批次号:" + branchNo, "branchNo", myfont, TextDirection.零度);
            string serialNoCN = UnicodeToZPL.UnCompressZPL("流水号:" + serialNo, "serialNo", myfont, TextDirection.零度);
            string qrCode = materialNo + separator + branchNo + separator + serialNo;
            sb.Append("^XA");
            sb.Append("^FO20,30^BQ,2,3^FDQA," + qrCode + "^FS");
            sb.Append("^FO95,40^A0,12,12^FD" + materialNoCN + "^FS");
            sb.Append("^FO95,70^A0,12,12^FD" + branchNoCN + "^FS");
            sb.Append("^FO95,100^A0,12,12^FD" + serialNoCN + "^FS");
            sb.Append("^FO300,30^BQ,2,3^FDQA," + qrCode + "^FS");
            sb.Append("^FO385,40^A0,12,12^FD" + materialNoCN + "^FS");
            sb.Append("^FO385,70^A0,12,12^FD" + branchNoCN + "^FS");
            sb.Append("^FO385,100^A0,12,12^FD" + serialNoCN + "^FS");
            sb.Append("^FO575,30^BQ,2,3^FDQA," + qrCode + "^FS");
            sb.Append("^FO655,40^A0,12,12^FD" + materialNoCN + "^FS");
            sb.Append("^FO655,70^A0,12,12^FD" + branchNoCN + "^FS");
            sb.Append("^FO655,100^A0,12,12^FD" + serialNoCN + "^FS");
            sb.Append("^XZ");
        }

        /// <summary>
        /// 检查打印机是否存在
        /// </summary>
        private bool CheckPrinter(string printername)
        {
            foreach (string sspirnt in PrinterSettings.InstalledPrinters)
            {

            }
            foreach (string sPrint in PrinterSettings.InstalledPrinters)//获取所有打印机名称
            {
                if (printername == sPrint)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 检查打印机是否在线
        /// </summary>
        private bool CheckPrinterInLine(string printername)
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");

            foreach (ManagementObject printer in searcher.Get())
            {
                if (printer["Name"].ToString() == printername && printer["WorkOffline"].ToString().ToLower().Equals("false"))
                    return true;
            }
            return false;
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}