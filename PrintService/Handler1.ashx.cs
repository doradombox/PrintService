using System.Configuration;
using System.Drawing;
using System.Drawing.Printing;
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
			string materialName = context.Request.QueryString["materialName"];
			//模板类型1、2、3，指的是打印单排、双排、三排。每一种的纸张尺寸是不一样，这里纸张单排的尺寸为60*40*1,双排为40*20*2,三排为30*20*3
			string templateType = context.Request.QueryString["templateType"];
            string separator = context.Request.QueryString["separator"];
			string productDrawNo = context.Request.QueryString["productDrawNo"];
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
					Print1Col(sb, materialNo, branchNo, serialNo, separator, materialName,productDrawNo);
                    break;
                case "2":
                    Print2Col(sb, materialNo, branchNo, serialNo,separator,materialName,productDrawNo);
                    break;
				case "4":
					Print1Col5030(sb, materialNo, branchNo, serialNo, separator, materialName,productDrawNo);
					break;
				default:
                    Print3Col(sb, materialNo, branchNo, serialNo, separator, materialName,productDrawNo);
                    break;
            }
            ZebraPrintHelper.PrintCommand(sb.ToString());
            Thread.Sleep(200);
        }

		/// <summary>
		/// 打印单排的二维码，尺寸为60mm*40mm。
		/// </summary>
		/// <param name="sb">完整的指令</param>
		/// <param name="materialNo">物料编号</param>
		/// <param name="branchNo">批次号</param>
		/// <param name="serialNo">流水号</param>
		/// <param name="materialName">物料名称</param>
		/// <param name="productDrawNo">产品图号</param>
		private void Print1Col(StringBuilder sb, string materialNo, string branchNo, string serialNo, string separator,string materialName,string productDrawNo)
        {
			Font myfont = new Font("黑体", 17, FontStyle.Bold);
			string branchNoSerialNo = branchNo + "-" + serialNo;//批次号+序号
			string productDrawNoCN = UnicodeToZPL.UnCompressZPL(productDrawNo, "productDrawNo", myfont, TextDirection.零度);
			string materialNameCN = UnicodeToZPL.UnCompressZPL(materialName, "materialName", myfont, TextDirection.零度);
			//string serialNoCN = UnicodeToZPL.UnCompressZPL(branchNoSerialNo, "branchNoSerialNo", myfont, TextDirection.零度);
			string qrCode = materialNo + separator + branchNo + separator + serialNo;
			sb.Append("^XA");
            sb.Append("^FO200,50,0^BQ,2,7^FDQA," + qrCode + "^FS");
            sb.Append("^FO450,75^A0,26,26^FDQA," + productDrawNoCN + "^FS");
            sb.Append("^FO450,105^A0,26,26^FDQA," + materialNameCN + "^FS");
            sb.Append("^FO450,135^A0,26,26^FD" + branchNoSerialNo + "^FS");
            sb.Append("^XZ");
        }

		/// <summary>
		/// 打印单排的二维码，尺寸为50mm*30mm。
		/// </summary>
		/// <param name="sb">完整的ZPL指令</param>
		/// <param name="materialNo">物料编号</param>
		/// <param name="branchNo">批次号</param>
		/// <param name="serialNo">序号</param>
		/// <param name="separator">分割符</param>
		/// <param name="materialName">物料名称</param>
		/// <param name="productDrawNo">物料图号</param>
		private void Print1Col5030(StringBuilder sb, string materialNo, string branchNo, string serialNo, string separator, string materialName,string productDrawNo)
		{
			Font myfont = new Font("黑体", 17, FontStyle.Bold);
			string branchNoSerialNo = branchNo + "-" + serialNo;//批次号+序号
			string productDrawNoCN = UnicodeToZPL.UnCompressZPL(productDrawNo, "productDrawNo", myfont, TextDirection.零度);
			string materialNameCN = UnicodeToZPL.UnCompressZPL(materialName, "materialName", myfont, TextDirection.零度);
			//string serialNoCN = UnicodeToZPL.UnCompressZPL(branchNoSerialNo, "branchNoSerialNo", myfont, TextDirection.零度);
			string qrCode = materialNo + separator + branchNo + separator + serialNo;
			sb.Append("^XA");
			sb.Append("^FO250,20,0^BQ,2,5^FDQA," + qrCode + "^FS");
			sb.Append("^FO450,60^A0,26,26^FDQA," + productDrawNoCN + "^FS");
			sb.Append("^FO450,90^A0,26,26^FDQA," + materialNameCN + "^FS");
			sb.Append("^FO450,120^A0,26,26^FD" + branchNoSerialNo + "^FS");
			sb.Append("^XZ");
		}



		/// <summary>
		/// 打印双排的zpl指令，双排:40*20mm。
		/// </summary>
		/// <param name="sb">完整的指令</param>
		/// <param name="materialNo">物料编号</param>
		/// <param name="branchNo">批次号</param>
		/// <param name="serialNo">流水号</param>
		/// <param name="materialName">物料名称</param>
		/// <param name="productDrawNo">产品图号</param>
		private void Print2Col(StringBuilder sb, string materialNo, string branchNo, string serialNo, string separator,string materialName,string productDrawNo)
        {
            Font myfont = new Font("黑体", 16, FontStyle.Bold);
			string branchNoSerialNo = branchNo + "-" + serialNo;//批次号+序号
			string productDrawNoCN = UnicodeToZPL.UnCompressZPL(productDrawNo, "productDrawNo", myfont, TextDirection.零度);
			string materialNameCN = UnicodeToZPL.UnCompressZPL(materialName, "materialName", myfont, TextDirection.零度);
			//string serialNoCN = UnicodeToZPL.UnCompressZPL(branchNoSerialNo, "branchNoSerialNo", myfont, TextDirection.零度);
			string qrCode = materialNo + separator + branchNo + separator + serialNo;
			sb.Append("^XA");
            sb.Append("^FO80,20^BQ,2,5^FDQA," + qrCode + "^FS");
            sb.Append("^FO185,30^A0,26,26^FD," + productDrawNoCN + "^FS");
            sb.Append("^FO185,70^A0,26,26^FD," + materialNameCN + "^FS");
            sb.Append("^FO185,110^A0,26,26^" + branchNoSerialNo + "^FS");
            sb.Append("^FO420,20^BQ,2,5^FDQA," + qrCode + "^FS");
            sb.Append("^FO530,30^A0,26,25^FD," + productDrawNoCN + "^FS");
            sb.Append("^FO530,70^A0,26,25^FD," + materialNameCN + "^FS");
            sb.Append("^FO530,110^A0,26,25^FD" + branchNoSerialNo + "^FS");
            sb.Append("^XZ");
        }

		/// <summary>
		/// 打印三排的zpl指令，尺寸为30*20mm
		/// </summary>
		/// <param name="sb">完整的指令</param>
		/// <param name="materialNo">物料编号</param>
		/// <param name="branchNo">批次号</param>
		/// <param name="serialNo">流水号</param>
		/// <param name="separator">分割符</param>
		/// <param name="materialName">物料名称</param>
		/// <param name="productDrawNo">产品图号</param>
		private void Print3Col(StringBuilder sb, string materialNo, string branchNo, string serialNo, string separator,string materialName,string productDrawNo)
        {
            Font myfont = new Font("黑体", 12, FontStyle.Bold);
			
			string productDrawNoCN = UnicodeToZPL.UnCompressZPL(productDrawNo, "productDrawNo", myfont, TextDirection.零度);
			string materialNameCN = UnicodeToZPL.UnCompressZPL(materialName, "materialName", myfont, TextDirection.零度);
			//string serialNoCN = UnicodeToZPL.UnCompressZPL(branchNoSerialNo, "branchNoSerialNo", myfont, TextDirection.零度);
			string branchNoSerialNo = branchNo + "-" + serialNo;//批次号+序号
			string qrCode = materialNo + separator + branchNo + separator + serialNo;
			sb.Append("^XA");
            sb.Append("^FO20,30^BQ,2,3^FDQA," + qrCode + "^FS");
            sb.Append("^FO95,40^A0,12,12^FD" + productDrawNoCN + "^FS");
            sb.Append("^FO95,70^A0,12,12^FD" + materialNameCN + "^FS");
            sb.Append("^FO95,100^A0,12,12^FD" + branchNoSerialNo + "^FS");
            sb.Append("^FO300,30^BQ,2,3^FDQA," + qrCode + "^FS");
            sb.Append("^FO385,40^A0,12,12^FD" + productDrawNoCN + "^FS");
            sb.Append("^FO385,70^A0,12,12^FD" + materialNameCN + "^FS");
            sb.Append("^FO385,100^A0,12,12^FD" + branchNoSerialNo + "^FS");
            sb.Append("^FO575,30^BQ,2,3^FDQA," + qrCode + "^FS");
            sb.Append("^FO655,40^A0,12,12^FD" + productDrawNoCN + "^FS");
            sb.Append("^FO655,70^A0,12,12^FD" + materialNameCN + "^FS");
            sb.Append("^FO655,100^A0,12,12^FD" + branchNoSerialNo + "^FS");
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