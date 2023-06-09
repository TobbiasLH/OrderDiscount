using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using BIG;
using System.IO;
using System.Configuration;

namespace OrderDiscount
{
    public class Events
    {
        internal BIG.Application bigApp;

        public static Configuration appConfig = ConfigurationManager.OpenExeConfiguration(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OrderDiscount.dll"));

        public static string discountProductNo = appConfig.AppSettings.Settings["DiscountProductNo"].Value;

        public static int discountUnit = Convert.ToInt32(appConfig.AppSettings.Settings["DiscountUnit"].Value);

        public static string connectionString = appConfig.ConnectionStrings.ConnectionStrings["vbConnection"].ConnectionString;

        public Events(BIG.Application bigApplication)
        {
            this.bigApp = bigApplication;

            this.bigApp.SessionLoggedOn += SessionLoggedOn;

            this.bigApp.ButtonClick += ButtonClick;
        }

        private void SessionLoggedOn()
        {
            // Logged on user
            BIG.Ts_ReadOnlyRowStringValue ts;
            string username = bigApp.User.GetStringValue((int)C_User.C_User_UserName, out ts);

            DebugLog("");
            DebugLog("   ---   NEW LOG   ---");
            DebugLog("");
            DebugLog("Logged on! - " + DateTime.Now.ToString() + " - " + username);
        }

        private void ButtonClick(BIG.Button button, ref bool SkipRecording)
        {
            try
            {
                if (button.Caption.ToLower() == "rabatt")
                {
                    int orderNo = 0;

                    // string client = bigApp.CurrentFirmNo.ToString("0000");

                    BIG.PageElement pageElement = button.PageElement;

                    BIG.Document bigDocument = pageElement.Document;

                    foreach (BIG.Table table in bigDocument.Tables)
                    {
                        //string title = table.Title.ToLower();

                        //DebugLog("Table: " + table.Title);

                        if (table.TableNo == (int)BIG.TableNo.TableNo_Order)  // Table NO = 127
                        {
                            Ts_RowGetValue ts_RowGetValue;

                            orderNo = table.ActiveRow.GetIntegerValue((int)C_Order.C_Order_OrderNo, out ts_RowGetValue);

                            if (orderNo > 0)
                            {
                                bigDocument.Save();  

                                double discount = GetDiscount(orderNo, discountProductNo);

                                if (discount > 0)
                                {
                                    bool result = AddOrderLineDiscountRow(bigDocument, table, orderNo, discountProductNo, discount * -1);
                                }

                                break;

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private bool AddOrderLineDiscountRow(Document bigDocument, Table orderTable, int orderNo, string productNo, double discount)
        {
            Ts_TableJoin ts_TableJoin;

            Ts_TableGetEmptyRow ts_TableGetEmptyRow;

            string message = "";

            // Add orderline table (OrdLn)
            var newOrderLineTable = orderTable.Join((int)Fk_OrderLine.Fk_OrderLine_Order, out ts_TableJoin);

            // Get an empty orderline
            var newOrderLineRow = newOrderLineTable.GetEmptyRow(out ts_TableGetEmptyRow);

            // Set values for new orderline

            // Let BIG suggest a value for new linenumber
            newOrderLineRow.SuggestValue((int)C_OrderLine.C_OrderLine_LineNo, out message);

            // Set other orderline values
            newOrderLineRow.SetStringValue((int)C_OrderLine.C_OrderLine_ProductNo, productNo, out message);
            newOrderLineRow.SetDecimalValue((int)C_OrderLine.C_OrderLine_Quantity, 1, out message);
            newOrderLineRow.SetDecimalValue((int)C_OrderLine.C_OrderLine_Price, discount, out message);


            bigDocument.Save();
            bigDocument.Refresh();

            return true;
        }

        public static void DebugLog(string message)
        {
            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter("DiscountDebugLog.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }

        public static double GetDiscount(int orderNo, string productNo)
        {
            double discount = 0;

            string commandText = @"
                -- Delete existing doscount lines
                delete from OrdLn where OrdNo = @OrderNo and ProdNo = @ProductNo;

                -- Find total kg of products    
                declare @TotalKg decimal = (select sum(NoInvoAb) from OrdLn where OrdNo = @OrderNo and Un = @DiscountUnit) 

                -- Find which month the order has
                declare @Month int = (select month(convert(datetime, convert(varchar(10), OrdDt))) from Ord where OrdNo = @OrderNo)

                select
                    case when @Month != 6  -- if month not equal to 6 (june)
	                then
	                    case when @TotalKg between 3000 and 6999 then @TotalKg * 0.10
		                     when @TotalKg between 7000 and 19999 then @TotalKg * 0.35
		                     when @TotalKg between 20000 and 49999 then @TotalKg * 0.55
		                     when @TotalKg between 50000 and 74999 then @TotalKg * 0.75
		                     when @TotalKg between 75000 and 99999 then @TotalKg * 0.80
		                     when @TotalKg between 100000 and 149999 then @TotalKg * 1
		                     when @TotalKg between 150000 and 199999 then @TotalKg * 1.05
		                     when @TotalKg > 200000 then @TotalKg * 1.10
		                     else 0
	                     end
	                else 0
	                end as TotalDiscount
                ";

            SqlCommand sqlCommand = new SqlCommand();

            sqlCommand.CommandText = commandText;
            sqlCommand.Parameters.AddWithValue("OrderNo", orderNo);
            sqlCommand.Parameters.AddWithValue("ProductNo", productNo);
            sqlCommand.Parameters.AddWithValue("DiscountUnit", discountUnit);

        DataTable dataTable = SelectData(sqlCommand);

            if (dataTable.Rows.Count > 0)
            {
                discount = Convert.ToDouble(dataTable.Rows[0]["TotalDiscount"]);
            }

            return discount;
        }

        static DataTable SelectData(SqlCommand sqlCommand)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                sqlConnection.Open();

                using (sqlCommand)
                {
                    sqlCommand.Connection = sqlConnection;

                    dataTable.Load(sqlCommand.ExecuteReader());
                }
            }

            return dataTable;
        }
    }
}
