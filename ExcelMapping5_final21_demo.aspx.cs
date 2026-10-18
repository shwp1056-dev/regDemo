using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.Globalization;
using System.Threading;
using System.Data.OleDb;
using ClosedXML.Excel;
using System.Net;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Web.UI.DataVisualization.Charting;
using System.ServiceModel.PeerResolvers;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

public class Itemhead
{
    public string Name { get; set; }
}

public partial class ExcelMapping5_final21 : System.Web.UI.Page
{
    cls_Common_cls cls_Common_cls = new cls_Common_cls();
    ConnectDB objdb = new ConnectDB();
    SqlConnection con = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString);

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Bind_ddlAct();
            Bind_Ddlcompany();

            btnSubmit.Visible = false;
        }

    }

    public void Bind_ddlAct()
    {
        using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
        {

            if (cnn.State == ConnectionState.Open)
                cnn.Close();
            cnn.Open();
            cls_Common_cls.F_BIND_DROPDOWN("select PK, [Act] from Act_master", DdlAct, cnn);

            cnn.Close();
        }
    }

    public void Bind_Ddlcompany()
    {
        using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
        {
            if (cnn.State == ConnectionState.Open)
                cnn.Close();
            cnn.Open();

            cls_Common_cls.F_BIND_DROPDOWN("select distinct cm.pk as cmpid,CompanyName as cmpname from [auditdb].[dbo].[tbl_CompanyMaster] cm left join [auditdb].[dbo].[tbl_CompanyBranchModuleMapping] cmm on cm.pk = cmm.Companyid left join [auditdb].[dbo].tbl_UserAccessMaster uam on uam.CompBranchModuleId = cmm.pk  where  uam.Userid = '" + Request.Cookies["UserInfo"]["id"].ToString() + "' and uam.Valid = 1", Ddlcompany, cnn);

            cnn.Close();
        }
    }

    private bool Validation()
    {
        if (!FileUpload1.HasFile)
        {
            cls_Common_cls.DisplayAJAXMessage(this, "Kindly Upload file before submit");
            return false;
        }
        else
        {
            HttpPostedFile files = FileUpload1.PostedFile;

            Thread.Sleep(10);
            if (files.ContentLength > 5000000)
            {
                cls_Common_cls.DisplayAJAXMessage(this, "File " + files.FileName + " cannot be upladed as the file size is greater than 5MB. kindly upload file less than 5MB");
                return false;
            }
        }

        return true;
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        try
        {
            string tablename = "";

            if (DdlAct.SelectedValue == "2")
            {
                tablename = "Maharashtra_shopestblishment";
            }
            else if (DdlAct.SelectedValue == "3")
            {
                tablename = "MHfactories_act";
            }
            else if (DdlAct.SelectedValue == "4")
            {
                tablename = "MHMaternity_act";
            }
            else if (DdlAct.SelectedValue == "5")
            {
                tablename = "MH_POG";
            }
            else if (DdlAct.SelectedValue == "6")
            {
                tablename = "MH_CLRA";
            }
            else if (DdlAct.SelectedValue == "7")
            {
                tablename = "tbl_MHBOCW";
            }
            else if (DdlAct.SelectedValue == "8")
            {
                tablename = "MH_POB";
            }
            else if (DdlAct.SelectedValue == "1002")
            {
                tablename = "MHfactories_act_otherDetails";
            }
            else if (DdlAct.SelectedValue == "1003")
            {
                tablename = "tblLeaveDetails";
            }
            else if (DdlAct.SelectedValue == "1004")
            {
                tablename = "tblFnF_Data";
            }
            else if (DdlAct.SelectedValue == "2004")
            {
                tablename = "MH_Sales_PromotionAct";
            }
            else if (DdlAct.SelectedValue == "1006")
            {
                tablename = "tblLeaveDataPeperfry";
            }
            foreach (ListViewItem hw in lstdc.Items)
            {
                string strexcel = ((Label)hw.FindControl("lblexcelcolumns")).Text;
                string strdb = ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text;
                string squencenumber = ((TextBox)hw.FindControl("txtSequenceNo")).Text;
                CheckBox ischeck = (CheckBox)hw.FindControl("chkMapping");
                bool ischecked = ischeck.Checked;
                //if(strexcel==)


                if (strdb != "Please Select")
                {

                    using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
                    {

                        if (cnn.State == ConnectionState.Open)
                            cnn.Close();
                        cnn.Open();
                        //string query1 = "select  [ExcelColumn],DBColumn from [ExcelDB_Mapping] where Companyname='" + Ddlcompany.SelectedValue + "' and Act='" + DdlAct.SelectedValue + "' and [ExcelColumn]='" + strexcel + "'  ";
                        //using (SqlDataAdapter sda = new SqlDataAdapter(query1, cnn))
                        //{
                        //    DataSet ds = new DataSet();
                        //    sda.Fill(ds);

                        //    string exl = ds.Tables[0].Rows[0]["ExcelColumn"].ToString();
                        //    string db1= ds.Tables[0].Rows[0]["DBColumn"].ToString();
                        //    if (ds.Tables[0].Rows.Count > 0)
                        //    { 
                        //        if(strexcel == exl)
                        //        {
                        //            //  cls_Common_cls.F_BIND_DROPDOWN1( "Already mapped!", con);
                        //            ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text = db1;
                        //        }
                        //        else
                        //        {

                        //        }
                        //    }
                        //}


                        //----------
                        string query = "select [ExcelColumn], [DBColumn] from [ExcelDB_Mapping] where Companyname='" + Ddlcompany.SelectedValue + "' and Act='" + DdlAct.SelectedValue + "'  and [DBColumn]='" + strdb + "' ";
                        using (SqlDataAdapter sda = new SqlDataAdapter(query, cnn))
                        {
                            DataSet ds = new DataSet();
                            sda.Fill(ds); 
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                //cls_Common_cls.DisplayAJAXMessage(this, " Already mapped !");
                                //    // ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text = ds.Tables[0].Rows[0]["DBColumn"].ToString();
                                //   // DropDownList ddlDBCOLUMNS = (e.Item.FindControl("ddlDbColumn") as DropDownList);
                                //  //  cls_Common_cls.F_BIND_DROPDOWN1(Queries, ddlDBCOLUMNS, con);

                                int affect = objdb.affect("update ExcelDB_Mapping set DBColumn='" + strdb + "',ExcelColumn='" + strexcel + "' , [SequenceNumber]='" + squencenumber + "' , [Checked]='" + ischecked + "' where [DBColumn]='" + strdb + "'  and Companyname='" + Ddlcompany.SelectedValue + "' ");
                            }
                            else
                            {
                                int affect = objdb.affect("INSERT INTO [dbo].[ExcelDB_Mapping] ([ExcelColumn] ,[DBColumn] ,[Act] ,[tablename] ,[Companyname],[SequenceNumber],[Checked]) VALUES ('" + strexcel + "' , '" + strdb + "' ,'" + DdlAct.SelectedValue + "' ,'" + tablename + "', '" + Ddlcompany.SelectedValue + "' ,'" + squencenumber + "','" + ischecked + "')");

                            }


                        }
                        bindgridview();

                        cnn.Close();
                    }

                    bindgridview();

                }
            }

            cls_Common_cls.DisplayAJAXMessage(this, "Data saved.");
        }
        catch (FormatException)
        {
            cls_Common_cls.DisplayAJAXMessage(this, "Error while saving, Kindly check all the numeric data.");
        }
        catch (Exception)
        {
            cls_Common_cls.DisplayAJAXMessage(this, "Error while uploading file, Kindly check and upload it again.");
        }

        btnSubmit.Enabled = true;
    }

    public MemoryStream GetStream(XLWorkbook excelWorkbook)
    {
        MemoryStream fs = new MemoryStream();
        excelWorkbook.SaveAs(fs);
        fs.Position = 0;
        return fs;
    }

    protected void lstdc_ItemDataBound(object sender, ListViewItemEventArgs e)
    {
        String Queries = "";

        if (DdlAct.SelectedValue == "2")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'Maharashtra_shopestblishment' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "3")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'MHfactories_act' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "4")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'MHMaternity_act' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "5")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'MH_POG' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "6")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'MH_CLRA' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "7")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'tbl_MHBOCW' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "8")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'MH_POB' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "1002")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'MHfactories_act_otherDetails' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        else if (DdlAct.SelectedValue == "1003")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'tblLeaveDetails' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        if (DdlAct.SelectedValue == "1004")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'tblFnF_Data' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        if (DdlAct.SelectedValue == "2004")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'MH_Sales_PromotionAct' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        if (DdlAct.SelectedValue == "1006")
        {
            Queries = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'tblLeaveDataPeperfry' and COLUMN_NAME not in ('PK', 'Month', 'Year', 'Company_name', 'Act') order by COLUMN_NAME";
        }
        DropDownList ddlDBCOLUMNS = (e.Item.FindControl("ddlDbColumn") as DropDownList);
        cls_Common_cls.F_BIND_DROPDOWN1(Queries, ddlDBCOLUMNS, con);


        foreach (ListViewItem hw in lstdc.Items)
        {
            string strexcel = ((Label)hw.FindControl("lblexcelcolumns")).Text;
            string strdb = ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text;

            //if(strexcel==)


            if (strdb != "Please Select")
            {

                using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
                {

                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                    cnn.Open();
                    string query1 = "select [ExcelColumn],[DBColumn] ,[SequenceNumber] from [ExcelDB_Mapping] where Companyname='" + Ddlcompany.SelectedValue + "' and Act='" + DdlAct.SelectedValue + "' and [ExcelColumn]='" + strexcel + "'  ";
                    using (SqlDataAdapter sda = new SqlDataAdapter(query1, cnn))
                    {
                        DataSet ds = new DataSet();
                        sda.Fill(ds);

                        string exl = ds.Tables[0].Rows[0]["ExcelColumn"].ToString();
                        string db1 = ds.Tables[0].Rows[0]["DBColumn"].ToString();
                        string seq = ds.Tables[0].Rows[0]["SequenceNumber"].ToString();
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            if (strexcel == exl)
                            {
                                //  cls_Common_cls.F_BIND_DROPDOWN1( "Already mapped!", con);
                                ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text = db1;
                                ((TextBox)hw.FindControl("txtSequenceNo")).Text = seq;
                            }
                            else
                            {
                                cls_Common_cls.F_BIND_DROPDOWN1(Queries, ddlDBCOLUMNS, con);

                            }

                        }
                        else
                        {
                            cls_Common_cls.F_BIND_DROPDOWN1(Queries, ddlDBCOLUMNS, con);

                        }
                    }

                }





                btnSubmit.Visible = true;

            }

        }
    }
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        bindgridview();
        btnDeleteMAPPING.Visible = true;

        if (Validation())
        {
            try
            {
                string filePath = Server.MapPath("~/Files/") + Path.GetFileName(FileUpload1.PostedFile.FileName);
                FileUpload1.SaveAs(filePath);

                using (XLWorkbook workBook = new XLWorkbook(filePath))
                {
                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    DataTable dt = new DataTable();

                    List<Itemhead> rowvalue = new List<Itemhead>();

                    foreach (IXLCell cell in workSheet.Row(1).Cells())
                    {
                        Itemhead row = new Itemhead();
                        row.Name = cell.Value.ToString();
                        rowvalue.Add(row);
                    }

                    lstdc.DataSource = rowvalue;
                    lstdc.DataBind();
                }
            }
            catch (FormatException)
            {
                cls_Common_cls.DisplayAJAXMessage(this, "Error while saving, Kindly check all the numeric data.");
            }
            catch (Exception)
            {
                cls_Common_cls.DisplayAJAXMessage(this, "Error while uploading file, Kindly check and upload it again.");
            }
        }


        // ==================================
        foreach (ListViewItem hw in lstdc.Items)
        {
            string strexcel = ((Label)hw.FindControl("lblexcelcolumns")).Text;
            string strdb = ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text;

            //if(strexcel==)


            //if (strdb != "Please Select")
            //{

            using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
            {

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                cnn.Open();
                string query1 = "select  [ExcelColumn],[DBColumn],[SequenceNumber] ,Checked from [ExcelDB_Mapping] where Companyname='" + Ddlcompany.SelectedValue + "' and Act='" + DdlAct.SelectedValue + "' and [ExcelColumn]='" + strexcel + "'   ";
                using (SqlDataAdapter sda = new SqlDataAdapter(query1, cnn))
                {
                    DataSet ds = new DataSet();
                    sda.Fill(ds);

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string exl = ds.Tables[0].Rows[0]["ExcelColumn"].ToString();
                        string db1 = ds.Tables[0].Rows[0]["DBColumn"].ToString();

                        string seq = ds.Tables[0].Rows[0]["SequenceNumber"].ToString();
                        string chk = ds.Tables[0].Rows[0]["Checked"].ToString();

                        if (strexcel == exl)
                        {
                            //  cls_Common_cls.F_BIND_DROPDOWN1( "Already mapped!", con);
                            ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text = db1;
                            ((TextBox)hw.FindControl("txtSequenceNo")).Text = seq;
                            if (chk != null)
                            {
                                if (chk == "True")
                                    ((CheckBox)hw.FindControl("chkMapping")).Checked = true;
                                else
                                    ((CheckBox)hw.FindControl("chkMapping")).Checked = false;

                            }
                        }
                        else
                        {
                            //cls_Common_cls.F_BIND_DROPDOWN1(Queries, ddlDBCOLUMNS, con);

                        }

                    }
                    else
                    {
                        // cls_Common_cls.F_BIND_DROPDOWN1(Queries, ddlDBCOLUMNS, con);

                    }
                }

            }





            btnSubmit.Visible = true;

            // }

        }
        //  =========================== =========




    }

    protected void ddlDbColumn_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    public void bindgridview()
    {

        using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
        {

            if (cnn.State == ConnectionState.Open)
                cnn.Close();
            cnn.Open();
            string query = "select pk,[ExcelColumn],[DBColumn] ,[SequenceNumber],[Checked] from [ExcelDB_Mapping] where Companyname='" + Ddlcompany.SelectedValue + "' and Act='" + DdlAct.SelectedValue + "'";
            using (SqlDataAdapter sda = new SqlDataAdapter(query, cnn))
            {   
                DataSet ds = new DataSet();
                sda.Fill(ds);
                grdview.DataSource = ds;
                grdview.DataBind();       

            }  
            cnn.Close();


        }

    }

    protected void grdview_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        if (e.CommandName == "DeleteRow")
        {
            using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
            {

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                cnn.Open();
                int index = Convert.ToInt32(e.CommandArgument);
                string query = "delete from ExcelDB_Mapping where pk='" + index + "'";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    //DataTable dt = new DataTable();
                    ////dt.Rows.RemoveAt(index);  
                    // dt.Rows[index].Delete();

                    //grdview.DataSource = dt;
                    //grdview.DataBind();

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();

                    cmd.Clone();
                    bindgridview();

                }

            }
        }
    }



    protected void btnUpdate_Click(object sender, EventArgs e)
    {
        try
        {
            string tablename = "";

            if (DdlAct.SelectedValue == "2")
            {
                tablename = "Maharashtra_shopestblishment";
            }
            else if (DdlAct.SelectedValue == "3")
            {
                tablename = "MHfactories_act";
            }
            else if (DdlAct.SelectedValue == "4")
            {
                tablename = "MHMaternity_act";
            }
            else if (DdlAct.SelectedValue == "5")
            {
                tablename = "MH_POG";
            }
            else if (DdlAct.SelectedValue == "6")
            {
                tablename = "MH_CLRA";
            }
            else if (DdlAct.SelectedValue == "7")
            {
                tablename = "tbl_MHBOCW";
            }
            else if (DdlAct.SelectedValue == "8")
            {
                tablename = "MH_POB";
            }

            foreach (ListViewItem hw in lstdc.Items)
            {
                string strexcel = ((Label)hw.FindControl("lblexcelcolumns")).Text;
                string strdb = ((DropDownList)hw.FindControl("ddlDbColumn")).SelectedItem.Text;
                string squencenumber = ((TextBox)hw.FindControl("txtSequenceNo")).Text;
                CheckBox ischeck = (CheckBox)hw.FindControl("chkMapping");
                bool ischecked = ischeck.Checked;
                //if(strexcel==)


                if (strdb != "Please Select")
                {

                    using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
                    {

                        if (cnn.State == ConnectionState.Open)
                            cnn.Close();
                        cnn.Open();
                        string query = "select [ExcelColumn], [DBColumn] from [ExcelDB_Mapping] where Companyname='" + Ddlcompany.SelectedValue + "' and Act='" + DdlAct.SelectedValue + "'  and  ExcelColumn='" + strexcel + "'  ";
                        using (SqlDataAdapter sda = new SqlDataAdapter(query, cnn))
                        {
                            DataSet ds = new DataSet();
                            sda.Fill(ds);
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                //cls_Common_cls.DisplayAJAXMessage(this, " Already mapped ! "  );
                                int affect = objdb.affect("update ExcelDB_Mapping set DBColumn='" + strdb + "', [SequenceNumber]='" + squencenumber + "' , [Checked]='" + ischecked + "' where ExcelColumn='" + strexcel + "' ");
                            }
                            else
                            {
                                int affect = objdb.affect("INSERT INTO [dbo].[ExcelDB_Mapping] ([ExcelColumn] ,[DBColumn] ,[Act] ,[tablename] ,[Companyname],SequenceNumber,Checked) VALUES ('" + strexcel + "' , '" + strdb + "' ,'" + DdlAct.SelectedValue + "' ,'" + tablename + "', '" + Ddlcompany.SelectedValue + "','" + squencenumber + "','" + ischecked + "')");
                                cls_Common_cls.DisplayAJAXMessage(this, "Data saved");
                            }


                        }
                        bindgridview();

                        cnn.Close();
                    }

                    bindgridview();

                }
            }

            cls_Common_cls.DisplayAJAXMessage(this, "Data Saved.");
        }
        catch (FormatException)
        {
            cls_Common_cls.DisplayAJAXMessage(this, "Error while saving, Kindly check all the numeric data.");
        }
        catch (Exception)
        {
            cls_Common_cls.DisplayAJAXMessage(this, "Error while uploading file, Kindly check and upload it again.");
        }
    }



    protected void grdview_ItemCommand(object sender, ListViewCommandEventArgs e)
    {
        if (e.CommandName == "DeleteRow")
        {
            using (SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
            {

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                cnn.Open();
                int index = Convert.ToInt32(e.CommandArgument);
                string query = "delete from ExcelDB_Mapping where pk='" + index + "'";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    //DataTable dt = new DataTable();
                    ////dt.Rows.RemoveAt(index);  
                    // dt.Rows[index].Delete();

                    //grdview.DataSource = dt;
                    //grdview.DataBind();

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();

                    cmd.Clone();
                    bindgridview();

                }

            }
        }
    }

    protected void Unnamed_Click(object sender, EventArgs e)
    {

    }

    protected void btnDeleteMAPPING_Click(object sender, EventArgs e)
    {
        using (SqlConnection conn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ESSConnectionString"].ConnectionString))
        {
            conn.Open();

            foreach (ListViewItem item in grdview.Items)
            {
                CheckBox chk = (CheckBox)item.FindControl("chkSelect");
                Label lblId = (Label)item.FindControl("lblexcelcolumns");

                if (chk != null && chk.Checked && lblId != null)
                {
                    string id = lblId.Text;

                    using (SqlCommand cmd = new SqlCommand("DELETE FROM ExcelDB_Mapping WHERE pk = @pk", conn))
                    {
                        cmd.Parameters.AddWithValue("@pk", id);
                        cmd.ExecuteNonQuery();
                        cls_Common_cls.DisplayAJAXMessage(this, "Deleted ");

                    }
                }
            }

            conn.Close();
            bindgridview();
        }
    }
}
