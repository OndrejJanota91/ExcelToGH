using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Special;
using Grasshopper.Kernel.Parameters;
using Grasshopper;
using System.Windows.Forms;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToGH
{
    public class ExcelToGHComponent : GH_Component
    {
        bool Rows = false;
        bool FromOutputUpdate = true;
        bool DeletedOutputs = false;
        bool SetResult = true;
        string WarningMessage = "";
        bool Wires = true;
        Dictionary<string, List<string>> Result = default;
        public ExcelToGHComponent()
          : base("ExcelToGH", "ExcelToGH",
            "Read excel rows/columns",
            "ExcelToGH", "ExcelToGH")
        {
        }
        public override void AppendAdditionalMenuItems(ToolStripDropDown menu)
        {
            base.AppendAdditionalMenuItems(menu);
            var ProgramToggle1 = Menu_AppendItem(menu, "Recompute.", Recompute, true);
            ProgramToggle1.ToolTipText = "Recompute";
            var ProgramToggle2 = Menu_AppendItem(menu, "Read by cols.", Menu_Clicked, true, !this.Rows);
            ProgramToggle2.ToolTipText = "Read by columns";
            var ProgramToggle3 = Menu_AppendItem(menu, "Read by rows", Menu_Clicked, true, this.Rows);
            ProgramToggle3.ToolTipText = "Read by rows";
            var ProgramToggle4 = Menu_AppendItem(menu, "Create wires", Wire, true, this.Wires);
            ProgramToggle4.ToolTipText = "Automatically create output wires";
        }
        protected void UpdateOutput()
        {
            for (int i = 0; i < Result.Count; i++)
            {
                string name = Result.ElementAt(i).Key;

                Param_GenericObject param = new Param_GenericObject();
                param.Name = Params.Input[2].VolatileData.get_Branch(0)[i].ToString();
                param.NickName = param.Name;
                param.Optional = true;
                param.Access = GH_ParamAccess.list;
                
                if (this.Rows)
                    param.Description = "Excel row";
                else
                    param.Description = "Excel column";

                Params.RegisterOutputParam(param);
            }
            this.FromOutputUpdate = true;
            ExpireSolution(true);
        }
        protected void RemoveParametersCallback(GH_Document document)
        {
            this.WarningMessage = "";
            if (Params.Output.Count > 0)
            {
                for (int i = Params.Output.Count - 1; i >= 0; i--)
                {
                    Params.UnregisterOutputParameter(Params.Output[i]);
                }
            }
            Params.OnParametersChanged();
            Instances.RedrawCanvas();
            this.FromOutputUpdate = false;
        }
        protected void UpdateWiresCallBack(GH_Document document)
        {
            foreach (IGH_DocumentObject docObject in document.Objects)
            {
                if (docObject as GH_Component != null)
                {
                    GH_Component comp = docObject as GH_Component;
                    foreach (IGH_Param param in comp.Params.Input)
                    {
                        if (this.Result.ContainsKey(param.NickName))
                        {
                            param.AddSource(this.Params.Output.Find(x => x.Name == param.NickName));
                        }
                    }
                }
                else if (docObject as IGH_Param != null)
                {
                    IGH_Param param = docObject as IGH_Param;
                    if (this.Result.ContainsKey(param.NickName))
                    {
                        param.AddSource(this.Params.Output.Find(x => x.Name == param.NickName));
                    }
                }
                else if (docObject as GH_Panel != null)
                {
                    GH_Panel panel = docObject as GH_Panel;
                    if (this.Result.ContainsKey(panel.NickName))
                    {
                        panel.AddSource(this.Params.Output.Find(x => x.Name == panel.NickName));
                    }
                }
            }
            Instances.RedrawCanvas();
        }
        protected void Recompute(object sender, EventArgs e)
        {
            this.FromOutputUpdate = true;
            this.DeletedOutputs = false;
            this.SetResult = true;
            ExpireSolution(true);
        }
        protected void Wire(object sender, EventArgs e)
        {
            this.Wires = !this.Wires;
        }
        protected void Menu_Clicked(object sender, EventArgs e)
        {
            ToolStripMenuItem senderObj = (sender as ToolStripMenuItem);
            if (senderObj.Text == "Read by cols.")
            {
                Rows = Rows != false ? false : true;
                ExpireSolution(true);
            }
            else if (senderObj.Text == "Read by rows")
            {
                Rows = Rows == false ? true : false;
                ExpireSolution(true);
            }
        }
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Path to excel", "Path", "Path to .xslx file", GH_ParamAccess.item);
            pManager.AddTextParameter("sheet name", "Sheet name", "Name of sheet in excel", GH_ParamAccess.item);
            pManager.AddTextParameter("Column/Row names which are read", "Column/Row names", "Column/Row names which are read from excel file", GH_ParamAccess.list);
        }
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            this.SetResult = true;
            if (!DeletedOutputs)
            {
                GH_Document thisDocument = this.OnPingDocument();
                this.DeletedOutputs = true;
                this.SetResult = false;
                thisDocument.ScheduleSolution(5, RemoveParametersCallback);
            }

            if (!FromOutputUpdate)
            {
                string ExcelPath = default;
                List<string> ColumnsNumbers = new List<string>();
                string SheetName = "";

                if (!DA.GetData(0, ref ExcelPath)) return;
                if (!DA.GetData(1, ref SheetName)) return;
                if (!DA.GetDataList(2, ColumnsNumbers)) return;

                Message = "Reading excel file.";
                Instances.RedrawCanvas();
                this.Result = new Dictionary<string, List<string>>();
                if (!Rows)
                {
                    Result = ReadXLSColumn(ExcelPath, SheetName, ColumnsNumbers);
                }
                else
                {
                    Result = ReadXLSRows(ExcelPath, SheetName, ColumnsNumbers);
                }
                UpdateOutput();
                this.SetResult = false;
            }
            if (SetResult)
            {
                this.FromOutputUpdate = false;
                if (Params.Output.Count > 0)
                {
                    foreach (IGH_Param Param in Params.Output)
                    {

                        if (Result.ContainsKey(Param.Name))
                        {
                            DA.SetDataList(Param.Name, Result[Param.Name]);
                        }
                    }
                }
                Message = "Component recomputed.";
                this.DeletedOutputs = false;
                if (this.WarningMessage != "")
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, this.WarningMessage);
                }
                Instances.RedrawCanvas();
                if (Wires)
                {
                    GH_Document thisDocument = this.OnPingDocument();
                    thisDocument.ScheduleSolution(5, UpdateWiresCallBack);
                }

            }
            else
            {
                this.SetResult = true;
                ExpireSolution(true);
            }
        }
        private Dictionary<string, List<string>> ReadXLSRows(string Path, string SheetName, List<string> RNames)
        {
            Excel.Application xlApp = default;
            Excel.Workbook xlWorkBook = default;
            Excel.Worksheet xlWorkSheet = default;
            try
            {
                Dictionary<string, int> RowNames = new Dictionary<string, int>();
                Dictionary<string, List<string>> Result = new Dictionary<string, List<string>>();


                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(Path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


                foreach (Excel.Worksheet xlSheet in xlWorkBook.Worksheets)
                {
                    if (xlSheet.Name == SheetName)
                    {
                        xlWorkSheet = xlSheet;
                        break;
                    }

                }
                if (xlWorkSheet != default)
                {
                    Excel.Range xlRange = xlWorkSheet.UsedRange;
                    xlRange.ClearFormats();
                    var row = 1;
                    while (row <= xlRange.Rows.Count)
                    {
                        if ((xlRange[row, 1] as Excel.Range).Value2 != null)
                        {
                            if (!RowNames.ContainsKey((xlRange[row, 1] as Excel.Range).Value2.ToString()))
                            {
                                RowNames.Add((xlRange[row, 1] as Excel.Range).Value2.ToString(), row);
                            }
                            else
                            {
                                xlApp.DisplayAlerts = false;
                                xlWorkBook.Close();
                                xlApp.Quit();
                                this.WarningMessage = "Rows in column A contains duplicate name.";
                                return new Dictionary<string, List<string>>();
                            }

                        }
                        row++;
                    }

                    foreach (string Name in RNames)
                    {
                        var column = 2;
                        if (RowNames.ContainsKey(Name))
                        {
                            List<string> Values = new List<string>();
                            while (column <= (xlRange.Rows[RowNames[Name], Type.Missing] as Excel.Range).Columns.Count)
                            {
                                if ((xlWorkSheet.Cells[RowNames[Name], column] as Excel.Range).Value2 != null)
                                {
                                    Values.Add((xlWorkSheet.Cells[RowNames[Name], column] as Excel.Range).Value2.ToString());
                                    column++;
                                }
                                else
                                {
                                    break;
                                }


                            }
                            Result.Add(Name, Values);
                        }
                        else
                        {
                            Result.Add(Name, new List<string>());
                        }
                    }

                    xlApp.DisplayAlerts = false;
                    xlWorkBook.Save();
                    xlWorkBook.Close();
                    xlApp.Quit();
                    return Result;
                }
                else
                {
                    xlApp.DisplayAlerts = false;
                    xlWorkBook.Close();
                    xlApp.Quit();
                    this.WarningMessage = "Sheet name was not found in excel file.";
                    return new Dictionary<string, List<string>>();
                }

            }
            catch
            {
                xlApp.DisplayAlerts = false;
                xlWorkBook.Close();
                xlApp.Quit();
                this.WarningMessage = "Something went wrong.";
                return new Dictionary<string, List<string>>();
            }


        }
        private Dictionary<string, List<string>> ReadXLSColumn(string Path, string SheetName, List<string> CNames)
        {
            Excel.Application xlApp = default;
            Excel.Workbook xlWorkBook = default;
            Excel.Worksheet xlWorkSheet = default;
            try
            {
                Dictionary<string, int> ColumnNames = new Dictionary<string, int>();
                Dictionary<string, List<string>> Result = new Dictionary<string, List<string>>();



                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(Path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                foreach (Excel.Worksheet xlSheet in xlWorkBook.Worksheets)
                {
                    if (xlSheet.Name == SheetName)
                    {
                        xlWorkSheet = xlSheet;
                        break;
                    }

                }
                if (xlWorkSheet != default)
                {
                    Excel.Range xlRange = xlWorkSheet.UsedRange;
                    xlRange.ClearFormats();
                    var column = 1;
                    while (column <= xlRange.Columns.Count)
                    {
                        if ((xlRange[1, column] as Excel.Range).Value2 != null)
                        {
                            if (!ColumnNames.ContainsKey((xlRange[1, column] as Excel.Range).Value2.ToString()))
                            {

                                ColumnNames.Add((xlRange[1, column] as Excel.Range).Value2.ToString(), column);
                            }
                            else
                            {
                                xlApp.DisplayAlerts = false;
                                xlWorkBook.Close();
                                xlApp.Quit();
                                this.WarningMessage = "Column in row 1 contains duplicate names";
                                return new Dictionary<string, List<string>>();
                            }

                        }
                        column++;
                    }

                    foreach (string Name in CNames)
                    {
                        var row = 2;
                        if (ColumnNames.ContainsKey(Name))
                        {
                            List<string> Values = new List<string>();
                            while (row <= (xlRange.Columns[ColumnNames[Name], Type.Missing] as Excel.Range).Rows.Count)
                            {
                                if ((xlWorkSheet.Cells[row, ColumnNames[Name]] as Excel.Range).Value2 != null)
                                {
                                    Values.Add((xlWorkSheet.Cells[row, ColumnNames[Name]] as Excel.Range).Value2.ToString());
                                    row++;
                                }
                                else
                                {
                                    break;
                                }


                            }
                            Result.Add(Name, Values);
                        }
                        else
                        {
                            Result.Add(Name, new List<string>());
                        }
                    }

                    xlApp.DisplayAlerts = false;
                    xlWorkBook.Save();
                    xlWorkBook.Close();
                    xlApp.Quit();
                    return Result;
                }
                else
                {
                    xlApp.DisplayAlerts = false;
                    xlWorkBook.Close();
                    xlApp.Quit();
                    this.WarningMessage = "Sheet name was not found in excel file.";
                    return new Dictionary<string, List<string>>();
                }
            }
            catch
            {
                xlApp.DisplayAlerts = false;
                xlWorkBook.Close();
                xlApp.Quit();
                this.WarningMessage = "Something went wrong.";
                return new Dictionary<string, List<string>>();
            }
        }
        #region VARIABLE COMPONENT INTERFACE IMPLEMENTATION
        public bool CanInsertParameter(GH_ParameterSide side, int index)
        {
            // Only insert parameters on input side. This can be changed if you like/need
            // side== GH_ParameterSide.Output
            if (side == GH_ParameterSide.Output)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool CanRemoveParameter(GH_ParameterSide side, int index)
        {
            // Only allowed to remove parameters if there are more than 1
            // from the input side
            if (side == GH_ParameterSide.Output && Params.Output.Count > 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public IGH_Param CreateParameter(GH_ParameterSide side, int index)
        {
            // Has to return a parameter object!
            Param_GenericObject param = new Param_GenericObject();

            int count = 0;
            for (int i = 0; i < Params.Output.Count; i++)
            {
                count += 1;
            }
            if (!this.Rows)
            {
                param.Name = "Column" + count.ToString();
                param.NickName = param.Name;
                param.Description = "Excel column";
                param.Optional = true;
                param.Access = GH_ParamAccess.list;
            }
            else
            {
                param.Name = "Row" + count.ToString();
                param.NickName = param.Name;
                param.Description = "Excel row";
                param.Optional = true;
                param.Access = GH_ParamAccess.list;
            }
            return param;
        }
        public bool DestroyParameter(GH_ParameterSide side, int index)
        {
            //This function will be called when a parameter is about to be removed. 
            //You do not need to do anything, but this would be a good time to remove 
            //any event handlers that might be attached to the parameter in question.
            return true;
        }
        public void VariableParameterMaintenance()
        {
            //This method will be called when a closely related set of variable parameter operations completes. 
            //This would be a good time to ensure all Nicknames and parameter properties are correct. This method will also be 
            //called upon IO operations such as Open, Paste, Undo and Redo.
        }
        #endregion

        protected override System.Drawing.Bitmap Icon => Properties.Resources.ExcelToGH;
        public override Guid ComponentGuid => new Guid("BB7F9E80-9BBA-427E-A862-1F034BCA6305");
    }
}

