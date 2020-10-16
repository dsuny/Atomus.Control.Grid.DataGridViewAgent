using Atomus.Diagnostics;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Atomus.Control.Grid
{
    public class DataGridViewAgent : IDataGridAgent
    {
        string[] menuListText;
        System.Windows.Forms.Control IDataGridAgent.GridControl { get; set; }
        Dictionary<TextBox, FilterAttribute> FilterControl { get; set; }
        int headerRowCount;
        int headerHeight;

        EditAble IDataGridAgent.EditAble
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).ReadOnly) ? EditAble.False : EditAble.True;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).ReadOnly = (value == EditAble.True) ? false : true;
            }
        }
        AddRows IDataGridAgent.AddRows
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).AllowUserToAddRows) ? AddRows.True : AddRows.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).AllowUserToAddRows = (value == AddRows.True) ? true : false;
            }
        }
        DeleteRows IDataGridAgent.DeleteRows
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).AllowUserToDeleteRows) ? DeleteRows.True : DeleteRows.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).AllowUserToDeleteRows = (value == DeleteRows.True) ? true : false;
            }
        }
        ResizeRows IDataGridAgent.ResizeRows
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).AllowUserToResizeRows) ? ResizeRows.True : ResizeRows.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).AllowUserToResizeRows = (value == ResizeRows.True) ? true : false;
            }
        }
        AutoSizeColumns IDataGridAgent.AutoSizeColumns
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).AutoSizeColumnsMode == DataGridViewAutoSizeColumnsMode.Fill) ? AutoSizeColumns.True : AutoSizeColumns.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).AutoSizeColumnsMode = (value == AutoSizeColumns.True) ? DataGridViewAutoSizeColumnsMode.Fill : DataGridViewAutoSizeColumnsMode.None;
            }
        }
        AutoSizeRows IDataGridAgent.AutoSizeRows
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).AutoSizeRowsMode == DataGridViewAutoSizeRowsMode.AllCells) ? AutoSizeRows.True : AutoSizeRows.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).AutoSizeRowsMode = (value == AutoSizeRows.True) ? DataGridViewAutoSizeRowsMode.AllCells : DataGridViewAutoSizeRowsMode.None;
            }
        }
        ColumnsHeadersVisible IDataGridAgent.ColumnsHeadersVisible
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).ColumnHeadersVisible) ? ColumnsHeadersVisible.True : ColumnsHeadersVisible.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).ColumnHeadersVisible = (value == ColumnsHeadersVisible.True) ? true : false;
            }
        }
        EnableMenu IDataGridAgent.EnableMenu {
            get
            {

                if (((IDataGridAgent)this).GridControl.ContextMenu != null)
                    return EnableMenu.False;
                else
                    return EnableMenu.False;
            }
            set
            {
                if (value == EnableMenu.True && ((IDataGridAgent)this).GridControl.ContextMenu == null)
                    SetContextMenuStrip((DataGridView)((IDataGridAgent)this).GridControl, this.menuListText);

            }
        }
        MultiSelect IDataGridAgent.MultiSelect
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).MultiSelect) ? MultiSelect.True : MultiSelect.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).MultiSelect = (value == MultiSelect.True) ? true : false;
            }
        }
        Alignment IDataGridAgent.HeadAlignment
        {
            get
            {
                return (Alignment)Enum.Parse(typeof(Alignment), ((DataGridView)((IDataGridAgent)this).GridControl).ColumnHeadersDefaultCellStyle.Alignment.ToString());
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).ColumnHeadersDefaultCellStyle.Alignment = (DataGridViewContentAlignment)Enum.Parse(typeof(DataGridViewContentAlignment), value.ToString());
            }
        }
        int IDataGridAgent.HeaderHeight
        {
            get
            {
                return this.headerHeight;
            }
            set
            {
                this.headerHeight = value;
                this.SetHeaderHeight();
            }
        }
        int IDataGridAgent.HeaderRowCount {
            get
            {
                return this.headerRowCount;
            }
            set
            {
                this.headerRowCount = value;
                this.SetHeaderHeight();
            }
        }
        private void SetHeaderHeight()
        {
            if (((IDataGridAgent)this).HeaderRowCount > 0)
                if (((IDataGridAgent)this).HeaderHeight > 0)//'행 높이
                    ((DataGridView)((IDataGridAgent)this).GridControl).ColumnHeadersHeight = ((IDataGridAgent)this).HeaderRowCount * ((IDataGridAgent)this).HeaderHeight;
                else
                    ((DataGridView)((IDataGridAgent)this).GridControl).ColumnHeadersHeight = ((IDataGridAgent)this).HeaderRowCount * ((IDataGridAgent)this).RowHeight;
        }

        int IDataGridAgent.RowHeight
        {
            get
            {
                return ((DataGridView)((IDataGridAgent)this).GridControl).RowTemplate.Height;
            }
            set
            {
                if (value >= 0)//'행 높이
                    ((DataGridView)((IDataGridAgent)this).GridControl).RowTemplate.Height = value;
            }
        }
        RowHeadersVisible IDataGridAgent.RowHeadersVisible
        {
            get
            {
                return (((DataGridView)((IDataGridAgent)this).GridControl).RowHeadersVisible) ? RowHeadersVisible.True : RowHeadersVisible.False;
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).RowHeadersVisible = (value == RowHeadersVisible.True) ? true : false;
            }
        }
        Selection IDataGridAgent.Selection
        {
            get
            {
                return (Selection)Enum.Parse(typeof(Selection), ((DataGridView)((IDataGridAgent)this).GridControl).SelectionMode.ToString());
            }
            set
            {
                ((DataGridView)((IDataGridAgent)this).GridControl).SelectionMode = (DataGridViewSelectionMode)Enum.Parse(typeof(DataGridViewSelectionMode), value.ToString());
            }
        }

        public DataGridViewAgent()
        {
            this.FilterControl = new Dictionary<TextBox, FilterAttribute>();

            try
            {
                this.menuListText = this.GetAttribute("MenuListText").Split(',').Translate();
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
            }
        }

        void IDataGridAgent.Init(EditAble editAble, AddRows allowAddRows, DeleteRows allowDeleteRows, ResizeRows allowResizeRows, AutoSizeColumns autoSizeColumns
            , AutoSizeRows autoSizeRows, ColumnsHeadersVisible columnsHeadersVisible, EnableMenu enableMenu, MultiSelect multiSelect, Alignment headAlign, int headerHeight, int headerRowCount
            , int rowHeight, RowHeadersVisible rowHeadersVisible, Selection selectionMode)
        {
            DataGridView dataGridView;

            dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);

            dataGridView.DefaultCellStyle.WrapMode = DataGridViewTriState.False; //자동 줄바꿈 금지
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;//컬럼헤더 높이 변경 금지

            ((IDataGridAgent)this).EditAble = editAble;
            ((IDataGridAgent)this).AddRows = allowAddRows;
            ((IDataGridAgent)this).DeleteRows = allowDeleteRows;
            ((IDataGridAgent)this).ResizeRows = allowResizeRows;
            ((IDataGridAgent)this).AutoSizeColumns = autoSizeColumns;
            ((IDataGridAgent)this).AutoSizeRows = autoSizeRows;
            ((IDataGridAgent)this).ColumnsHeadersVisible = columnsHeadersVisible;
            ((IDataGridAgent)this).EnableMenu = enableMenu;
            ((IDataGridAgent)this).MultiSelect = multiSelect;
            ((IDataGridAgent)this).HeadAlignment = headAlign;
            ((IDataGridAgent)this).RowHeight = rowHeight;

            ((IDataGridAgent)this).HeaderRowCount = headerRowCount;//'컬럼헤더 수
            ((IDataGridAgent)this).HeaderHeight = headerHeight;

            ((IDataGridAgent)this).RowHeadersVisible = rowHeadersVisible;
            ((IDataGridAgent)this).Selection = selectionMode;

            dataGridView.DoubleBuffered(true);

            dataGridView.DataSourceChanged -= this.DataGridView_DataSourceChanged;
            dataGridView.DataSourceChanged += this.DataGridView_DataSourceChanged;

            dataGridView.DataBindingComplete -= this.DataGridView_DataBindingComplete;
            dataGridView.DataBindingComplete += this.DataGridView_DataBindingComplete;

            this.SetSkin(dataGridView);
        }

        void IDataGridAgent.AddColumn(int width, ColumnVisible visible, EditAble editAble, Filter allowFilter, Merge allowMerge, Sort sortMode, object editControl, Alignment textAlign, string format, string name, params string[] caption)
        {
            DataGridViewColumn dataGridViewColumn;
            DataGridView dataGridView;

            dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);

            if (dataGridView.Columns[name] != null)//'기존에 컬럼이 있으면 가져오기
            {

                dataGridViewColumn = dataGridView.Columns[name];

                if (editControl != null) //'컬럼을 교체를 하면 기존에꺼를 제거함
                {
                    dataGridView.Columns.Remove(dataGridViewColumn);
                    dataGridViewColumn = (DataGridViewColumn)editControl;
                }
            }
            else
                if (editControl != null) // '교체할 컬럼이 있으면 컬럼을 교체
                dataGridViewColumn = (DataGridViewColumn)editControl;
            else //'교체할 컬럼이 없으면 생성
                dataGridViewColumn = new DataGridViewTextBoxColumn();

            dataGridViewColumn.DataPropertyName = name;//'Data 컬럼명
            dataGridViewColumn.Name = name;

            caption = caption.Translate();

            if (((IDataGridAgent)this).HeaderRowCount == 1)//'컬럼 헤더 수가 1개면
                dataGridViewColumn.HeaderText = caption[caption.Length - 1];
            else
            {
                dataGridViewColumn.HeaderText = "";

                for (int i = 0; i <= ((IDataGridAgent)this).HeaderRowCount - 1; i++)
                    if (i < caption.Length)
                        dataGridViewColumn.HeaderText += caption[i] + Environment.NewLine;
                    else
                        dataGridViewColumn.HeaderText += Environment.NewLine;

                if (dataGridViewColumn.HeaderText.Length > 0)
                    dataGridViewColumn.HeaderText = dataGridViewColumn.HeaderText.Substring(0, dataGridViewColumn.HeaderText.Length - 2);// '마지막 vbCrLf 제거
            }

            dataGridViewColumn.Width = width;
            dataGridViewColumn.Visible = (visible == ColumnVisible.True);
            dataGridViewColumn.ReadOnly = !(editAble == EditAble.True);

            //_AllowFilter(구현이 안되었음)

            //_AllowMerge(구현이 안되었음)


            dataGridViewColumn.SortMode = (DataGridViewColumnSortMode)Enum.Parse(typeof(DataGridViewColumnSortMode), sortMode.ToString());
            dataGridViewColumn.DefaultCellStyle.Alignment = (DataGridViewContentAlignment)Enum.Parse(typeof(DataGridViewContentAlignment), textAlign.ToString());
            dataGridViewColumn.DefaultCellStyle.Format = format;

            if (dataGridView.Columns[name] == null)
                dataGridView.Columns.Add(dataGridViewColumn);
        }

        void IDataGridAgent.Clear()
        {
            DataGridView dataGridView;

            dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);

            dataGridView.Columns.Clear();
            dataGridView.DataSource = null;
        }

        void IDataGridAgent.RemoveColumn(string name)
        {
            DataGridView dataGridView;

            dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);

            dataGridView.Columns.Remove(name);
        }

        void IDataGridAgent.RemoveColumn(int index)
        {
            DataGridView dataGridView;

            dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);

            dataGridView.Columns.Remove(dataGridView.Columns[index]);
        }

        private void SetContextMenuStrip(DataGridView dataGridView, string[] menuListText)
        {
            ContextMenuStrip contextMenuStrip;
            ToolStripMenuItem toolStripMenuItem;

            if (this.menuListText == null)
                return;

            contextMenuStrip = new ContextMenuStrip();

            contextMenuStrip.Opening += ContextMenuStrip_Opening; 

            toolStripMenuItem = null;

            for (int i = 0; i <= this.menuListText.Length - 1; i++)
            {
                if (i == 2)
                    continue;

                if (this.menuListText[i] != "")
                {
                    toolStripMenuItem = new ToolStripMenuItem(this.menuListText[i], null, ToolStripMenuItem_Click);
                    contextMenuStrip.Items.Add(toolStripMenuItem);
                }
                else
                    contextMenuStrip.Items.Add(new ToolStripSeparator());

                switch (i)
                {
                    case 5://행추가
                        if (!dataGridView.AllowUserToAddRows)
                            toolStripMenuItem.Enabled = false;
                        break;

                    case 6://행삭제
                        if (!dataGridView.AllowUserToDeleteRows)
                            toolStripMenuItem.Enabled = false;
                        break;

                    case 7://행복사
                        if (!dataGridView.AllowUserToAddRows)
                            toolStripMenuItem.Enabled = false;
                        break;

                    //case 9 ://Sum
                    //    _ToolStripMenuItem.Enabled = false;
                    //    break;
                    //
                    //case 10://Avg
                    //    _ToolStripMenuItem.Enabled = false;
                    //    break;
                }
            }

            dataGridView.ContextMenuStrip = contextMenuStrip;
        }

        private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ContextMenuStrip contextMenuStrip;

            DataGridView dataGridView;
            decimal sum;
            decimal count;
            decimal avg;

            sum = 0;
            count = 0;
            avg = 0;

            try
            {
                dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);
                contextMenuStrip = (ContextMenuStrip)sender;

                foreach (DataGridViewCell dataGridViewCell in dataGridView.SelectedCells)
                {
                    try
                    {
                        if (!dataGridViewCell.Visible)
                            continue;

                        sum += dataGridViewCell.Value.ToString().ToDecimal();
                        count += 1;
                    }
                    catch (Exception exception)
                    {
                        DiagnosticsTool.MyTrace(exception);
                    }
                }

                avg = sum / count;

                contextMenuStrip.Items[8].Text = string.Format("Sum : {0}", sum);
                contextMenuStrip.Items[9].Text = string.Format("Avg : {0}", avg);

                contextMenuStrip.Items[8].Tag = sum;
                contextMenuStrip.Items[9].Tag = avg;
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);

                contextMenuStrip = (ContextMenuStrip)sender;
                contextMenuStrip.Items[8].Text = string.Format("Sum : {0}", 0);
                contextMenuStrip.Items[9].Text = string.Format("Avg : {0}", 0);

                contextMenuStrip.Items[8].Tag = 0;
                contextMenuStrip.Items[9].Tag = 0;
            }

        }

        private void ToolStripMenuItem_Click(Object _sender, EventArgs e)
        {
            ContextMenuStrip contextMenuStrip;
            ToolStripMenuItem toolStripMenuItem;
            SaveFileDialog fileDialog;
            DataGridView dataGridView;
            Process process;
            System.Data.DataRowView dataRowView;

            toolStripMenuItem = (ToolStripMenuItem)_sender;
            contextMenuStrip = (ContextMenuStrip)toolStripMenuItem.Owner;
            dataGridView = (DataGridView)contextMenuStrip.SourceControl;

            //엑셀 저장, 엑셀 저장 & 열기
            if (toolStripMenuItem.Text.Equals(this.menuListText[0]) || toolStripMenuItem.Text.Equals(this.menuListText[1]))
            {
                fileDialog = new SaveFileDialog()
                {
                    DefaultExt = "*.xls",
                    Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*"
                };

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (!ExportExcel(fileDialog.FileName, dataGridView))
                        return;

                    if (toolStripMenuItem.Text.Equals(this.menuListText[1])) //'엑셀 저장 & 열기
                    {
                        process = new Process();
                        process = Process.Start(fileDialog.FileName);
                    }
                }
                return;
            }

            //출력
            if (toolStripMenuItem.Text.Equals(menuListText[2]))
            {
                //'_PrintDocument = New Printing.PrintDocument
                //'_PrintDocument.PrinterSettings.DefaultPageSettings.Margins = New Printing.Margins(30, 40, 30, 30)
                //'AddHandler _PrintDocument.BeginPrint, AddressOf Document_BeginPrint
                //'AddHandler _PrintDocument.PrintPage, AddressOf Document_PrintPage
                //
                //'_PrintPreviewDialog = New PrintPreviewDialog
                //'_PrintPreviewDialog.Document = _PrintDocument
                //'_PrintPreviewDialog.ShowDialog(_DataGridView.FindForm)
                return;
            }

            //복사
            if (toolStripMenuItem.Text.Equals(menuListText[3]))
            {
                Clipboard.SetDataObject(dataGridView.GetClipboardContent());
                return;
            }


            //행추가
            if (toolStripMenuItem.Text.Equals(menuListText[5]))
            {
                if (dataGridView.DataSource != null && dataGridView.DataSource is System.Data.DataView)
                {
                    if (dataGridView.AllowUserToAddRows)
                    {
                        ((System.Data.DataView)dataGridView.DataSource).AddNew().EndEdit();
                    }
                }

                return;
            }


            //행삭제
            if (toolStripMenuItem.Text.Equals(menuListText[6]))
            {
                if (dataGridView.DataSource != null && dataGridView.DataSource is System.Data.DataView)
                    if (dataGridView.AllowUserToDeleteRows)
                        ((System.Data.DataView)dataGridView.DataSource).Delete(dataGridView.CurrentRow.Index);

                return;
            }

            dataRowView = null;

            //행복사
            if (toolStripMenuItem.Text.Equals(menuListText[7]))
            {
                if (dataGridView.DataSource is System.Data.DataView)
                {
                    if (dataGridView.AllowUserToAddRows)
                    {
                        try
                        {
                            dataRowView = ((System.Data.DataView)dataGridView.DataSource).AddNew();

                            for (int i = 0; i < dataGridView.CurrentRow.Cells.Count; i++)
                            {
                                dataRowView.Row[i] = dataGridView.CurrentRow.Cells[i].Value;
                            }
                        }
                        finally
                        {
                            dataRowView?.EndEdit();
                        }
                    }
                }

                return;
            }

            //Sum
            if (toolStripMenuItem.Equals(contextMenuStrip.Items[8]))
            {
                Clipboard.SetText(((decimal)toolStripMenuItem.Tag).ToString());
            }

            //Avg
            if (toolStripMenuItem.Equals(contextMenuStrip.Items[9]))
            {
                Clipboard.SetText(((decimal)toolStripMenuItem.Tag).ToString());
            }
        }

        private static bool ExportExcel(string path, DataGridView dataGridView)
        {
            Microsoft.Office.Interop.Excel.Application application;
            Microsoft.Office.Interop.Excel._Workbook workbook;
            Microsoft.Office.Interop.Excel._Worksheet worksheet;
            object[,] data;

            data = new string[dataGridView.RowCount, dataGridView.ColumnCount];
            application = null;
            workbook = null;

            try
            {
                application = new Microsoft.Office.Interop.Excel.Application();
                workbook = application.Workbooks.Add(Type.Missing);
                worksheet = workbook.Sheets[1];

                for (int i = 0; i <= dataGridView.ColumnCount - 1; i++)//'컬럼 헤더 값
                    worksheet.Cells[1, i + 1] = dataGridView.Columns[i].HeaderText;

                for (int j = 0; j <= dataGridView.RowCount - 1; j++)
                    for (int i = 0; i <= dataGridView.ColumnCount - 1; i++)
                        if (dataGridView.Columns[i].Visible)
                            data[j, i] = dataGridView.Rows[j].Cells[i].Value.ToString();

                worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[dataGridView.RowCount + 1, dataGridView.ColumnCount]].Value2 = data;
                //_Worksheet.Range(_Worksheet.Cells(1, 1), _Worksheet.Cells(1, _DataGridView.ColumnCount)).Font.Bold = _DataGridView.ColumnHeadersDefaultCellStyle.Font.Bold
                worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, dataGridView.ColumnCount]].EntireColumn.AutoFit();
                workbook.SaveAs(path);

                return true;
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
                return false;
            }
            finally
            {
                worksheet = null;

                if (workbook != null)
                {
                    workbook.Close();
                    workbook = null;
                }

                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
            }
        }

        void IDataGridAgent.AddColumnFiter(SearchAll searchAll, StartsWith startsWith, AutoComplete autoComplete, string name, params System.Windows.Forms.Control[] controls)
        {
            DataGridView dataGridView;
            FilterAttribute fiterAttribute;
            TextBox textBox;

            dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);

            fiterAttribute = new FilterAttribute()
            {
                DataGridView = dataGridView,
                ColumnName = name,
                IsSearchAll = (searchAll == SearchAll.True),
                IsStartsWith = (startsWith == StartsWith.True),
                AutoCompleteMode = (AutoCompleteMode)Enum.Parse(typeof(AutoCompleteMode), autoComplete.ToString()),
                AutoCompleteStringCollection = new AutoCompleteStringCollection()
            };

            foreach (System.Windows.Forms.Control control in controls)
            {
                if (control is TextBox)
                {
                    textBox = (TextBox)control;

                    this.FilterControl.Add(textBox, fiterAttribute);

                    if (fiterAttribute.AutoCompleteMode != AutoCompleteMode.None)
                    {
                        textBox.TextChanged -= this.Control_TextChanged;
                        textBox.TextChanged += this.Control_TextChanged;

                        textBox.AutoCompleteMode = fiterAttribute.AutoCompleteMode;
                        textBox.AutoCompleteCustomSource = fiterAttribute.AutoCompleteStringCollection;
                        textBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    }

                    textBox.DoubleBuffered(true);
                }
            }
        }
        void IDataGridAgent.AddColumnFiter(SearchAll searchAll, StartsWith _StartsWith, AutoComplete autoComplete, int _Index, params System.Windows.Forms.Control[] controls)
        {
            DataGridView dataGridView;

            dataGridView = ((DataGridView)((IDataGridAgent)this).GridControl);

            ((IDataGridAgent)this).AddColumnFiter(searchAll, _StartsWith, autoComplete, dataGridView.Columns[_Index].Name, controls);
        }

        void IDataGridAgent.RemoveColumnFiter(params System.Windows.Forms.Control[] controls)
        {
            if (controls is TextBox[])
                foreach (TextBox textBox in (TextBox[])controls)
                {
                    if (this.FilterControl.ContainsKey(textBox))
                    {
                        if (this.FilterControl[textBox].AutoCompleteMode != AutoCompleteMode.None)
                        {
                            textBox.TextChanged -= this.Control_TextChanged;
                            textBox.AutoCompleteMode = AutoCompleteMode.None;
                            textBox.AutoCompleteCustomSource = null;
                            textBox.AutoCompleteSource = AutoCompleteSource.None;
                        }

                        this.FilterControl.Remove(textBox);
                    }
                }
        }

        void Control_TextChanged(object sender, EventArgs e)
        {
            DataGridView dataGridView;
            System.Data.DataView dataView;
            TextBox textBox;
            FilterAttribute filterAttribute;
            StringBuilder stringBuilder;
            //string _Tmp;
            decimal decimalTmp;
            string text;

            textBox = (TextBox)sender;
            filterAttribute = this.FilterControl[textBox];
            dataGridView = filterAttribute.DataGridView;

            if (dataGridView.DataSource == null)
                return;

            dataView = null;

            if (dataGridView.DataSource is System.Data.DataView)
                dataView = (System.Data.DataView)dataGridView.DataSource;

            if (dataGridView.DataSource is System.Data.DataSet)
            {
                if (dataGridView.DataMember != null && dataGridView.DataMember != "")
                    dataView = ((System.Data.DataSet)dataGridView.DataSource).Tables[dataGridView.DataMember].DefaultView;
                else
                    dataView = ((System.Data.DataSet)dataGridView.DataSource).Tables[0].DefaultView;
            }

            if (dataGridView.DataSource is System.Data.DataTable)
                dataView = ((System.Data.DataTable)dataGridView.DataSource).DefaultView;

            if (dataView == null)
                return;

            text = textBox.Text;
            text = text.Replace("[", "[[").Replace("]", "]]").Replace("[[", "[[]").Replace("]]", "[]]").Replace("*", "[*]").Replace("%", "[%]").Replace("'", "''");

            stringBuilder = new StringBuilder();
            if (!text.Equals(""))
            {
                foreach (DataGridViewColumn _DataGridViewColumn in dataGridView.Columns)
                {
                    if (!_DataGridViewColumn.Visible)//보이는 컬럼만 검색
                        continue;

                    if (_DataGridViewColumn.Name == filterAttribute.ColumnName && !filterAttribute.IsSearchAll)
                    {
                        if (_DataGridViewColumn.DefaultCellStyle.Alignment == DataGridViewContentAlignment.MiddleRight)
                        {
                            if (text.ToTryDecimal(out decimalTmp))
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '" + _Text + "%' ";
                                else
                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '%{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '%" + _Text + "%' ";
                            }
                            else
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                                else
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                            }
                        }
                        else
                        {
                            if (filterAttribute.IsStartsWith)
                                stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                            else
                                stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                        }
                        break;
                    }

                    if (filterAttribute.IsSearchAll)
                    {
                        if (_DataGridViewColumn.DefaultCellStyle.Alignment == DataGridViewContentAlignment.MiddleRight)
                        {
                            if (text.ToTryDecimal(out decimalTmp))
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '" + _Text + "%' ";
                                else

                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '%{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '%" + _Text + "%' ";
                            }
                            else
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                                else
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                            }
                        }
                        else
                        {
                            if (filterAttribute.IsStartsWith)
                                stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                            else
                                stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataPropertyName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                        }
                    }
                }
            }

            try
            {
                if (stringBuilder.ToString().StartsWith("OR "))
                    dataView.RowFilter = stringBuilder.ToString(3, stringBuilder.Length - 3);
                else
                    dataView.RowFilter = stringBuilder.ToString();
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
                dataView.RowFilter = "";
            }
        }

        private void DataGridView_DataSourceChanged(object sender, EventArgs e)
        {
            DataGridView dataGridView;
            System.Data.DataView dataView;

            dataGridView = (DataGridView)sender;

            if (dataGridView.DataSource == null)
                return;

            if (this.FilterControl == null)
                return;

            dataView = null;

            if (dataGridView.DataSource is System.Data.DataView)
                dataView = (System.Data.DataView)dataGridView.DataSource;

            if (dataGridView.DataSource is System.Data.DataSet)
                dataView = ((System.Data.DataSet)dataGridView.DataSource).Tables[0].DefaultView;

            if (dataGridView.DataSource is System.Data.DataTable)
                dataView = ((System.Data.DataTable)dataGridView.DataSource).DefaultView;

            if (dataView == null)
                return;

            foreach (FilterAttribute filterAttribute in this.FilterControl.Values)
            {
                //_FilterAttribute.AutoCompleteStringCollection.Clear();

                for (int i = 0; i < dataView.Table.Rows.Count; i++)
                {
                    if (dataView.Table.Columns.Contains(filterAttribute.ColumnName)
                        && !filterAttribute.AutoCompleteStringCollection.Contains(dataView.Table.Rows[i][filterAttribute.ColumnName].ToString()))
                        filterAttribute.AutoCompleteStringCollection.Add(dataView.Table.Rows[i][filterAttribute.ColumnName].ToString());
                }
            }

        }

        private void DataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView dataGridView;
            System.Data.DataView dataView;

            dataGridView = (DataGridView)sender;

            if (dataGridView.DataSource == null)
                return;

            dataView = null;

            if (dataGridView.DataSource is System.Data.DataView)
                dataView = (System.Data.DataView)dataGridView.DataSource;

            if (dataGridView.DataSource is System.Data.DataSet)
                if (dataGridView.DataMember != null && dataGridView.DataMember != "")
                    dataView = ((System.Data.DataSet)dataGridView.DataSource).Tables[dataGridView.DataMember].DefaultView;
                else
                    dataView = ((System.Data.DataSet)dataGridView.DataSource).Tables[0].DefaultView;

            if (dataGridView.DataSource is System.Data.DataTable)
                dataView = ((System.Data.DataTable)dataGridView.DataSource).DefaultView;

            if (dataView == null)
                return;

            foreach (DataGridViewRow dataGridViewRow in dataGridView.Rows)
            {
                if (dataGridViewRow.IsNewRow) continue;

                if (dataGridViewRow.HeaderCell.Value != null)
                    break;

                dataGridViewRow.HeaderCell.Value = (dataGridViewRow.Index + 1).ToString();
            }

            dataGridView.TopLeftHeaderCell.Value = dataGridView.Rows.Count.ToString() + "/" + dataView.Table.Rows.Count.ToString();

            dataGridView.RowHeadersWidth = 30 + (dataGridView.Rows.Count.ToString().Length * 10);
            dataGridView.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void SetSkin(DataGridView dataGridView)
        {
            string skinName;
            string tmp;
            System.Drawing.Color color;
            System.Drawing.Font font;
            //string[] _Tmps;

            skinName = (string)Config.Client.GetAttribute("SkinName");

            if (skinName == null || skinName == "")
                try
                {
                    skinName = this.GetAttribute("SkinName");
                }
                catch (Exception ex)
                {
                }

            if (skinName == null || skinName == "")
                return;

            dataGridView.EnableHeadersVisualStyles = false;

            color = this.GetAttributeColor(string.Format("{0}.BackgroundColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.BackgroundColor = color;

            color = this.GetAttributeColor(string.Format("{0}.ColumnHeadersDefaultCellStyle.BackColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.ColumnHeadersDefaultCellStyle.BackColor = color;

            color = this.GetAttributeColor(string.Format("{0}.ColumnHeadersDefaultCellStyle.ForeColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = color;

            font = this.GetAttributeFont(dataGridView.Font, string.Format("{0}.ColumnHeadersDefaultCellStyle.Font", skinName));
            if (font != null)
            {
                dataGridView.ColumnHeadersDefaultCellStyle.Font = font;
                dataGridView.RowHeadersDefaultCellStyle.Font = font;
            }

            tmp = this.GetAttribute(string.Format("{0}.ColumnHeadersBorderStyle", skinName));
            if (tmp != null)
                dataGridView.ColumnHeadersBorderStyle = (DataGridViewHeaderBorderStyle)Enum.Parse(typeof(DataGridViewHeaderBorderStyle), tmp);

            color = this.GetAttributeColor(string.Format("{0}.RowHeadersDefaultCellStyle.BackColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.RowHeadersDefaultCellStyle.BackColor = color;

            color = this.GetAttributeColor(string.Format("{0}.RowHeadersDefaultCellStyle.ForeColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.RowHeadersDefaultCellStyle.ForeColor = color;

            tmp = this.GetAttribute(string.Format("{0}.RowHeadersBorderStyle", skinName));
            if (tmp != null)
                dataGridView.RowHeadersBorderStyle = (DataGridViewHeaderBorderStyle)Enum.Parse(typeof(DataGridViewHeaderBorderStyle), tmp);

            color = this.GetAttributeColor(string.Format("{0}.RowsDefaultCellStyle.SelectionBackColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.RowsDefaultCellStyle.SelectionBackColor = color;

            color = this.GetAttributeColor(string.Format("{0}.RowsDefaultCellStyle.SelectionForeColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.RowsDefaultCellStyle.SelectionForeColor = color;

            color = this.GetAttributeColor(string.Format("{0}.AlternatingRowsDefaultCellStyle.BackColor", skinName));
            if (color != System.Drawing.Color.Empty)
                dataGridView.AlternatingRowsDefaultCellStyle.BackColor = color;

            //_DataGridView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            //_DataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            dataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
        }

        protected struct FilterAttribute
        {
            public DataGridView DataGridView { get; set; }
            public string ColumnName { get; set; }
            public bool IsSearchAll { get; set; }
            public bool IsStartsWith { get; set; }
            public AutoCompleteMode AutoCompleteMode { get; set; }
            public AutoCompleteStringCollection AutoCompleteStringCollection { get; set; }
        }
    }
}