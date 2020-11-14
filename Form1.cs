using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.Text;
using System.Diagnostics;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private IWorkbook DeliveryWorkbook = null;
        private IWorkbook FormatWorkbook = null;
        private IWorkbook DetailWorkbook = null;
        private IWorkbook MaterialWorkbook = null;
        private IWorkbook Format2Workbook = null;
        private IWorkbook CalculationWorkBook = null;
        private IWorkbook ReplaceWorkBook = null;
        private IWorkbook ToReplaceWorkBook = null;


        private List<KeyValuePair<string,IWorkbook>> ProjectIWorkbooks = new List<KeyValuePair<string, IWorkbook>>();

        delegate void DelSetStatus(String Msg,int Row=0);

        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 处理出库表Excel
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="SearStr"></param>
        private void HandleDeliveryExcelData(ISheet UseCollectionSheet,ISheet WithDrawCollectionSheet,ISheet MaterialConsumptionSheet)
        {
            IRow Row;
            List<int> WithDrawIndexlst = new List<int>();
            List<int> TotalConsumptionIndexlst = new List<int>();   //甲供情况1
            List<int> TotalConsumptionIndex2lst = new List<int>(); //甲供情况2
            List<int> SupplySumIndexlst = new List<int>();               //乙供
            DelSetStatus fn = SetStatus;
            //未退库扣材料款数
            int UnWithDrawCount = 0;
            //已退库数
            int HasWithDrawCount = 0;
            //工程用量
            int TotalConsumptionCount = 0;
            //乙供
            int SupplySumCount = 0;


            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { "正在寻找未退库扣材料款和已退库行",0});
            }
            else
            {
                fn("正在寻找未退库扣材料款和已退库行");
            }

            for (int i = 4; i < UseCollectionSheet.LastRowNum; i++)
            {
                Row = UseCollectionSheet.GetRow(i);
                //double WithDrawNum = Row.Cells[7].NumericCellValue;//已退库
                //double UnWithDrawNum = Row.Cells[10].NumericCellValue;//未退库和扣款
                //double TotalConsumption = Row.Cells[9].NumericCellValue;
                //double SupplySum = Row.Cells[11].NumericCellValue;
                if(Row.Cells[0].CellType == CellType.Blank)
                {
                    break;
                }
                if ((Row.Cells[10].CellType == CellType.Numeric) && Row.Cells[10].NumericCellValue!=0)
                {

                    IRow wRow = WithDrawCollectionSheet.GetRow(3 + UnWithDrawCount);
                    wRow.Cells[0].SetCellValue(HasWithDrawCount + UnWithDrawCount + 1);
                    wRow.Cells[1].SetCellValue(Row.Cells[2].ToString());//物料名称
                    wRow.Cells[2].SetCellValue(Row.Cells[3].ToString());//物料规格
                    wRow.Cells[3].SetCellValue(Row.Cells[4].ToString());//物料型号
                    wRow.Cells[4].SetCellValue(Row.Cells[5].ToString());//物料单位
                    wRow.Cells[5].SetCellValue(Row.Cells[10].NumericCellValue);//数量
                    wRow.Cells[6].SetCellValue("未退库和扣款");//备注
                    wRow.Cells[7].SetCellValue(Row.Cells[12].ToString());//单据号
                    wRow.Cells[8].SetCellValue(Row.Cells[13].NumericCellValue);//单价
                    wRow.Cells[9].SetCellFormula(wRow.Cells[9].CellFormula);//计算公式

                    UnWithDrawCount++;
                    

                }

 
                //记录已入库行
                else if ((Row.Cells[7].CellType == CellType.Numeric) && Row.Cells[7].NumericCellValue!=0)
                {
                    WithDrawIndexlst.Add(i);
                }
                //工程用量非空、乙供空，复制工程用量到工程用量列   : 甲供
                //工程用量非空、乙供非空、 出库总数非空，复制出库数到工程用量列   : 甲供
                //工程用量非空、乙供非空、 复制乙供到工程用量列   : 乙供


                //记录工程用量非空>0且乙供空或非空
                if ((Row.Cells[9].CellType == CellType.Numeric) && Row.Cells[9].NumericCellValue > 0 
                   // !(Row.Cells[11].CellType == CellType.String) && Row.Cells[11].NumericCellValue == 0
                    )
                {
                    //乙供空
                    if(!(Row.Cells[11].CellType == CellType.Numeric) && Row.Cells[11].StringCellValue == "")
                    {
                        TotalConsumptionIndexlst.Add(i);
                        
                    }
                    //乙供非空
                    else if ((Row.Cells[11].CellType == CellType.Numeric))
                    {
                        //出库总数非空
                        if ((Row.Cells[8].CellType == CellType.Numeric) && Row.Cells[8].NumericCellValue != 0)
                        {
                            TotalConsumptionIndex2lst.Add(i);
                        }

                        SupplySumIndexlst.Add(i);
                    }

                }
                else
                {
                    
                    continue;
                }
            }

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { string.Format("共处理{0}行未退库和扣款", UnWithDrawCount), 0 });
            }
            else
            {
                fn(string.Format("共处理未退库和扣款{0}行", UnWithDrawCount));
            }


            foreach (var i in WithDrawIndexlst)
            {
                Row = UseCollectionSheet.GetRow(i);
                IRow wRow = WithDrawCollectionSheet.GetRow(3 + HasWithDrawCount + UnWithDrawCount);
                wRow.Cells[0].SetCellValue(HasWithDrawCount + 1);
                wRow.Cells[1].SetCellValue(Row.Cells[2].ToString());//物料名称
                wRow.Cells[2].SetCellValue(Row.Cells[3].ToString());//物料规格
                wRow.Cells[3].SetCellValue(Row.Cells[4].ToString());//物料型号
                wRow.Cells[4].SetCellValue(Row.Cells[5].ToString());//物料单位
                wRow.Cells[5].SetCellValue(Row.Cells[7].NumericCellValue);//数量
                wRow.Cells[6].SetCellValue("已退库");//备注
                wRow.Cells[7].SetCellValue(Row.Cells[12].ToString());//单据号


                HasWithDrawCount++;
            }

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { string.Format("共处理{0}行已退库", HasWithDrawCount), 0 });
            }
            else
            {
                fn(string.Format("共处理{0}行已退库", HasWithDrawCount));
            }

            foreach (var i in TotalConsumptionIndexlst)
            {
                Row = UseCollectionSheet.GetRow(i);
                IRow wRow = MaterialConsumptionSheet.GetRow(3 + TotalConsumptionCount);
                wRow.Cells[0].SetCellValue(TotalConsumptionCount + 1);
                wRow.Cells[1].SetCellValue(Row.Cells[2].ToString());//物料名称
                wRow.Cells[2].SetCellValue(Row.Cells[3].ToString());//物料规格
                wRow.Cells[3].SetCellValue(Row.Cells[4].ToString());//物料型号
                wRow.Cells[4].SetCellValue(Row.Cells[5].ToString());//物料单位
                wRow.Cells[5].SetCellValue(Row.Cells[9].NumericCellValue);//工程用量
                wRow.Cells[6].SetCellValue("甲供");//备注
                TotalConsumptionCount++;
            }
            foreach (var i in TotalConsumptionIndex2lst)
            {
                Row = UseCollectionSheet.GetRow(i);
                IRow wRow = MaterialConsumptionSheet.GetRow(3 + TotalConsumptionCount);
                wRow.Cells[0].SetCellValue(TotalConsumptionCount + 1);
                wRow.Cells[1].SetCellValue(Row.Cells[2].ToString());//物料名称
                wRow.Cells[2].SetCellValue(Row.Cells[3].ToString());//物料规格
                wRow.Cells[3].SetCellValue(Row.Cells[4].ToString());//物料型号
                wRow.Cells[4].SetCellValue(Row.Cells[5].ToString());//物料单位
                wRow.Cells[5].SetCellValue(Row.Cells[8].NumericCellValue);//出库数
                wRow.Cells[6].SetCellValue("甲供");//备注
                TotalConsumptionCount++;
            }
            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { string.Format("共处理{0}行工程用量(甲供)", TotalConsumptionCount), 0 });
            }
            else
            {
                fn(string.Format("共处理{0}行工程用量(甲供)", TotalConsumptionCount));
            }

            foreach (var i in SupplySumIndexlst)
            {
                Row = UseCollectionSheet.GetRow(i);
                IRow wRow = MaterialConsumptionSheet.GetRow(3 + TotalConsumptionCount + SupplySumCount);
                wRow.Cells[0].SetCellValue(TotalConsumptionCount + SupplySumCount + 1);
                wRow.Cells[1].SetCellValue(Row.Cells[2].ToString());//物料名称
                wRow.Cells[2].SetCellValue(Row.Cells[3].ToString());//物料规格
                wRow.Cells[3].SetCellValue(Row.Cells[4].ToString());//物料型号
                wRow.Cells[4].SetCellValue(Row.Cells[5].ToString());//物料单位
                wRow.Cells[5].SetCellValue(Row.Cells[11].NumericCellValue);//乙供
                wRow.Cells[6].SetCellValue("乙供");//备注
                SupplySumCount++;
            }

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { string.Format("共处理{0}行乙供", SupplySumCount), 0 });
            }
            else
            {
                fn(string.Format("共处理{0}行乙供", SupplySumCount));
            }








        }

        /// <summary>
        /// 处理明细表
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="SearStr"></param>
        private void HandleDeliveryStoreExcelData(ISheet DetailSheet,ISheet FormatSheet)
        {
            IRow Row;
            Dictionary<string, int> DicMaterialCountB = new Dictionary<string, int>();
            DelSetStatus fn = SetStatus;
            int Count = 0;

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { "正在将明细表复制到格式表",0});
            }
            else
            {
                fn("正在将明细表复制到格式表");
            }

            //8-13
            for (int i = 6; i <= DetailSheet.LastRowNum; i++)
            {
                Row = DetailSheet.GetRow(i);


                if (!string.IsNullOrEmpty(Row.Cells[0].ToString()))
                {
                    FormatSheet.GetRow(1).Cells[2].SetCellValue(Row.Cells[2].ToString()); // 复制CEA
                    FormatSheet.GetRow(1).Cells[12].SetCellValue(Row.Cells[3].ToString()); // 复制工程名称
                    IRow wRow = FormatSheet.GetRow(4 + Count);
                    if(Row.Cells[14].CellType!=CellType.Numeric)
                    {
                        continue;
                    }

                    double sum = Row.Cells[14].NumericCellValue;

                    //是否是负数
                    if(sum<0)
                    {
                        wRow.Cells[7].SetCellValue(Convert.ToDouble(sum));//已退货
                    }
                    else
                    {
                        wRow.Cells[6].SetCellValue(Convert.ToDouble(sum));//实发主数量
                    }



                    wRow.Cells[13].SetCellValue(Row.Cells[15].NumericCellValue);//单价
 

                    wRow.Cells[0].SetCellValue(Count + 1);
                    wRow.Cells[12].SetCellValue(Row.Cells[4].ToString());//单据号
                    wRow.Cells[1].SetCellValue(Row.Cells[9].ToString());//物料代码
                    wRow.Cells[2].SetCellValue(Row.Cells[10].ToString());//物料名称
                    wRow.Cells[3].SetCellValue(Row.Cells[11].ToString());//物料规格
                    wRow.Cells[4].SetCellValue(Row.Cells[12].ToString());//型号
                    wRow.Cells[5].SetCellValue(Row.Cells[13].ToString());//物料单位

                    Count++;

                }
                else
                {
                    break;
                }


            }

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { string.Format("共复制明细表到格式表{0}行",Count), 0 });
            }
            else
            {
                fn(string.Format("共复制明细表到格式表{0}行", Count));
            }


        }


        /// <summary>
        /// 复制项目名称和替换序号数据
        /// </summary>
        /// <param name="ToReplaceSheet"></param>
        /// <param name="ReplaceSheet"></param>
        private void HandleReplaceExcelData(ISheet ToReplaceSheet, ISheet ReplaceSheet)
        {
            IRow Row;
            DelSetStatus fn = SetStatus;
            int ReplaceCount = 0;
            int CopyCount = 0;

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { "正在复制项目名称和替换", 0 });
            }
            else
            {
                fn("正在复制项目名称和替换");
            }


            for (int i = 4; i < ToReplaceSheet.LastRowNum; i++)
            {
                Row = ToReplaceSheet.GetRow(i);
                if (Row.GetCell(1, MissingCellPolicy.RETURN_NULL_AND_BLANK).StringCellValue == "小计")
                {
                    break;
                }
                //复制项目名称到最后一列
                Row.GetCell(10, MissingCellPolicy.RETURN_NULL_AND_BLANK).SetCellValue(Row.GetCell(1,MissingCellPolicy.RETURN_NULL_AND_BLANK).StringCellValue);
                ICell Cell = Row.GetCell(2, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                

                if (Cell == null)
                {
                    continue;
                }

                //获取清单序号
                string FindSerialNumber = Cell.CellType == CellType.String ? Cell.StringCellValue : Cell.NumericCellValue.ToString();

                if (Cell.CellType == CellType.Blank)
                {
                    continue;
                }

                //假设未找到
                Cell.SetCellValue("未找到");


                CopyCount++;

                for (int j = 1; j < ReplaceSheet.LastRowNum; j++)
                {
                    IRow RplRow = ReplaceSheet.GetRow(j);
                    ICell RplCell = RplRow.GetCell(0, MissingCellPolicy.RETURN_NULL_AND_BLANK);

                    if (RplCell == null)
                    {
                        continue;
                    }

                    string SerialNumber = RplCell.CellType == CellType.String? RplCell.StringCellValue:RplCell.NumericCellValue.ToString();



                    if(FindSerialNumber == SerialNumber)
                    {
                        ICell ValCell = RplRow.GetCell(2, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        //复制物料编码到原来清单序号
                        Row.GetCell(2, MissingCellPolicy.RETURN_NULL_AND_BLANK).SetCellValue(
                            ValCell.CellType == CellType.String?Convert.ToDouble(ValCell.StringCellValue):ValCell.NumericCellValue);

                        //复制名字
                        Row.GetCell(1, MissingCellPolicy.RETURN_NULL_AND_BLANK).SetCellValue(
                            RplRow.GetCell(1, MissingCellPolicy.RETURN_NULL_AND_BLANK).StringCellValue);
                        //复制单位
                        Row.GetCell(3, MissingCellPolicy.RETURN_NULL_AND_BLANK).SetCellValue(
                             RplRow.GetCell(3, MissingCellPolicy.RETURN_NULL_AND_BLANK).StringCellValue);

                        ValCell = RplRow.GetCell(5, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        //复制单价
                        Row.GetCell(7, MissingCellPolicy.RETURN_NULL_AND_BLANK).SetCellValue(
                            ValCell.CellType == CellType.String ? Convert.ToDouble(ValCell.StringCellValue) : ValCell.NumericCellValue);

                        ReplaceCount++;
                    }

                }

                if(Cell.CellType == CellType.String && Cell.StringCellValue == "未找到")
                {
                    if (this.InvokeRequired)
                    {
                        this.Invoke(fn, new object[] { $"清单序号{FindSerialNumber}未找到", 0 });
                    }
                    else
                    {
                        fn($"清单序号{FindSerialNumber}未找到");
                    }
                }


            }

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { $"共复制项目名称{CopyCount}个,共替换{ReplaceCount}个", 0 });
            }
            else
            {
                fn($"共复制项目名称{CopyCount}个,共替换{ReplaceCount}个");
            }


        }








        /// <summary>
        /// 处理工程平料表
        /// </summary>
        /// <param name="MaterialSheet"></param>
        /// <param name="ProjectSheet"></param>
        /// <param name="Index">第几个工程</param>
        private void HandleProjectMaterial(ISheet MaterialSheet, ISheet ProjectSheet,string ProjectName,int Index)
        {
            IRow Row;
            //物料编码实发主数量相加字典 double[0]存行号、[2]存实发主数量
            Dictionary<string, double[]> ActualSendCountDic = new Dictionary<string, double[]>();

            DelSetStatus fn = SetStatus;
            int Count = 0;

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { string.Format("正在处理工程{0}",ProjectName), 0 });
            }
            else
            {
                fn(string.Format("正在处理工程{0}", ProjectName));
            }

            
            for (int i = 6; i <= ProjectSheet.LastRowNum; i++)
            {
                Row = ProjectSheet.GetRow(i);
                if (!string.IsNullOrWhiteSpace(Row.Cells[0].StringCellValue) && Row.Cells[14].CellType == CellType.Numeric)
                {
                    //是否存在该物料编码
                    if(!ActualSendCountDic.ContainsKey(Row.Cells[9].StringCellValue))
                    {
                        ActualSendCountDic.Add(Row.Cells[9].StringCellValue, new double[] { i,Row.Cells[14].NumericCellValue });
                    }
                    else
                    {
                        //实发主数量相加
                        ActualSendCountDic[Row.Cells[9].StringCellValue][1] += Row.Cells[14].NumericCellValue;
                    }
                }
                else
                {
                    break;
                }
            }
            int SetCellIndex = 6 + Index * 3;
            MissingCellPolicy policy = MissingCellPolicy.CREATE_NULL_AS_BLANK;
            //新建工程列名
            MaterialSheet.GetRow(1).GetCell(SetCellIndex+1, policy).SetCellValue(ProjectName);
            MaterialSheet.GetRow(2).GetCell(SetCellIndex + 1, policy).SetCellValue("工程量");
            MaterialSheet.GetRow(2).GetCell(SetCellIndex + 2, policy).SetCellValue("差异数");

            for (int i = 3; i <= MaterialSheet.LastRowNum; i++)
            {
                Row = MaterialSheet.GetRow(i);
                if (Row == null) break;
                if(!string.IsNullOrWhiteSpace(Row.Cells[0].StringCellValue) && ActualSendCountDic.ContainsKey(Row.Cells[0].StringCellValue))
                {
                    
                    Count++;
                    Row.GetCell(SetCellIndex, policy).SetCellValue(ActualSendCountDic[Row.Cells[0].StringCellValue][1]); //实发主数量
                    Row.GetCell(SetCellIndex+1, policy).SetCellValue(0); //工程量，手填，默认0

                    string Formula 
                        = string.Format("IF({1}{0}-{2}{0}=0,\"\",{1}{0}-{2}{0})", i + 1, GetRowCodeByRowIndex(SetCellIndex), GetRowCodeByRowIndex(SetCellIndex + 1));
                    Row.GetCell(SetCellIndex+2, policy).SetCellFormula
                        (Formula); //差异数 = 实发主数量-工程量

                    ActualSendCountDic[Row.Cells[0].StringCellValue][1] = -1;

                }
            }

            //寻找哪些物料是平料表没有的，插入到对应行中(物料编号排序)
            
            if (ActualSendCountDic.Keys.Count>Count)
            {
                List<double> NewMaterialCode = ActualSendCountDic.Where(m => m.Value[1] != -1).Select(m => m.Value[0]).ToList();

                    for (int i = 3; i <= MaterialSheet.LastRowNum; i++)
                    {
                        Row = MaterialSheet.GetRow(i);
                        if (!string.IsNullOrWhiteSpace(Row.Cells[0].StringCellValue))
                        {
                            int Code;
                            if (int.TryParse(Row.Cells[0].StringCellValue, out Code))
                            {
                                for(int j = 0;j< NewMaterialCode.Count;j++)
                                {
                                    int NewMaterialIndex = (int)NewMaterialCode[j];
                                     //出库表中的新材料所在行
                                    IRow NewRow = ProjectSheet.GetRow(NewMaterialIndex);
                                    IRow CreateRow = null;
                                    ICellStyle SrcCellStyle= Row.Cells[0].CellStyle;
                                    string NewCodeStr = NewRow.Cells[9].StringCellValue;
                                    int NewRowIndex = 0;
                                    int NewCode;

                                    if (int.TryParse(NewRow.Cells[9].StringCellValue, out NewCode))
                                    {
                                        //判断新材料是插前还是插后
                                        if (Code > NewCode)
                                        {
                                            MaterialSheet.ShiftRows(i, MaterialSheet.LastRowNum, 1,true,false);
                                            CreateRow = MaterialSheet.CreateRow(i);
                                            CreateRow.RowStyle = Row.RowStyle;  //设置为当前行风格
                                                
                                            NewMaterialCode.RemoveAt(j);
                                            NewRowIndex = i;
                                        }
                                        //else if (Code <= NewCode - 1)
                                        //{
                                        //    CreateRow = MaterialSheet.CreateRow(i + 1);
                                        //    NewMaterialCode.RemoveAt(j);
                                        //    NewRowIndex = i + 1;
                                        //}
                                    }
                                    ////材料编码为英文则插入到最后
                                    //else
                                    //{
                                    //    CreateRow = MaterialSheet.CreateRow(MaterialSheet.LastRowNum - 1);
                                    //}
                                         if(CreateRow!=null)
                                        {
                                            ICell NewCell = CreateRow.CreateCell(0);
                                            NewCell.CellStyle = SrcCellStyle;
                                            NewCell.SetCellValue(NewCodeStr);

                                            NewCell = CreateRow.CreateCell(1);
                                            NewCell.CellStyle = SrcCellStyle;
                                            NewCell.SetCellValue(NewRow.Cells[10].StringCellValue);

                                            NewCell = CreateRow.CreateCell(2);
                                            NewCell.CellStyle = SrcCellStyle;
                                            NewCell.SetCellValue(NewRow.Cells[11].StringCellValue);

                                            NewCell = CreateRow.CreateCell(3);
                                            NewCell.CellStyle = SrcCellStyle;
                                            NewCell.SetCellValue(NewRow.Cells[12].StringCellValue);

                                                    //CreateRow.CreateCell(152);
                                                    //CreateRow.CreateCell(4).SetCellFormula(string.Format("FB", NewRowIndex));
                                            if (this.InvokeRequired)
                                            {
                                                this.Invoke(fn, new object[] { string.Format("发现新物料编码{0}，已经插入到对应位置", NewCodeStr), 0 });
                                            }
                                            else
                                            {
                                                fn(string.Format("发现新物料编码{0},已经插入到对应位置", NewCodeStr));
                                            }
                                         }

                                
                                }

                            }
                        }   

                    }

                //新的物料编号是英文或无法插入到对应位置，插入到最后一行
                for (int j = 0; j < NewMaterialCode.Count; j++)
                {
                    int NewMaterialIndex = (int)NewMaterialCode[j];
                    IRow NewRow = ProjectSheet.GetRow(NewMaterialIndex);
                    IRow CreateRow = MaterialSheet.CreateRow(MaterialSheet.LastRowNum);
                    string NewCodeStr = NewRow.Cells[9].StringCellValue;
                    CreateRow.CreateCell(0).SetCellValue(NewCodeStr);
                    CreateRow.CreateCell(1).SetCellValue(NewRow.Cells[10].StringCellValue);
                    CreateRow.CreateCell(2).SetCellValue(NewRow.Cells[11].StringCellValue);
                    CreateRow.CreateCell(3).SetCellValue(NewRow.Cells[12].StringCellValue);
                    //CreateRow.CreateCell(152);
                    //CreateRow.CreateCell(4).SetCellFormula(string.Format("FB", MaterialSheet.LastRowNum-1));
                    if (this.InvokeRequired)
                    {
                        this.Invoke(fn, new object[] { string.Format("发现新物料编码{0}，无法插入到对应位置，故插入到末尾", NewCodeStr), 0 });
                    }
                    else
                    {
                        fn(string.Format("发现新物料编码{0}，无法插入到对应位置，故插入到末尾", NewCodeStr));
                    }
                }


            }



            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { string.Format("{0}共处理{1}项物料", ProjectName,Count), 0 });
            }
            else
            {
                fn(string.Format("{0}共处理{1}项物料", ProjectName, Count));
            }


        }


        private void HandleCombination(ISheet DetailSheet)
        {
            IRow Row;
            DelSetStatus fn = SetStatus;
            Dictionary<string, int[]> CombinationRowLogDic = new Dictionary<string, int[]>();

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { "正在合并", 0 });
            }
            else
            {
                fn("正在合并");
            }

            
            for (int i = 3; i < DetailSheet.LastRowNum; i++)
            {

                Row = DetailSheet.GetRow(i);

                if (Row.Cells[0].CellType == CellType.Blank)
                {
                    break;
                }

                string materialCode = Row.Cells[1].ToString();

                if (!string.IsNullOrEmpty(materialCode))
                {
                    if (CombinationRowLogDic.ContainsKey(materialCode))
                    {
                        //记录要合并的行
                        CombinationRowLogDic[materialCode][1] = i;
                    }
                    else
                    {
                        int[] LoggedRowArr = new int[2];
                        LoggedRowArr[0] = i;
                        CombinationRowLogDic[materialCode] = LoggedRowArr;
                    }
                }
            }

            foreach (var kvp in CombinationRowLogDic)
            {
                //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                if(kvp.Value[0]<kvp.Value[1])
                {
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 1, 1));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 2, 2));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 3, 3));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 4, 4));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 5, 5));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 8, 8));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 9, 9));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 10, 10));
                    DetailSheet.AddMergedRegion(new CellRangeAddress(kvp.Value[0], kvp.Value[1], 11, 11));
                }

            }

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { "合并完成", 0 });
            }
            else
            {
                fn("合并完成");
            }


        }

        /// <summary>
        /// 列出差异数
        /// </summary>
        /// <param name="DetailSheet"></param>
        private void HandleContrast(ISheet ContrastSheet,ISheet NewSheet)
        {
            IRow Row;
            DelSetStatus fn = SetStatus;
            //领
            Dictionary<string, List<int>> ProjectContrastPlusNum = new Dictionary<string, List<int>>();
            //退
            Dictionary<string, List<int>> ProjectContrastMinusNum = new Dictionary<string, List<int>>();
            //工程名列表
            Dictionary<string,int> ProjectList = new Dictionary<string, int>();

            int Index = 0;
            int WriteIndex = 0;

            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { "正在列出差异数", 0 });
            }
            else
            {
                fn("正在列出差异数");
            }

            MissingCellPolicy policy = MissingCellPolicy.RETURN_NULL_AND_BLANK;


            for(int j=0;j<ContrastSheet.GetRow(0).LastCellNum-7;j++)
            {

                //工程名
                int ProjectNameIndex = 7 + j;
                string ProjectName = ContrastSheet.GetRow(1).GetCell(ProjectNameIndex, policy).StringCellValue;

                if(!string.IsNullOrWhiteSpace(ProjectName))
                {
                    //差异数列
                    int ContrastCellIndex = ProjectNameIndex+1;

                    ProjectList[ProjectName] = ContrastCellIndex;

                    //该工程名的所有物料差异数
                    for (int i = 3; i < ContrastSheet.LastRowNum; i++)
                    {
                        Row = ContrastSheet.GetRow(i);

                        var ContrastCell = Row.GetCell(ContrastCellIndex, policy);

                        if ((ContrastCell.CellType == CellType.Formula || ContrastCell.CellType == CellType.Numeric) && ContrastCell.NumericCellValue != 0)
                        {
                            if (ContrastCell.NumericCellValue > 0)
                            {
                                if (ProjectContrastPlusNum.ContainsKey(ProjectName))
                                {
                                    ProjectContrastPlusNum[ProjectName].Add(i);
                                }
                                else
                                {
                                    List<int> lst = new List<int>();
                                    lst.Add(i);
                                    ProjectContrastPlusNum[ProjectName] = lst;
                                }
                            }
                            else
                            {
                                if (ProjectContrastMinusNum.ContainsKey(ProjectName))
                                {
                                    ProjectContrastMinusNum[ProjectName].Add(i);
                                }
                                else
                                {
                                    List<int> lst = new List<int>();
                                    lst.Add(i);
                                    ProjectContrastMinusNum[ProjectName] = lst;
                                }
                            }
                        }
                    }
                }
                Index++;
            }

            Index = 0;

            foreach(var item in ProjectList)
            {
                //填写工程名
                NewSheet.CreateRow(WriteIndex++).CreateCell(4).SetCellValue(item.Key);
                //列出所有负数差异数
                if(ProjectContrastMinusNum.ContainsKey(item.Key))
                {
                    NewSheet.CreateRow(WriteIndex++).CreateCell(4).SetCellValue("领");

                    var MinusRowIndexList = ProjectContrastMinusNum[item.Key];

                    foreach(var index in MinusRowIndexList)
                    {
                        Row = ContrastSheet.GetRow(index);
                        var wRow = NewSheet.CreateRow(WriteIndex++);
                        wRow.CreateCell(0).SetCellValue(Row.GetCell(0, policy).StringCellValue);
                        wRow.CreateCell(1).SetCellValue(Row.GetCell(1, policy).StringCellValue);
                        wRow.CreateCell(2).SetCellValue(Row.GetCell(2, policy).StringCellValue);
                        wRow.CreateCell(3).SetCellValue(Row.GetCell(3, policy).StringCellValue);
                       //var dd = Row.GetCell(item.Value, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        wRow.CreateCell(4).SetCellValue(Row.GetCell(item.Value, policy).NumericCellValue);
                    }
                    if (this.InvokeRequired)
                    {
                        this.Invoke(fn, new object[] { string.Format("{0}领:{1}行",item.Key,MinusRowIndexList.Count), 0 });
                    }
                    else
                    {
                        fn(string.Format("{0}领:{1}行", item.Key, MinusRowIndexList.Count));
                    }
                }

                //列出所有正数差异数
                if (ProjectContrastPlusNum.ContainsKey(item.Key))
                {
                    NewSheet.CreateRow(WriteIndex++).CreateCell(4).SetCellValue("退");

                    var PlusRowIndexList = ProjectContrastPlusNum[item.Key];

                    foreach (var index in PlusRowIndexList)
                    {
                        Row = ContrastSheet.GetRow(index);
                        var wRow = NewSheet.CreateRow(WriteIndex++);
                        wRow.CreateCell(0).SetCellValue(Row.GetCell(0, policy).StringCellValue);
                        wRow.CreateCell(1).SetCellValue(Row.GetCell(1, policy).StringCellValue);
                        wRow.CreateCell(2).SetCellValue(Row.GetCell(2, policy).StringCellValue);
                        wRow.CreateCell(3).SetCellValue(Row.GetCell(3, policy).StringCellValue);
                        //double dd = Row.GetCell(item.Value, MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue;
                        wRow.CreateCell(4).SetCellValue(Row.GetCell(item.Value, policy).NumericCellValue);
                    }
                    if (this.InvokeRequired)
                    {
                        this.Invoke(fn, new object[] { string.Format("{0}退:{1}行", item.Key, PlusRowIndexList.Count), 0 });
                    }
                    else
                    {
                        fn(string.Format("{0}退:{1}行", item.Key, PlusRowIndexList.Count));
                    }
                }

                Index++;
                //空行隔开
                NewSheet.CreateRow(WriteIndex++) ;

            }


            if (this.InvokeRequired)
            {
                this.Invoke(fn, new object[] { "列出差异数完成", 0 });
            }
            else
            {
                fn("列出差异数完成");
            }


        }


        /// <summary>
        /// 打开Excel，保存到IWorkbook
        /// </summary>
        /// <param name="workbook"></param>
        private IWorkbook OpenExcel(Label lb)
        {
            IWorkbook workbook = null;
            OpenFileDialog OFDialog = new OpenFileDialog();
            OFDialog.Filter = "Excel文件(*.xls,*.xlsx)|*.xls;*.xlsx";
            OFDialog.Title = "选择一个Excel文件";
            OFDialog.RestoreDirectory = true;

            if (OFDialog.ShowDialog() == DialogResult.OK && OFDialog.FileName != null)
            {
                FileStream Fs = null;
                lb.Text = Path.GetFileName(OFDialog.FileName);
                lb.Tag = OFDialog.FileName;

                try
                {
                    Fs = new FileStream(OFDialog.FileName, FileMode.Open, FileAccess.Read);

                    if (OFDialog.FileName.IndexOf(".xlsx") > 0)
                    {
                        workbook = new XSSFWorkbook(Fs);
                    }
                    else if (OFDialog.FileName.IndexOf(".xls") > 0)
                    {
                        workbook = new HSSFWorkbook(Fs);
                    }
                    if (workbook.NumberOfSheets == 0)
                    {
                        throw new Exception("该excel没有工作表");
                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (Fs != null)
                    {
                        Fs.Close();
                    }
                }

            }

            return workbook;
        }


        /// <summary>
        /// 打开多个Excel，保存到IWorkbook
        /// </summary>
        /// <param name="workbook"></param>
        private void OpenExcels(Label lb)
        {
            IWorkbook workbook = null;
            OpenFileDialog OFDialog = new OpenFileDialog();
            OFDialog.Filter = "Excel文件(*.xls)|*.xls|Excel文件(*.xlsx)|*.xlsx";
            OFDialog.Title = "选择Excel文件";
            OFDialog.RestoreDirectory = true;
            OFDialog.Multiselect = true;

            if (OFDialog.ShowDialog() == DialogResult.OK && OFDialog.FileName != null)
            {
                FileStream Fs = null;
                lb.Text = OFDialog.FileName;

                    foreach(string fileName in OFDialog.FileNames)
                    {
                        try
                        {
                            Fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);

                            if (OFDialog.FileName.IndexOf(".xlsx") > 0)
                            {
                                workbook = new XSSFWorkbook(Fs);
                            }
                            else if (OFDialog.FileName.IndexOf(".xls") > 0)
                            {
                                workbook = new HSSFWorkbook(Fs);
                            }
                            if (workbook.NumberOfSheets == 0)
                            {
                                throw new Exception(fileName+"没有工作表");
                            }
                            ProjectIWorkbooks.Add(new KeyValuePair<string, IWorkbook>(Path.GetFileName(fileName),workbook));

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            if (Fs != null)
                            {
                                Fs.Close();
                            }
                        }

                    }
            }
        }


        /// <summary>
        /// 打开出库表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_SelExcel_Click(object sender, EventArgs e)
        {
            DeliveryWorkbook = OpenExcel(LBDeliveryFilePath);
            if (DeliveryWorkbook != null)
            {
                Task t = new Task(() =>
                {
                    HandleDeliveryExcelData(DeliveryWorkbook.GetSheetAt(2),DeliveryWorkbook.GetSheetAt(4)
                        , DeliveryWorkbook.GetSheetAt(3));
                    using (var fs = File.OpenWrite(LBDeliveryFilePath.Tag.ToString()))
                    {
                        DeliveryWorkbook.Write(fs);
                    }
                    
                    DeliveryWorkbook.Close();
                });

                t.Start();

            }

        }


        /// <summary>
        /// 选择明细表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void button1_Click(object sender, EventArgs e)
        {   
            if(MaterialWorkbook == null)
            {
                MessageBox.Show("请先选择平料表","错误",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else
            {
                OpenExcels(LBProjectFilePath);

                if (ProjectIWorkbooks.Count > 0)
                {
                    int Index = 0;
                    List<Task> tasks = new List<Task>();

                    foreach (var item in ProjectIWorkbooks)
                    {
                        //Task t = new Task(() =>
                        //{
                        //    HandleProjectMaterial(MaterialWorkbook.GetSheetAt(0), item.Value.GetSheetAt(0),item.Key, Index++);

                        //});
                        //t.Start();
                        //tasks.Add(t);
                        await Task.Run(() =>
                        {
                            HandleProjectMaterial(MaterialWorkbook.GetSheetAt(0), item.Value.GetSheetAt(0), item.Key, Index++);

                        });
                    }

                    //await Task.WhenAll(tasks);

                    ProjectIWorkbooks.Clear();

                    string[] fileName = LBMaterialFilePath.Tag.ToString().Split('.');
                    using (var fs = File.OpenWrite(fileName[0] + "_OK." + fileName[1]))
                    {
                        MaterialWorkbook.Write(fs);
                    }
                    //MaterialWorkbook.Close();
                }
            }
            
        }


        private void button6_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = OpenExcel(LBContrastFilePath);
            if (workbook != null)
            {
                Task t = new Task(() =>
                {
                    IWorkbook wb = new HSSFWorkbook();
                    HandleContrast(workbook.GetSheetAt(0), wb.CreateSheet());
                    string[] fileName = LBContrastFilePath.Tag.ToString().Split('.');

                    using (var fs = File.OpenWrite(fileName[0] + "_差异数." + fileName[1]))
                    {
                        wb.Write(fs);
                    }
                    wb.Close();
                    workbook.Close();
                });

                t.Start();

            }
        }

        /// <summary>
        /// 选择格式表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if(FormatWorkbook!=null)
            {
                FormatWorkbook.Close();
            }

            FormatWorkbook = OpenExcel(LBFormatFilePath);
        }

        private void SetStatus(String Msg,int Row=0)
        {
            if(Row==0)
            {
                lstStatus.Items.Add(Msg);
            }
            else
            {
                lstStatus.Items.RemoveAt(Row);
                lstStatus.Items.Insert(Row, Msg);
            }
            
        }
        /// <summary>
        /// 合并明细表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (DetailWorkbook != null)
            {
                DetailWorkbook.Close();
            }

            DetailWorkbook = OpenExcel(LBDeliveryStoreFilePath);
                if (DetailWorkbook != null)
                {
                    Task t = new Task(() =>
                    {
                        HandleCombination(DetailWorkbook.GetSheetAt(2));
                        //string[] fileName = LBDeliveryStoreFilePath.Text.Split('.');
                        using (var fs = File.OpenWrite(LBDeliveryStoreFilePath.Tag.ToString()/*fileName[0].Replace("OK", "结算书")+ fileName[1]*/))
                        {
                            DetailWorkbook.Write(fs);
                        }

                        DetailWorkbook.Close();
                    });

                    t.Start();

                }
            
        }



        private void button4_Click(object sender, EventArgs e)
        {
            if (MaterialWorkbook != null)
            {
                MaterialWorkbook.Close();
            }

            MaterialWorkbook = OpenExcel(LBMaterialFilePath);
        }

        private string GetRowCodeByRowIndex(int RowIndex)
        {
            string table = "0ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string strRowIndex = RowIndex.ToString();
            StringBuilder RowCode = new StringBuilder();
            foreach(var r in strRowIndex)
            {
                RowCode.Append(table[int.Parse(r.ToString())]);
            }
            return RowCode.ToString();
        }

        private void BtnDetailExcel_Click(object sender, EventArgs e)
        {
            if (DetailWorkbook != null)
            {
                DetailWorkbook.Close();
            }
            if (FormatWorkbook == null)
            {
                MessageBox.Show("请先选择格式表", "错误", MessageBoxButtons.OK);
            }
            else
            {
                DetailWorkbook = OpenExcel(LBDeliveryStoreFilePath);
                if (DetailWorkbook != null)
                {
                    Task t = new Task(() =>
                    {
                        HandleDeliveryStoreExcelData(DetailWorkbook.GetSheetAt(0), FormatWorkbook.GetSheetAt(2));
                        string fileName = Path.GetDirectoryName(LBFormatFilePath.Tag.ToString())+"\\"+Path.GetFileNameWithoutExtension(LBDeliveryStoreFilePath.Text);

                        using (var fs = File.OpenWrite(fileName + "结算书" + Path.GetExtension(LBDeliveryStoreFilePath.Text)))
                        {
                            FormatWorkbook.Write(fs);
                        }

                        FormatWorkbook.Close();
                        DetailWorkbook.Close();
                    });

                    t.Start();

                }
            }
        }

        //格式表拖拽功能打开
        private void button2_DragDrop(object sender, DragEventArgs e)
        {
            if (MaterialWorkbook != null)
            {
                MaterialWorkbook.Close();
            }
            string FileName = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            Stream Fs = null;
            try
            {
                Fs = new FileStream(FileName, FileMode.Open, FileAccess.Read);

                if (FileName.IndexOf(".xlsx") > 0)
                {
                    MaterialWorkbook = new XSSFWorkbook(Fs);
                }
                else if (FileName.IndexOf(".xls") > 0)
                {
                    MaterialWorkbook = new HSSFWorkbook(Fs);
                }
                if (MaterialWorkbook.NumberOfSheets == 0)
                {
                    throw new Exception("该excel没有工作表");
                }

                LBMaterialFilePath.Text = FileName;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (Fs != null)
                {
                    Fs.Close();
                }
            }
        }

        private void button2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        //明细表拖拽功能打开
        private void button1_DragDrop(object sender, DragEventArgs e)
        {
            if (DetailWorkbook != null)
            {
                DetailWorkbook.Close();
            }
            if (FormatWorkbook == null)
            {
                MessageBox.Show("请先选择格式表", "错误", MessageBoxButtons.OK);
            }
            else
            {
                string FileName = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

                Stream Fs = null;
                try
                {
                    Fs = new FileStream(FileName, FileMode.Open, FileAccess.Read);

                    if (FileName.IndexOf(".xlsx") > 0)
                    {
                        DetailWorkbook = new XSSFWorkbook(Fs);
                    }
                    else if (FileName.IndexOf(".xls") > 0)
                    {
                        DetailWorkbook = new HSSFWorkbook(Fs);
                    }
                    if (DetailWorkbook.NumberOfSheets == 0)
                    {
                        throw new Exception("该excel没有工作表");
                    }

                    LBDeliveryStoreFilePath.Text = FileName;

                    if (DetailWorkbook != null)
                    {
                        Task t = new Task(() =>
                        {
                            HandleDeliveryStoreExcelData(DetailWorkbook.GetSheetAt(0), FormatWorkbook.GetSheetAt(2));
                            string fileName = Path.GetDirectoryName(LBFormatFilePath.Text) + "\\" + Path.GetFileNameWithoutExtension(LBDeliveryStoreFilePath.Text);

                            using (var fs = File.OpenWrite(fileName + "_OK" + Path.GetExtension(LBDeliveryStoreFilePath.Text)))
                            {
                                FormatWorkbook.Write(fs);
                            }

                            FormatWorkbook.Close();
                            DetailWorkbook.Close();
                        });

                        t.Start();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (Fs != null)
                    {
                        Fs.Close();
                    }
                }
            }
            
        }

        private void button3_DragDrop(object sender, DragEventArgs e)
        {
            if (DetailWorkbook != null)
            {
                DetailWorkbook.Close();
            }
                string FileName = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

                Stream Fs = null;
                try
                {
                    Fs = new FileStream(FileName, FileMode.Open, FileAccess.Read);

                    if (FileName.IndexOf(".xlsx") > 0)
                    {
                        DetailWorkbook = new XSSFWorkbook(Fs);
                    }
                    else if (FileName.IndexOf(".xls") > 0)
                    {
                        DetailWorkbook = new HSSFWorkbook(Fs);
                    }
                    if (DetailWorkbook.NumberOfSheets == 0)
                    {
                        throw new Exception("该excel没有工作表");
                    }

                    LBDeliveryStoreFilePath.Text = FileName;


                    if (DetailWorkbook != null)
                    {
                        Task t = new Task(() =>
                        {
                            HandleCombination(DetailWorkbook.GetSheetAt(2));
                            string[] fileName = LBDeliveryStoreFilePath.Text.Split('.');
                            using (var fs = File.OpenWrite(fileName[0].Replace("OK", "结算书.") + fileName[1]))
                            {
                                DetailWorkbook.Write(fs);
                            }

                            DetailWorkbook.Close();
                        });

                        t.Start();

                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (Fs != null)
                    {
                        Fs.Close();
                    }
                }
         }

        private void Btn_SelDeliveryExcel_DragDrop(object sender, DragEventArgs e)
        {

            if (DeliveryWorkbook != null)
            {
                DeliveryWorkbook.Close();
            }

                string FileName = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

                Stream Fs = null;
                try
                {
                    Fs = new FileStream(FileName, FileMode.Open, FileAccess.Read);

                    if (FileName.IndexOf(".xlsx") > 0)
                    {
                        DeliveryWorkbook = new XSSFWorkbook(Fs);
                    }
                    else if (FileName.IndexOf(".xls") > 0)
                    {
                        DeliveryWorkbook = new HSSFWorkbook(Fs);
                    }
                    if (DeliveryWorkbook.NumberOfSheets == 0)
                    {
                        throw new Exception("该excel没有工作表");
                    }

                    LBDeliveryFilePath.Text = FileName;



                    if (DeliveryWorkbook != null)
                    {
                        Task t = new Task(() =>
                        {
                            HandleDeliveryExcelData(DeliveryWorkbook.GetSheetAt(2), DeliveryWorkbook.GetSheetAt(4)
                                , DeliveryWorkbook.GetSheetAt(3));
                            using (var fs = File.OpenWrite(LBDeliveryFilePath.Text))
                            {
                                DeliveryWorkbook.Write(fs);
                            }

                            DeliveryWorkbook.Close();
                        });

                        t.Start();

                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (Fs != null)
                    {
                        Fs.Close();
                    }
                }
            




        }

        private async void button5_DragDrop(object sender, DragEventArgs e)
        {

            if (MaterialWorkbook == null)
            {
                MessageBox.Show("请先选择平料表", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                FileStream Fs = null;
                string FileName = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

                LBProjectFilePath.Text = FileName;

                IWorkbook workbook = null;

                    try
                    {
                        Fs = new FileStream(FileName, FileMode.Open, FileAccess.Read);

                        if (FileName.IndexOf(".xlsx") > 0)
                        {
                            workbook = new XSSFWorkbook(Fs);
                        }
                        else if (FileName.IndexOf(".xls") > 0)
                        {
                            workbook = new HSSFWorkbook(Fs);
                        }
                        if (workbook.NumberOfSheets == 0)
                        {
                            throw new Exception(FileName + "没有工作表");
                        }
                        ProjectIWorkbooks.Add(new KeyValuePair<string, IWorkbook>(Path.GetFileName(FileName), workbook));

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        if (Fs != null)
                        {
                            Fs.Close();
                        }
                    }

                

                if (ProjectIWorkbooks.Count > 0)
                {
                    int Index = 0;
                    List<Task> tasks = new List<Task>();

                    foreach (var item in ProjectIWorkbooks)
                    {
                        //Task t = new Task(() =>
                        //{
                        //    HandleProjectMaterial(MaterialWorkbook.GetSheetAt(0), item.Value.GetSheetAt(0),item.Key, Index++);

                        //});
                        //t.Start();
                        //tasks.Add(t);
                        await Task.Run(() =>
                        {
                            HandleProjectMaterial(MaterialWorkbook.GetSheetAt(0), item.Value.GetSheetAt(0), item.Key, Index++);

                        });
                    }

                    //await Task.WhenAll(tasks);

                    ProjectIWorkbooks.Clear();

                    string[] fileName = LBMaterialFilePath.Text.Split('.');
                    using (var fs = File.OpenWrite(fileName[0] + "_OK." + fileName[1]))
                    {
                        MaterialWorkbook.Write(fs);
                    }
                    //MaterialWorkbook.Close();
                }
            }
        }

        private void button6_DragDrop(object sender, DragEventArgs e)
        {
            IWorkbook workbook = OpenExcel(LBContrastFilePath);
            if (workbook != null)
            {
                Task t = new Task(() =>
                {
                    IWorkbook wb = new HSSFWorkbook();
                    HandleContrast(workbook.GetSheetAt(0), wb.CreateSheet());
                    string[] fileName = LBContrastFilePath.Text.Split('.');

                    using (var fs = File.OpenWrite(fileName[0] + "_差异数." + fileName[1]))
                    {
                        wb.Write(fs);
                    }
                    wb.Close();
                    workbook.Close();
                });

                t.Start();

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //private void button7_Click(object sender, EventArgs e)
        //{
        //    if (CalculationWorkBook != null)
        //    {
        //        CalculationWorkBook.Close();
        //    }
        //    if (Format2Workbook == null)
        //    {
        //        MessageBox.Show("请先选择格式表", "错误", MessageBoxButtons.OK);
        //    }
        //    else
        //    {
        //        CalculationWorkBook = OpenExcel(LBCalcFilePath);
        //        if (CalculationWorkBook != null)
        //        {
        //            Task t = new Task(() =>
        //            {
        //                HandleReplaceExcelData(CalculationWorkBook.GetSheetAt(1), Format2Workbook.GetSheetAt(1));
        //                string fileName = Path.GetDirectoryName(LBFormat2FilePath.Text) + "\\" + Path.GetFileNameWithoutExtension(LBCalcFilePath.Text);

        //                using (var fs = File.OpenWrite(fileName + "结算书_OK" + Path.GetExtension(LBFormat2FilePath.Text)))
        //                {
        //                    Format2Workbook.Write(fs);
        //                }

        //                Format2Workbook.Close();
        //                CalculationWorkBook.Close();
        //            });

        //            t.Start();

        //        }
        //    }
        //}

        //private void button8_Click(object sender, EventArgs e)
        //{
        //    if (Format2Workbook != null)
        //    {
        //        Format2Workbook.Close();
        //    }

        //    Format2Workbook = OpenExcel(LBFormat2FilePath);
        //}

        private void button9_Click(object sender, EventArgs e)
        {
            if (ReplaceWorkBook != null)
            {
                ReplaceWorkBook.Close();
            }

            ReplaceWorkBook = OpenExcel(LBReplaceFilePath);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (ToReplaceWorkBook != null)
            {
                ToReplaceWorkBook.Close();
            }
            if (ReplaceWorkBook == null)
            {
                MessageBox.Show("请先选择替换Excel", "错误", MessageBoxButtons.OK);
            }
            else
            {
                ToReplaceWorkBook = OpenExcel(LBToReplaceFilePath);
                if (ToReplaceWorkBook != null)
                {
                    Task t = new Task(() =>
                    {
                        HandleReplaceExcelData(ToReplaceWorkBook.GetSheetAt(1), ReplaceWorkBook.GetSheetAt(0));
                        string fileName = LBToReplaceFilePath.Tag.ToString();

                        using (var fs = File.OpenWrite(fileName))
                        {
                            ToReplaceWorkBook.Write(fs);
                        }

                        ReplaceWorkBook.Close();
                        ToReplaceWorkBook.Close();
                    });

                    t.Start();

                }
            }



        }
    }
}
