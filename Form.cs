using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Diagnostics;
using static WDF2CSV.WDFAPI;

namespace WDF2CSV
{
    public partial class Form : System.Windows.Forms.Form
    {
        private const string WDF_EXTENSION_FORMAT    = "WDF(*.WDF)|*.WDF";
        private const string WDF_EXTENSION           = "*.WDF";
        private const string DLL_EXTENSION           = ".dll";
        private const string WDF_DATA_FORMAT_TYPE    = "%WDF";
        private const string ADD_FOLDERPATH = "DLL/";
        private const string OUTPUT_FOLDERPATH = "exported";


        private const string FAILED_TO_FIND_MODEL    = "Failed to find model's DLL";

        string[] modelType = { "DLM3000", "DLM5000HD", "DLM5000" };

        private WDFAPI wdfAPI;

        public Form()
        {
            InitializeComponent();

            LogTextBox.AppendText("Supported model : DLM Series only.." + System.Environment.NewLine);
        }

        private void onOpenbuttonClick(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = WDF_EXTENSION_FORMAT;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    if (CheckFileFormat(ofd.FileName))
                    {
                        convertWaveToCSV(ofd.FileName);
                        MessageBox.Show("Done!", "Parser");
                    }
                    else
                    {
                        MessageBox.Show("Wrong File Format!", "Parser");
                    }

                }
            }
        }

        private bool CheckFileFormat(string filePath)
        {
            bool res = false;
            string str = "";

            using (StreamReader sr = new StreamReader(filePath))
            {
                str = sr.ReadLine();
            }

            if (str.Contains(WDF_DATA_FORMAT_TYPE))
            {
                res = true;
            }

            return res;
        }

        private void onCloseButtonClick(object sender, EventArgs e)
        {
            Close();
        }


        private string searchProperDLL(string filePath, string dllSearchPath)
        {
            string dllPath = "";
            string dllName = "";
            string str = "";

            using (StreamReader sr = new StreamReader(filePath))
            {
                str = sr.ReadLine();
            }

            foreach (string model in modelType)
            {
                if (str.Contains(model))
                {
                    dllName = model + DLL_EXTENSION;
                    break;
                }
            }
            if (dllName == string.Empty)
            {
                MessageBox.Show(FAILED_TO_FIND_MODEL, "parser");
            }

            dllPath = Path.Combine(dllSearchPath, dllName);

            return dllPath;
        }

        private void loadDLL(string filePath)
        {
            string dllSearchPath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "DLL");

            string dllPath = searchProperDLL(filePath, dllSearchPath);

            try
            {
                wdfAPI = new WDFAPI();
                wdfAPI.tryLoadDLL(dllPath);
            }
            catch (DllNotFoundException e)
            {
                Debug.WriteLine(e.Message + e.StackTrace.ToString());
                MessageBox.Show("failed to find .dll file, Locate dll under /DLL");
                this.Close();
            }
        }

        private void convertWaveToCSV(string filePath)
        {
            int handle;

            loadDLL(filePath);


            WDFAPI.WDFAccessParam param = new WDFAPI.WDFAccessParam();

            int result = wdfAPI.openFileEx(out handle, filePath);

            uint traceNumber;
            wdfAPI.getTraceNumber(handle, out traceNumber);



            for (uint num = 0; num < traceNumber; num++)
            {
                uint blockNumber = 0;
                Double vResolution = 0;
                Double hResolution = 0;
                Double vOffset = 0;
                string vUnit = "";
                Double hOffset = 0;
                string hUnit = "";
                WDFAPI.WDFVDataType vDataType = WDFAPI.WDFVDataType.wdfDataTypeSINT16;

                int blockSize = 0;
                

                string outputCSVname = "";
                StringBuilder tn = new StringBuilder();
                wdfAPI.getTraceName(handle, num, tn);

                Directory.CreateDirectory(OUTPUT_FOLDERPATH);
                
                outputCSVname = Path.Combine(OUTPUT_FOLDERPATH, System.DateTime.Now.ToString("yyyyMMddHHmmss_") + tn.ToString() + ".csv");

                FileStream fs = new FileStream(outputCSVname, FileMode.Create);
                using (System.IO.StreamWriter csvFile = new System.IO.StreamWriter(fs))
                {

                    wdfAPI.getVResolution(handle, num, blockNumber, out vResolution);

                    wdfAPI.getHResolution(handle, num, blockNumber, out hResolution);

                    wdfAPI.getVOffset(handle, num, blockNumber, out vOffset);

                    wdfAPI.getHOffset(handle, num, blockNumber, out hOffset);

                    wdfAPI.getBlockSize(handle, num, blockNumber, out blockSize);

                    wdfAPI.getVDataType(handle, num, blockNumber, out vDataType);


                    StringBuilder sb = new StringBuilder();

                    wdfAPI.getVUnit(handle, num, blockNumber, sb);
                    vUnit = sb.ToString();
                    sb.Clear();

                    wdfAPI.getHUnit(handle, num, blockNumber, sb);
                    hUnit = sb.ToString();
                    sb.Clear();

                    param.version = (uint)WDFAPI.WDFDefaultValue.WDF_DEFAULT_ACSPRM_VERSION;
                    param.trace = num;
                    param.block = 0;
                    param.start = 0;
                    param.count = blockSize;
                    param.ppRate = (int)WDFAPI.WDFDefaultValue.WDF_DEFAULT_ACSPRM_PPRATE;
                    param.waveType = (int)WDFAPI.WDFDefaultValue.WDF_DEFAULT_ACSPRM_WAVETYPE;
                    param.dataType = (int)WDFAPI.WDFDefaultValue.WDF_DEFAULT_ACSPRM_DATATYPE;
                    param.cntOut = 0;
                    param.box = 0;
                    param.compMode = (int)WDFAPI.WDFDefaultValue.WDF_DEFAULT_ACSPRM_COMPMODE;

                    byte[] buff = new byte[blockSize * 4];
                    GCHandle h = GCHandle.Alloc(buff, GCHandleType.Pinned);
                    IntPtr p = Marshal.UnsafeAddrOfPinnedArrayElement(buff, 0);
                    param.dst = p;

                    result = wdfAPI.getScaleWave(handle, out param);

                    List<double> vvalue = new List<double>(blockSize);
                    List<double> hvalue = new List<double>(blockSize);

                    if(vDataType == WDFVDataType.wdfDataTypeLOGIC32 || vDataType == WDFVDataType.wdfDataTypeLOGIC16)
                    {
                        MessageBox.Show("Logic datatype not supported. you may get wrong result");
                        //throw new InvalidDataException("Wrong Datatype..");
                    }

                    GetPhysicalVValue(buff, (int)blockSize, (double)vOffset, vResolution, vDataType, out vvalue);
                    GetPhysicalHValue((int)blockSize, (double)hOffset, hResolution, out hvalue);


                    csvFile.WriteLine("{0},{1}", hUnit, vUnit);

                    for (int i = 0; i < hvalue.Count; i++)
                    {
                        csvFile.WriteLine("{0},{1}", hvalue[i], vvalue[i]);
                    }
                    LogTextBox.AppendText("Exported : " + outputCSVname + System.Environment.NewLine);
                    csvFile.Close();
                }
            }

            wdfAPI.closeFile(out handle);
        }


        private void GetPhysicalVValue(in byte[] data, in int dataPoints, in double vOffset, in double vResolution, in WDFAPI.WDFVDataType vDatatype, out List<double> physicalValue)
        {
            physicalValue = new List<double>(dataPoints);

            int dataCount = 0;

            double? tempData = null;

            for (int i = 0; i < dataPoints; i++)
            {
                switch (vDatatype)
                {
                    case WDFVDataType.wdfDataTypeUINT16:
                        tempData = GetWaveDataU16(data, ref dataCount);
                        break;
                    case WDFVDataType.wdfDataTypeSINT16:
                        tempData = GetWaveData16(data, ref dataCount);
                        break;
                    case WDFVDataType.wdfDataTypeUINT32:
                        tempData = GetWaveDataU32(data, ref dataCount);
                        break;
                    case WDFVDataType.wdfDataTypeSINT32:
                        tempData = GetWaveData32(data, ref dataCount);
                        break;
                    case WDFVDataType.wdfDataTypeFLOAT:
                        tempData = GetWaveData32_float(data, ref dataCount);
                        break;
                }

                if (tempData.HasValue)
                {
                    physicalValue.Add((tempData.Value * vResolution) + vOffset);
                }
            }
        }

        private void GetPhysicalHValue(int dataPoints, double hOffset, double hResolution, out List<double> physicalValue)
        {
            physicalValue = new List<double>(dataPoints);

            for (int i = 0; i < dataPoints; i++)
            {
                double t = (double)i * hResolution + hOffset;
                physicalValue.Add(t);
            }
        }

        private void GetMinMaxData(double[] data, int dataPoints, out double minValue, out double maxValue)
        {
            minValue = 0.0;
            maxValue = 0.0;
            for (int i = 0; i < dataPoints; i++)
            {
                double value = data[i];

                if (value > maxValue)
                {
                    maxValue = value;
                }

                if (value < minValue)
                {
                    minValue = value;
                }
            }
        }

        private void Form_Load(object sender, EventArgs e)
        {

        }

        private int GetWaveData32(byte[] mpBuff, ref int dataCount)
        {
            int data = BitConverter.ToInt32(mpBuff, dataCount);
            dataCount += 4;
            return data;
        }

        private uint GetWaveDataU32(byte[] mpBuff, ref int dataCount)
        {
            uint data = BitConverter.ToUInt32(mpBuff, dataCount);
            dataCount += 4;
            return data;
        }

        private float GetWaveData32_float(byte[] mpBuff, ref int dataCount)
        {
            float data = BitConverter.ToSingle(mpBuff, dataCount);
            dataCount += 4;
            return data;
        }

        private long GetWaveData64(byte[] mpBuff, ref int dataCount)
        {
            long data = BitConverter.ToInt64(mpBuff, dataCount);
            dataCount += 8;
            return data;
        }

        private short GetWaveData16(byte[] mpBuff, ref int dataCount)
        {
            short data = BitConverter.ToInt16(mpBuff, dataCount);
            dataCount += 2;
            return data;
        }

        private ushort GetWaveDataU16(byte[] mpBuff, ref int dataCount)
        {
            ushort data = BitConverter.ToUInt16(mpBuff, dataCount);
            dataCount += 2;
            return data;
        }


    }
}

