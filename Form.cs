using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Diagnostics;

namespace WDF2CSV
{
    public partial class Form : System.Windows.Forms.Form
    {
        private const string WDF_EXTENSION_FORMAT    = "WDF(*.WDF)|*.WDF";
        private const string WDF_EXTENSION           = "*.WDF";
        private const string DLL_EXTENSION           = ".dll";
        private const string WDF_DATA_FORMAT_TYPE    = "%WDF";
        private const string ADD_FOLDERPATH          = "DLL/";

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
                MessageBox.Show(e.Message);
                this.Close();
            }
        }

        private void convertWaveToCSV(string filePath)
        {
            int handle;
            uint blockNumber = 0;
            Double vResolution = 0;
            Double vOffset = 0;
            Double hOffset = 0;
            int blockSize = 0;

            loadDLL(filePath);


            WDFAPI.WDFAccessParam param = new WDFAPI.WDFAccessParam();

            int result = wdfAPI.openFileEx(out handle, filePath);

            uint traceNumber;
            wdfAPI.getTraceNumber(handle, out traceNumber);



            for (uint num = 0; num < traceNumber; num++)
            {
                string outputCSVname = "";
                StringBuilder tn = new StringBuilder();
                wdfAPI.getTraceName(handle,num, tn);

                outputCSVname = System.DateTime.Now.ToString("yyyyMMddHHmmss_")+ tn.ToString() +  ".csv";

                FileStream fs = new FileStream(outputCSVname, FileMode.Create);
                using (System.IO.StreamWriter csvFile = new System.IO.StreamWriter(fs))
                {

                    wdfAPI.getVResolution(handle, num, blockNumber, out vResolution);

                    wdfAPI.getVOffset(handle, num, blockNumber, out vOffset);

                    wdfAPI.getHOffset(handle, num, blockNumber, out hOffset);

                    wdfAPI.getBlockSize(handle, num, blockNumber, out blockSize);

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

                    GetPhysicalVValue(buff, (int)blockSize, (double)vOffset, vResolution, ref vvalue);
                    GetPhysicalHValue((int)blockSize, (double)hOffset, vResolution, ref hvalue);

                    csvFile.WriteLine("{0},{1}", "t","V");

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


        private void GetPhysicalVValue(byte[] data, int dataPoints, double vOffset, double vResolution, ref List<double> physicalValue)
        {
            if(physicalValue == null) physicalValue = new List<double>(dataPoints);
            physicalValue.Clear();

            int dataCount = 0;

            for (int i = 0; i < dataPoints; i++)
            {
                short v = BitConverter.ToInt16(data, dataCount);
                dataCount += 2;
                double t = ((double)v * vResolution) + vOffset;
                physicalValue.Add(t);
            }
        }

        private void GetPhysicalHValue(int dataPoints, double hOffset, double hResolution, ref List<double> physicalValue)
        {
            if (physicalValue == null) physicalValue = new List<double>(dataPoints);
            physicalValue.Clear();

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
    }
}

