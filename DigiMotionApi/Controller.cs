using System;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;

namespace DigiMotionApi.Controller
{
    /// <summary>
    /// 串口数据位列表
    /// </summary>
    public enum SerialPortDatabits
    {
        FiveBits = 5,
        SixBits = 6,
        SeventBits = 7,
        EightBits = 8
    }

    /// <summary>
    /// 串口波特率列表。
    /// </summary>
    public enum SerialPortBaudRates
    {
        BaudRate_75 = 75,
        BaudRate_110 = 110,
        BaudRate_150 = 150,
        BaudRate_300 = 300,
        BaudRate_600 = 600,
        BaudRate_1200 = 1200,
        BaudRate_2400 = 2400,
        BaudRate_4800 = 4800,
        BaudRate_9600 = 9600,
        BaudRate_14400 = 14400,
        BaudRate_19200 = 19200,
        BaudRate_28800 = 28800,
        BaudRate_38400 = 38400,
        BaudRate_56000 = 56000,
        BaudRate_57600 = 57600,
        BaudRate_115200 = 115200,
        BaudRate_128000 = 128000,
        BaudRate_230400 = 230400,
        BaudRate_256000 = 256000
    }

    /// <summary>
    /// 串口辅助类
    /// </summary>
    public class Controller
    {
        /// <summary>
        /// 接收事件是否有效 false表示有效
        /// </summary>
        public bool ReceiveEventFlag = false;

        /// <summary>
        /// 条码读取触发委托
        /// </summary>
        /// <param name="barcode"></param>
        public delegate void DelegateReadBarcode(string barcode);

        /// <summary>
        /// 条码读取触发事件
        /// </summary>
        public event DelegateReadBarcode OnReadBarcode;

        #region 变量属性

        private string _portName; //串口号
        private SerialPortBaudRates _baudRate; //波特率
        private Parity _parity; //校验位
        private StopBits _stopBits; //停止位
        private SerialPortDatabits _dataBits; //数据位

        private readonly SerialPort _port = new SerialPort();

        /// <summary>
        /// 串口号
        /// </summary>
        public string PortName
        {
            get { return _portName; }
            set { _portName = value; }
        }
        /// <summary>
        /// 扫描枪端口
        /// </summary>
        public SerialPort Port
        {
            get { return _port; }
        }

        /// <summary>
        /// 波特率
        /// </summary>
        public SerialPortBaudRates BaudRate
        {
            get { return _baudRate; }
            set { _baudRate = value; }
        }

        /// <summary>
        /// 奇偶校验位
        /// </summary>
        public Parity Parity
        {
            get { return _parity; }
            set { _parity = value; }
        }

        /// <summary>
        /// 数据位
        /// </summary>
        public SerialPortDatabits DataBits
        {
            get { return _dataBits; }
            set { _dataBits = value; }
        }

        /// <summary>
        /// 停止位
        /// </summary>
        public StopBits StopBits
        {
            get { return _stopBits; }
            set { _stopBits = value; }
        }

        #endregion

        #region 构造函数

        /// <summary>
        /// 默认构造函数
        /// </summary>
        public Controller()
        {
            _portName = "COM4";
            _baudRate = SerialPortBaudRates.BaudRate_115200;
            _parity = Parity.None;
            _dataBits = SerialPortDatabits.EightBits;
            _stopBits = StopBits.One;

            _port.DataReceived += DataReceived;
            _port.ErrorReceived += ErrorReceived;
        }

        /// <summary>
        /// 参数构造函数（使用枚举参数构造）
        /// </summary>
        /// <param name="baud">波特率</param>
        /// <param name="par">奇偶校验位</param>
        /// <param name="sBits">停止位</param>
        /// <param name="dBits">数据位</param>
        /// <param name="name">串口号</param>
        public Controller(string name, SerialPortBaudRates baud, Parity par, SerialPortDatabits dBits, StopBits sBits)
        {
            _portName = name;
            _baudRate = baud;
            _parity = par;
            _dataBits = dBits;
            _stopBits = sBits;

            _port.DataReceived += DataReceived;
            _port.ErrorReceived += ErrorReceived;
        }

        /// <summary>
        /// 参数构造函数（使用字符串参数构造）
        /// </summary>
        /// <param name="baud">波特率</param>
        /// <param name="par">奇偶校验位</param>
        /// <param name="sBits">停止位</param>
        /// <param name="dBits">数据位</param>
        /// <param name="name">串口号</param>
        public Controller(string name, string baud, string par, string dBits, string sBits)
        {
            _portName = name;
            _baudRate = (SerialPortBaudRates)Enum.Parse(typeof(SerialPortBaudRates), baud);
            _parity = (Parity)Enum.Parse(typeof(Parity), par);
            _dataBits = (SerialPortDatabits)Enum.Parse(typeof(SerialPortDatabits), dBits);
            _stopBits = (StopBits)Enum.Parse(typeof(StopBits), sBits);

            _port.DataReceived += DataReceived;
            _port.ErrorReceived += ErrorReceived;
        }

        #endregion

        /// <summary>
        /// 端口是否已经打开
        /// </summary>
        public bool IsOpen
        {
            get { return _port.IsOpen; }
        }

        /// <summary>
        /// 打开端口
        /// </summary>
        /// <returns></returns>
        //public void OpenPort()
        //{
        //    if (_port.IsOpen) _port.Close();

        //    _port.PortName = _portName;
        //    _port.BaudRate = (int)_baudRate;
        //    _port.Parity = _parity;
        //    _port.DataBits = (int)_dataBits;
        //    _port.StopBits = _stopBits;

        //    //_port.ReceivedBytesThreshold = 8;
        //    //_port.RtsEnable = true;
        //    _port.ReadTimeout = 500;
        //    _port.WriteTimeout = 500;
        //    _port.ReadBufferSize = 256;
        //    _port.WriteBufferSize = 256;
        //    _port.RtsEnable = true;
        //    _port.Open();
        //}
        public bool OpenPort()
        {
            try
            {
                if (_port.IsOpen) _port.Close();
                _port.PortName = _portName;
                _port.BaudRate = (int)_baudRate;
                _port.Parity = _parity;
                _port.DataBits = (int)_dataBits;
                _port.StopBits = _stopBits;
                //_port.ReceivedBytesThreshold = 8;
                //_port.RtsEnable = true;
                _port.ReadTimeout = 500;
                _port.WriteTimeout = 500;
                _port.ReadBufferSize = 256;
                _port.WriteBufferSize = 256;
                _port.RtsEnable = true;
                _port.Open();
            }
            catch
            {
                return false;
            }
            return true;
        }


        /// <summary>
        /// 关闭端口
        /// </summary>
        public void ClosePort()
        {
            DiscardBuffer();
            if (_port.IsOpen) _port.Close();
        }

        /// <summary>
        /// 丢弃来自串行驱动程序的接收和发送缓冲区的数据
        /// </summary>
        public void DiscardBuffer()
        {
            _port.DiscardInBuffer();
            _port.DiscardOutBuffer();
        }

        /// <summary>
        /// 数据接收处理
        /// </summary>
        private void DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //禁止接收事件时直接退出
            if (ReceiveEventFlag) return;

            //延时以保证接收数据完整
            Thread.Sleep(50);

            //获取字节长度
            int bytes = Port.BytesToRead;

            //读取缓冲区的数据到数组
            byte[] receiveBuff = new byte[bytes];
            Port.Read(receiveBuff, 0, bytes);

            ////触发整条记录的处理
            //if (DataReceived != null)
            //{
            //    DataReceived(new DataReceivedEventArgs(receiveBuff));
            //}

            string barcode = Encoding.Default.GetString(receiveBuff);
            //Code39起始/结束符去除
            barcode = barcode.Replace("*", string.Empty);
            if (OnReadBarcode == null)
            {
                //If There is No barcode Message Break Operation
            }
            try
            {
                OnReadBarcode(barcode);
            }
            catch { }
        }

        /// <summary>
        /// 错误处理函数
        /// </summary>
        private void ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            //if (Error != null)
            //{
            //    Error(sender, e);
            //}
        }

        #region 数据写入操作

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <param name="msg"></param>
        public void WriteData(string msg)
        {
            if (!(_port.IsOpen)) _port.Open();

            _port.DiscardInBuffer();
            _port.Write(msg);
        }

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <param name="msg">包含要写入端口的字节数组</param>
        /// <param name="offset">参数从0字节开始的字节偏移量</param>
        /// <param name="count">要写入的字节数</param>
        public void WriteData(byte[] msg, int offset, int count)
        {
            if (!(_port.IsOpen)) _port.Open();

            _port.DiscardInBuffer();
            _port.Write(msg, offset, count);
        }

        /// <summary>
        /// 发送串口指令
        /// </summary>
        /// <param name="sendData">发送数据</param>
        /// <param name="receiveData">接收数据</param>
        /// <param name="overtime">重复次数</param>
        /// <returns></returns>
        public int SendCommand(byte[] sendData, ref byte[] receiveData, int overtime)
        {
            if (!(_port.IsOpen)) _port.Open();

            //关闭接收事件
            ReceiveEventFlag = true;

            //清空接收缓冲区
            _port.DiscardInBuffer();
            _port.Write(sendData, 0, sendData.Length);

            int num = 0, ret = 0;
            while (num++ < overtime)
            {
                if (_port.BytesToRead >= receiveData.Length) break;
                Thread.Sleep(1);
            }

            if (_port.BytesToRead >= receiveData.Length)
            {
                ret = _port.Read(receiveData, 0, receiveData.Length);
            }

            //打开事件
            ReceiveEventFlag = false;
            return ret;
        }

        /// <summary>
        /// 发送串口指令
        /// </summary>
        /// <param name="sendData">发送数据</param>
        /// <returns></returns>
        public bool SendCommand(byte[] sendData)
        {
            try
            {
                if (!(_port.IsOpen)) _port.Open();

                //清空接收缓冲区
                _port.DiscardInBuffer();

                //发送指令
                _port.Write(sendData, 0, sendData.Length);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        /// <summary>
        /// 检查端口名称是否存在
        /// </summary>
        /// <param name="portName"></param>
        /// <returns></returns>
        public static bool Exists(string portName)
        {
            return SerialPort.GetPortNames().Any(port => port == portName);
        }

        /// <summary>
        /// 格式化端口相关属性
        /// </summary>
        /// <param name="port"></param>
        /// <returns></returns>
        public static string Format(SerialPort port)
        {
            return String.Format("{0} ({1},{2},{3},{4},{5})",
                port.PortName, port.BaudRate, port.DataBits, port.StopBits, port.Parity, port.Handshake);
        }
    }

    public class DataReceivedEventArgs : EventArgs
    {
        public byte[] DataReceived;

        public DataReceivedEventArgs(byte[] receivedData)
        {
            DataReceived = receivedData;
        }
    }
}
