using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using MahApps.Metro.Controls;
using NAudio.Wave;
using WORLD.NET;
using NAudio.CoreAudioApi;
using System.IO;

namespace VoiceChangerTest
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : MetroWindow
    {

        public bool play { get; set; }
        public List<string> inputDeviceList { get; set; }
        public List<string> outputDeviceList { get; set; }
        public int inputDeviceNumber { get; set; }
        public int outputDeviceNumber { get; set; }
        public bool outputIsSelectable { get; set; } = false;
        public float prate { get; set; } = 1.8f;
        public double srate { get; set; } = 1.3;

        public MainWindow()
        {
            InitializeComponent();
            GetInputDevice();
            GetOutputDevice();
            WorldConfig.D4C_threshold = 0;
            WorldConfig.f0_floor = 60;
            WorldConfig.f0_ceil = 600;
            this.DataContext = this;
        }

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopRecording();
        }


        private void GetDeviceButton_Click(object sender, RoutedEventArgs e)
        {
            GetInputDevice();
            GetOutputDevice();
        }

        private void GetInputDevice()
        {
            List<string> deviceList = new List<string>();
            for (int i = 0; i < WaveIn.DeviceCount; i++)
            {
                var capabilities = WaveIn.GetCapabilities(i);
                deviceList.Add(capabilities.ProductName);
            }
            inputDeviceList = deviceList;
        }
        private void GetOutputDevice()
        {
            List<string> deviceList = new List<string>();
            for (int i = 0; i < WaveOut.DeviceCount; i++)
            {
                var capabilities = WaveOut.GetCapabilities(i);
                deviceList.Add(capabilities.ProductName);
            }
            outputDeviceList = deviceList;
        }

        private void RecordingButton_Click(object sender, RoutedEventArgs e)
        {
            if (play)
            {
                StartRecording();
            }
            else
            {
                StopRecording();
            }
        }


        WaveInEvent waveIn;
        WorldConverter converter = new WorldConverter(0.005);
        readonly WaveFormat waveFormat = new WaveFormat(44100, 16, 1);
        IWavePlayer wavPlayer;
        ConcurrentQueue<WP[]> buffer;
        WaveFileWriter waveWriter_In;
        WaveFileWriter waveWriter_Out;

        private void StartRecording()
        {

            buffer = new ConcurrentQueue<WP[]>();

            // マイク側の準備
            string filePath_In = Path.Combine(Environment.CurrentDirectory, "マイク.wav");
            waveWriter_In = new WaveFileWriter(filePath_In, waveFormat);
            waveIn = new WaveInEvent()
            {
                DeviceNumber = inputDeviceNumber,
                WaveFormat = waveFormat,
            };
            waveIn.DataAvailable += ReceiveVoice;
            waveIn.RecordingStopped += async (_, __) =>
            {
                // 録音が止まったときの処理
            };
            waveIn.StartRecording();

            // スピーカ側の準備
            string filePath_Out = Path.Combine(Environment.CurrentDirectory, "スピーカ.wav");
            waveWriter_Out = new WaveFileWriter(filePath_Out, waveFormat);
            BufferedWaveProvider provider = new BufferedWaveProvider(new WaveFormat(WorldConfig.fs, 16, 1)); //16bit１チャンネルの音源を想定
            if (outputIsSelectable)
            {
                wavPlayer = new WaveOut()
                {
                    DeviceNumber = outputDeviceNumber,
                };
            }
            else
            {
                MMDevice mmDevice = new MMDeviceEnumerator()
                    .GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);  //再生デバイスと出力先を設定
                wavPlayer = new WasapiOut(mmDevice, AudioClientShareMode.Shared, false, 100);
            }
            wavPlayer.Init(provider);
            wavPlayer.Play();
            Task t = SendVoice(provider);
        }

        private void StopRecording()
        {
            waveIn?.StopRecording();
            waveIn?.Dispose();
            waveIn = null;
            wavPlayer?.Stop();
            wavPlayer?.Dispose();
            wavPlayer = null;
            waveWriter_In?.Close();
            waveWriter_In = null;
            waveWriter_Out?.Close();
            waveWriter_Out = null;
        }

        private void ReceiveVoice(object obj, WaveInEventArgs ee)
        {
            try
            {
                waveWriter_In.Write(ee.Buffer, 0, ee.Buffer.Length);
                waveWriter_In.Flush();

                WP[] wp = null;
                float[] signal = ByteToFloat(ee.Buffer);

                converter.SignalToParameter.AddSignal(signal);
                converter.SignalToParameter.Analyze();    // この行はコメントアウトできます。
                wp = converter.SignalToParameter.ReadParameter();

                for (int i = 0; i < wp.Length; i++)
                {
                    wp[i].f0 *= prate;
                    float[] temp = new float[wp[i].spectrogram.Length];
                    for (int j = 0; j < wp[i].spectrogram.Length; j++)
                    {
                        temp[j] = wp[i].spectrogram[(int)Math.Min(j / srate, wp[i].spectrogram.Length-1)];
                    }
                    wp[i].spectrogram = temp;
                }

                buffer.Enqueue(wp);

            }
            catch (Exception e)
            {
                ErrorLog("音声分析時に失敗しました。",e);
            }
        }

        async Task SendVoice(BufferedWaveProvider provider)
        {
            try
            {
                while (play) //再生可能な間ずっと
                {
                    if(buffer.Count > 0)
                    {
                        WP[] wp;
                        buffer.TryDequeue(out wp);
                        if (wp == null)
                        {
                            continue;
                        }
                        converter.ParameterToSignal.AddParameter(wp);
                        converter.ParameterToSignal.Synthsize();    // この行はコメントアウトできます。
                        float[] Fbuffer = converter.ParameterToSignal.ReadSignal();
                        byte[] Bbuffer = FloatToByte(Fbuffer);
                        provider.AddSamples(Bbuffer, 0, Bbuffer.Length); // バッファーを渡す

                        waveWriter_Out.Write(Bbuffer, 0, Bbuffer.Length);
                        waveWriter_Out.Flush();
                    }
                    else
                    {
                        await Task.Delay(30);
                    }
                }
            }
            catch(Exception e)
            {
                ErrorLog("音声合成時に失敗しました。", e);
            }
        }

        /// <summary>
        /// byteで表現された音をfloatに直す
        /// 16bitを想定
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        private float[] ByteToFloat(byte[] bytes)
        {
            float[] values = new float[bytes.Length/2];
            float pow = (float)Math.Pow(2, 15);
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = BitConverter.ToInt16(bytes, 2 * i) / pow; // -1～1に正規化
            }
            return values;
        }
        public static byte[] FloatToByte(float[] Fbuffer)
        {
            byte[] temp;
            byte[] Bbuffer = new byte[2 * Fbuffer.Length];
            for (int i = 0; i < Fbuffer.Length; i++)
            {
                temp = BitConverter.GetBytes((Int16)(Fbuffer[i] * Math.Pow(2, 15)));  // byteデータに変換
                Array.Copy(temp, 0, Bbuffer, 2 * i, 2);
            }
            return Bbuffer;
        }


        
        private void ErrorLog(string s, Exception e)
        {
            Log(s);
            Log(e.ToString());
            Log("[message] \r\n" + e.Message);
            Log("[source] \r\n" + e.Source);
            Log("[stacktrace] \r\n" + e.StackTrace);
        }
        
        private void Log(string s)
        {
            try
            {
                string logPath = Path.Combine(Directory.GetCurrentDirectory(), "log");
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
                DirectoryInfo dirInfo = new DirectoryInfo(logPath);
                // 読み取り専用か確認する
                if ((dirInfo.Attributes & FileAttributes.ReadOnly) > 0)
                {
                    // 読み取り専用を解除する
                    dirInfo.Attributes = dirInfo.Attributes & ~FileAttributes.ReadOnly;
                }
                string date = DateTime.Now.ToString().Split(' ')[0];
                date = date.Replace("/", "_");
                logPath = Path.Combine(logPath, date + ".txt");
                if (!File.Exists(logPath))
                {
                    File.Create(logPath);
                }
                {
                    FileAttributes fa = File.GetAttributes(logPath);
                    // 読み取り専用属性を削除（他の属性は変更しない）
                    fa = fa & ~FileAttributes.ReadOnly;
                    File.SetAttributes(logPath, fa);
                }
                using (StreamWriter sw = new StreamWriter(logPath, true, Encoding.GetEncoding("shift_jis")))
                {
                    sw.WriteLine(DateTime.Now.ToString("yyyyMMdd:hh:mm:ss.fff") + " " + s);
                }
            }
            catch
            {

            }
        }

    }
}
