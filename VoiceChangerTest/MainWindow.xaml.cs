using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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
        #region ビューにバインドされてる変数
        /// <summary>
        /// 再生ボタンが押されているかどうか
        /// true:押されている　false:押されていない
        /// </summary>
        public bool play { get; set; }
        /// <summary>
        /// マイクのコンボボックスに表示される文字列
        /// </summary>
        public List<string> inputDeviceList { get; set; }
        /// <summary>
        /// スピーカーのコンボボックスに表示される文字列
        /// </summary>
        public List<string> outputDeviceList { get; set; }
        /// <summary>
        /// マイクのコンボボックスで選択されている番号
        /// </summary>
        public int inputDeviceNumber { get; set; }
        /// <summary>
        /// スピーカーのコンボボックスで選択されている番号
        /// </summary>
        public int outputDeviceNumber { get; set; }
        /// <summary>
        /// 「出力先を選択する」にチェックがついているがどうか
        /// true:チェックがついている　false:チェックがついていない
        /// </summary>
        public bool outputIsSelectable { get; set; } = false;
        /// <summary>
        /// 「ピッチ倍率」の値
        /// </summary>
        public float prate { get; set; } = 2.0f;
        /// <summary>
        /// 「フォルマント」の値
        /// </summary>
        public double srate { get; set; } = 1.25;
        #endregion

        /// <summary>
        /// コンストラクタなので初期化と設定をしている
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            GetInputDevice();   // 入力デバイスの取得
            GetOutputDevice();  // 出力デバイスの取得
            WorldConfig.D4C_threshold = 0;  // D4Cの閾値 0～0.3くらいが安定する印象、よくわからない。WORLDでは0.85。
            WorldConfig.f0_floor = 60;  // ピッチ推定の下限値[Hz]。男性の声を想定して60Hz。
            WorldConfig.f0_ceil = 600;  // ピッチ推定の上限値[Hz]。
            this.DataContext = this;    // 自分自身をバインド
        }

        /// <summary>
        /// 画面をとじるときには停止処理もしておく。
        /// 
        /// </summary>
        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopRecording();
        }

        /// <summary>
        /// 「マイク・スピーカーの更新」ボタンが押されたときの処理
        /// </summary>
        private void GetDeviceButton_Click(object sender, RoutedEventArgs e)
        {
            GetInputDevice();
            GetOutputDevice();
        }

        /// <summary>
        /// 入力デバイスの取得
        /// </summary>
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
        /// <summary>
        /// 出力デバイスの取得
        /// </summary>
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

        /// <summary>
        /// 再生・停止してくれそうなトグルボタンが押されたときの処理
        /// </summary>
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



        /*---------------------------------------------------------*
         * ここから下はモデル部分（裏で動いてる信号処理とか）です。
         * 本来であれば別クラスにした方がいいですが、
         * 可読性重視？のため全てひとまとめにしてあります。
         *---------------------------------------------------------*/



        #region プライベートな変数
        /// <summary>
        /// マイクからの入力を処理してくれるやつ
        /// </summary>
        WaveInEvent waveIn;
        /// <summary>
        /// スピーカーなどの出力の処理をしてくれるやつ
        /// </summary>
        IWavePlayer wavPlayer;
        /// <summary>
        /// 音声の分析と合成をしてくれるやつ
        /// </summary>
        WorldConverter converter = new WorldConverter(0.005);
        /// <summary>
        /// 音声の分析結果を一時的に保持してくれるキュー（FIFO）
        /// </summary>
        ConcurrentQueue<WP[]> buffer;
        /// <summary>
        /// マイクからの入力（未加工の音声）をファイルに保存してくれるやつ
        /// </summary>
        WaveFileWriter waveWriter_In;
        /// <summary>
        /// スピーカーへの出力（加工後の音声）をファイルに保存してくれるやつ
        /// </summary>
        WaveFileWriter waveWriter_Out;
        /// <summary>
        /// 音声のフォーマット
        /// 44.1kHz, 16bit, 1チャンネル
        /// </summary>
        readonly WaveFormat waveFormat = new WaveFormat(WorldConfig.fs, 16, 1);
        #endregion

        /// <summary>
        /// ボイスチェンジャーの開始
        /// </summary>
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
            waveIn.DataAvailable += ReceiveVoice;   // マイク入力に対する処理を登録
            waveIn.StartRecording();    // マイク入力の受け入れ開始

            // スピーカ側の準備
            string filePath_Out = Path.Combine(Environment.CurrentDirectory, "スピーカ.wav");
            waveWriter_Out = new WaveFileWriter(filePath_Out, waveFormat);
            BufferedWaveProvider provider = new BufferedWaveProvider(waveFormat);
            if (outputIsSelectable)
            {
                wavPlayer = new WaveOut()
                {
                    DeviceNumber = outputDeviceNumber,
                };
            }
            else
            {
                MMDevice mmDevice = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                wavPlayer = new WasapiOut(mmDevice, AudioClientShareMode.Shared, false, 100);
            }
            wavPlayer.Init(provider);   // これをして、providerに信号を突っ込むと、wavPlayerが音声を再生してくれる
            wavPlayer.Play();   // スピーカーの再生開始
            Task t = SendVoice(provider);   // 音声を流し込む処理を開始
        }

        /// <summary>
        /// ボイスチェンジャーの停止
        /// </summary>
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

        /// <summary>
        /// マイク入力に対する処理
        /// </summary>
        /// <param name="ee">ここに音声信号が入ってくる</param>
        private void ReceiveVoice(object obj, WaveInEventArgs ee)
        {
            try
            {
                // 音声をファイル（マイク.wav）に保存
                waveWriter_In.Write(ee.Buffer, 0, ee.Buffer.Length);
                waveWriter_In.Flush();

                float[] signal = ByteToFloat(ee.Buffer);    // マイクからの音声信号
                converter.SignalToParameter.AddSignal(signal);
                converter.SignalToParameter.Analyze();    // 音声信号の分析（この行をコメントアウトしても、同じ処理が次の行のReadParameterで実行されます）
                WP[] wp = converter.SignalToParameter.ReadParameter();  // 分析結果

                // 音声の加工　ピッチとフォルマントの変更
                for (int i = 0; i < wp.Length; i++)
                {
                    wp[i].f0 *= prate;
                    float[] temp = new float[wp[i].spectrogram.Length];
                    for (int j = 0; j < wp[i].spectrogram.Length; j++)
                    {
                        temp[j] = wp[i].spectrogram[(int)Math.Min(j / srate, wp[i].spectrogram.Length - 1)];
                    }
                    wp[i].spectrogram = temp;
                }

                buffer.Enqueue(wp); // 分析結果（加工済み）をバッファに入れておく SendVoiceで処理される
            }
            catch (Exception e)
            {
                ErrorLog("音声分析時に失敗しました。", e);
            }
        }

        /// <summary>
        /// スピーカーへ出力する処理
        /// </summary>
        /// <param name="provider"></param>
        /// <returns></returns>
        async Task SendVoice(BufferedWaveProvider provider)
        {
            while (play) //再生可能な間ずっとループ処理
            {
                try
                {
                    buffer.TryDequeue(out WP[] wp); // 分析結果を取り出す
                    if (wp == null)
                    {
                        // 再生するものが無ければ少し待つ
                        await Task.Delay(20);
                        continue;
                    }
                    converter.ParameterToSignal.AddParameter(wp);
                    converter.ParameterToSignal.Synthsize();    // 音声信号の合成（この行をコメントアウトしても、同じ処理が次の行のReadSignalで実行されます）
                    float[] signal = converter.ParameterToSignal.ReadSignal();
                    byte[] Bbuffer = FloatToByte(signal);
                    provider.AddSamples(Bbuffer, 0, Bbuffer.Length); // providerに合成結果を流し込む

                    // 音声をファイル（スピーカー.wav）に保存
                    waveWriter_Out.Write(Bbuffer, 0, Bbuffer.Length);
                    waveWriter_Out.Flush();
                }
                catch (Exception e)
                {
                    ErrorLog("音声合成時に失敗しました。", e);
                }
            }
        }

        /// <summary>
        /// 16bitのbyte列で表現された音をfloatに直す
        /// </summary>
        /// <param name="bytes">16bitの信号</param>
        /// <returns>floatの信号 -1～+1</returns>
        private static float[] ByteToFloat(byte[] bytes)
        {
            float[] values = new float[bytes.Length/2];
            float pow = (float)Math.Pow(2, 15);
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = BitConverter.ToInt16(bytes, 2 * i) / pow; // -1～1に正規化
            }
            return values;
        }
        /// <summary>
        /// floatで表現された音を16bitのbyteに直す
        /// </summary>
        /// <param name="Fbuffer">floatの信号　-1～+1</param>
        /// <returns>16bitのbyte列で表現された信号</returns>
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

        /// <summary>
        /// エラー内容をログ（log/日付.txt）に残す
        /// </summary>
        /// <param name="s">任意の文字列</param>
        /// <param name="e">エラー</param>
        private static void ErrorLog(string s, Exception e)
        {
            Log(s);
            Log(e.ToString());
            Log("[message] \r\n" + e.Message);
            Log("[source] \r\n" + e.Source);
            Log("[stacktrace] \r\n" + e.StackTrace);
        }
        
        /// <summary>
        /// ログに書き込む
        /// </summary>
        /// <param name="s">任意の文字列</param>
        private static void Log(string s)
        {
            try
            {
                string logPath = Path.Combine(Directory.GetCurrentDirectory(), "log");
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
                DirectoryInfo dirInfo = new DirectoryInfo(logPath);
                if ((dirInfo.Attributes & FileAttributes.ReadOnly) > 0)
                {
                    dirInfo.Attributes = dirInfo.Attributes & ~FileAttributes.ReadOnly;
                }
                string date = DateTime.Now.ToString().Split(' ')[0].Replace("/", "_");
                logPath = Path.Combine(logPath, date + ".txt");
                if (!File.Exists(logPath))
                {
                    File.Create(logPath);
                }
                FileAttributes fa = File.GetAttributes(logPath);
                fa = fa & ~FileAttributes.ReadOnly;
                File.SetAttributes(logPath, fa);
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
