using Accord.Video;
using Accord.Video.DirectShow;
using Accord.Video.FFMPEG;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace AutoLiquid_GenScript_Single_Handling.Utils
{
    /// <summary>
    /// 摄像/拍照工具（依赖 Accord.Video.DirectShow + Accord.Video.FFMPEG）
    /// </summary>
    public static class CameraUtils
    {
        private static FilterInfoCollection _videoDevices;
        private static VideoCaptureDevice _videoSource;
        private static VideoFileWriter _videoWriter;
        private static bool _isPhotoMode;
        private static string _outputDir;
        private static string _excelBaseName;
        private static System.Timers.Timer _midTimer;
        private static readonly object _frameLock = new object();
        private static volatile bool _saveNextFrameAsStart;
        private static volatile bool _saveNextFrameAsMid;
        private static volatile bool _saveNextFrameAsEnd;
        private static volatile bool _videoWriterInitialized;
        private static readonly int _videoFps = 20;

        /// <summary>
        /// 启动捕获（在 Run 开始时调用）
        /// </summary>
        /// <param name="excelBaseName">导入的 Excel 名称（无扩展名）</param>
        /// <param name="photoMode">true：拍照模式；false：视频模式</param>
        /// <param name="delayMidSeconds">拍照模式下中间延时（秒）</param>
        public static void Start(string excelBaseName, bool photoMode, int delayMidSeconds)
        {
            try
            {
                Stop(); // 先确保之前没有残留

                _isPhotoMode = photoMode;
                _excelBaseName = SanitizeFileName(excelBaseName ?? "UnknownExcel");
                _outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    (string)Application.Current.FindResource("ProcessesName"),
                    "Records",
                    DateTime.Now.ToString("yyyy-MM-dd"));
                Directory.CreateDirectory(_outputDir);

                // 查找第一个摄像头设备
                _videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                if (_videoDevices == null || _videoDevices.Count == 0)
                    return;

                var deviceInfo = _videoDevices[0];
                _videoSource = new VideoCaptureDevice(deviceInfo.MonikerString);
                _videoSource.NewFrame += VideoSource_NewFrame;

                // 启动设备
                _videoSource.Start();

                if (_isPhotoMode)
                {
                    // 拍照模式：下一帧保存为 start
                    _saveNextFrameAsStart = true;

                    // 中间延时触发
                    if (_midTimer != null)
                    {
                        _midTimer.Stop();
                        _midTimer.Dispose();
                    }
                    _midTimer = new System.Timers.Timer(Math.Max(1000, delayMidSeconds * 1000));
                    _midTimer.AutoReset = false;
                    _midTimer.Elapsed += (s, e) =>
                    {
                        _saveNextFrameAsMid = true;
                    };
                    _midTimer.Start();
                }
                else
                {
                    // 视频模式：初始化视频写入在收到第一帧时完成（依赖帧大小）
                    _videoWriter = new VideoFileWriter();
                    _videoWriterInitialized = false;
                }
            }
            catch (Exception ex)
            {
                // 不抛异常，调用方通过日志记录
                LogHelper.Error("CameraUtils.Start error: " + ex.Message);
            }
        }

        /// <summary>
        /// 在 Run 结束时调用：拍照模式会触发 end 拍照；视频模式会结束并保存视频文件
        /// </summary>
        public static void Stop()
        {
            try
            {
                if (_isPhotoMode)
                {
                    if (_videoSource != null && _videoSource.IsRunning)
                    {
                        // 触发结束照片保存；等候保存后再停止（最多等待 4s）
                        _saveNextFrameAsEnd = true;
                        WaitFrameSave(4000);
                    }
                }
                else
                {
                    // 视频：停止并关闭 writer
                    if (_videoSource != null && _videoSource.IsRunning)
                    {
                        // 等待 writer flush（在 NewFrame 中写入）
                        // 标记停止：先 stop video source so NewFrame stops coming, then close writer
                        _videoSource.SignalToStop();
                        _videoSource.WaitForStop();
                    }

                    if (_videoWriter != null && _videoWriter.IsOpen)
                    {
                        try
                        {
                            _videoWriter.Close();
                        }
                        catch (Exception) { }
                        finally
                        {
                            _videoWriter.Dispose();
                            _videoWriter = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error("CameraUtils.Stop error: " + ex.Message);
            }
            finally
            {
                // 清理
                if (_midTimer != null)
                {
                    _midTimer.Stop();
                    _midTimer.Dispose();
                    _midTimer = null;
                }

                if (_videoSource != null)
                {
                    try
                    {
                        if (_videoSource.IsRunning)
                        {
                            _videoSource.SignalToStop();
                            _videoSource.WaitForStop();
                        }
                        _videoSource.NewFrame -= VideoSource_NewFrame;
                        _videoSource = null;
                    }
                    catch { _videoSource = null; }
                }

                _videoWriterInitialized = false;
                _saveNextFrameAsStart = _saveNextFrameAsMid = _saveNextFrameAsEnd = false;
            }
        }

        private static void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            // 该事件在摄像头线程中触发，必须快速完成
            try
            {
                Bitmap frame = (Bitmap)eventArgs.Frame.Clone();
                lock (_frameLock)
                {
                    if (_isPhotoMode)
                    {
                        if (_saveNextFrameAsStart)
                        {
                            SaveBitmapToJpeg(frame, MakeFileName("_start"));
                            _saveNextFrameAsStart = false;
                        }
                        else if (_saveNextFrameAsMid)
                        {
                            SaveBitmapToJpeg(frame, MakeFileName("_mid"));
                            _saveNextFrameAsMid = false;
                        }
                        else if (_saveNextFrameAsEnd)
                        {
                            SaveBitmapToJpeg(frame, MakeFileName("_end"));
                            _saveNextFrameAsEnd = false;
                        }
                    }
                    else
                    {
                        // 视频模式：初始化 writer（第一次）
                        if (!_videoWriterInitialized)
                        {
                            try
                            {
                                var videoFile = Path.Combine(_outputDir, _excelBaseName + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".avi");
                                _videoWriter.Open(videoFile, frame.Width, frame.Height, _videoFps, VideoCodec.MPEG4);
                                _videoWriterInitialized = true;
                            }
                            catch (Exception ex)
                            {
                                LogHelper.Error("Video writer open failed: " + ex.Message);
                                _videoWriterInitialized = false;
                            }
                        }

                        if (_videoWriterInitialized && _videoWriter != null)
                        {
                            try
                            {
                                _videoWriter.WriteVideoFrame(frame);
                            }
                            catch (Exception ex)
                            {
                                LogHelper.Error("WriteVideoFrame error: " + ex.Message);
                            }
                        }
                    }
                }

                frame.Dispose();
            }
            catch (Exception ex)
            {
                LogHelper.Error("VideoSource_NewFrame error: " + ex.Message);
            }
        }

        private static void SaveBitmapToJpeg(Bitmap bmp, string fileName)
        {
            try
            {
                var path = Path.Combine(_outputDir, fileName + ".jpg");
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            catch (Exception ex)
            {
                LogHelper.Error("SaveBitmapToJpeg error: " + ex.Message);
            }
        }

        private static string MakeFileName(string suffix)
        {
            return _excelBaseName + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + suffix;
        }

        private static void WaitFrameSave(int timeoutMs)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                // 只要有任何保存标志被清除就退出
                if (!_saveNextFrameAsEnd)
                    break;
                Thread.Sleep(50);
            }
        }

        private static string SanitizeFileName(string name)
        {
            var invalids = new string(Path.GetInvalidFileNameChars());
            var r = new Regex(string.Format("[{0}]", Regex.Escape(invalids)));
            return r.Replace(name, "_");
        }
    }
}
