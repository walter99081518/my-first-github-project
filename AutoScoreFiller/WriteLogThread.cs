using System;
using System.Threading;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;

namespace AutoScoreFiller {
    public delegate void EmitCustomizedMessage(string msg);
    public class WriteLogThread {
        private Thread thread = null;
        private const int size = 128;
        private static List<string> msglist = new List<string>();
        private static StreamWriter sw = new StreamWriter(@"./appstatus.log", true);
        public static event EmitCustomizedMessage EmitMessage;

        public WriteLogThread() {
            thread = new Thread(new ThreadStart(Run));
            thread.IsBackground = true;
        }

        private static void Run() {
            bool fatal = false;
            while (true) {
                try {
                    Monitor.Enter(msglist);
                    if (msglist.Count > 0) {
                        sw.WriteLine(msglist[0]);
                        sw.Flush();
                        msglist.RemoveAt(0);
                    }
                    Thread.Sleep(500);

                }
                catch (Exception ex) {
                    fatal = true;
                    EmitMessage(ex.Message);
                }
                finally {
                    Monitor.Exit(msglist);
                }
                if (fatal) {
                    break;
                }
            }
        }

        public void Append(string msg, bool timestamp = true) {
            try {
                Monitor.Enter(msglist);
                if (msglist.Count <= size) {
                    if (timestamp) {
                        string tmstp = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss.fff");
                        msglist.Add(tmstp + " " + msg);
                    }
                    else {
                        msglist.Add(msg);
                    }
                }
            }
            catch (Exception ex) {
                EmitMessage("添加程序运行日志时发生错误\n" + ex.Message);
            }
            finally {
                Monitor.Exit(msglist);
            }
        }

        public void Start() {
            thread.Start();
        }
    }
}
