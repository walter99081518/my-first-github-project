using System;
using System.Threading;
using System.Collections;

public class WriteLogThread
{
    private Thread thread = null;
    private static ArrayList msglist;

    public WriteLogThread()
    {
        thread = new Thread(new ThreadStart(run));
        thread.IsBackground = true;
    }

    private static void run()
    {
        while (true)
        {
            try
            {
                Monitor.Enter(msglist);
            }
            catch (Exception ex)
            {

            }
            finally
            {
                Monitor.Exit(msglist);
            }
            Thread.Sleep(500);
        }
    }

    public void start() 
    {
        thread.Start(); 
    } 

}
