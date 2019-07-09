using System;
using System.Text;
using System.Net;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Drawing;
using System.Text.RegularExpressions;
using Codeplex.Data;
using System.Linq;
using System.Collections.Generic;

namespace WebRequestSample
{
    /// <summary>
    /// 访问网页，获取html代码
    /// 分析html，获取相关资源地址
    /// 例，下载bing的背景图片图片
    /// Bing主页使用了完整路径表示背景图片地址
    /// 
    /// Update by Supegg Rao on 2019-6-24
    /// 1.增加玉皇顶Ending
    /// 
    /// Update by Supegg Rao on 2019-6-21
    /// 1.工程类型变更为Windows Application，隐藏dos窗口
    /// 
    /// Update by Supegg Rao on 2019-6-17
    /// 1.修复本地文件名没有时间标记的问题
    /// 2.变更json解析类至 NewtonSoft.Json
    /// 
    /// Update by Supegg Rao on 2019-5-1
    /// 1.调整正则表达式适应新的数据结构
    /// 2.变更json解析类至 NewtonSoft.Json
    /// 
    /// Update by Supegg Rao on 2017-1-12
    /// 1.调整正则表达式适应动态的json数据规则
    /// 2.暴露Dynamic JSON的JsonXml字段，动态查找jpg结尾的对象
    /// 
    /// Update by Supegg Rao on 2017-1-9
    /// 1.调整正则表达式适应新的json数据规则
    /// 2.增加Dynamic JSON直接解析json数据
    /// 
    /// Update by Supegg Rao on 2014-2-19
    /// 1.将图片下载到Pics目录，bing.txt移到Pics目录
    /// 
    /// Update by Supegg Rao on 2014-2-19
    /// 1.返回的数据格式因浏览器不同而不同，ie的是完整路径，其他返回的路径不包括主机地址
    /// 增加对这一类型数据的支持
    /// Regex regex1 = new Regex(@"\/az\/hprichbg\/.+?\.jpg"); //?,非贪婪匹配
    /// Match match1 = regex1.Match(ResponseText);
    /// 
    /// Update by Supegg Rao on 2012-11-22
    /// 1.直接访问服务获取json格式的数据，用正则表达式解析图片地址
    /// 
    /// Update by Supegg Rao on 2012-11-29
    /// 原来使用的非贪婪匹配，匹配的地址过长。//Regex regex = new Regex("http.+.jpg");
    /// Regex regex = new Regex(@"http.+?\.jpg"); //?,非贪婪匹配，匹配到第一个".jpg"
    /// 
    /// Update by Supegg Rao on 2012-11-15
    /// 1.Bing主页直接使用了完整路径
    /// 2.为避免Bing主页再改回去，增加了对完整路径、相对路径的支持
    /// 
    /// Update by Supegg Rao on 2012-8-10
    /// 1.将读写超时设定为10S，即重试周期为10S
    /// 
    /// Update by Supegg Rao on 2012-7-10
    /// 1.修正了开机启动时bing.txt的文件路径问题
    /// 
    /// Update by Supegg Rao on 2012-7-9
    /// 1.修正了开机启动时文件保存至C:\Documents and Settings\user 的bug。
    /// 修改后，用全文件名保存文件。
    /// tips：程序路径和程序运行环境路径是两个不同的概念。
    /// 在自启动时程序运行环境路径是C:\Documents and Settings\user
    /// 
    /// Update by Supegg Rao on 2012-6-28
    /// 1.修正了自启动路径可能出错的bug
    /// 
    /// Update by Supegg Rao on 2012-6-15
    /// 1.修正了开机没联网时程序出错的bug
    /// 
    /// Update by Supegg Rao on 2012-6-12
    /// 1.修正了一个因不同电脑日期格式不同而产生的路径bug
    /// 
    /// Update by Supegg Rao on 2012-6-11
    /// 1.新增显示文件下载进度
    /// 
    /// Update by Supegg Rao on 2012-6-8
    /// 1.修正第一次运行创建bing.txt的bug
    /// 
    /// Update by Supegg Rao on 2012-6-6
    /// 1.新增openPicture功能，调用默认图片程序打开图片
    /// 2.新增随系统启动功能
    /// 
    /// Update by Supegg Rao on 2012-5-28
    /// 1.由显示整个html文档改为显示图片地址
    /// 2.将每日图片地址记录入bing.txt文档
    /// </summary>
    class Program
    {
        static Timer Timer = new Timer(rand, null, 5000, 5000);
        static int cnt = 100;

        static void Main(string[] args)
        {
            //Environment.CurrentDirectory + "\\" + "BingTheme.exe"); 启动环境的路径
            string fileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            string path = fileName.Substring(0,fileName.LastIndexOf('\\')+1) +"Pics\\"; //下载目录
            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            Console.WriteLine("乖乖赏美图...");
            
            GetBackground(path);//下载并打开图片

            //添加到启动项
            SetAutoRun("BingTheme", fileName);

            while(cnt>0)
            {
                Thread.Sleep(1000);
                Console.WriteLine($"Count down {cnt--}.......");
            }

            Timer.Change(0, Timeout.Infinite);
            openPicture("MountTai.jpg");
            Thread.Sleep(1000);
        }

        private static void rand(object state)
        {
            string fileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            string path = fileName.Substring(0, fileName.LastIndexOf('\\') + 1) + "Pics\\";
            List<string> pics = new List<string>();

            foreach(var file in Directory.GetFiles(path, "*.jpg"))
            {
                pics.Add(file);
            }

            openPicture(pics[(new Random()).Next(0, pics.Count)]);
        }


        static void GetBackground(string path)
        {
            string ResponseText = "";
            //int startIndex;
            //int endIndex;
            string backgroundUrl = null;//背景图片url
            string bgPath; //完整文件路径

            //访问网页
            //HttpWebRequest是WebRequest的子类，实现了更多的方法和属性
            string uri = "http://cn.bing.com/HPImageArchive.aspx?format=js&n=1";
            Uri url = new Uri(uri);
            HttpWebResponse res = null;
            StreamReader ReaderText =null;
            bool success = false;
            do{
                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                    req.Timeout = 10000;//10秒超时
                    req.ReadWriteTimeout = 10000;
                    //req.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.2; .NET4.0C; .NET4.0E; rv:11.0) like Gecko ";
                    res = (HttpWebResponse)req.GetResponse();
                    //req.Method = "Post";
                    ReaderText = new StreamReader(res.GetResponseStream(), Encoding.UTF8);
                    ResponseText = ReaderText.ReadToEnd();//从流中读取html
                    Console.WriteLine(ResponseText);
                    success = true;
                }
                catch
                {
                    success = false;
                    Console.WriteLine("重试...");
                }
            }while(!success);

            //获取背景图片的地址并下载
            //方法一，uri= "http://cn.bing.com/"
            //string backgroundStart = "g_img={url:'";
            //string backgroundEnd = ".jpg";
            //startIndex = ResponseText.IndexOf(backgroundStart)+backgroundStart.Length;
            //backgroundUrl = ResponseText.Substring(startIndex);
            //endIndex = backgroundUrl.IndexOf(backgroundEnd) + backgroundEnd.Length;
            //string backgroundUrlOld = uri + backgroundUrl.Substring(0, endIndex);//获取最终地址
            //backgroundUrl =backgroundUrl.Substring(0, endIndex);//获取最终地址
            ///////************end of 方法一************************************//////

            //方法二，uri = "http://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&nc=1385084200455&video=1"
            //Regex regex = new Regex("http.+.jpg");
            Regex regex = new Regex(@"(?<="")?[\w-/]+?\.jpg");// new Regex("(?<=\"image\":\").+?\\.jpg"); //?,非贪婪匹配
            Match match = regex.Match(ResponseText);
            if (match.Success)
            {
                backgroundUrl = "http://cn.bing.com" + match.Value;
                Console.WriteLine(backgroundUrl);
            }
            else
            {
                Console.WriteLine("Ooops,download failed.");
                Console.Read();
                return;
            }
            ///////************end of 方法二************************************//////

            #region use  json
            //DynamicJson json = DynamicJson.Parse(ResponseText);
            var json = (Newtonsoft.Json.Linq.JObject)Newtonsoft.Json.JsonConvert.DeserializeObject(ResponseText);
            Console.WriteLine(json);
            //backgroundUrl = "http://cn.bing.com" + json.images[0].vid.image;
            //backgroundUrl = "http://cn.bing.com" + json.JsonXml.Descendants().Where(x => x.Value.EndsWith("jpg")).First().Value;
            backgroundUrl = "http://cn.bing.com/" + json["images"][0]["url"].ToString();
            
            Console.WriteLine(backgroundUrl);

            #endregion

            //bgPath = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "_" + backgroundUrl.Substring(backgroundUrl.LastIndexOf("/")+1); //文件名
            bgPath = path + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "_" + Regex.Match(json["images"][0]["url"].ToString(), @"(?<=id=).+?(?=\.jpg)").Value + ".jpg";//完整路径名

            //如果bing.txt不存在，则新建
            StreamReader sr =null;
            if (!File.Exists(path + "bing.txt"))
            {
                StreamWriter sw = new StreamWriter(path + "bing.txt");
                sw.Close();

            }
            sr= new StreamReader(path + "bing.txt");

            string txt = sr.ReadToEnd();
            sr.Close();
            if (txt.IndexOf(backgroundUrl) == -1)//如果没有记录，就记录地址
            {
                //StreamWriter sw = new StreamWriter("bing.txt");
                StreamWriter sw = File.AppendText(path +"bing.txt");
                try
                {
                    sw.Write(DateTime.Now + ":" + backgroundUrl + "\r\n");
                }
                catch (Exception)
                { }
                finally
                {
                    sw.Close();
                }
            }

            try
            {
                downFile(bgPath, backgroundUrl);//下载
            }
            catch
            {
                downFile(bgPath, backgroundUrl);//下载
            }
            

            openPicture(bgPath); //打开图片

            //设置桌面背景
            //Console.Write("是否设为桌面背景(Y/N)。  ");
            //if (Console.ReadLine().ToUpper()=="Y")
            //    SetDestPicture(bgPath);
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="path">本机路径</param>
        /// <param name="url">网络地址</param>
        static void downFile( string path, string url)
        {
            Console.WriteLine(url);

            Uri uri = new Uri(url);
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(uri);
            HttpWebResponse res = (HttpWebResponse)req.GetResponse();
            Int32 length = (Int32)res.ContentLength;
            Stream stream = res.GetResponseStream();

            FileStream fileWriter = new FileStream(path, FileMode.Create, FileAccess.Write);
            byte[] buff = new byte[1024];//缓冲数组
            int count = 0; //当次循环实际读取的字节数
            int written = 0;//已经写出的长度
            int progress = -10;//进度

            //从流中循环读出
            while ((count = stream.Read(buff, 0, buff.Length)) > 0) //当网络非常不好时，read会触发一个异常，尚未解决。
            {
                written += count;
                fileWriter.Write(buff, 0, count);//写入文件流
                if (written * 100 / length - progress > 10)
                {
                    progress = written * 100 / length;
                    Console.WriteLine("下载{0}%...",progress);
                }
            }

            //释放资源
            fileWriter.Close();
            fileWriter.Dispose();
            stream.Close();
            stream.Dispose();
            res.Close();
        }

        /// <summary>
        /// 打开刚下载的图片
        /// </summary>
        /// <param name="path">文件路径</param>
        static void openPicture(string path)
        {
            //建立新的系统进程    
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            //设置文件名，此处为图片的真实路径+文件名    
            process.StartInfo.FileName = path;
            //此为关键部分。设置进程运行参数，此时为最大化窗口显示图片。    
            process.StartInfo.Arguments = "rundll32.exe C://WINDOWS//system32//shimgvw.dll,ImageView_Fullscreen";
            //此项为是否使用Shell执行程序，因系统默认为true，此项也可不设，但若设置必须为true    
            process.StartInfo.UseShellExecute = true;
            //此处可以更改进程所打开窗体的显示样式，可以不设    
            process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            process.Start();
            //process.Close()方法可以不写
            //可以换成process.WaitForExit()不过后者将会使程序暂停，直到打开的窗体被关闭，程序才会继续进行。
            process.Close();  
        }

        #region 设置桌面背景
        /// <summary>  
        /// 设置桌面背景图片  
        /// </summary>  
        /// <param name="picture">图片路径</param>  
        static void SetDestPicture(string picture)
        {
            if (File.Exists(picture))
            {
                if (Path.GetExtension(picture).ToLower() != "bmp")
                {
                    // 其它格式文件先转换为bmp再设置  
                    string tempFile = @"C:\test.bmp";
                    Image image = Image.FromFile(picture);
                    image.Save(tempFile, System.Drawing.Imaging.ImageFormat.Bmp);
                    picture = tempFile;
                    setBMPAsDesktop(picture);
                    File.Delete(tempFile);
                }
                else
                {
                    setBMPAsDesktop(picture);
                }

            }
        }

        [DllImport("user32.dll", EntryPoint = "SystemParametersInfo")]
        public static extern int SystemParametersInfo(
            int uAction,
            int uParam,
            string lpvParam,
            int fuWinIni
        );
        /// <summary>  
        /// 设置BMP格式的背景图片  
        /// </summary>  
        /// <param name="bmp">图片路径</param>  
        static void setBMPAsDesktop(string bmp)
        {
            SystemParametersInfo(20, 0, bmp, 0x10);
        }
        #endregion

        public static bool SetAutoRun(string keyName, string filePath)
        {
            try
            {
                RegistryKey runKey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
                runKey.SetValue(keyName, filePath);
                runKey.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }

    }
}
