//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//!!!                                                   !!!
//!!!  memlib32.net на C#.        Автор: A.Б.Корниенко  !!!
//!!!  v0.4.1.0                             07.02.2026  !!!
//!!!                                                   !!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace memlib {
  [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
  public interface ITask {
    void OnEnded(object ret);
    void OnError(string cMethod);
  }

  [ComSourceInterfaces(typeof(ITask))]
  [ClassInterface(ClassInterfaceType.None)]
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  [ProgId("VFP.memlib")]    //!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  public class memlib {
    public delegate void OnTaskCompletedDelegate(object ret);
    public delegate void OnTaskErrordDelegate(string cMethod);
    public event OnTaskCompletedDelegate OnEnded;
    public event OnTaskErrordDelegate OnError;

    const int k0=0, k1=1, k_1=-1, k9=23000, n2=2048, nbuf=4096, maxInt=67108832;
    int[] n1 = new int[] { k0, n2 };
    Dictionary<string, object> di;
    Process[] pu = new Process[2];
    bool lWrite, tsEnd = true;
    int lenS = k0, pi = k0;
    Queue<object> fifo;
    StringWriter sw;
    StringReader sr;
    Task<object> ts;
    string cMethod;
    Encoding eDos;
    object[] ar;
    string eOut;

    // метод записи в поток
    public void Write(string x) {
      if(!(sw != null)) {
        sw = new StringWriter();
        lenS = k0;
      }
      lenS += x.Length;
      lWrite = true;
      sw.Write(x);
    }

    // метод чтения блока
    public string Read(int count) {
      if(StartRead()){
        if(count>k0) {
          char[] buf = new char[count];
          sr.Read(buf, k0, count);
          return new string(buf);
        } else {
          return string.Empty;
        }
      } else {
        return null;
      }
    }

    // метод чтения кода символа
    public int Asc() {
      return StartRead()? sr.Peek() : k_1;
    }

    // метод чтения записи из записанного потока
    public string ReadLine() {
      return StartRead()? sr.ReadLine() : null;
    }

    // метод чтения записи из записанного потока
    public string ReadToEnd() {
      return StartRead()? sr.ReadToEnd() : null;
    }

    bool StartRead() {
      if(lWrite) CloseSr();
      if(!(sr != null)) {
        if(sw != null) {
          sr = new StringReader(sw.ToString());
          lWrite = false;
          CloseSw();
        } else {
          return false;
        }
      }
      return true;
    }

    public int LenStream() {
      return lenS;
    }

    void CloseSw() {
      if(sw != null) {
        sw.Close();
        sw.Dispose();
      }
      sw = null;
    }

    void CloseSr() {
      if(sr != null) {
        sr.Close();
        sr.Dispose();
      }
      sr = null;
    }

    public void CloseStream() {
      CloseSw();
      CloseSr();
    }

    // Создать массив
    public int NewArray(int n) {
      try {
        ar = new object[n];
        return n;
      } catch(Exception) {
        return k0;
      }
    }

    public int LenArray() {
      return ar != null? ar.Length : k0;
    }

    // Присвоить значение
    public void PutArray(int j, object a) {
      if(ar != null) if(ar.Length>j) ar[j] = a;
    }

    // Вернуть значение
    public object GetArray(int j) {
      if(ar != null) if(ar.Length>j) return ar[j];
      return null;
    }

    // Удалить массив
    public void CloseArray() {
      ar = null;
    }

    public int LenDic() {
      return di != null? di.Count : k0;
    }

    // Добавить значение
    public void PutDic(string key, object val) {
      if(!(di != null)) di = new Dictionary<string, object>();
      try{
        di.Add(key, val);
      } catch(Exception) {
        di.Remove(key);
        di.Add(key, val);
      }
    }

    // Вернуть значение
    public object GetDic(string key) {
      object val;
      if(di != null) if(di.TryGetValue(key, out val)) {
        return val;
      }
      return null;
    }

    // Удалить значение
    public void DelDic(string key) {
      if(di != null) di.Remove(key);
    }

    // Удалить словарь
    public void CloseDic() {
      if(di != null) di.Clear();
      di = null;
    }

    public int LenFIFO() {
      return fifo != null? fifo.Count : k0;
    }

    // Добавить значение
    public void PutFIFO(object o) {
      if(!(fifo != null)) fifo = new Queue<object>();
      fifo.Enqueue(o);
    }

    // Извлечь значение
    public object GetFIFO() {
      if(fifo.Count>k0 && fifo != null) {
        return fifo.Dequeue();
      } else {
        return null;
      }
    }

    // Посмотреть значение без извлечения
    public object PeekFIFO() {
      if(fifo.Count>k0 && fifo != null) {
        return fifo.Peek();
      } else {
        return null;
      }
    }

    // Удалить очередь
    public void CloseFIFO() {
      if(fifo != null) fifo.Clear();
      fifo = null;
    }

    // Выполнить метод  COM асинхронно имея до 10 параметров
    public int DoAsync(object com, string method, object p1 = null,
                       object p2 = null, object p3 = null, object p4 = null,
                       object p5 = null, object p6 = null, object p7 = null,
                       object p8 = null, object p9 = null, object p10 = null) {
      object[] pars;
      if(!(p1 != null)) {
        pars = new object[] { };
      } else if(!(p2 != null)) {
        pars = new object[] { p1 };
      } else if(!(p3 != null)) {
        pars = new object[] { p1,p2 };
      } else if(!(p4 != null)) {
        pars = new object[] { p1,p2,p3 };
      } else if(!(p5 != null)) {
        pars = new object[] { p1,p2,p3,p4 };
      } else if(!(p6 != null)) {
        pars = new object[] { p1,p2,p3,p4,p5 };
      } else if(!(p7 != null)) {
        pars = new object[] { p1,p2,p3,p4,p5,p6 };
      } else if(!(p8 != null)) {
        pars = new object[] { p1,p2,p3,p4,p5,p6,p7 };
      } else if(!(p9 != null)) {
        pars = new object[] { p1,p2,p3,p4,p5,p6,p7,p8 };
      } else if(!(p10 != null)) {
        pars = new object[] { p1,p2,p3,p4,p5,p6,p7,p8,p9 };
      } else {
        pars = new object[] { p1,p2,p3,p4,p5,p6,p7,p8,p9,p10 };
      }
      return runTaskAsync(com, method, pars);
    }

    // Выполнить метод  COM асинхронно с любым чисом параметров
    public int DoAsyncN(object com, string method, ref object[] pars) {
      return runTaskAsync(com, method, pars);
    }

    // Выполнить метод COM асинхронно
    int runTaskAsync(object com, string method, object[] pars) {
      if(tsEnd) CloseTask();
      if(ts != null) {
        return k_1;
      } else {
        ts = Task<object>.Run(() => {
           object ret = null;
           try {
             ret = com.GetType().InvokeMember(
                 method,BindingFlags.InvokeMethod|BindingFlags.Instance|BindingFlags.Public,
                 null, com, pars);
             tsEnd = true;
             OnEnded(ret);
           } catch(Exception) {
             tsEnd = true;
             OnError(method);
           }
           return ret;
        });
        cMethod = method;
        tsEnd = false;
        return k0;
      }
    }

    // Получить результат асинхронной задачи
    public object WaitTask() {
      if(ts != null) {
        ts.Wait();
        tsEnd = true;
        object ret = null;
        try {
          ret = ts.Result !=null? ts.Result : string.Empty;
        } catch(Exception) {
          OnError(cMethod);
        }
        return ret;
      } else {
        return null;
      }
    }

    public memlib() {
      tsEnd = true;
      OnEnded = (object ret) => { };
      OnError = (string cMethod) => { };
    }

    // Освободить ресурсы, занятые задачей
    public void CloseTask() {
      if(ts != null) try { ts.Dispose(); } catch(Exception) { }
      ts = null;
    }

    // Запустить утилиту
    public object RunAsync(string util, object arg = null, object cp = null) {
      byte[] buf = new byte[maxInt];
      bool larg = true;
      string[] args;
      string[] utils = util.Split('|');
      eDos = Encoding.GetEncoding(866);

      if(cp != null) {
        try { eDos = Encoding.GetEncoding((int)cp); } catch(Exception) { }
      } else if(arg != null) {
          try { eDos = Encoding.GetEncoding((int)arg); larg = false;
          } catch(Exception) { };
      } else {
        larg = false;
      }
      if(larg) {
        args = (eDos.GetString(Encoding.GetEncoding(1251).GetBytes((string)arg))).Split('|');
      } else {
        args = new string[] {string.Empty};
      }
      int i, p1, q = Math.Max(utils.Length,args.Length);
      CloseUtil();
      eOut = null;
      p1 = k1;
      for (i = k0; i<q; i++) {
        pi = p1==k0? k1:k0;
        pu[pi] = new Process();
        pu[pi].StartInfo.CreateNoWindow = true;
        pu[pi].StartInfo.UseShellExecute = false;
        pu[pi].StartInfo.RedirectStandardError = true;
        pu[pi].StartInfo.RedirectStandardInput = true;
        pu[pi].StartInfo.RedirectStandardOutput = true;
        pu[pi].StartInfo.StandardOutputEncoding = eDos;
        pu[pi].ErrorDataReceived += new DataReceivedEventHandler((sender, e) =>
                      { eOut += "\r\n" + e.Data; });
        pu[pi].StartInfo.FileName = utils[i<utils.Length? i:(utils.Length-1)];
        pu[pi].StartInfo.Arguments = args[i<args.Length? i:(args.Length-1)];
        try { pu[pi].Start(); } catch(Exception) { return false; }
        if (i > 0) pu[pi].StandardInput.BaseStream.Write(buf, 0,
                   pu[p1].StandardOutput.BaseStream.Read(buf,0,buf.Length));
        pu[pi].StandardInput.Close();
        p1 = pi;
      }
      return true;
    }

    // Прочитать из стандартного вывода утилиты всё или заданное количество символов
    public object ReadUtil(int n = k0) {
      bool l = false;
      if(n==k0) {
         n = maxInt;
         l = true;
      }else if(n>maxInt) {
         n = maxInt;
      }
      //pu[pi].WaitForExit(k9);
      if (pu[pi].StandardOutput.Peek() < k0) {
        return string.Empty;
      } else if (l) {
        pu[pi].BeginErrorReadLine();
        return pu[pi].StandardOutput.ReadToEnd() + eOut;
      } else {
        char[] buf = new char[n];
        return new string(buf,0,pu[pi].StandardOutput.Read(buf,0,n));
      }
    }

    public void CloseUtil() {
      for (int i = 0; i<=k1; i++) {
        if(pu[i] != null) {
          try {
            pu[i].StandardInput.Close();
            pu[i].StandardOutput.Close();
          } catch(Exception) { }
          pu[i] = null;
        }
      }
    }

    // Удалить все объекты
    public void CloseAll() {
      CloseStream();
      CloseArray();
      CloseFIFO();
      CloseTask();
      CloseUtil();
      CloseDic();
    }
  }
}
