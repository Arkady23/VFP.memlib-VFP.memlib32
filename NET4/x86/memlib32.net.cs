//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//!!!                                                   !!!
//!!!  memlib32.net на C#.        Автор: A.Б.Корниенко  !!!
//!!!  v0.3.0.0                             20.12.2025  !!!
//!!!                                                   !!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
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
  [ProgId("VFP.memlib32")]  //!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  public class memlib {
    public delegate void OnTaskCompletedDelegate(object ret);
    public delegate void OnTaskErrordDelegate(string cMethod);
    public event OnTaskCompletedDelegate OnEnded;
    public event OnTaskErrordDelegate OnError;

    const int k0=0, k_1=-1, maxInt=16777184;
    Encoding eDos = Encoding.GetEncoding(866);
    Dictionary<string, object> di;
    bool lWrite, tsEnd = true;
    Queue<object> fifo;
    Task<object> ts;
    StringWriter sw;
    StringReader sr;
    string cMethod;
    int lenS = k0;
    object[] ar;
    Process pu;

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
          return "";
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
          ret = ts.Result !=null? ts.Result : "";
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
    public object RunAsync(string util, object arg = null) {
      var ps = new ProcessStartInfo();
      string par = string.Empty;
      try {
        par = (string)arg;
      } catch(Exception) {
        par = string.Empty;
      }
      ps.FileName = util;
      ps.CreateNoWindow = true;
      ps.UseShellExecute = false;
      ps.RedirectStandardInput = true;
      ps.RedirectStandardOutput = true;
      ps.Arguments = par;
      CloseUtil();
      try {
        pu = Process.Start(ps);
      } catch(Exception) {
        return false;
      }
      return true;
    }

    // Записать что-то в стандартный ввод утилиты
    public void WriteUtil(string x) {
      if(x.Length>0) {
         byte[] buf = eDos.GetBytes(x);
         pu.StandardInput.BaseStream.Write(buf,0,buf.Length);
      }
    }

    // Прочитать из стандартного вывода утилиты всё или заданное количество символов
    public object ReadUtil(int n = 0) {
      if(n==0 || n>maxInt) n = maxInt;
      byte[] buf = new byte[n];
      return eDos.GetString(buf,0,
             pu.StandardOutput.BaseStream.Read(buf,0,n));
    }

    public void CloseUtil() {
      if(pu != null) {
        try {
          pu.StandardInput.Close();
          pu.StandardOutput.Close();
        } catch(Exception) { }
        pu = null;
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
