//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//!!!                                                   !!!
//!!!  memlib32.net на C#.        Автор: A.Б.Корниенко  !!!
//!!!  v0.6.0.0                             22.05.2026  !!!
//!!!                                                   !!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace memlib32 {
  
  [ComVisible(true)]
  [Guid("E3D12675-9C4B-4E38-B91A-FDF319A648A1")]   // фиксированный GUID для интерфейса
  [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
  public interface ITask {
    void OnEnded(object ret);
    void OnError(string cMethod);
  }

  [ComVisible(true)]
  [Guid("7C9F1A4E-BD62-4C71-AF83-1D29B4E3FA72")]   // фиксированный GUID для класса
  [ComSourceInterfaces(typeof(ITask))]
  [ClassInterface(ClassInterfaceType.None)]
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  [ProgId("VFP.memlib32")]  //!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  public class memlib32 {
    public delegate void OnTaskCompletedDelegate(object ret);
    public delegate void OnTaskErrordDelegate(string cMethod);
    public event OnTaskCompletedDelegate OnEnded;
    public event OnTaskErrordDelegate OnError;

    const int k0=0, k1=1, k_1=-1, n2=2048, nbuf=4096, maxInt=16777184;
    Encoding vfpw = Encoding.GetEncoding(1251);
    CancellationTokenSource ctc, cts;
    int[] n1 = new int[] { k0, n2 };
    Dictionary<string, object> di;
    Process[] pu = new Process[2];
    bool lWrite, tcEnd = true;
    int lenS = k0, pi = k0;
    Queue<object> fifo;
    MemoryStream ms;
    Task<object> tc;
    string cMethod;
    Encoding eDos;
    Task<bool> tu;
    object[] ar;
    string eOut;
    object oCom;
    Task ts;

    // Метод записи в поток
    public void Write(string b) {
      lWrite = true;
      if(ms == null) {
        ms = new MemoryStream();
        lenS = k0;
      }
      if (b != null) {
        // Раскладываем UTF-8 строку FoxPro на чистые байты
        byte[] bytes = vfpw.GetBytes(b);

        lenS += bytes.Length;
        ms.Write(bytes, k0, bytes.Length);
      }
    }

    // Метод чтения блока
    public object Read(int count) {
      if (StartRead()) {
        if(count > k0) {
          // Защита от выхода за границы потока
          int available = (int)(ms.Length - ms.Position);
          int bytesToRead = Math.Min(count, available);
          if (bytesToRead <= 0) return new byte[0];

          byte[] b = new byte[bytesToRead];
          ms.Read(b, k0, bytesToRead);
          return b;  // в FoxPro будет тип Blob/Varbinary
        } else {
          return new byte[0];
        }
      } else {
        return null;
      }
    }

    // Метод чтения кода символа
    public int Asc() {
      if (StartRead() && ms.Position < ms.Length) {
        return ms.ReadByte();   // Быстрое чтение одного байта
      }
      return k_1;
    }

    // Подготовка потока к чтению
    bool StartRead() {
      if (ms == null) return false;
      
      if (lWrite) {
        // Перематываем поток в начало для чтения, если до этого в него писали
        ms.Position = 0; 
        lWrite = false;
      }
      return true;
    }

    // Метод чтения записи из записанного потока
    public object ReadLine() {
      if(StartRead()) {
        if (ms.Position >= ms.Length) return new byte[k0];
        byte[] buf = ms.GetBuffer();
        int k = (int)ms.Position;                  // текущая позиция чтения
        int len = (int)(ms.Length - ms.Position);  // сколько байт осталось прочитать

        // Быстрый поиск байта 10 (\n) во всем остатке массива
        int i = Array.IndexOf(buf, (byte)10, k, len);

        int m1 = k0; // длина результирующей строки в байтах
        if(i >= k0) {

          // Ситуация 1: Нашли символ 10. Проверяем, нет ли перед ним символа 13 (\r)
          if(i > k && buf[i - 1] == 13) {
            m1 = i - k - 1; // Длина строки без \r и без \n
          } else {
            m1 = i - k;     // Длина строки без \n
          }
      
          // Сдвигаем курсор потока ms за символ 10, к началу следующей строки
          ms.Position = i + 1;
        } else {

          // Ситуация 2: Символ 10 не найден. Значит, это последняя строка без переноса на конце
          m1 = len;
      
          // Сдвигаем курсор в самый конец потока
          ms.Position = ms.Length;
        }

        // Если строка оказалась пустой (например, пустая строка между \n\n)
        if(m1 <= k0) return new byte[k0];

        // Вырезаем и возвращаем чистый блок байт
        byte[] result = new byte[m1];
        Array.Copy(buf, k, result, k0, m1);

        return result; // В FoxPro улетает Varbinary
      } else {
        return null;
      }
    }

    // метод чтения записи из записанного потока
    public object ReadToEnd() {
      if(StartRead()) {
        if (ms.Length <= 0) return new byte[0];
        byte[] b = ms.ToArray();
        ms.SetLength(0); // Очищаем буфер
        return b; // В FoxPro прилетит Varbinary
      }
      return null;
    }

    public int LenStream() {
      return lenS;
    }

    public void CloseStream() {
      lenS = k0;
      lWrite = false;
      if (ms != null) {
        ms.Close();
        ms.Dispose();
        ms = null;
      }
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

    // Послать сигнал через заданное время
    public void SignalAsync(int tsig, object sig= null) {
      CloseSignal();
      if(!(sig!=null)) sig="Signal";
      cts = new CancellationTokenSource();
      ts= Task.Run(async () => {
            await Task.Delay(tsig, cts.Token);
            OnEnded(sig);
      }, cts.Token);
    }

    // Освободить ресурсы, занятые задачей
    public void CloseSignal() {
      if(ts != null) try {
         cts.Cancel();
         ts.Dispose();
      } catch(Exception) { }
      ts= null;
    }

    // Задержать выполнение на tsig мс
    public void Wait(int tsig) {
      Task.Delay(tsig).Wait();
    }

    // Выполнить метод  COM асинхронно имея до 10 параметров
    public object DoAsync(object com, string method, object p1 = null,
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
    public object DoAsyncN(object com, string method, ref object[] pars) {
      return runTaskAsync(com, method, pars);
    }

    // Выполнить метод COM асинхронно
    bool runTaskAsync(object com, string method, object[] pars) {
      if(tcEnd) CloseTask();
      if(tc != null) {
        return false;
      } else {
        oCom= com;
        ctc = new CancellationTokenSource();
        tc= Task<object>.Run(() => {
           object ret = null;
           try {
             ret = com.GetType().InvokeMember(
                 method,BindingFlags.InvokeMethod|BindingFlags.Instance|BindingFlags.Public,
                 null, com, pars);
             OnEnded(ret);
           } catch(Exception e) {
             if(!ctc.IsCancellationRequested) {
                 Exception realException = e.InnerException ?? e;
                 OnError(method + ": " + realException.Message);
             }
           }
           tcEnd = true;
           return ret;
        }, ctc.Token);
        cMethod = method;
        tcEnd = false;
        return true;
      }
    }

    // Получить результат асинхронной задачи
    public object WaitTask() {
      if(tc != null) {
        tc.Wait();
        tcEnd = true;
        object ret = null;
        try {
          ret = tc.Result !=null? tc.Result : string.Empty;
        } catch(Exception) {
          OnError(cMethod);
        }
        return ret;
      } else {
        return null;
      }
    }

    public memlib32() {
      tcEnd = true;
      OnEnded = (object ret) => { };
      OnError = (string cMethod) => { };
    }

    // Освободить ресурсы, занятые задачей
    public void CloseTask() {
      if(tc != null) try {
        ctc.Cancel();
        if(Marshal.IsComObject(oCom)) Marshal.FinalReleaseComObject(oCom);
        oCom= null;
        tc.Dispose();
      }
      catch(Exception) { }
      tc= null;
      GC.Collect();
      GC.WaitForPendingFinalizers();
    }

    // Запустить утилиту
    public object RunAsync(string util, object arg = null, object cp = null) {
      byte[] buf = new byte[maxInt];
      bool larg = true, lret = true;
      string[] args;
      string[] utils = util.Split('|');
      int kp = k0;
      eDos = Encoding.GetEncoding(866);

      if(cp != null) {
        try{ kp = (int)cp; } catch(Exception) { lret = false; }
      } else if(arg != null) {
        try{ kp = (int)arg; larg = false; } catch(Exception) { }
      } else {
        larg = false;
      }
      if(lret && kp>k0) {
        try { eDos = Encoding.GetEncoding(kp); } catch(Exception) { lret = false; }
      }
      if(lret) {
        if(larg) {
          args = (eDos.GetString(vfpw.GetBytes((string)arg))).Split('|');
        } else {
          args = new string[] {string.Empty};
        }
        int i, p1, q = Math.Max(utils.Length,args.Length);
        CloseUtil();

        tu = Task<bool>.Run(() => {
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
             pu[pi].StartInfo.FileName = utils[i<utils.Length? i:(utils.Length-k1)];
             pu[pi].StartInfo.Arguments = args[i<args.Length? i:(args.Length-k1)];
             try { pu[pi].Start(); } catch(Exception) { return false; }
             if (i > k0) {
               pu[pi].StandardInput.BaseStream.Write(buf,k0,pu[p1].StandardOutput.BaseStream.Read(buf,k0,buf.Length));
               pu[pi].StandardInput.Close();
             }
             p1 = pi;
           }
           return true;
        });

      }
      return lret;
    }

    // Записать что-то в стандартный ввод утилиты
    public void WriteUtil(string x) {
      if(tu != null) {
         if(x.Length>k0) {
            int i = 23;
            bool l = true;
            byte[] buf = eDos.GetBytes(x);
            while (l) {
               try {
                 pu[pi].StandardInput.BaseStream.Write(buf,k0,buf.Length);
                 l = false;
               } catch(Exception) {
                 Task.Delay(4).Wait();
                 if(--i < k0) l = false;
               }
            }
         }
         pu[pi].StandardInput.Close();
      }
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
      if (tu != null) {
        tu.Wait();
        if(tu.Result) {
          if (l) {
            if(pu[pi].StandardOutput.Peek() < k0) {
               return string.Empty;
            } else {
              pu[pi].BeginErrorReadLine();
              return pu[pi].StandardOutput.ReadToEnd() + eOut;
            }
          } else {
            char[] buf = new char[n];
            return new string(buf,k0,pu[pi].StandardOutput.ReadBlock(buf,k0,n));
          }
        } else {
          return "Utility not found";
        }
      } else {
        return string.Empty;
      }
    }

    public void CloseUtil() {
      for (int i = k0; i<=k1; i++) {
        if(pu[i] != null) {
          try {
            pu[i].StandardInput.Close();
            pu[i].StandardOutput.Close();
          } catch(Exception) { }
          pu[i] = null;
        }
      }
      if(tu != null) try { tu.Dispose(); } catch(Exception) { }
      tu = null;
    }

    // Удалить все объекты
    public void CloseAll() {
      CloseStream();
      CloseSignal();
      CloseArray();
      CloseFIFO();
      CloseUtil();
      CloseTask();
      CloseDic();
    }
  }
}
