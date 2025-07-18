# VFP.memlib и VFP.memlib32
### Оглавление
[Назначение](#Назначение)  
[Регистрация COM-сервера в реестре Windows](#Регистрация-COM-сервера-в-реестре-Windows)  
[Создание объекта VFP.memlib и VFP.memlib32](#Создание-объекта-VFPmemlib-и-VFPmemlib32)  
[Объект Stream](#Объект-Stream)  
&emsp; [Write(str)](#Writestr)  
&emsp; [LenStream()](#LenStream)  
&emsp; [Asc()](#Asc)  
&emsp; [Read(count)](#Readcount)  
&emsp; [ReadLine()](#ReadLine)  
&emsp; [ReadToEnd()](#ReadToEnd)  
&emsp; [CloseStream()](#CloseStream)  
[Объект Array](#Объект-Array)  
&emsp; [NewArray(n)](#newArrayn)  
&emsp; [LenArray()](#LenArray)  
&emsp; [PutArray(i, val)](#PutArrayi-val)  
&emsp; [GetArray(i)](#GetArrayi)  
&emsp; [CloseArray()](#CloseArray)  
[Объект Dictionary](#Объект-Dictionary)  
&emsp; [LenDic()](#LenDic)  
&emsp; [PutDic(strKey, val)](#PutDicstrKey-val)  
&emsp; [GetDic(strKey)](#GetDicstrKey)  
&emsp; [DelDic(strKey)](#DelDicstrKey)  
&emsp; [СloseDic()](#СloseDic)  
[Объект Queue/FIFO](#Объект-QueueFIFO)  
&emsp; [LenFIFO()](#LenFIFO)  
&emsp; [PutFIFO(val)](#PutFIFOval)  
&emsp; [PeekFIFO()](#GetFIFO)  
&emsp; [GetFIFO()](#GetFIFO)  
&emsp; [СloseFIFO()](#СloseFIFO)  
[Объект Task](#Объект-Task)  
&emsp; [DoAsync(com, method[, v1[, v2[, v3[, v4[, v5[, v6[, v7[, v8[, v9[, v10]]]]]]]]]])](#DoAsynccom-method-v1-v2-v3-v4-v5-v6-v7-v8-v9-v10)  
&emsp; [DoAsyncN(com, method, @vals)](#DoAsyncNcom-method-vals)  
&emsp; [WaitTask()](#WaitTask)  
&emsp; [ITask_OnEnded(ret)](#ITask_OnEndedret)  
&emsp; [ITask_OnError(errCode, errMsg)](#ITask_OnErrorerrCode-errMsg)  
&emsp; [CloseTask()](#CloseTask)  
&emsp; [Примеры использования асинхронной задачи на языке Visual FoxPro](#Примеры-использования-асинхронной-задачи-на-языке-Visual-FoxPro)  
[СloseAll()](#СloseAll)  

[Обсуждение](#Обсуждение)  

[История версий](#История-версий)  
### Назначение
Библиотеки memlib32.net.dll и memlib.net.dll реализуют COM-сервер для VFP9 или VFPA, который в принципе может использоваться и в любых других языках, поддерживающих COM технологию обмена данными.  

Microsoft VFP имеет ряд ограничений, связанных с использованием ОП. Теоретически в VFP cуществуют способы использовать ОП до 2 Гб, но при этом возникают очень большие утечки памяти, приводящие к блокировке работы VFP.  

При создании COM-объекта VFP.memlib или VFP.memlib32, выделяемая ОП не занимает место в области VFP. Эта память выделяется в объекте VFP.memlib/VFP.memlib32.  

В memlib.net.dll/memlib32.net.dll реализованы:
- объект Stream, который по используемым методам напоминает работу с файлом, находящимся в ОП;
- объект Array, представляющий одномерный массив с числовым индексом;
- объект Dictionary, представляющий одномерный массив с текстовым индексом (словарь);
- объект Queue/FIFO, очередь, в которцю помещаются любые переменные и объекты;
- объект Task — асинхронная задача, создаваемыя из метода любого COM-объекта.  
### Регистрация COM-сервера в реестре Windows
#### Для VFPA и другого 64-х разрядного ПО
Чтобы объект VFP.memlib был доступен в разрабатываемых программах 64-х битных версий, его нужно зарегистрировать в ОС с помощью утилиты regasm.exe с ключами /codebase и /tlb. Например:
```PowerShell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\memlib.net.dll /codebase /tlb
```
Предварительно поместите файл memlib.net.dll в удобную для вас папку, например, в папку, где находятся другие библиотеки Microsoft VFP.  

Для удаления регистрации из Windows используйте ключ /unregister. Например:
```PowerShell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\memlib.net.dll /unregister
```
Для выполнения вышеуказанных команд требуются права администратора.
#### Для VFP9 и другого 32-х разрядного ПО
Используйте утилиту регистрации для 32-х разрядных программ, находящуюся по другому пути:
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe. Команды на регистрацию и удаление регистрации аналогичны командам
для 64-х разрядного ПО. Например, регистрация:
```PowerShell
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\memlib32.net.dll /codebase /tlb
```
и удаление регистрации:
```PowerShell
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\memlib32.net.dll /unregister
```
## Создание объекта VFP.memlib и VFP.memlib32
Текст кода на VFP:
```xBase
oMem = CreateObject('VFP.memlib')
```
и
```xBase
oMem = CreateObject('VFP.memlib32')
```
## Объект Stream
Объект Stream имеет ниже следующие методы.
### Write(str)
Метод записывает указанную параметром строку в объект Stream.
```xBase
oMem.Write("Привет, Мир!"+chr(10))
```
Указатель записи всегда находится в конце записанного потока. При первом же запросе любой операции на чтение, запись в этот же поток становится невозможной. В этом случае следующая операция записи откроет новый поток.
### LenStream()
Метод возвращает 32-х разрядное положительное число, содержащее длину потока или ноль, если запись в объект еще не производилась.
```xBase
len = oMem.LenStream()
```
### Asc()
Метод читает код первого за указателем чтения символа, но не перемещает указаль.
```xBase
kod = oMem.Asc()
```
Если поток пустой или достигнут конец потока, то метод возвращает значение -1.
### Read(count)
Метод читает заданное параметром количество символов из Stream и возвращает прочитанную строку. Указатель чтения переносится за последний прочитанный символ.
```xBase
str = oMem.Read(10)
```
Если начальная позиция указателя находилась в коце потока, то вместо строки возвращается значение NULL.
### ReadLine()
Метод читает из Stream строку до символа перевода строки. Символы CR и LF не включаются в возвращаемое методом значение. Указатель чтения переносится за последний прочитанный символ включая CR и LF.
```xBase
str = oMem.ReadLine()
```
Если начальная позиция указателя находилась в коце потока, то вместо строки возвращается значение NULL.
### ReadToEnd()
Метод читает из Stream строку до конца потока. Указатель чтения переносится в конец.
```xBase
str = oMem.ReadToEnd()
```
Если начальная позиция указателя находилась в коце потока, то вместо строки возвращается значение NULL.
### CloseStream()
Метод удаляет объект Stream из памяти. Память освобождается. Используйте этот метод, если вам больше не нужен объект Stream. Объект Stream создается при первом использовании метода Write(str).
```xBase
oMem.CloseStream()
```
## Объект Array
Объект Array имеет ниже следующие методы.
### NewArray(n)
Метод создает массив размером, передаваемым числовым параметром. В случае успешного создания массива метод возвращает число созданных элементов массива n. В противном случае возвращается ноль.
### LenArray()
Метод возвращает длинну массива.
### PutArray(i, val)
Метод присваивает элементу массива под номером первого числового параметра значение произвольного типа, переданное вторым параметром.
### GetArray(i)
Метод возвращает значение элемента массива с номером переданным числовым параметром.
### CloseArray()
Метод уничтожает массив.
## Объект Dictionary
Объект Dictionary имеет ниже следующие методы.
### LenDic()
Метод возвращает число пар ключ-значение.
### PutDic(strKey, val)
Метод присваивает элементу словаря со строковым ключем, заданным первым параметром, значение произвольного типа, переданное вторым параметром.
### GetDic(strKey)
Метод возвращает значение элемента словаря со строковым ключем, заданным параметром.
### DelDic(strKey)
Метод удаляет элемент словаря ключ-значение, заданное строковым параметром, содержащим значение ключа.
### СloseDic()
Метод уничтожает словарь и освобождает память.
## Объект Queue/FIFO
Объект FIFO представляет собой очередь значений помещаемых и извлекаемых по принципу первый помещен, первый извлечен. Объект имеет ниже следующие методы.
### LenFIFO()
Метод возвращает число значений в очереди.
### PutFIFO(val)
Метод помещает значение произвольного типа в очередь.
### PeekFIFO()
Метод возвращает первое значение в очереди, но не извлекает его.
### GetFIFO()
Метод возвращает первое значение в очереди и извлекает его.
### СloseFIFO()
Метод уничтожает очередь и освобождает память.
## Объект Task
Объект Task представляеет собой асинхронную задачу. Задача создается на основе метода открытого COM-объекта, такого как Excel.application, com.sun.star.frame.Desktop, VisualFoxPro.Application, ADODB.Connection и т.д. Метод представлятся строкой. Верхний или нижний регистр символов строки с методом имеют значение. Необходимо указывать точное название метода. Объект Task имеет ниже следующие методы и интефейс для обратного вызова после завершения задачи.
### DoAsync(com, method[, v1[, v2[, v3[, v4[, v5[, v6[, v7[, v8[, v9[, v10]]]]]]]]]])
Метод запускает асинхронную задачу без параметров или, если неоходимо — с дополнительными параметрами в количестве до 10-ти значений. Метод возвращает 0 или -1 в случае если задача не создана. Первым параметром передается открытый COM-объект, вторым параметром — текстовое наименование метода.
### DoAsyncN(com, method, @vals)
Метод запускает асинхронную задачу со специально подготовленным массивом требуемых параметров, возвращает 0 или -1 в случае если задача не создана. Первым параметром передается открытый COM-объект, вторым параметром — текстовое наименование метода. Нумерация элементов массива с параметрами начинается от нуля. Массив передается по ссылке. Смотрите ниже приведенный пример на языке Visual FoxPro.
### WaitTask()
Если требуется, метод ожидает завершения задачи. Метод возвращает сформированный результат. Если задача не формирует результат, то возвращается пустая строка.
### ITask_OnError(errCode, errMsg)
Метод обеспечивает обратный вызов через интефейс ITask и событие OnError. Параметрами являются номер ошибки и строка, содержащая текст описания ошибки, возникшей в асинхронной задаче. Если обнаруживается этот метод на вызывающей стороне, то он получает управление при возникновении ошибки. Смотрите ниже пример 3 на языке Visual FoxPro.  

Примечание. VFPA версий до 2024 года не поддерживает работу интерфесов с обратными вызовами.  
### ITask_OnEnded(ret)
Метод обеспечивает обратный вызов через интефейс ITask и событие OnEnded. Параметром является сформированный результат закончившейся задачи или пустая строка. Если обнаруживается этот метод на вызывающей стороне, то он получает управление после завершения задачи. Смотрите ниже пример 2 на языке Visual FoxPro.  

Примечание. VFPA версий до 2024 года не поддерживает работу интерфесов с обратными вызовами.  

Благодарность. Интерфейс ITask и его события добавлены благодоря усилиям [Дмитрия Чунихина](https://github.com/dmitriychunikhin).
### CloseTask()
Метод освобождает сформированные в объекте VFP.memlib ресурсы задачи.
### Примеры использования асинхронной задачи на языке Visual FoxPro
Пример 1. Без использования обратного вызова:
```xBase
oMem=CreateO('VFP.memlib')
oVFP=CreateO('VisualFoxPro.Application')

* Эмитируем работу, выполняемую в течении 123.4 секунд в паралельном процессе.
* Метод DoCmd имеет один параметр - выполняемую команду:
oMem.DoAsync(oVFP,"DoCmd","wait wind '' time 123.4")

* Тем временем вычисляем и возвращаем сумму чисел в текущем процессе.
? 2+2*2                 &&   6

* Формируем массив параметров размером в 1 элемент:
dime vals(1)
vals(1)="m.A+m.A*m.A"   && выражение

* Указываем тип массива с нулевого элемента для COM-объекта oMem:
ComArray(oMem,10)

* Завершаем запущенную асинхронно задачу:
oMem.WaitTask()   && Перешли в синхронное ожидание асинхронной задачи

* Заносим значение в переменную m.A синхронно с помощью метода SetVar.
* Метод SetVar требует два параметра:
oVFP.SetVar("A",2)

* Выполняем вычисление асинхронно с использованием массива параметров:
oMem.DoAsyncN(oVFP,"Eval", @vals)

* Получаем результат вычисления асинхронной задачи:
? oMem.WaitTask()       &&   6

* Закрываем процесс VFP синхронно. Метод Quit не имеет параметров:
oVFP.Quit()

oMem.CloseTask()
rele oVFP
```

Пример 2. C использованием обратного вызова:
```xBase
oMem = CreateO('VFP.memlib')
EventHandler(oMem, NewO("Callback"))
oVFP = CreateO('VisualFoxPro.Application')

oMem.DoAsync(oVFP,"DoCMD","wait wind '' time 12.3")

? tran(seco())+" Старт"
wait wind '' time 10
? tran(seco())          && Через 2.3 секунды должен быть выведен
                        && результат через метод обратного вызова
wait wind '' time 10
? tran(seco())+" Конец"

* Закрываем процесс VFP синхронно.
oVFP.Quit()

* Результат все еще можно вернуть:
? oMem.WaitTask()       && В нашем случае пустая строка

* Удаляем объект задачи:
oMem.CloseTask()

rele oVFP

DEFINE CLASS Callback as Session
  IMPLEMENTS ITask IN 'VFP.memlib'

  * Метод, получающий обратный вызов:
  PROC ITask_OnEnded(ret)
    ? tran(seco())+" Возвращено значение типа "+type('m.ret')+" {"+m.ret+"}"
  ENDPROC
ENDDEFINE
```
Пример 3. C контролем возникновения ошибок:
```xBase
oMem = CreateO('VFP.memlib')
EventHandler(oMem, NewO("Callback"))
oVFP = CreateO('VisualFoxPro.Application')

* Ошибка в команде:
oMem.DoAsync(oVFP,"DoCMD","wait wind '' :-) time 12.3")

read even

DEFINE CLASS Callback as Session
  IMPLEMENTS ITask IN 'VFP.memlib'

  * Метод, получающий обратный вызов:
  PROC ITask_OnEnded(ret)
    ? " Возвращено значение типа: "+type('m.ret')
    this.Finish()
  ENDPROC

  * Метод, обработки ошибки в асинхронной задаче:
  PROC ITask_OnError(errCode, errMsg)
    ? "ОШИБКА: "+tran(m.errCode)
    ? m.errMsg
    this.Finish()
  ENDPROC

  * Завершение в любом случае
  PROC Finish()
    oMem.CloseTask()
    oVFP.Quit()
    clea even
  ENDPROC
ENDDEFINE
```
### СloseAll()
Метод закрывает все объекты COM-сервера и максимально освобождает всю память.
### Обсуждение
Задать вопрос или обсудить тему, касающуюся VFP/VFPA, вы можете в разделе проекта `Issues` > `New issue`.
### История версий
0.0.0.0. 08.03.2025. Опубликована первая рабочая версия с объектами Stream, Array и Dictionary.  
0.0.1.0. 09.03.2025. Добавлен метод CloseAll().  
0.0.2.0. 09.03.2025. Добавлена автоматическая очистка потока записи после формирования потока чтения.  
0.0.2.1. 26.03.2025. Добавлен объект FIFO и объект асинхронной задачи.  
0.0.2.2. 28.03.2025. Объект Task переведен из памяти вызывающей задачи в память VFP.memlib.  
0.0.3.0. 03.04.2025. Добавлен интерфейс ITask, обеспечивающий обратный вызов.  
0.1.0.0. 05.04.2025. Устанвлена видимость интерфейса ITask по умолчанию.  
0.1.1.0. 06.04.2025. Методы DoAsyncX теперь возвращают 0 в случае успеха. Предусмотрена обработка замеченных исключений.  
0.1.2.0. 09.04.2025. Добавлен обратный вызов при возникновении ошибки в асинхронной задаче.  
0.1.3.0. 19.04.2025. Вместо нескольких методов создания асинхронной задачи с разным количеством параметров сделан один универсальный метод с числом параметров от 0 до 10.  
