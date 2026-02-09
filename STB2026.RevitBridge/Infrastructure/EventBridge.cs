using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace STB2026.RevitBridge.Infrastructure
{
    /// <summary>
    /// Мост для выполнения кода на потоке Revit API.
    /// 
    /// Проблема: Named Pipe работает в фоновом потоке,
    /// но Revit API доступен ТОЛЬКО из UI-потока.
    /// 
    /// Решение: ExternalEvent + очередь запросов.
    /// Фоновый поток кладёт задачу в очередь → ExternalEvent срабатывает
    /// на UI потоке → задача выполняется → результат возвращается.
    /// </summary>
    public sealed class EventBridge : IExternalEventHandler
    {
        private readonly ConcurrentQueue<WorkItem> _queue = new ConcurrentQueue<WorkItem>();
        private ExternalEvent _externalEvent;
        private UIApplication _uiApp;

        public string GetName() => "STB2026.EventBridge";

        /// <summary>Инициализация — вызывается один раз при старте плагина.</summary>
        public void Initialize(UIControlledApplication app)
        {
            _externalEvent = ExternalEvent.Create(this);
        }

        /// <summary>
        /// Выполнить функцию на потоке Revit API и вернуть результат.
        /// Вызывается из фонового потока Named Pipe.
        /// Таймаут: 30 секунд.
        /// </summary>
        public async Task<object> ExecuteOnRevitThread(
            Func<UIApplication, object> action, 
            int timeoutMs = 30000)
        {
            var item = new WorkItem(action);
            _queue.Enqueue(item);
            _externalEvent.Raise();

            // Ждём завершения (ExternalEvent сработает на UI потоке)
            var completed = await Task.Run(() => item.Completed.Wait(timeoutMs));
            if (!completed)
                throw new TimeoutException("Revit не ответил за 30 секунд");

            if (item.Exception != null)
                throw item.Exception;

            return item.Result;
        }

        /// <summary>
        /// Вызывается Revit на UI-потоке при Raise().
        /// Обрабатывает ВСЕ задачи в очереди за один вызов.
        /// </summary>
        public void Execute(UIApplication app)
        {
            _uiApp = app;

            while (_queue.TryDequeue(out var item))
            {
                try
                {
                    item.Result = item.Action(app);
                }
                catch (Exception ex)
                {
                    item.Exception = ex;
                }
                finally
                {
                    item.Completed.Set();
                }
            }
        }

        /// <summary>Единица работы в очереди.</summary>
        private class WorkItem
        {
            public Func<UIApplication, object> Action { get; }
            public object Result { get; set; }
            public Exception Exception { get; set; }
            public ManualResetEventSlim Completed { get; } = new ManualResetEventSlim(false);

            public WorkItem(Func<UIApplication, object> action) => Action = action;
        }
    }
}
