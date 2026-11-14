using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Suite.Core.Batching
{
    public enum BatchItemStatus { Pending, Succeeded, Blocked, Failed, Cancelled }

    public sealed class BatchItemResult<T>
    {
        internal BatchItemResult(T item) { Item = item; Status = BatchItemStatus.Pending; Message = string.Empty; }
        public T Item { get; private set; }
        public BatchItemStatus Status { get; internal set; }
        public string Message { get; internal set; }
        public string Output { get; internal set; }
    }

    public sealed class BatchExecutionResult<T>
    {
        internal BatchExecutionResult(IList<BatchItemResult<T>> items) { Items = items.ToList().AsReadOnly(); }
        public IReadOnlyList<BatchItemResult<T>> Items { get; private set; }
        public int Succeeded { get { return Items.Count(x => x.Status == BatchItemStatus.Succeeded); } }
        public int Failed { get { return Items.Count(x => x.Status == BatchItemStatus.Failed); } }
        public int Blocked { get { return Items.Count(x => x.Status == BatchItemStatus.Blocked); } }
        public int Cancelled { get { return Items.Count(x => x.Status == BatchItemStatus.Cancelled); } }
        public bool IsCompleteSuccess { get { return Items.All(x => x.Status == BatchItemStatus.Succeeded); } }
    }

    public sealed class BatchExecutor
    {
        public BatchExecutionResult<T> Execute<T>(IEnumerable<T> items, Func<T, string> precheck, Func<T, CancellationToken, string> execute, IProgress<string> progress, CancellationToken cancellationToken)
        {
            if (items == null) throw new ArgumentNullException("items");
            if (execute == null) throw new ArgumentNullException("execute");
            var results = items.Select(x => new BatchItemResult<T>(x)).ToList();
            bool cancelled = false;
            foreach (var result in results)
            {
                if (cancelled || cancellationToken.IsCancellationRequested)
                {
                    result.Status = BatchItemStatus.Cancelled;
                    result.Message = "预检前操作已取消。";
                    cancelled = true;
                    continue;
                }
                try
                {
                    string reason = precheck == null ? null : precheck(result.Item);
                    if (!string.IsNullOrEmpty(reason))
                    {
                        result.Status = BatchItemStatus.Blocked;
                        result.Message = reason;
                    }
                }
                catch (Exception ex)
                {
                    result.Status = BatchItemStatus.Blocked;
                    result.Message = "预检失败：" + ex.Message;
                }
            }

            int completed = 0;
            foreach (var result in results)
            {
                if (result.Status == BatchItemStatus.Blocked) { completed++; continue; }
                if (result.Status == BatchItemStatus.Cancelled) { completed++; continue; }
                if (cancelled || cancellationToken.IsCancellationRequested)
                {
                    result.Status = BatchItemStatus.Cancelled;
                    result.Message = "执行前操作已取消。";
                    cancelled = true;
                    completed++;
                    continue;
                }
                try
                {
                    result.Output = execute(result.Item, cancellationToken);
                    result.Status = BatchItemStatus.Succeeded;
                    result.Message = "成功";
                }
                catch (OperationCanceledException)
                {
                    result.Status = BatchItemStatus.Cancelled;
                    result.Message = "执行中操作已取消。";
                    cancelled = true;
                }
                catch (Exception ex)
                {
                    result.Status = BatchItemStatus.Failed;
                    result.Message = ex.Message;
                }
                completed++;
                if (progress != null) progress.Report(string.Format("已处理 {0}/{1}", completed, results.Count));
            }
            return new BatchExecutionResult<T>(results);
        }
    }
}
