using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace DeepExcel.AddIn.Performance
{
    /// <summary>
    /// Token预算管理器 - 防止API调用过度消耗
    /// </summary>
    public class TokenBudgetManager
    {
        private static TokenBudgetManager _instance;
        public static TokenBudgetManager Instance => _instance ??= new TokenBudgetManager();

        private readonly Dictionary<string, DailyBudget> _budgets = new();
        private const int DefaultDailyLimit = 500000;  // 50万token/天

        public event Action<string, int, int> OnBudgetWarning;
        public event Action<string, int> OnBudgetExceeded;

        /// <summary>
        /// 设置每日预算
        /// </summary>
        public void SetDailyLimit(string providerKey, int limit)
        {
            if (!_budgets.ContainsKey(providerKey))
                _budgets[providerKey] = new DailyBudget();
            _budgets[providerKey].DailyLimit = limit;
        }

        /// <summary>
        /// 记录Token消耗
        /// </summary>
        public bool RecordUsage(string providerKey, int inputTokens, int outputTokens)
        {
            if (!_budgets.ContainsKey(providerKey))
                _budgets[providerKey] = new DailyBudget(DefaultDailyLimit);

            var budget = _budgets[providerKey];
            var today = DateTime.Now.Date;

            // 跨天重置
            if (budget.LastDate != today)
            {
                budget.LastDate = today;
                budget.TodayUsage = 0;
            }

            var total = inputTokens + outputTokens;
            budget.TodayUsage += total;

            // 预警（80%）
            if (!budget.WarningTriggered && budget.TodayUsage >= budget.DailyLimit * 0.8)
            {
                budget.WarningTriggered = true;
                OnBudgetWarning?.Invoke(providerKey, budget.TodayUsage, budget.DailyLimit);
            }

            // 超限
            if (budget.TodayUsage > budget.DailyLimit)
            {
                OnBudgetExceeded?.Invoke(providerKey, budget.TodayUsage);
                return false;
            }

            return true;
        }

        /// <summary>
        /// 获取当前预算状态
        /// </summary>
        public BudgetStatus GetStatus(string providerKey)
        {
            if (!_budgets.ContainsKey(providerKey))
                return new BudgetStatus { Remaining = DefaultDailyLimit };

            var budget = _budgets[providerKey];
            return new BudgetStatus
            {
                Used = budget.TodayUsage,
                Limit = budget.DailyLimit,
                Remaining = Math.Max(0, budget.DailyLimit - budget.TodayUsage),
                Percentage = (double)budget.TodayUsage / budget.DailyLimit * 100
            };
        }

        /// <summary>
        /// 重置当日预算
        /// </summary>
        public void ResetDaily(string providerKey)
        {
            if (_budgets.ContainsKey(providerKey))
            {
                _budgets[providerKey].TodayUsage = 0;
                _budgets[providerKey].WarningTriggered = false;
            }
        }
    }

    public class DailyBudget
    {
        public DateTime LastDate { get; set; } = DateTime.MinValue;
        public int DailyLimit { get; set; }
        public int TodayUsage { get; set; }
        public bool WarningTriggered { get; set; }

        public DailyBudget(int limit = 500000)
        {
            DailyLimit = limit;
        }
    }

    public class BudgetStatus
    {
        public int Used { get; set; }
        public int Limit { get; set; }
        public int Remaining { get; set; }
        public double Percentage { get; set; }
    }

    /// <summary>
    /// 大数据处理器 - 分批处理避免Excel卡顿
    /// </summary>
    public class BatchProcessor
    {
        public const int DefaultBatchSize = 1000;
        public const int MaxBatchSize = 10000;

        public event Action<int, int> OnBatchProgress;

        /// <summary>
        /// 分批处理
        /// </summary>
        public void ProcessInBatches<T>(List<T> items, Action<List<T>> processAction, int batchSize = DefaultBatchSize)
        {
            if (items == null || items.Count == 0) return;

            batchSize = Math.Min(batchSize, MaxBatchSize);
            int total = items.Count;
            int processed = 0;

            for (int i = 0; i < total; i += batchSize)
            {
                var batch = items.Skip(i).Take(batchSize).ToList();
                processAction(batch);

                processed += batch.Count;
                OnBatchProgress?.Invoke(processed, total);
            }
        }

        /// <summary>
        /// 异步分批处理
        /// </summary>
        public async System.Threading.Tasks.Task ProcessInBatchesAsync<T>(
            List<T> items,
            Func<List<T>, System.Threading.Tasks.Task> processAction,
            int batchSize = DefaultBatchSize)
        {
            if (items == null || items.Count == 0) return;

            batchSize = Math.Min(batchSize, MaxBatchSize);
            int total = items.Count;
            int processed = 0;

            for (int i = 0; i < total; i += batchSize)
            {
                var batch = items.Skip(i).Take(batchSize).ToList();
                await processAction(batch);

                processed += batch.Count;
                OnBatchProgress?.Invoke(processed, total);
            }
        }

        /// <summary>
        /// 并行分批处理（适合独立操作）
        /// </summary>
        public void ProcessInParallelBatches<T>(List<T> items, Action<List<T>> processAction, int batchSize = DefaultBatchSize)
        {
            if (items == null || items.Count == 0) return;

            batchSize = Math.Min(batchSize, MaxBatchSize);
            var batches = new List<List<T>>();

            for (int i = 0; i < items.Count; i += batchSize)
            {
                batches.Add(items.Skip(i).Take(batchSize).ToList());
            }

            System.Threading.Tasks.Parallel.ForEach(batches, batch =>
            {
                processAction(batch);
            });
        }
    }

    /// <summary>
    /// 缓存管理器 - 减少重复数据读取
    /// </summary>
    public class ExcelCache
    {
        private static ExcelCache _instance;
        public static ExcelCache Instance => _instance ??= new ExcelCache();

        private readonly Dictionary<string, CacheEntry> _cache = new();
        private readonly object _lock = new();
        private const int DefaultCacheMinutes = 5;

        /// <summary>
        /// 获取缓存数据
        /// </summary>
        public T Get<T>(string key)
        {
            lock (_lock)
            {
                if (_cache.ContainsKey(key) && !_cache[key].Expired)
                {
                    return (T)_cache[key].Value;
                }
                return default;
            }
        }

        /// <summary>
        /// 设置缓存数据
        /// </summary>
        public void Set<T>(string key, T value, int minutes = DefaultCacheMinutes)
        {
            lock (_lock)
            {
                _cache[key] = new CacheEntry
                {
                    Value = value,
                    ExpiresAt = DateTime.Now.AddMinutes(minutes)
                };
            }
        }

        /// <summary>
        /// 移除缓存
        /// </summary>
        public void Remove(string key)
        {
            lock (_lock)
            {
                _cache.Remove(key);
            }
        }

        /// <summary>
        /// 清理过期缓存
        /// </summary>
        public int Cleanup()
        {
            lock (_lock)
            {
                var expired = _cache.Where(kv => kv.Value.Expired).ToList();
                foreach (var kv in expired)
                {
                    _cache.Remove(kv.Key);
                }
                return expired.Count;
            }
        }

        /// <summary>
        /// 清空所有缓存
        /// </summary>
        public void Clear()
        {
            lock (_lock)
            {
                _cache.Clear();
            }
        }

        /// <summary>
        /// 获取缓存大小
        /// </summary>
        public int Count => _cache.Count;
    }

    public class CacheEntry
    {
        public object Value { get; set; }
        public DateTime ExpiresAt { get; set; }
        public bool Expired => DateTime.Now > ExpiresAt;
    }
}
