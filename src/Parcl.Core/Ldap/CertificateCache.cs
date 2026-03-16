using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Parcl.Core.Models;

namespace Parcl.Core.Ldap
{
    public class CertificateCache
    {
        private readonly ConcurrentDictionary<string, CachedEntry> _cache = new();
        private readonly string _cacheFile;
        private readonly int _expirationHours;
        private readonly int _maxEntries;

        public CertificateCache(int expirationHours = 24, int maxEntries = 500)
        {
            _expirationHours = expirationHours;
            _maxEntries = maxEntries;
            _cacheFile = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Parcl", "cert-cache.json");
            Load();
        }

        public void Add(string email, List<CertificateInfo> certs)
        {
            var entry = new CachedEntry
            {
                Email = email.ToLowerInvariant(),
                Certificates = certs,
                CachedAt = DateTime.UtcNow
            };

            _cache.AddOrUpdate(entry.Email, entry, (_, _) => entry);
            Evict();
            Save();
        }

        public List<CertificateInfo>? Get(string email)
        {
            var key = email.ToLowerInvariant();
            if (!_cache.TryGetValue(key, out var entry))
                return null;

            if (entry.CachedAt.AddHours(_expirationHours) < DateTime.UtcNow)
            {
                _cache.TryRemove(key, out _);
                return null;
            }

            return entry.Certificates;
        }

        public void Remove(string email)
        {
            _cache.TryRemove(email.ToLowerInvariant(), out _);
            Save();
        }

        public void Clear()
        {
            _cache.Clear();
            Save();
        }

        private void Evict()
        {
            if (_cache.Count <= _maxEntries) return;

            var oldest = _cache.OrderBy(kv => kv.Value.CachedAt)
                .Take(_cache.Count - _maxEntries)
                .Select(kv => kv.Key)
                .ToList();

            foreach (var key in oldest)
                _cache.TryRemove(key, out _);
        }

        private void Load()
        {
            if (!File.Exists(_cacheFile)) return;
            try
            {
                var json = File.ReadAllText(_cacheFile);
                var entries = JsonConvert.DeserializeObject<List<CachedEntry>>(json);
                if (entries == null) return;
                foreach (var entry in entries)
                    _cache.TryAdd(entry.Email, entry);
            }
            catch { }
        }

        private void Save()
        {
            try
            {
                var dir = Path.GetDirectoryName(_cacheFile)!;
                Directory.CreateDirectory(dir);
                var json = JsonConvert.SerializeObject(_cache.Values.ToList(), Formatting.Indented);
                File.WriteAllText(_cacheFile, json);
            }
            catch { }
        }

        private class CachedEntry
        {
            public string Email { get; set; } = string.Empty;
            public List<CertificateInfo> Certificates { get; set; } = new();
            public DateTime CachedAt { get; set; }
        }
    }
}
