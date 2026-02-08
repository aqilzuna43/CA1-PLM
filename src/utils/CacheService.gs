/**
 * Caching service for performance optimization.
 * Wraps Google Apps Script CacheService with convenience methods.
 *
 * Uses UTIL_CONFIG.CACHE from Constants.gs for default duration and key names.
 */

class CacheManager {
  constructor() {
    this.cache = CacheService.getScriptCache();
  }

  /**
   * Get cached data.
   * @param {string} key Cache key.
   * @returns {*} Parsed cached data or null if not found / expired.
   */
  get(key) {
    try {
      const cached = this.cache.get(key);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (e) {
      // JSON.parse failure or cache read error — treat as miss
      Logger.log(`Cache GET error for "${key}": ${e.message}`);
    }
    return null;
  }

  /**
   * Set cache data.
   * @param {string} key Cache key.
   * @param {*} data Data to cache (must be JSON-serializable).
   * @param {number} [duration] Cache duration in seconds (default from UTIL_CONFIG).
   */
  set(key, data, duration) {
    const ttl = duration || UTIL_CONFIG.CACHE.DURATION;
    try {
      const serialized = JSON.stringify(data);
      // CacheService has a 100 KB limit per key. Skip silently if too large.
      if (serialized.length > 100000) {
        Logger.log(`Cache SET skipped for "${key}": data exceeds 100 KB limit (${(serialized.length / 1024).toFixed(1)} KB)`);
        return;
      }
      this.cache.put(key, serialized, ttl);
    } catch (e) {
      Logger.log(`Cache SET error for "${key}": ${e.message}`);
    }
  }

  /**
   * Clear a specific cache key.
   * @param {string} key Cache key to remove.
   */
  clear(key) {
    this.cache.remove(key);
  }

  /**
   * Clear all known cache keys.
   */
  clearAll() {
    this.cache.removeAll(Object.values(UTIL_CONFIG.CACHE.KEYS));
  }
}

// Global singleton instance
const cacheManager = new CacheManager();
