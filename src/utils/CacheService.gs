/**
 * Caching service for performance optimization
 */

class CacheManager {
  constructor() {
    this.cache = CacheService.getScriptCache();
  }
  
  /**
   * Get cached data
   * @param {string} key - Cache key
   * @returns {*} Cached data or null
   */
  get(key) {
    const cached = this.cache.get(key);
    if (cached) {
      Logger.log(`Cache HIT: ${key}`);
      return JSON.parse(cached);
    }
    Logger.log(`Cache MISS: ${key}`);
    return null;
  }
  
  /**
   * Set cache data
   * @param {string} key - Cache key
   * @param {*} data - Data to cache
   * @param {number} duration - Cache duration in seconds
   */
  set(key, data, duration = CONFIG.CACHE.DURATION) {
    this.cache.put(key, JSON.stringify(data), duration);
    Logger.log(`Cache SET: ${key} (${duration}s)`);
  }
  
  /**
   * Clear specific cache key
   * @param {string} key - Cache key
   */
  clear(key) {
    this.cache.remove(key);
    Logger.log(`Cache CLEAR: ${key}`);
  }
  
  /**
   * Clear all caches
   */
  clearAll() {
    this.cache.removeAll(Object.values(CONFIG.CACHE.KEYS));
    Logger.log('Cache CLEAR ALL');
  }
}

// Global instance
const cacheManager = new CacheManager();