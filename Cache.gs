// Cache.gs
// Caching utilities for improved performance

/**
 * Cache duration in seconds
 * @constant {number}
 */
const CACHE_DURATION = 300; // 5 minutes

/**
 * Maximum number of cache entries
 * @constant {number}
 */
const MAX_CACHE_SIZE = 100;

/**
 * Get cached data with automatic expiration
 * @function getCachedData
 * @param {string} key - The cache key
 * @returns {*} The cached value or null if not found/expired
 * @description Retrieves data from script cache with automatic expiration check
 */
function getCachedData(key) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    
    if (cached) {
      const data = JSON.parse(cached);
      if (data.timestamp && (Date.now() - data.timestamp) < (CACHE_DURATION * 1000)) {
        return data.value;
      }
    }
    return null;
  } catch (error) {
    console.error('Cache get error:', error.message);
    return null;
  }
}

/**
 * Set cached data with timestamp
 * @function setCachedData
 * @param {string} key - The cache key
 * @param {*} value - The value to cache
 * @description Stores data in script cache with timestamp for expiration
 */
function setCachedData(key, value) {
  try {
    const cache = CacheService.getScriptCache();
    const data = {
      value: value,
      timestamp: Date.now()
    };
    
    cache.put(key, JSON.stringify(data), CACHE_DURATION);
  } catch (error) {
    console.error('Cache set error:', error.message);
  }
}

/**
 * Clear specific cache entry
 * @function clearCachedData
 * @param {string} key - The cache key to remove
 * @description Removes a specific entry from the script cache
 */
function clearCachedData(key) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(key);
  } catch (error) {
    console.error('Cache clear error:', error.message);
  }
}

/**
 * Clear all cache
 * @function clearAllCache
 * @description Removes all entries from the script cache
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll();
  } catch (error) {
    console.error('Cache clear all error:', error.message);
  }
}
