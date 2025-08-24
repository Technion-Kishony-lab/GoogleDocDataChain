// Cache.gs
// Caching utilities for improved performance

const CACHE_DURATION = 300; // 5 minutes
const MAX_CACHE_SIZE = 100;

/**
 * Get cached data with automatic expiration
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
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll();
  } catch (error) {
    console.error('Cache clear all error:', error.message);
  }
}
