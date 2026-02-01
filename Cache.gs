/**
 * =================================================================================
 * ENHANCED CACHE SYSTEM WITH UNIFIEDDATAACCESS INTEGRATION (v4.0 - CLEAN)
 * Advanced cache system optimized for UnifiedDataAccess with intelligent invalidation
 * =================================================================================
 */

var ADVANCED_CACHE = {
  memory: {},
  scriptCache: CacheService.getScriptCache(),
  documentProperties: PropertiesService.getDocumentProperties(),
  
  // 🆕 ENHANCED INVALIDATION GROUPS optimized for UnifiedDataAccess
  invalidationGroups: {
    'SHEET_DATA': ['unified_data_', 'sheet_data_', 'column_map_'],
    'UNIFIED_ACCESS': ['unified_data_', 'unified_column_map_'],
    'DASHBOARD_DATA': ['dashboard_', 'mtd_'],
    'YTD_DATA': ['ytd_'],
    'CLIENT_DATA': ['client_']
  },
  
  // Invalidate specific groups of cache instead of everything
  invalidateGroup: function(groupName) {
    var patterns = this.invalidationGroups[groupName] || [];
    
    patterns.forEach(function(pattern) {
      // Remove from memory cache
      Object.keys(this.memory).forEach(function(key) {
        if (key.startsWith(pattern)) {
          delete this.memory[key];
        }
      }.bind(this));
      
      // Note: ScriptCache doesn't support pattern deletion easily without key tracking.
      // For production simplicity/speed, we rely on memory clearing + TTL expiry for ScriptCache,
      // or specific key removal where possible.
    }.bind(this));
  },
  
  // Get data from cache with fallback layers
  get: function(key) {
    // Layer 1: Memory cache (fastest)
    if (this.memory[key] && this.memory[key].expiry > new Date().getTime()) {
      return this.memory[key].data;
    }
    
    // Layer 2: Script cache
    if (CONFIG.CACHE_CONFIG.ENABLE_CACHE) {
      try {
        var scriptCacheData = this.scriptCache.get(key);
        if (scriptCacheData) {
          var parsedData = JSON.parse(scriptCacheData);
          this.memory[key] = {
            data: parsedData,
            expiry: new Date().getTime() + (CONFIG.CACHE_CONFIG.DEFAULT_TTL * 1000)
          };
          return parsedData;
        }
      } catch (error) {
        // Silent fail on cache error
      }
    }
    
    return null;
  },
  
  // 🆕 ENHANCED SET METHOD - with size limits
  set: function(key, data, ttl) {
    var expiry = new Date().getTime() + ((ttl || CONFIG.CACHE_CONFIG.DEFAULT_TTL) * 1000);
    
    try {
      // Layer 1: Memory cache
      this.memory[key] = {
        data: data,
        expiry: expiry
      };
      
      // Layer 2: Script cache (with size limits)
      if (CONFIG.CACHE_CONFIG.ENABLE_CACHE) {
        var dataString = JSON.stringify(data);
        if (dataString.length < 100000) { // 100KB limit
          this.scriptCache.put(key, dataString, ttl || CONFIG.CACHE_CONFIG.DEFAULT_TTL);
        }
      }
      
      this.manageCacheSize();
      
    } catch (error) {
      // Silent fail
    }
  },
  
  // Remove data from all cache layers
  remove: function(key) {
    delete this.memory[key];
    try {
      this.scriptCache.remove(key);
    } catch (error) {
      // Silent fail
    }
  },
  
  // Clear all cache layers
  clearAll: function() {
    this.memory = {};
    try {
      this.scriptCache.removeAll(Object.keys(this.memory)); // Attempt to remove known keys
      // Also try to remove common keys explicitly
      this.scriptCache.removeAll(['unified_data', 'dashboard_data', 'ytd_data']);
    } catch (error) {
      // Silent fail
    }
  },
  
  // Manage cache size to prevent memory issues
  manageCacheSize: function() {
    var keys = Object.keys(this.memory);
    if (keys.length > CONFIG.CACHE_CONFIG.MAX_CACHE_KEYS) {
      // Simple FIFO eviction if too many keys
      var keysToRemove = keys.slice(0, keys.length - CONFIG.CACHE_CONFIG.MAX_CACHE_KEYS);
      keysToRemove.forEach(function(key) {
        delete this.memory[key];
      }.bind(this));
    }
  },
  
  // Get cache statistics
  getStats: function() {
    return {
      memoryKeys: Object.keys(this.memory).length,
      cacheEnabled: CONFIG.CACHE_CONFIG.ENABLE_CACHE,
      maxKeys: CONFIG.CACHE_CONFIG.MAX_CACHE_KEYS
    };
  }
};

/**
 * UTILITY: Clear System Caches
 */
function clearAllSystemCaches() {
  ADVANCED_CACHE.clearAll();
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ System caches cleared", "Success", 3);
}
