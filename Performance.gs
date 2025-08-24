// Performance.gs
// Performance monitoring and optimization utilities

const PERFORMANCE_LOG = [];

/**
 * Start performance timer
 */
function startTimer(operation) {
  return {
    operation: operation,
    startTime: Date.now()
  };
}

/**
 * End performance timer and log result
 */
function endTimer(timer) {
  const duration = Date.now() - timer.startTime;
  PERFORMANCE_LOG.push({
    operation: timer.operation,
    duration: duration,
    timestamp: new Date().toISOString()
  });
  
  // Log slow operations
  if (duration > 1000) {
    console.warn(`Slow operation: ${timer.operation} took ${duration}ms`);
  }
  
  return duration;
}

/**
 * Get performance statistics
 */
function getPerformanceStats() {
  if (PERFORMANCE_LOG.length === 0) {
    return { message: 'No performance data available' };
  }
  
  const operations = {};
  
  PERFORMANCE_LOG.forEach(log => {
    if (!operations[log.operation]) {
      operations[log.operation] = {
        count: 0,
        totalTime: 0,
        avgTime: 0,
        minTime: Infinity,
        maxTime: 0
      };
    }
    
    const op = operations[log.operation];
    op.count++;
    op.totalTime += log.duration;
    op.avgTime = op.totalTime / op.count;
    op.minTime = Math.min(op.minTime, log.duration);
    op.maxTime = Math.max(op.maxTime, log.duration);
  });
  
  return operations;
}

/**
 * Clear performance log
 */
function clearPerformanceLog() {
  PERFORMANCE_LOG.length = 0;
}

/**
 * Batch operations for better performance
 */
function batchOperation(operations, batchSize = 10) {
  const results = [];
  
  for (let i = 0; i < operations.length; i += batchSize) {
    const batch = operations.slice(i, i + batchSize);
    const batchResults = batch.map(op => {
      try {
        return op();
      } catch (error) {
        console.error('Batch operation error:', error.message);
        return null;
      }
    });
    results.push(...batchResults);
  }
  
  return results;
}

/**
 * Debounce function calls
 */
function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}
