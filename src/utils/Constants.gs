/**
 * Shared constants for CA1-PLM utilities
 *
 * NOTE: The primary configuration lives in CA1_MINI_BOM.gs as BOM_CONFIG.
 * This file provides supplementary constants for the utility layer (cache, performance).
 * Do NOT redeclare BOM_CONFIG or COL here — they are defined in the legacy module.
 */

const UTIL_CONFIG = {
  // Cache settings
  CACHE: {
    DURATION: 600, // 10 minutes in seconds
    KEYS: {
      ITEMS: 'items_cache',
      AML: 'aml_cache',
      MASTER: 'master_cache'
    }
  },

  // Performance settings
  PERFORMANCE: {
    BATCH_SIZE: 100,
    MAX_RETRIES: 3
  }
};
