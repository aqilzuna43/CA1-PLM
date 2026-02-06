/**
 * Configuration constants for CA1-PLM
 */

const CONFIG = {
  // Spreadsheet IDs
  MASTER_SHEET_ID: '12016271_CA1_Mini_BoM', // Update with actual ID
  
  // Sheet names
  SHEETS: {
    ITEMS: 'ITEMS',
    AML: 'AML',
    MASTER: 'MASTER'
  },
  
  // Column indices (0-based)
  COLUMNS: {
    MASTER: {
      LEVEL: 0,
      PART_NUMBER: 1,
      DESCRIPTION: 2,
      QUANTITY: 3,
      REF_DES: 4,
      REVISION: 5
    },
    ITEMS: {
      PART_NUMBER: 0,
      DESCRIPTION: 1,
      REVISION: 2,
      LIFECYCLE: 3
    },
    AML: {
      PART_NUMBER: 0,
      MFR_NAME: 1,
      MFR_PART_NUMBER: 2
    }
  },
  
  // ECR Actions
  ECR_ACTIONS: {
    ADD: 'ADD',
    REMOVE: 'REMOVE',
    QTY_CHANGE: 'QTY CHANGE',
    REV_ROLL: 'REV ROLL'
  },
  
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