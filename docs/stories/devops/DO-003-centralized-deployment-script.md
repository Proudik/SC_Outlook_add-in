# DO-003: Centralized Deployment Script

**Story ID:** DO-003
**Story Points:** 5
**Epic Link:** DevOps & Infrastructure
**Status:** Ready for Development

## Description

Create a centralized, reusable deployment script that handles building, validating, and deploying the Office Outlook Add-in to different environments (development, staging, production). The script should be usable both locally by developers and in CI/CD pipelines, providing consistent deployment logic with proper error handling, rollback capabilities, and deployment verification.

This script abstracts away environment-specific complexity and provides a single entry point for all deployments, reducing errors and ensuring consistency.

## Acceptance Criteria

1. **Single Deployment Command**
   - One script handles all environments: `npm run deploy:dev`, `npm run deploy:staging`, `npm run deploy:production`
   - Environment detected from command argument or environment variable
   - Script validates prerequisites before deployment (credentials, dependencies, etc.)
   - Provides clear, actionable error messages on failure
   - Supports dry-run mode for validation without deployment

2. **Build Process**
   - Clean previous build artifacts
   - Generate environment-specific manifest
   - Validate manifest XML structure
   - Build application bundle (webpack production mode)
   - Optimize assets (minification, tree-shaking, compression)
   - Generate source maps (optional, configurable per environment)
   - Verify build output integrity (file sizes, required files)

3. **Deployment Process**
   - Upload files to static hosting (Azure Blob Storage, AWS S3, or similar)
   - Set appropriate cache headers (long-term for assets, no-cache for HTML/manifest)
   - Upload manifest separately with no-cache headers
   - Support incremental deployments (only changed files)
   - Provide progress indicators during upload
   - Verify deployment success with health checks

4. **Rollback Support**
   - Create backup of current deployment before deploying new version
   - Store backups with timestamps
   - Provide rollback command: `npm run rollback:production`
   - Restore previous deployment from backup
   - Validate rollback success with health checks
   - Clean up old backups (keep last 5 versions)

5. **Validation & Health Checks**
   - Verify all required files are present in build output
   - Check file sizes against limits (max 5MB per bundle)
   - Validate manifest.xml structure and URLs
   - Test deployment URL accessibility (HTTP 200 status)
   - Verify asset loading (icons, CSS, JavaScript)
   - Report validation results before deployment

6. **Logging & Output**
   - Clear progress indicators for each step
   - Detailed error messages with troubleshooting hints
   - Summary report after deployment (files uploaded, URLs, duration)
   - Support verbose mode for debugging
   - Machine-readable output for CI/CD integration
   - Save deployment logs to file

## Technical Requirements

### Directory Structure

```
project-root/
├── scripts/
│   ├── deploy.js                  # Main deployment script
│   ├── deploy-azure.js            # Azure Blob Storage deployer
│   ├── deploy-aws.js              # AWS S3 deployer (optional)
│   ├── rollback.js                # Rollback script
│   ├── health-check.js            # Deployment validation
│   └── utils/
│       ├── logger.js              # Logging utility
│       ├── file-utils.js          # File operations
│       └── config.js              # Deployment configuration
├── .env.dev
├── .env.staging
├── .env.production
└── package.json
```

### Main Deployment Script

**scripts/deploy.js**
```javascript
#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const logger = require('./utils/logger');
const { loadConfig } = require('./utils/config');
const healthCheck = require('./health-check');

/**
 * Centralized deployment script for Office Add-in
 *
 * Usage:
 *   npm run deploy:dev
 *   npm run deploy:staging
 *   npm run deploy:production
 *   node scripts/deploy.js [environment] [--dry-run] [--verbose]
 */

class Deployer {
  constructor(environment, options = {}) {
    this.environment = environment;
    this.options = options;
    this.config = loadConfig(environment);
    this.startTime = Date.now();
  }

  async deploy() {
    try {
      logger.info(`Starting deployment to ${this.environment}...`);

      // Step 1: Validate prerequisites
      await this.validatePrerequisites();

      // Step 2: Clean previous build
      await this.cleanBuild();

      // Step 3: Generate manifest
      await this.generateManifest();

      // Step 4: Build application
      await this.buildApplication();

      // Step 5: Validate build output
      await this.validateBuild();

      // Step 6: Create backup (staging/production only)
      if (this.environment !== 'dev') {
        await this.createBackup();
      }

      // Step 7: Deploy to hosting
      if (!this.options.dryRun) {
        await this.deployToHosting();
      } else {
        logger.info('[DRY RUN] Skipping actual deployment');
      }

      // Step 8: Health check
      if (!this.options.dryRun) {
        await this.performHealthCheck();
      }

      // Step 9: Summary
      this.printSummary();

      logger.success(`Deployment completed successfully!`);
      process.exit(0);
    } catch (error) {
      logger.error('Deployment failed:', error.message);
      if (this.options.verbose) {
        console.error(error.stack);
      }
      process.exit(1);
    }
  }

  async validatePrerequisites() {
    logger.step('Validating prerequisites...');

    // Check Node.js version
    const nodeVersion = process.version;
    const requiredVersion = 'v18.0.0';
    if (nodeVersion < requiredVersion) {
      throw new Error(`Node.js ${requiredVersion} or higher required (current: ${nodeVersion})`);
    }

    // Check npm dependencies
    if (!fs.existsSync(path.join(__dirname, '../node_modules'))) {
      throw new Error('Dependencies not installed. Run: npm install');
    }

    // Check environment configuration
    if (!this.config.baseUrl) {
      throw new Error(`BASE_URL not configured for ${this.environment} environment`);
    }

    // Check deployment credentials
    if (!this.config.storageAccount || !this.config.storageKey) {
      throw new Error(`Storage credentials not configured for ${this.environment}`);
    }

    logger.success('Prerequisites validated');
  }

  async cleanBuild() {
    logger.step('Cleaning previous build...');

    const distPath = path.join(__dirname, '../dist');
    if (fs.existsSync(distPath)) {
      fs.rmSync(distPath, { recursive: true, force: true });
    }

    logger.success('Build directory cleaned');
  }

  async generateManifest() {
    logger.step('Generating manifest...');

    const manifestScript = this.environment === 'production'
      ? 'manifest:production'
      : this.environment === 'staging'
      ? 'manifest:staging'
      : 'manifest:dev';

    execSync(`npm run ${manifestScript}`, {
      cwd: path.join(__dirname, '..'),
      stdio: this.options.verbose ? 'inherit' : 'pipe',
    });

    logger.success('Manifest generated');
  }

  async buildApplication() {
    logger.step('Building application...');

    const buildEnv = {
      NODE_ENV: this.environment === 'dev' ? 'development' : 'production',
      PUBLIC_URL: this.config.baseUrl,
    };

    execSync('npm run build', {
      cwd: path.join(__dirname, '..'),
      stdio: this.options.verbose ? 'inherit' : 'pipe',
      env: { ...process.env, ...buildEnv },
    });

    logger.success('Application built successfully');
  }

  async validateBuild() {
    logger.step('Validating build output...');

    const distPath = path.join(__dirname, '../dist');
    const manifestPath = path.join(__dirname, '../manifest.xml');

    // Check required files
    const requiredFiles = [
      'taskpane.html',
      'taskpane.js',
      'commands.html',
      'commands.js',
    ];

    for (const file of requiredFiles) {
      const filePath = path.join(distPath, file);
      if (!fs.existsSync(filePath)) {
        throw new Error(`Required file missing: ${file}`);
      }
    }

    // Check manifest exists
    if (!fs.existsSync(manifestPath)) {
      throw new Error('manifest.xml not found');
    }

    // Validate manifest XML
    try {
      execSync('npm run manifest:validate', {
        cwd: path.join(__dirname, '..'),
        stdio: 'pipe',
      });
    } catch (error) {
      throw new Error('Manifest validation failed');
    }

    // Check bundle sizes
    const maxSize = 5 * 1024 * 1024; // 5MB
    const taskpaneJs = path.join(distPath, 'taskpane.js');
    if (fs.existsSync(taskpaneJs)) {
      const size = fs.statSync(taskpaneJs).size;
      if (size > maxSize) {
        throw new Error(`Bundle size (${(size / 1024 / 1024).toFixed(2)}MB) exceeds maximum (5MB)`);
      }
    }

    logger.success('Build output validated');
  }

  async createBackup() {
    logger.step('Creating backup of current deployment...');

    const backupName = `backup-${Date.now()}`;

    try {
      // Use Azure CLI or AWS CLI to create backup
      if (this.config.provider === 'azure') {
        const { createBackup } = require('./deploy-azure');
        await createBackup(this.config, backupName);
      } else if (this.config.provider === 'aws') {
        const { createBackup } = require('./deploy-aws');
        await createBackup(this.config, backupName);
      }

      logger.success(`Backup created: ${backupName}`);
      this.backupName = backupName;
    } catch (error) {
      logger.warn('Backup creation failed (continuing anyway):', error.message);
    }
  }

  async deployToHosting() {
    logger.step('Deploying to hosting provider...');

    const distPath = path.join(__dirname, '../dist');
    const manifestPath = path.join(__dirname, '../manifest.xml');

    if (this.config.provider === 'azure') {
      const { deployToAzure } = require('./deploy-azure');
      await deployToAzure(this.config, distPath, manifestPath);
    } else if (this.config.provider === 'aws') {
      const { deployToAWS } = require('./deploy-aws');
      await deployToAWS(this.config, distPath, manifestPath);
    } else {
      throw new Error(`Unknown hosting provider: ${this.config.provider}`);
    }

    logger.success('Files uploaded successfully');
  }

  async performHealthCheck() {
    logger.step('Performing health check...');

    await healthCheck(this.config.baseUrl);

    logger.success('Health check passed');
  }

  printSummary() {
    const duration = ((Date.now() - this.startTime) / 1000).toFixed(2);

    logger.info('\n=== Deployment Summary ===');
    logger.info(`Environment: ${this.environment}`);
    logger.info(`Base URL: ${this.config.baseUrl}`);
    logger.info(`Duration: ${duration}s`);

    if (this.backupName) {
      logger.info(`Backup: ${this.backupName}`);
    }

    logger.info(`\nDeployment URL: ${this.config.baseUrl}/taskpane.html`);
    logger.info(`Manifest URL: ${this.config.baseUrl}/manifest.xml`);
  }
}

// Parse command-line arguments
const args = process.argv.slice(2);
const environment = args[0] || process.env.NODE_ENV || 'dev';
const options = {
  dryRun: args.includes('--dry-run'),
  verbose: args.includes('--verbose'),
};

// Run deployment
const deployer = new Deployer(environment, options);
deployer.deploy();
```

### Azure Blob Storage Deployer

**scripts/deploy-azure.js**
```javascript
const { BlobServiceClient } = require('@azure/storage-blob');
const fs = require('fs');
const path = require('path');
const logger = require('./utils/logger');

async function deployToAzure(config, distPath, manifestPath) {
  const connectionString = `DefaultEndpointsProtocol=https;AccountName=${config.storageAccount};AccountKey=${config.storageKey};EndpointSuffix=core.windows.net`;

  const blobServiceClient = BlobServiceClient.fromConnectionString(connectionString);
  const containerClient = blobServiceClient.getContainerClient(config.containerName);

  // Ensure container exists
  await containerClient.createIfNotExists({ access: 'blob' });

  // Upload dist files
  const files = getAllFiles(distPath);

  for (const file of files) {
    const relativePath = path.relative(distPath, file);
    const blobName = relativePath.replace(/\\/g, '/');
    const blockBlobClient = containerClient.getBlockBlobClient(blobName);

    const contentType = getContentType(file);
    const cacheControl = getCacheControl(file);

    logger.verbose(`Uploading: ${blobName}`);

    await blockBlobClient.uploadFile(file, {
      blobHTTPHeaders: {
        blobContentType: contentType,
        blobCacheControl: cacheControl,
      },
    });
  }

  // Upload manifest (no caching)
  const manifestBlobClient = containerClient.getBlockBlobClient('manifest.xml');
  await manifestBlobClient.uploadFile(manifestPath, {
    blobHTTPHeaders: {
      blobContentType: 'application/xml',
      blobCacheControl: 'no-cache, no-store, must-revalidate',
    },
  });

  logger.success(`Deployed ${files.length + 1} files to Azure Blob Storage`);
}

async function createBackup(config, backupName) {
  const connectionString = `DefaultEndpointsProtocol=https;AccountName=${config.storageAccount};AccountKey=${config.storageKey};EndpointSuffix=core.windows.net`;

  const blobServiceClient = BlobServiceClient.fromConnectionString(connectionString);
  const sourceContainer = blobServiceClient.getContainerClient(config.containerName);
  const backupContainer = blobServiceClient.getContainerClient(backupName);

  // Create backup container
  await backupContainer.createIfNotExists();

  // Copy all blobs
  for await (const blob of sourceContainer.listBlobsFlat()) {
    const sourceBlobClient = sourceContainer.getBlobClient(blob.name);
    const destBlobClient = backupContainer.getBlobClient(blob.name);
    await destBlobClient.beginCopyFromURL(sourceBlobClient.url);
  }

  logger.success(`Backup created: ${backupName}`);
}

function getAllFiles(dirPath, arrayOfFiles = []) {
  const files = fs.readdirSync(dirPath);

  files.forEach((file) => {
    const filePath = path.join(dirPath, file);
    if (fs.statSync(filePath).isDirectory()) {
      arrayOfFiles = getAllFiles(filePath, arrayOfFiles);
    } else {
      arrayOfFiles.push(filePath);
    }
  });

  return arrayOfFiles;
}

function getContentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const contentTypes = {
    '.html': 'text/html',
    '.js': 'application/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.xml': 'application/xml',
  };
  return contentTypes[ext] || 'application/octet-stream';
}

function getCacheControl(filePath) {
  const ext = path.extname(filePath).toLowerCase();

  // Long-term caching for assets with hashes
  if (['.js', '.css', '.png', '.jpg', '.svg'].includes(ext)) {
    return 'public, max-age=31536000, immutable';
  }

  // No caching for HTML files
  if (ext === '.html') {
    return 'no-cache, no-store, must-revalidate';
  }

  return 'public, max-age=3600';
}

module.exports = { deployToAzure, createBackup };
```

### Health Check Script

**scripts/health-check.js**
```javascript
const https = require('https');
const logger = require('./utils/logger');

async function healthCheck(baseUrl) {
  const urlsToCheck = [
    `${baseUrl}/taskpane.html`,
    `${baseUrl}/commands.html`,
    `${baseUrl}/manifest.xml`,
  ];

  logger.info('Running health checks...');

  for (const url of urlsToCheck) {
    await checkUrl(url);
  }
}

function checkUrl(url) {
  return new Promise((resolve, reject) => {
    logger.verbose(`Checking: ${url}`);

    https.get(url, (res) => {
      if (res.statusCode === 200) {
        logger.success(`✓ ${url} (${res.statusCode})`);
        resolve();
      } else {
        reject(new Error(`${url} returned ${res.statusCode}`));
      }
    }).on('error', (error) => {
      reject(new Error(`${url} failed: ${error.message}`));
    });
  });
}

module.exports = healthCheck;
```

### Rollback Script

**scripts/rollback.js**
```javascript
#!/usr/bin/env node

const { BlobServiceClient } = require('@azure/storage-blob');
const logger = require('./utils/logger');
const { loadConfig } = require('./utils/config');
const healthCheck = require('./health-check');

/**
 * Rollback deployment to previous version
 *
 * Usage:
 *   npm run rollback:production [backup-name]
 *   node scripts/rollback.js production [backup-name]
 */

async function rollback(environment, backupName) {
  try {
    logger.info(`Rolling back ${environment} to ${backupName}...`);

    const config = loadConfig(environment);
    const connectionString = `DefaultEndpointsProtocol=https;AccountName=${config.storageAccount};AccountKey=${config.storageKey};EndpointSuffix=core.windows.net`;

    const blobServiceClient = BlobServiceClient.fromConnectionString(connectionString);
    const sourceContainer = blobServiceClient.getContainerClient(backupName);
    const destContainer = blobServiceClient.getContainerClient(config.containerName);

    // Check backup exists
    const backupExists = await sourceContainer.exists();
    if (!backupExists) {
      throw new Error(`Backup not found: ${backupName}`);
    }

    // Delete current deployment
    logger.step('Removing current deployment...');
    for await (const blob of destContainer.listBlobsFlat()) {
      await destContainer.deleteBlob(blob.name);
    }

    // Restore from backup
    logger.step('Restoring from backup...');
    let fileCount = 0;
    for await (const blob of sourceContainer.listBlobsFlat()) {
      const sourceBlobClient = sourceContainer.getBlobClient(blob.name);
      const destBlobClient = destContainer.getBlobClient(blob.name);
      await destBlobClient.beginCopyFromURL(sourceBlobClient.url);
      fileCount++;
    }

    logger.success(`Restored ${fileCount} files from backup`);

    // Health check
    logger.step('Performing health check...');
    await healthCheck(config.baseUrl);

    logger.success('Rollback completed successfully!');
    process.exit(0);
  } catch (error) {
    logger.error('Rollback failed:', error.message);
    process.exit(1);
  }
}

// Parse arguments
const environment = process.argv[2] || 'production';
const backupName = process.argv[3];

if (!backupName) {
  logger.error('Usage: npm run rollback:production <backup-name>');
  process.exit(1);
}

rollback(environment, backupName);
```

### Configuration Utility

**scripts/utils/config.js**
```javascript
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, `../../.env.${process.env.NODE_ENV}`) });

function loadConfig(environment) {
  const envFile = path.join(__dirname, `../../.env.${environment}`);
  require('dotenv').config({ path: envFile });

  return {
    environment,
    baseUrl: process.env.MANIFEST_BASE_URL,
    storageAccount: process.env.AZURE_STORAGE_ACCOUNT,
    storageKey: process.env.AZURE_STORAGE_KEY,
    containerName: process.env.AZURE_CONTAINER_NAME,
    provider: process.env.HOSTING_PROVIDER || 'azure',
  };
}

module.exports = { loadConfig };
```

### NPM Scripts

Add to `package.json`:
```json
{
  "scripts": {
    "deploy:dev": "node scripts/deploy.js dev",
    "deploy:staging": "node scripts/deploy.js staging",
    "deploy:production": "node scripts/deploy.js production",
    "deploy:dry-run": "node scripts/deploy.js production --dry-run",
    "rollback:production": "node scripts/rollback.js production",
    "health-check": "node scripts/health-check.js"
  },
  "devDependencies": {
    "@azure/storage-blob": "^12.17.0",
    "dotenv": "^16.3.1"
  }
}
```

## Environment Configuration

**.env.production**
```env
NODE_ENV=production
MANIFEST_BASE_URL=https://addin.singlecase.com
AZURE_STORAGE_ACCOUNT=singlecaseprod
AZURE_STORAGE_KEY=<your-production-key>
AZURE_CONTAINER_NAME=outlook-addin-prod
HOSTING_PROVIDER=azure
```

## Usage Examples

### Deploy to Development
```bash
npm run deploy:dev
```

### Deploy to Staging (with verbose output)
```bash
node scripts/deploy.js staging --verbose
```

### Deploy to Production (dry run first)
```bash
npm run deploy:dry-run
npm run deploy:production
```

### Rollback Production
```bash
npm run rollback:production backup-1234567890
```

### Health Check Only
```bash
npm run health-check
```

## Dependencies

- **Depends on**: DO-001 (Environment-Specific Manifests)
- **Used by**: DO-002 (CI/CD Pipeline Setup)

## Notes

1. **Idempotent Deployments**: Script can be run multiple times safely
2. **Atomic Operations**: Either all files deploy successfully or none
3. **Progress Indicators**: Clear feedback during long operations
4. **Error Recovery**: Automatic rollback on critical errors

## Definition of Done

- [ ] `scripts/deploy.js` implemented with all deployment steps
- [ ] Azure Blob Storage deployer implemented
- [ ] Health check script validates deployment success
- [ ] Rollback script restores previous deployment
- [ ] Configuration utility loads environment-specific settings
- [ ] Logging utility provides clear output
- [ ] NPM scripts added for all deployment commands
- [ ] Dry-run mode tested and working
- [ ] Verbose mode provides detailed debugging output
- [ ] Prerequisites validation catches common errors
- [ ] Build validation checks required files and sizes
- [ ] Health checks verify deployment accessibility
- [ ] Backup creation tested (staging/production)
- [ ] Rollback tested successfully
- [ ] Documentation written for deployment process
- [ ] Team trained on deployment scripts
