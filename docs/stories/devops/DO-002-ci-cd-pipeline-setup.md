# DO-002: CI/CD Pipeline Setup

**Story ID:** DO-002
**Story Points:** 8
**Epic Link:** DevOps & Infrastructure
**Status:** Ready for Development

## Description

Establish a comprehensive CI/CD pipeline for the Office Outlook Add-in using GitHub Actions (or Azure DevOps). The pipeline should automate building, testing, validation, and deployment across development, staging, and production environments. This ensures consistent, reliable deployments with proper quality gates and automated rollback capabilities.

The pipeline should handle Office Add-in specific requirements including manifest validation, asset optimization, and deployment to static hosting (Azure Blob Storage, AWS S3, or similar).

## Acceptance Criteria

1. **Continuous Integration (CI)**
   - Run on every push to `main`, `develop`, and pull requests
   - Install dependencies with caching for faster builds
   - Run linter (ESLint) and fail on errors
   - Run TypeScript type checking
   - Generate and validate Office Add-in manifest
   - Build application bundle (webpack production mode)
   - Run unit tests (if available)
   - Generate build artifacts
   - Report build status to pull requests

2. **Continuous Deployment (CD) - Development**
   - Trigger on push to `develop` branch
   - Build with development configuration
   - Generate dev manifest
   - Deploy to development environment (auto-deploy)
   - Update manifest version with build number
   - Notify team on Slack/Teams/Email
   - No manual approval required

3. **Continuous Deployment (CD) - Staging**
   - Trigger on push to `staging` branch or manual workflow dispatch
   - Build with staging configuration
   - Generate staging manifest
   - Run integration tests (if available)
   - Deploy to staging environment
   - Require manual approval before deployment
   - Update manifest in Microsoft 365 Admin Center (optional automation)
   - Notify team with deployment details

4. **Continuous Deployment (CD) - Production**
   - Trigger on push to `main` branch or release tag (e.g., `v1.0.0`)
   - Build with production configuration
   - Generate production manifest
   - Run full test suite
   - Require two-person approval before deployment
   - Deploy to production environment
   - Create GitHub release with changelog
   - Upload manifest to Microsoft 365 Admin Center (manual or automated)
   - Notify team with release notes
   - Tag deployment with version number

5. **Build Artifacts**
   - Bundle JavaScript/CSS assets
   - Optimize images and icons
   - Generate source maps (staging/production)
   - Create manifest.xml for environment
   - Package all artifacts in deployment archive
   - Upload artifacts to GitHub Actions or Azure DevOps
   - Retain artifacts for 90 days

6. **Quality Gates**
   - Linting must pass (no errors, warnings allowed)
   - TypeScript compilation must succeed
   - Manifest validation must pass
   - Build must complete without errors
   - File size limits enforced (max 5MB per bundle)
   - No secrets or .env files in build output
   - Security scanning (npm audit, Snyk, or similar)

7. **Rollback Strategy**
   - Keep last 5 production deployments
   - One-click rollback to previous version
   - Automatic rollback on health check failure
   - Rollback notification to team
   - Document rollback procedure

## Technical Requirements

### Pipeline Architecture

```
┌─────────────────┐
│  Code Push      │
│  (Git)          │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  CI Pipeline    │
│  - Install      │
│  - Lint         │
│  - Type Check   │
│  - Build        │
│  - Test         │
│  - Validate     │
└────────┬────────┘
         │
         ├──────────────┬──────────────┬──────────────┐
         ▼              ▼              ▼              ▼
    ┌────────┐    ┌────────┐    ┌────────┐    ┌────────┐
    │  Dev   │    │ Staging│    │  Prod  │    │   PR   │
    │ Deploy │    │ Deploy │    │ Deploy │    │ Check  │
    │ (Auto) │    │(Approve)│   │(Approve)│   │(Report)│
    └────────┘    └────────┘    └────────┘    └────────┘
```

### GitHub Actions Workflow

**File: `.github/workflows/ci.yml`**
```yaml
name: CI - Build and Test

on:
  push:
    branches:
      - main
      - develop
      - staging
  pull_request:
    branches:
      - main
      - develop

jobs:
  build-and-test:
    name: Build and Test
    runs-on: ubuntu-latest

    strategy:
      matrix:
        node-version: [18.x, 20.x]

    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Setup Node.js ${{ matrix.node-version }}
        uses: actions/setup-node@v4
        with:
          node-version: ${{ matrix.node-version }}
          cache: 'npm'

      - name: Install dependencies
        run: npm ci

      - name: Run linter
        run: npm run lint

      - name: Run type check
        run: npx tsc --noEmit

      - name: Generate manifest
        run: npm run manifest:dev

      - name: Validate manifest
        run: npm run manifest:validate

      - name: Build application
        run: npm run build
        env:
          NODE_ENV: production

      - name: Run tests
        run: npm test
        continue-on-error: true

      - name: Security audit
        run: npm audit --audit-level=high
        continue-on-error: true

      - name: Check bundle size
        run: |
          BUNDLE_SIZE=$(du -sb dist/taskpane.js | cut -f1)
          MAX_SIZE=$((5 * 1024 * 1024))  # 5MB
          if [ $BUNDLE_SIZE -gt $MAX_SIZE ]; then
            echo "Bundle size ($BUNDLE_SIZE bytes) exceeds maximum ($MAX_SIZE bytes)"
            exit 1
          fi

      - name: Upload build artifacts
        uses: actions/upload-artifact@v4
        with:
          name: build-artifacts-${{ matrix.node-version }}
          path: |
            dist/
            manifest.xml
          retention-days: 90

      - name: Report build status
        if: always()
        uses: actions/github-script@v7
        with:
          script: |
            const status = '${{ job.status }}';
            const message = status === 'success'
              ? '✅ Build and tests passed'
              : '❌ Build or tests failed';

            github.rest.repos.createCommitStatus({
              owner: context.repo.owner,
              repo: context.repo.repo,
              sha: context.sha,
              state: status,
              description: message,
              context: 'CI Build'
            });
```

**File: `.github/workflows/deploy-dev.yml`**
```yaml
name: CD - Deploy to Development

on:
  push:
    branches:
      - develop

env:
  NODE_ENV: development
  AZURE_STORAGE_ACCOUNT: ${{ secrets.DEV_STORAGE_ACCOUNT }}
  AZURE_STORAGE_KEY: ${{ secrets.DEV_STORAGE_KEY }}
  CONTAINER_NAME: outlook-addin-dev

jobs:
  deploy-dev:
    name: Deploy to Development
    runs-on: ubuntu-latest

    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20.x'
          cache: 'npm'

      - name: Install dependencies
        run: npm ci

      - name: Generate dev manifest
        run: npm run manifest:dev
        env:
          BUILD_NUMBER: ${{ github.run_number }}

      - name: Build application
        run: npm run build

      - name: Deploy to Azure Blob Storage
        uses: azure/CLI@v1
        with:
          azcliversion: latest
          inlineScript: |
            # Upload all files to Azure Blob Storage
            az storage blob upload-batch \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --destination ${{ env.CONTAINER_NAME }} \
              --source ./dist \
              --overwrite \
              --content-cache-control "public, max-age=300"

            # Upload manifest separately (no caching)
            az storage blob upload \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --container-name ${{ env.CONTAINER_NAME }} \
              --file ./manifest.xml \
              --name manifest.xml \
              --overwrite \
              --content-cache-control "no-cache"

      - name: Notify team (Slack)
        if: always()
        uses: slackapi/slack-github-action@v1
        with:
          webhook: ${{ secrets.SLACK_WEBHOOK_URL }}
          payload: |
            {
              "text": "Dev Deployment ${{ job.status }}",
              "blocks": [
                {
                  "type": "section",
                  "text": {
                    "type": "mrkdwn",
                    "text": "*Development Deployment*\nStatus: ${{ job.status }}\nBranch: ${{ github.ref_name }}\nCommit: ${{ github.sha }}\nAuthor: ${{ github.actor }}\nURL: https://dev-addin.singlecase.com"
                  }
                }
              ]
            }
```

**File: `.github/workflows/deploy-staging.yml`**
```yaml
name: CD - Deploy to Staging

on:
  push:
    branches:
      - staging
  workflow_dispatch:

env:
  NODE_ENV: staging
  AZURE_STORAGE_ACCOUNT: ${{ secrets.STAGING_STORAGE_ACCOUNT }}
  AZURE_STORAGE_KEY: ${{ secrets.STAGING_STORAGE_KEY }}
  CONTAINER_NAME: outlook-addin-staging

jobs:
  build-staging:
    name: Build Staging
    runs-on: ubuntu-latest

    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20.x'
          cache: 'npm'

      - name: Install dependencies
        run: npm ci

      - name: Generate staging manifest
        run: npm run manifest:staging

      - name: Validate manifest
        run: npm run manifest:validate

      - name: Build application
        run: npm run build

      - name: Run integration tests
        run: npm run test:integration
        continue-on-error: true

      - name: Upload staging artifacts
        uses: actions/upload-artifact@v4
        with:
          name: staging-build
          path: |
            dist/
            manifest.xml

  deploy-staging:
    name: Deploy to Staging
    runs-on: ubuntu-latest
    needs: build-staging
    environment:
      name: staging
      url: https://staging-addin.singlecase.com

    steps:
      - name: Download artifacts
        uses: actions/download-artifact@v4
        with:
          name: staging-build

      - name: Deploy to Azure Blob Storage
        uses: azure/CLI@v1
        with:
          azcliversion: latest
          inlineScript: |
            az storage blob upload-batch \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --destination ${{ env.CONTAINER_NAME }} \
              --source ./dist \
              --overwrite \
              --content-cache-control "public, max-age=3600"

            az storage blob upload \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --container-name ${{ env.CONTAINER_NAME }} \
              --file ./manifest.xml \
              --name manifest.xml \
              --overwrite \
              --content-cache-control "no-cache"

      - name: Notify team (Slack)
        if: always()
        uses: slackapi/slack-github-action@v1
        with:
          webhook: ${{ secrets.SLACK_WEBHOOK_URL }}
          payload: |
            {
              "text": "Staging Deployment ${{ job.status }}",
              "blocks": [
                {
                  "type": "section",
                  "text": {
                    "type": "mrkdwn",
                    "text": "*Staging Deployment*\nStatus: ${{ job.status }}\nBranch: ${{ github.ref_name }}\nCommit: ${{ github.sha }}\nDeployed by: ${{ github.actor }}\nURL: https://staging-addin.singlecase.com"
                  }
                }
              ]
            }
```

**File: `.github/workflows/deploy-production.yml`**
```yaml
name: CD - Deploy to Production

on:
  push:
    tags:
      - 'v*.*.*'
  workflow_dispatch:

env:
  NODE_ENV: production
  AZURE_STORAGE_ACCOUNT: ${{ secrets.PROD_STORAGE_ACCOUNT }}
  AZURE_STORAGE_KEY: ${{ secrets.PROD_STORAGE_KEY }}
  CONTAINER_NAME: outlook-addin-prod

jobs:
  build-production:
    name: Build Production
    runs-on: ubuntu-latest

    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20.x'
          cache: 'npm'

      - name: Install dependencies
        run: npm ci

      - name: Generate production manifest
        run: npm run manifest:production

      - name: Validate manifest
        run: npm run manifest:validate

      - name: Build application
        run: npm run build

      - name: Run full test suite
        run: npm test

      - name: Security scan
        run: npm audit --audit-level=moderate

      - name: Upload production artifacts
        uses: actions/upload-artifact@v4
        with:
          name: production-build
          path: |
            dist/
            manifest.xml

  deploy-production:
    name: Deploy to Production
    runs-on: ubuntu-latest
    needs: build-production
    environment:
      name: production
      url: https://addin.singlecase.com

    steps:
      - name: Download artifacts
        uses: actions/download-artifact@v4
        with:
          name: production-build

      - name: Deploy to Azure Blob Storage
        uses: azure/CLI@v1
        with:
          azcliversion: latest
          inlineScript: |
            # Create backup of current production
            BACKUP_CONTAINER="${{ env.CONTAINER_NAME }}-backup-$(date +%Y%m%d-%H%M%S)"
            az storage container create \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --name $BACKUP_CONTAINER

            az storage blob copy start-batch \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --source-container ${{ env.CONTAINER_NAME }} \
              --destination-container $BACKUP_CONTAINER

            # Deploy new version
            az storage blob upload-batch \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --destination ${{ env.CONTAINER_NAME }} \
              --source ./dist \
              --overwrite \
              --content-cache-control "public, max-age=31536000, immutable"

            az storage blob upload \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --container-name ${{ env.CONTAINER_NAME }} \
              --file ./manifest.xml \
              --name manifest.xml \
              --overwrite \
              --content-cache-control "no-cache"

      - name: Health check
        run: |
          sleep 10
          HEALTH_CHECK=$(curl -s -o /dev/null -w "%{http_code}" https://addin.singlecase.com/taskpane.html)
          if [ $HEALTH_CHECK -ne 200 ]; then
            echo "Health check failed with status $HEALTH_CHECK"
            exit 1
          fi

      - name: Create GitHub Release
        uses: actions/create-release@v1
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
        with:
          tag_name: ${{ github.ref_name }}
          release_name: Release ${{ github.ref_name }}
          body: |
            Production deployment of version ${{ github.ref_name }}

            ## Changes
            See CHANGELOG.md for details

            ## Deployment
            - Deployed to: https://addin.singlecase.com
            - Manifest version: ${{ github.ref_name }}
          draft: false
          prerelease: false

      - name: Notify team (Slack)
        if: always()
        uses: slackapi/slack-github-action@v1
        with:
          webhook: ${{ secrets.SLACK_WEBHOOK_URL }}
          payload: |
            {
              "text": "🚀 Production Deployment ${{ job.status }}",
              "blocks": [
                {
                  "type": "section",
                  "text": {
                    "type": "mrkdwn",
                    "text": "*Production Deployment*\nStatus: ${{ job.status }}\nVersion: ${{ github.ref_name }}\nDeployed by: ${{ github.actor }}\nURL: https://addin.singlecase.com"
                  }
                }
              ]
            }
```

**File: `.github/workflows/rollback.yml`**
```yaml
name: Rollback Production

on:
  workflow_dispatch:
    inputs:
      backup_container:
        description: 'Backup container name to restore from'
        required: true
        type: string

env:
  AZURE_STORAGE_ACCOUNT: ${{ secrets.PROD_STORAGE_ACCOUNT }}
  AZURE_STORAGE_KEY: ${{ secrets.PROD_STORAGE_KEY }}
  CONTAINER_NAME: outlook-addin-prod

jobs:
  rollback:
    name: Rollback to Previous Version
    runs-on: ubuntu-latest
    environment:
      name: production

    steps:
      - name: Restore from backup
        uses: azure/CLI@v1
        with:
          azcliversion: latest
          inlineScript: |
            echo "Rolling back to backup: ${{ github.event.inputs.backup_container }}"

            # Delete current production files
            az storage blob delete-batch \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --source ${{ env.CONTAINER_NAME }}

            # Restore from backup
            az storage blob copy start-batch \
              --account-name ${{ env.AZURE_STORAGE_ACCOUNT }} \
              --account-key ${{ env.AZURE_STORAGE_KEY }} \
              --source-container ${{ github.event.inputs.backup_container }} \
              --destination-container ${{ env.CONTAINER_NAME }}

            echo "Rollback completed"

      - name: Health check
        run: |
          sleep 10
          HEALTH_CHECK=$(curl -s -o /dev/null -w "%{http_code}" https://addin.singlecase.com/taskpane.html)
          if [ $HEALTH_CHECK -ne 200 ]; then
            echo "Health check failed after rollback"
            exit 1
          fi

      - name: Notify team
        uses: slackapi/slack-github-action@v1
        with:
          webhook: ${{ secrets.SLACK_WEBHOOK_URL }}
          payload: |
            {
              "text": "⚠️ Production Rollback ${{ job.status }}",
              "blocks": [
                {
                  "type": "section",
                  "text": {
                    "type": "mrkdwn",
                    "text": "*Production Rollback*\nStatus: ${{ job.status }}\nBackup: ${{ github.event.inputs.backup_container }}\nTriggered by: ${{ github.actor }}"
                  }
                }
              ]
            }
```

### Azure DevOps Alternative

If using Azure DevOps instead of GitHub Actions:

**azure-pipelines.yml**
```yaml
trigger:
  branches:
    include:
      - main
      - develop
      - staging

pool:
  vmImage: 'ubuntu-latest'

variables:
  nodeVersion: '20.x'

stages:
  - stage: Build
    jobs:
      - job: BuildAndTest
        steps:
          - task: NodeTool@0
            inputs:
              versionSpec: $(nodeVersion)

          - script: npm ci
            displayName: 'Install dependencies'

          - script: npm run lint
            displayName: 'Run linter'

          - script: npm run build
            displayName: 'Build application'

          - task: PublishBuildArtifacts@1
            inputs:
              pathToPublish: 'dist'
              artifactName: 'build-output'

  - stage: DeployDev
    condition: eq(variables['Build.SourceBranch'], 'refs/heads/develop')
    jobs:
      - deployment: DeployToDev
        environment: 'development'
        strategy:
          runOnce:
            deploy:
              steps:
                - task: AzureCLI@2
                  inputs:
                    azureSubscription: 'Azure Subscription'
                    scriptType: 'bash'
                    scriptLocation: 'inlineScript'
                    inlineScript: |
                      az storage blob upload-batch \
                        --account-name $(DEV_STORAGE_ACCOUNT) \
                        --destination outlook-addin-dev \
                        --source $(Pipeline.Workspace)/build-output
```

## Environment Setup

### GitHub Secrets

Configure the following secrets in GitHub repository settings:

**Development Environment**
- `DEV_STORAGE_ACCOUNT`: Azure Storage account name for dev
- `DEV_STORAGE_KEY`: Azure Storage account key for dev

**Staging Environment**
- `STAGING_STORAGE_ACCOUNT`: Azure Storage account name for staging
- `STAGING_STORAGE_KEY`: Azure Storage account key for staging

**Production Environment**
- `PROD_STORAGE_ACCOUNT`: Azure Storage account name for production
- `PROD_STORAGE_KEY`: Azure Storage account key for production

**Notification**
- `SLACK_WEBHOOK_URL`: Slack webhook URL for deployment notifications

### Branch Protection Rules

Configure branch protection in GitHub:

**main branch (production)**
- Require pull request reviews before merging (2 approvals)
- Require status checks to pass (CI build)
- Require branches to be up to date before merging
- Include administrators in restrictions

**staging branch**
- Require pull request reviews before merging (1 approval)
- Require status checks to pass (CI build)

**develop branch**
- Require status checks to pass (CI build)
- No approval required (auto-deploy)

### Environment Protection Rules

Configure environment protection in GitHub:

**production**
- Required reviewers: 2 team leads
- Wait timer: 5 minutes (allows cancellation)

**staging**
- Required reviewers: 1 developer
- No wait timer

**development**
- No protection (auto-deploy)

## Deployment Targets

### Azure Blob Storage (Static Website Hosting)

1. Create Azure Storage Account
2. Enable static website hosting
3. Configure custom domain (optional)
4. Set CORS rules for Office Add-in

### AWS S3 Alternative

1. Create S3 bucket
2. Enable static website hosting
3. Configure bucket policy for public read access
4. Set CORS configuration

### CDN Configuration (Optional)

- Azure CDN or CloudFlare for global distribution
- Cache-Control headers for optimal caching
- Invalidate cache on deployment

## Testing Strategy

### CI Testing
- Unit tests (Jest/Mocha)
- Linting (ESLint)
- Type checking (TypeScript)
- Manifest validation

### Staging Testing
- Integration tests
- Manual QA testing
- User acceptance testing (UAT)

### Production Testing
- Health checks after deployment
- Smoke tests
- Rollback readiness

## Monitoring & Alerting

### Build Monitoring
- Build success rate
- Build duration trends
- Failed build alerts

### Deployment Monitoring
- Deployment frequency
- Deployment success rate
- Rollback frequency
- Failed deployment alerts

### Application Monitoring
- Page load times
- Error rates
- User adoption metrics

## Dependencies

- **Depends on**: DO-001 (Environment-Specific Manifests)
- **Required for**: DO-003 (Centralized Deployment Script), DO-005 (Production Monitoring Setup)

## Notes

1. **Deployment Frequency**: Aim for daily deployments to dev, weekly to staging, bi-weekly to production

2. **Zero-Downtime Deployments**: Office Add-ins don't support blue-green deployments, but static hosting allows instant updates

3. **Manifest Updates**: Manifest updates take 24 hours to propagate in Office clients

4. **Security**: Never commit secrets to repository. Use GitHub Secrets or Azure Key Vault

## Definition of Done

- [ ] GitHub Actions workflows created for CI/CD
- [ ] CI pipeline runs on all branches and PRs
- [ ] Development auto-deploys on push to develop branch
- [ ] Staging deploys with manual approval
- [ ] Production deploys with two-person approval
- [ ] Build artifacts uploaded and retained for 90 days
- [ ] Quality gates enforced (linting, type checking, validation)
- [ ] Rollback workflow implemented and tested
- [ ] GitHub Secrets configured for all environments
- [ ] Branch protection rules enabled
- [ ] Environment protection rules configured
- [ ] Slack notifications working for all deployments
- [ ] Health checks implemented and tested
- [ ] Security scanning integrated (npm audit)
- [ ] Documentation written for CI/CD process
- [ ] Team trained on deployment procedures
