# DO-001: Environment-Specific Manifests

**Story ID:** DO-001
**Story Points:** 3
**Epic Link:** DevOps & Infrastructure
**Status:** Ready for Development

## Description

Create environment-specific Office Add-in manifest files for development, staging, and production environments. Office Add-ins require XML manifest files that define metadata, permissions, URLs, and resource locations. Currently, the project uses a single `manifest.xml` with hardcoded localhost URLs, which is not suitable for production deployment.

This story establishes a foundation for environment-based deployments by creating separate manifest configurations and a build process to generate the appropriate manifest for each environment.

## Acceptance Criteria

1. **Multiple Manifest Templates**
   - Create `manifest.dev.xml` for local development (localhost:3000)
   - Create `manifest.staging.xml` for staging environment
   - Create `manifest.production.xml` for production environment
   - Each manifest uses environment-specific URLs and settings
   - Maintain consistent add-in ID across environments (or use environment-specific IDs if needed)

2. **Environment-Specific Configuration**
   - Development: Uses `https://localhost:3000` for all URLs
   - Staging: Uses `https://staging-addin.singlecase.com` (or similar)
   - Production: Uses `https://addin.singlecase.com` (or production domain)
   - Each environment has appropriate AppDomains configured
   - Version numbers increment independently per environment

3. **Manifest Generation Script**
   - Create `scripts/generate-manifest.js` to process manifest templates
   - Support environment variable substitution (URLs, version, app ID)
   - Validate generated manifest XML structure
   - Copy generated manifest to project root as `manifest.xml`
   - Add npm scripts for manifest generation per environment

4. **Version Management**
   - Use semantic versioning (MAJOR.MINOR.PATCH)
   - Development: Auto-increment patch version on each build
   - Staging: Manual version updates before deployment
   - Production: Strict version control with changelog
   - Version must match between `manifest.xml` and `package.json`

5. **Asset URL Configuration**
   - Icon URLs point to correct environment domains
   - TaskPane URLs include environment-specific paths
   - Commands HTML file URLs are environment-specific
   - Support for CDN URLs in production (optional)
   - Add cache-busting query parameters (e.g., `?v=1.0.0`)

6. **Validation & Testing**
   - Run Office manifest validation for each environment
   - Verify all URLs are reachable before deployment
   - Ensure HTTPS is enforced for staging and production
   - Test manifest sideloading in Outlook (dev environment)
   - Document manifest differences between environments

## Technical Requirements

### Directory Structure

```
project-root/
├── manifests/
│   ├── manifest.dev.xml          # Development template
│   ├── manifest.staging.xml      # Staging template
│   └── manifest.production.xml   # Production template
├── scripts/
│   ├── generate-manifest.js      # Manifest generation script
│   └── validate-manifest.js      # Manifest validation script
├── manifest.xml                   # Generated manifest (gitignored)
└── .env.{environment}             # Environment variables
```

### Manifest Template Format

**manifest.dev.xml** (Development)
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xsi:type="MailApp">

  <Id>{{APP_ID}}</Id>
  <Version>{{VERSION}}</Version>
  <ProviderName>SingleCase Outlook Add-in</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>

  <DisplayName DefaultValue="SingleCase (Dev)"/>
  <Description DefaultValue="Attach emails and attachments from Outlook directly to cases in SingleCase (Development Environment)."/>

  <IconUrl DefaultValue="{{BASE_URL}}/assets/sc.png"/>
  <HighResolutionIconUrl DefaultValue="{{BASE_URL}}/assets/sc.png"/>
  <SupportUrl DefaultValue="https://support.singlecase.com"/>

  <AppDomains>
    <AppDomain>{{BASE_URL}}</AppDomain>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>

  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.12"/>
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="{{BASE_URL}}/taskpane.html?v={{VERSION}}"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteMailbox</Permissions>

  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
  </Rule>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">

      <Requirements>
        <bt:Sets DefaultMinVersion="1.12">
          <bt:Set Name="Mailbox"/>
        </bt:Sets>
      </Requirements>

      <Hosts>
        <Host xsi:type="MailHost">

          <Runtimes>
            <Runtime resid="WebViewRuntime.Url" lifetime="short" />
          </Runtimes>

          <DesktopFormFactor>
            <FunctionFile resid="Commands.Url"/>

            <ExtensionPoint xsi:type="LaunchEvent">
              <LaunchEvents>
                <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock"/>
              </LaunchEvents>
              <SourceLocation resid="WebViewRuntime.Url"/>
            </ExtensionPoint>

            <ExtensionPoint xsi:type="MessageReadCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="msgReadGroup">
                  <Label resid="GroupLabel"/>
                  <Control xsi:type="Button" id="msgReadOpenPaneButton">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>true</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>

            <ExtensionPoint xsi:type="MessageComposeCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="msgComposeGroup">
                  <Label resid="GroupLabel"/>
                  <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>true</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>

          </DesktopFormFactor>
        </Host>
      </Hosts>

      <Resources>
        <bt:Images>
          <bt:Image id="Icon.16x16" DefaultValue="{{BASE_URL}}/assets/sc.png"/>
          <bt:Image id="Icon.32x32" DefaultValue="{{BASE_URL}}/assets/sc.png"/>
          <bt:Image id="Icon.80x80" DefaultValue="{{BASE_URL}}/assets/sc.png"/>
        </bt:Images>

        <bt:Urls>
          <bt:Url id="Commands.Url" DefaultValue="{{BASE_URL}}/commands.html?v={{VERSION}}"/>
          <bt:Url id="WebViewRuntime.Url" DefaultValue="{{BASE_URL}}/commands.html?v={{VERSION}}"/>
          <bt:Url id="Taskpane.Url" DefaultValue="{{BASE_URL}}/taskpane.html?v={{VERSION}}"/>
        </bt:Urls>

        <bt:ShortStrings>
          <bt:String id="GroupLabel" DefaultValue="SingleCase Add-in"/>
          <bt:String id="TaskpaneButton.Label" DefaultValue="Show Task Pane"/>
        </bt:ShortStrings>

        <bt:LongStrings>
          <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open SingleCase panel"/>
        </bt:LongStrings>
      </Resources>

    </VersionOverrides>
  </VersionOverrides>

</OfficeApp>
```

**manifest.staging.xml** (Staging)
- Same structure as dev, but with:
  - `DisplayName`: "SingleCase (Staging)"
  - `BASE_URL`: `https://staging-addin.singlecase.com`
  - Different version number if needed

**manifest.production.xml** (Production)
- Same structure as dev, but with:
  - `DisplayName`: "SingleCase"
  - `BASE_URL`: `https://addin.singlecase.com`
  - Production version number
  - No environment indicator in display name

### Manifest Generation Script

**scripts/generate-manifest.js**
```javascript
#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

/**
 * Generate environment-specific Office Add-in manifest
 *
 * Usage:
 *   node scripts/generate-manifest.js [environment]
 *
 * Environment: dev, staging, production (default: dev)
 */

const environments = {
  dev: {
    APP_ID: '163ed3c6-8878-484e-8b35-37d4279aa769',
    BASE_URL: 'https://localhost:3000',
    VERSION: getVersion('dev'),
  },
  staging: {
    APP_ID: '163ed3c6-8878-484e-8b35-37d4279aa769',
    BASE_URL: 'https://staging-addin.singlecase.com',
    VERSION: getVersion('staging'),
  },
  production: {
    APP_ID: '163ed3c6-8878-484e-8b35-37d4279aa769',
    BASE_URL: 'https://addin.singlecase.com',
    VERSION: getVersion('production'),
  },
};

function getVersion(env) {
  const packageJson = JSON.parse(
    fs.readFileSync(path.join(__dirname, '../package.json'), 'utf8')
  );

  let version = packageJson.version;

  // Auto-increment patch version for dev builds
  if (env === 'dev') {
    const [major, minor, patch] = version.split('.').map(Number);
    version = `${major}.${minor}.${patch + 1}`;
  }

  return version;
}

function generateManifest(environment) {
  const env = environment || 'dev';

  if (!environments[env]) {
    console.error(`Error: Unknown environment "${env}"`);
    console.error('Valid environments: dev, staging, production');
    process.exit(1);
  }

  const templatePath = path.join(__dirname, '../manifests', `manifest.${env}.xml`);
  const outputPath = path.join(__dirname, '../manifest.xml');

  if (!fs.existsSync(templatePath)) {
    console.error(`Error: Template not found at ${templatePath}`);
    process.exit(1);
  }

  console.log(`Generating manifest for ${env} environment...`);

  // Read template
  let manifestContent = fs.readFileSync(templatePath, 'utf8');

  // Replace placeholders
  const config = environments[env];
  Object.keys(config).forEach((key) => {
    const placeholder = new RegExp(`{{${key}}}`, 'g');
    manifestContent = manifestContent.replace(placeholder, config[key]);
  });

  // Write output
  fs.writeFileSync(outputPath, manifestContent, 'utf8');

  console.log(`✓ Manifest generated successfully: ${outputPath}`);
  console.log(`  Environment: ${env}`);
  console.log(`  Version: ${config.VERSION}`);
  console.log(`  Base URL: ${config.BASE_URL}`);
}

// Run
const environment = process.argv[2] || process.env.NODE_ENV || 'dev';
generateManifest(environment);
```

**scripts/validate-manifest.js**
```javascript
#!/usr/bin/env node

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

/**
 * Validate Office Add-in manifest
 * Uses office-addin-manifest validation tools
 */

const manifestPath = path.join(__dirname, '../manifest.xml');

if (!fs.existsSync(manifestPath)) {
  console.error('Error: manifest.xml not found. Run generate-manifest.js first.');
  process.exit(1);
}

console.log('Validating manifest.xml...');

try {
  // Use Office Add-in CLI to validate
  execSync('office-addin-manifest validate manifest.xml', {
    stdio: 'inherit',
    cwd: path.join(__dirname, '..'),
  });

  console.log('✓ Manifest validation passed');
  process.exit(0);
} catch (error) {
  console.error('✗ Manifest validation failed');
  process.exit(1);
}
```

### NPM Scripts

Add to `package.json`:
```json
{
  "scripts": {
    "manifest:dev": "node scripts/generate-manifest.js dev",
    "manifest:staging": "node scripts/generate-manifest.js staging",
    "manifest:production": "node scripts/generate-manifest.js production",
    "manifest:validate": "node scripts/validate-manifest.js",
    "prebuild": "npm run manifest:dev && npm run manifest:validate",
    "prebuild:staging": "npm run manifest:staging && npm run manifest:validate",
    "prebuild:production": "npm run manifest:production && npm run manifest:validate"
  }
}
```

### Environment Variables

**.env.dev**
```env
NODE_ENV=development
MANIFEST_BASE_URL=https://localhost:3000
MANIFEST_APP_ID=163ed3c6-8878-484e-8b35-37d4279aa769
```

**.env.staging**
```env
NODE_ENV=staging
MANIFEST_BASE_URL=https://staging-addin.singlecase.com
MANIFEST_APP_ID=163ed3c6-8878-484e-8b35-37d4279aa769
```

**.env.production**
```env
NODE_ENV=production
MANIFEST_BASE_URL=https://addin.singlecase.com
MANIFEST_APP_ID=163ed3c6-8878-484e-8b35-37d4279aa769
```

### Git Configuration

Update `.gitignore`:
```gitignore
# Generated manifest (use templates instead)
/manifest.xml

# Keep template manifests
!/manifests/manifest.*.xml
```

## Deployment Process

### Development

1. Generate dev manifest:
   ```bash
   npm run manifest:dev
   ```

2. Start dev server:
   ```bash
   npm run dev-server
   ```

3. Sideload manifest in Outlook:
   ```bash
   npm start
   ```

### Staging Deployment

1. Update version in `package.json` (if needed)

2. Generate staging manifest:
   ```bash
   npm run manifest:staging
   ```

3. Build application:
   ```bash
   npm run build
   ```

4. Deploy to staging server (see DO-003 for deployment script)

5. Upload manifest to Microsoft 365 Admin Center (Integrated Apps)

### Production Deployment

1. Update version in `package.json` (required)

2. Generate production manifest:
   ```bash
   npm run manifest:production
   ```

3. Validate manifest:
   ```bash
   npm run manifest:validate
   ```

4. Build application:
   ```bash
   npm run build
   ```

5. Deploy to production server (see DO-003 for deployment script)

6. Upload manifest to Microsoft 365 Admin Center or AppSource

## Manifest Versioning Strategy

### Version Number Format

Use semantic versioning: `MAJOR.MINOR.PATCH`

- **MAJOR**: Breaking changes, incompatible API changes
- **MINOR**: New features, backward-compatible
- **PATCH**: Bug fixes, backward-compatible

### Version Increment Rules

1. **Development**
   - Auto-increment patch on every build
   - Example: `1.0.0` → `1.0.1` → `1.0.2`

2. **Staging**
   - Manual version update before deployment
   - Use pre-release tags: `1.1.0-rc.1`

3. **Production**
   - Strict version control
   - No pre-release tags
   - Update CHANGELOG.md before release

### Manifest Update Behavior

Office Add-ins check for manifest updates:
- **Office Desktop**: Checks every 24 hours
- **Office Online**: Checks on browser refresh
- **Force Update**: Remove and re-add the add-in

To force users to update:
1. Increment `<Version>` in manifest
2. Deploy new manifest to Microsoft 365 Admin Center
3. Users will receive update notification within 24 hours

## Testing Checklist

- [ ] Generate dev manifest and verify URLs
- [ ] Generate staging manifest and verify URLs
- [ ] Generate production manifest and verify URLs
- [ ] Validate all manifests with office-addin-manifest tool
- [ ] Test sideloading dev manifest in Outlook Desktop
- [ ] Test sideloading dev manifest in Outlook Web
- [ ] Verify all icon URLs are accessible
- [ ] Verify all HTML resource URLs are accessible
- [ ] Test version auto-increment for dev builds
- [ ] Verify manifest.xml is gitignored
- [ ] Test npm scripts (manifest:dev, manifest:staging, etc.)
- [ ] Verify AppDomains include all necessary domains
- [ ] Test cache-busting query parameters

## Dependencies

- **Required for**: DO-002 (CI/CD Pipeline Setup), DO-003 (Centralized Deployment Script)
- **Depends on**: None (foundational story)

## References

- [Office Add-ins Manifest Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests)
- [Office Add-in Versioning](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/update-your-add-in)
- [Manifest Validation Tools](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest)

## Notes

1. **Manifest Caching**: Office clients cache manifests for 24 hours. During development, clear cache by:
   - Outlook Desktop: Close Outlook, delete cache folder, restart
   - Outlook Web: Clear browser cache

2. **App ID Considerations**:
   - **Same ID across environments**: Users see one add-in, updates smoothly
   - **Different IDs per environment**: Users can have dev and prod installed simultaneously
   - **Recommendation**: Use same ID for staging/production, different for dev

3. **SSL Certificates**:
   - Development: Use office-addin-dev-certs for localhost
   - Staging/Production: Use valid SSL certificate (Let's Encrypt, DigiCert, etc.)
   - Office Add-ins REQUIRE HTTPS (except localhost)

4. **Manifest Upload Locations**:
   - **Personal Development**: Sideload via Outlook → Get Add-ins → My Add-ins → Custom
   - **Organizational Deployment**: Microsoft 365 Admin Center → Integrated Apps
   - **Public Deployment**: Submit to AppSource (requires validation)

## Definition of Done

- [ ] `manifests/` directory created with dev/staging/production templates
- [ ] `scripts/generate-manifest.js` implemented and tested
- [ ] `scripts/validate-manifest.js` implemented and tested
- [ ] NPM scripts added for manifest generation
- [ ] Environment-specific .env files created
- [ ] `.gitignore` updated to ignore generated manifest.xml
- [ ] All three manifest templates validated with office-addin-manifest
- [ ] Dev manifest tested with local sideloading
- [ ] Version auto-increment working for dev environment
- [ ] Cache-busting query parameters added to all resource URLs
- [ ] Documentation written for manifest generation process
- [ ] README.md updated with manifest deployment instructions
