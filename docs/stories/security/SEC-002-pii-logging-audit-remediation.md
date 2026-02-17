# SEC-002: PII Logging Audit & Remediation

**Story ID:** SEC-002
**Story Points:** 8
**Epic Link:** Security & Compliance
**Status:** Ready for Development

## Description

Conduct comprehensive audit of application codebase to identify and remediate all instances where Personally Identifiable Information (PII) is being logged, stored, or transmitted insecurely. Implement enterprise-grade logging practices that prevent PII exposure while maintaining effective debugging and monitoring capabilities.

This story addresses GDPR Article 25 (Data Protection by Design), CCPA compliance requirements, and OWASP A09:2021 (Security Logging and Monitoring Failures).

## Acceptance Criteria

1. **PII Identification & Classification**
   - Identify all PII data points: email addresses, names, phone numbers, IP addresses, user IDs
   - Classify PII sensitivity levels: high (SSN, passwords), medium (email, name), low (user IDs)
   - Document all locations where PII is processed or stored
   - Create PII data flow diagram
   - Maintain PII inventory for compliance reporting

2. **Logging Audit**
   - Audit all console.log, console.error, console.warn statements
   - Audit all error reporting (stack traces, error messages)
   - Audit all debugging statements
   - Audit all analytics/telemetry data
   - Audit all API request/response logging
   - Identify PII in log messages, variable names, or data structures

3. **PII Removal from Logs**
   - Remove or redact email addresses from logs
   - Remove or redact user names from logs
   - Remove or redact authentication tokens from logs
   - Remove or redact API keys and secrets from logs
   - Remove or redact workspace hosts from logs (if contains customer identifiers)
   - Replace PII with sanitized placeholders (e.g., "user_abc123" instead of "john.doe@example.com")

4. **Secure Logging Implementation**
   - Implement structured logging framework
   - Create PII-safe logger utility
   - Automatic PII detection and redaction
   - Log sanitization before transmission
   - Separate error logs from debug logs
   - Implement log levels (DEBUG, INFO, WARN, ERROR)

5. **Error Handling Without PII**
   - Never include PII in error messages
   - Use error codes instead of descriptive messages with PII
   - Sanitize stack traces before logging
   - Remove sensitive request/response data from errors
   - Implement user-friendly error messages (no technical details)

6. **Analytics & Telemetry**
   - Remove PII from analytics events
   - Hash or pseudonymize user identifiers
   - Aggregate data before transmission
   - Implement consent management for telemetry
   - Document what data is collected and why

7. **Monitoring & Alerting**
   - Implement automated PII detection in logs
   - Alert on PII exposure in production logs
   - Regular audit schedule (quarterly)
   - Compliance reporting dashboard
   - Incident response plan for PII leaks

8. **Developer Training**
   - Document PII handling guidelines
   - Create developer training materials
   - Implement pre-commit hooks to detect PII in code
   - Code review checklist for PII compliance
   - Automated static analysis for PII detection

## Technical Requirements

### PII-Safe Logger Implementation

1. **Structured Logger Service**
   ```typescript
   // src/services/logger.ts

   export enum LogLevel {
     DEBUG = 0,
     INFO = 1,
     WARN = 2,
     ERROR = 3,
   }

   export interface LogContext {
     component?: string;
     action?: string;
     userId?: string; // Pseudonymized
     sessionId?: string;
     requestId?: string;
     [key: string]: any;
   }

   class Logger {
     private minLevel: LogLevel = LogLevel.INFO;
     private sensitivePatterns: RegExp[] = [
       // Email addresses
       /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g,
       // Phone numbers (various formats)
       /\b(\+\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b/g,
       // Credit card numbers
       /\b\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b/g,
       // SSN
       /\b\d{3}-\d{2}-\d{4}\b/g,
       // API keys (common patterns)
       /\b[A-Za-z0-9]{32,}\b/g,
       // JWT tokens
       /\beyJ[A-Za-z0-9_-]+\.eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b/g,
       // Bearer tokens
       /Bearer\s+[A-Za-z0-9_-]+/gi,
       // IP addresses (if considered PII)
       /\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b/g,
     ];

     constructor() {
       // Set log level from environment
       const envLevel = process.env.LOG_LEVEL || 'INFO';
       this.minLevel = LogLevel[envLevel as keyof typeof LogLevel] || LogLevel.INFO;
     }

     debug(message: string, context?: LogContext): void {
       this.log(LogLevel.DEBUG, message, context);
     }

     info(message: string, context?: LogContext): void {
       this.log(LogLevel.INFO, message, context);
     }

     warn(message: string, context?: LogContext): void {
       this.log(LogLevel.WARN, message, context);
     }

     error(message: string, error?: Error, context?: LogContext): void {
       const sanitizedError = error ? this.sanitizeError(error) : undefined;
       this.log(LogLevel.ERROR, message, {
         ...context,
         error: sanitizedError,
       });
     }

     private log(level: LogLevel, message: string, context?: LogContext): void {
       if (level < this.minLevel) {
         return;
       }

       // Sanitize message and context
       const sanitizedMessage = this.sanitize(message);
       const sanitizedContext = context ? this.sanitizeContext(context) : undefined;

       // Create structured log entry
       const logEntry = {
         timestamp: new Date().toISOString(),
         level: LogLevel[level],
         message: sanitizedMessage,
         ...sanitizedContext,
       };

       // Output to console (in production, send to logging service)
       this.output(level, logEntry);
     }

     /**
      * Sanitize string to remove PII
      */
     private sanitize(text: string): string {
       let sanitized = text;

       // Replace sensitive patterns with redacted placeholders
       for (const pattern of this.sensitivePatterns) {
         sanitized = sanitized.replace(pattern, '[REDACTED]');
       }

       return sanitized;
     }

     /**
      * Sanitize log context object
      */
     private sanitizeContext(context: LogContext): LogContext {
       const sanitized: LogContext = {};

       for (const [key, value] of Object.entries(context)) {
         // Skip known sensitive fields
         if (this.isSensitiveField(key)) {
           sanitized[key] = '[REDACTED]';
           continue;
         }

         // Sanitize string values
         if (typeof value === 'string') {
           sanitized[key] = this.sanitize(value);
         }
         // Recursively sanitize objects
         else if (typeof value === 'object' && value !== null) {
           sanitized[key] = this.sanitizeObject(value);
         }
         // Keep primitive values
         else {
           sanitized[key] = value;
         }
       }

       return sanitized;
     }

     /**
      * Sanitize nested objects
      */
     private sanitizeObject(obj: any): any {
       if (Array.isArray(obj)) {
         return obj.map(item =>
           typeof item === 'object' ? this.sanitizeObject(item) : this.sanitize(String(item))
         );
       }

       const sanitized: any = {};
       for (const [key, value] of Object.entries(obj)) {
         if (this.isSensitiveField(key)) {
           sanitized[key] = '[REDACTED]';
         } else if (typeof value === 'string') {
           sanitized[key] = this.sanitize(value);
         } else if (typeof value === 'object' && value !== null) {
           sanitized[key] = this.sanitizeObject(value);
         } else {
           sanitized[key] = value;
         }
       }
       return sanitized;
     }

     /**
      * Check if field name indicates sensitive data
      */
     private isSensitiveField(fieldName: string): boolean {
       const sensitiveFields = [
         'password',
         'token',
         'secret',
         'key',
         'auth',
         'authorization',
         'cookie',
         'session',
         'ssn',
         'credit_card',
         'creditcard',
         'cvv',
         'pin',
       ];

       const lowerField = fieldName.toLowerCase();
       return sensitiveFields.some(sensitive => lowerField.includes(sensitive));
     }

     /**
      * Sanitize error objects
      */
     private sanitizeError(error: Error): any {
       return {
         name: error.name,
         message: this.sanitize(error.message),
         // Sanitize stack trace
         stack: error.stack ? this.sanitizeStackTrace(error.stack) : undefined,
       };
     }

     /**
      * Sanitize stack traces to remove PII from file paths
      */
     private sanitizeStackTrace(stack: string): string {
       // Remove file paths that might contain usernames
       let sanitized = stack.replace(/\/Users\/[^/]+\//g, '/Users/[USER]/');
       sanitized = sanitized.replace(/C:\\Users\\[^\\]+\\/g, 'C:\\Users\\[USER]\\');

       // Remove any detected PII
       sanitized = this.sanitize(sanitized);

       return sanitized;
     }

     /**
      * Output log entry
      */
     private output(level: LogLevel, logEntry: any): void {
       const formatted = JSON.stringify(logEntry);

       switch (level) {
         case LogLevel.DEBUG:
           console.debug(formatted);
           break;
         case LogLevel.INFO:
           console.info(formatted);
           break;
         case LogLevel.WARN:
           console.warn(formatted);
           break;
         case LogLevel.ERROR:
           console.error(formatted);
           break;
       }
     }
   }

   // Export singleton instance
   export const logger = new Logger();
   ```

2. **PII Detection Utility**
   ```typescript
   // src/utils/piiDetection.ts

   export interface PIIMatch {
     type: 'email' | 'phone' | 'ssn' | 'credit_card' | 'token' | 'ip_address';
     value: string;
     position: number;
   }

   export class PIIDetector {
     private patterns = {
       email: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g,
       phone: /\b(\+\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b/g,
       ssn: /\b\d{3}-\d{2}-\d{4}\b/g,
       credit_card: /\b\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b/g,
       token: /\beyJ[A-Za-z0-9_-]+\.eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b/g,
       ip_address: /\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b/g,
     };

     /**
      * Detect PII in text
      */
     detectPII(text: string): PIIMatch[] {
       const matches: PIIMatch[] = [];

       for (const [type, pattern] of Object.entries(this.patterns)) {
         const regex = new RegExp(pattern.source, pattern.flags);
         let match;

         while ((match = regex.exec(text)) !== null) {
           matches.push({
             type: type as PIIMatch['type'],
             value: match[0],
             position: match.index,
           });
         }
       }

       return matches;
     }

     /**
      * Check if text contains PII
      */
     containsPII(text: string): boolean {
       return this.detectPII(text).length > 0;
     }

     /**
      * Redact PII from text
      */
     redactPII(text: string, replacement: string = '[REDACTED]'): string {
       let redacted = text;

       for (const pattern of Object.values(this.patterns)) {
         redacted = redacted.replace(pattern, replacement);
       }

       return redacted;
     }

     /**
      * Hash PII for pseudonymization
      */
     async hashPII(pii: string): Promise<string> {
       const encoder = new TextEncoder();
       const data = encoder.encode(pii);
       const hash = await crypto.subtle.digest('SHA-256', data);
       const hashArray = Array.from(new Uint8Array(hash));
       return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
     }

     /**
      * Pseudonymize user identifier
      */
     async pseudonymizeUserId(email: string): Promise<string> {
       const hash = await this.hashPII(email);
       return `user_${hash.substring(0, 16)}`;
     }
   }

   export const piiDetector = new PIIDetector();
   ```

3. **API Request/Response Sanitization**
   ```typescript
   // src/services/apiSanitizer.ts

   import { piiDetector } from '../utils/piiDetection';

   export class APISanitizer {
     /**
      * Sanitize API request for logging
      */
     sanitizeRequest(request: {
       url: string;
       method: string;
       headers: Record<string, string>;
       body?: any;
     }): any {
       return {
         url: this.sanitizeURL(request.url),
         method: request.method,
         headers: this.sanitizeHeaders(request.headers),
         // Never log request body (may contain PII)
         body: '[REDACTED]',
       };
     }

     /**
      * Sanitize API response for logging
      */
     sanitizeResponse(response: {
       status: number;
       headers: Record<string, string>;
       body?: any;
     }): any {
       return {
         status: response.status,
         headers: this.sanitizeHeaders(response.headers),
         // Never log full response body (may contain PII)
         // Only log metadata
         bodySize: response.body ? JSON.stringify(response.body).length : 0,
       };
     }

     /**
      * Sanitize URL (remove query params that may contain PII)
      */
     private sanitizeURL(url: string): string {
       try {
         const urlObj = new URL(url, window.location.origin);

         // Keep only safe query params
         const safeParams = ['page', 'limit', 'sort', 'filter'];
         const params = new URLSearchParams(urlObj.search);
         const sanitizedParams = new URLSearchParams();

         for (const param of safeParams) {
           if (params.has(param)) {
             sanitizedParams.set(param, params.get(param)!);
           }
         }

         urlObj.search = sanitizedParams.toString();
         return urlObj.toString();
       } catch {
         return '[INVALID_URL]';
       }
     }

     /**
      * Sanitize HTTP headers
      */
     private sanitizeHeaders(headers: Record<string, string>): Record<string, string> {
       const sanitized: Record<string, string> = {};
       const sensitiveHeaders = [
         'authorization',
         'cookie',
         'set-cookie',
         'authentication',
         'x-api-key',
       ];

       for (const [key, value] of Object.entries(headers)) {
         if (sensitiveHeaders.includes(key.toLowerCase())) {
           sanitized[key] = '[REDACTED]';
         } else {
           sanitized[key] = piiDetector.redactPII(value);
         }
       }

       return sanitized;
     }
   }

   export const apiSanitizer = new APISanitizer();
   ```

4. **Pre-Commit Hook for PII Detection**
   ```bash
   #!/bin/bash
   # .git/hooks/pre-commit

   echo "Running PII detection on staged files..."

   # Find all staged TypeScript/JavaScript files
   staged_files=$(git diff --cached --name-only --diff-filter=ACM | grep -E '\.(ts|tsx|js|jsx)$')

   if [ -z "$staged_files" ]; then
     exit 0
   fi

   # PII patterns to detect
   patterns=(
     "console\\.log.*@"  # Email in console.log
     "console\\.log.*Bearer"  # Token in console.log
     "console\\.log.*password"  # Password in console.log
     "\+1[0-9]{10}"  # Phone numbers
     "[0-9]{3}-[0-9]{2}-[0-9]{4}"  # SSN
   )

   found_pii=false

   for file in $staged_files; do
     for pattern in "${patterns[@]}"; do
       if grep -nE "$pattern" "$file" > /dev/null 2>&1; then
         echo "⚠️  WARNING: Potential PII detected in $file"
         grep -nE "$pattern" "$file" --color=always
         found_pii=true
       fi
     done
   done

   if [ "$found_pii" = true ]; then
     echo ""
     echo "❌ Commit blocked: Potential PII detected in code"
     echo "Please review and remove any sensitive information before committing"
     exit 1
   fi

   echo "✅ No PII detected"
   exit 0
   ```

5. **Automated PII Audit Script**
   ```typescript
   // scripts/auditPII.ts

   import * as fs from 'fs';
   import * as path from 'path';
   import { glob } from 'glob';

   interface PIIViolation {
     file: string;
     line: number;
     content: string;
     type: string;
   }

   const PII_PATTERNS = [
     { name: 'Email in console.log', pattern: /console\.log.*@[\w.-]+\.\w+/g },
     { name: 'Token in console.log', pattern: /console\.log.*Bearer/gi },
     { name: 'Password in console.log', pattern: /console\.log.*password/gi },
     { name: 'API Key in code', pattern: /api[_-]?key\s*[:=]\s*['"][^'"]+['"]/gi },
     { name: 'Email hardcoded', pattern: /['"][A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}['"]/g },
   ];

   async function auditDirectory(dir: string): Promise<PIIViolation[]> {
     const violations: PIIViolation[] = [];

     // Find all source files
     const files = await glob('**/*.{ts,tsx,js,jsx}', {
       cwd: dir,
       ignore: ['node_modules/**', 'dist/**', 'build/**'],
     });

     for (const file of files) {
       const filePath = path.join(dir, file);
       const content = fs.readFileSync(filePath, 'utf-8');
       const lines = content.split('\n');

       for (let i = 0; i < lines.length; i++) {
         const line = lines[i];

         for (const { name, pattern } of PII_PATTERNS) {
           if (pattern.test(line)) {
             violations.push({
               file: filePath,
               line: i + 1,
               content: line.trim(),
               type: name,
             });
           }
         }
       }
     }

     return violations;
   }

   async function main() {
     console.log('🔍 Starting PII audit...\n');

     const violations = await auditDirectory(process.cwd());

     if (violations.length === 0) {
       console.log('✅ No PII violations detected!\n');
       return;
     }

     console.log(`❌ Found ${violations.length} potential PII violations:\n`);

     for (const violation of violations) {
       console.log(`File: ${violation.file}:${violation.line}`);
       console.log(`Type: ${violation.type}`);
       console.log(`Content: ${violation.content}`);
       console.log('');
     }

     process.exit(1);
   }

   main();
   ```

### Environment Variables

```env
# .env

# Logging Configuration
LOG_LEVEL=INFO  # DEBUG, INFO, WARN, ERROR
ENABLE_PII_DETECTION=true
ENABLE_LOG_SANITIZATION=true

# Analytics (with consent)
ENABLE_ANALYTICS=false
PSEUDONYMIZE_USER_IDS=true
```

## OWASP Top 10 Coverage

### A09:2021 - Security Logging and Monitoring Failures
- **Mitigation**: Implement structured logging without PII
- **Mitigation**: Automatic PII detection and redaction
- **Mitigation**: Log sanitization before transmission
- **Mitigation**: Separate sensitive data from logs

### A01:2021 - Broken Access Control
- **Mitigation**: Never log user credentials or tokens
- **Mitigation**: Redact authorization headers from logs

### A04:2021 - Insecure Design
- **Mitigation**: Privacy by design - PII protection built into logging
- **Mitigation**: Developer training on PII handling

## GDPR Compliance

### Article 25 - Data Protection by Design and by Default
- Implement PII detection and redaction by default
- Privacy-preserving logging architecture
- Minimize data collection (only log necessary information)

### Article 32 - Security of Processing
- Technical measures to protect PII (encryption, pseudonymization)
- Regular security audits for PII exposure

### Article 33/34 - Breach Notification
- Incident response plan for PII leaks in logs
- Automated detection of PII exposure

## Testing Strategy

1. **PII Detection Tests**
   - Test email detection in various formats
   - Test phone number detection (multiple formats)
   - Test credit card number detection
   - Test token detection (JWT, Bearer)
   - Test false positive rate

2. **Sanitization Tests**
   - Verify PII is redacted from logs
   - Verify tokens are never logged
   - Verify error messages don't contain PII
   - Verify stack traces are sanitized

3. **Integration Tests**
   - Test logger with real log scenarios
   - Test API request/response sanitization
   - Test pre-commit hook with PII in code

4. **Compliance Tests**
   - Audit all log outputs for PII
   - Verify GDPR compliance
   - Generate compliance reports

## Audit Checklist

- [ ] Audit all console.log statements in codebase
- [ ] Audit all console.error statements
- [ ] Audit all console.warn statements
- [ ] Audit error handling code for PII
- [ ] Audit API request logging
- [ ] Audit API response logging
- [ ] Audit analytics/telemetry code
- [ ] Audit debugging statements
- [ ] Audit error messages shown to users
- [ ] Audit stack traces
- [ ] Check for hardcoded emails in code
- [ ] Check for hardcoded tokens in code
- [ ] Check for PII in variable names
- [ ] Check for PII in comments

## Definition of Done

- [ ] PII audit completed for entire codebase
- [ ] All PII removed or redacted from logs
- [ ] PII-safe logger service implemented
- [ ] Structured logging framework deployed
- [ ] PII detection utility implemented
- [ ] API request/response sanitizer implemented
- [ ] Pre-commit hook for PII detection installed
- [ ] Automated PII audit script created
- [ ] Developer training materials created
- [ ] PII handling guidelines documented
- [ ] Code review checklist updated
- [ ] Unit tests for PII detection (>95% coverage)
- [ ] Integration tests for logging
- [ ] Compliance report generated
- [ ] GDPR compliance verified
- [ ] Security audit passed
- [ ] No PII in production logs (verified)

## Dependencies

- ESLint plugin for PII detection
- Git hooks for pre-commit validation
- Log aggregation service (production)

## References

- [GDPR Article 25 - Data Protection by Design](https://gdpr-info.eu/art-25-gdpr/)
- [OWASP Logging Cheat Sheet](https://cheatsheetseries.owasp.org/cheatsheets/Logging_Cheat_Sheet.html)
- [NIST SP 800-122 - Guide to Protecting PII](https://csrc.nist.gov/publications/detail/sp/800-122/final)
- [CWE-532: Insertion of Sensitive Information into Log File](https://cwe.mitre.org/data/definitions/532.html)

## Notes

1. **Common PII in Logs**
   - Email addresses in user login flows
   - Tokens in authentication headers
   - Workspace hosts containing customer names
   - Error messages with user input
   - Stack traces with file paths containing usernames

2. **Safe Logging Practices**
   - Log actions, not data (e.g., "User logged in" not "john@example.com logged in")
   - Use pseudonymized IDs (e.g., "user_abc123" not email)
   - Log metadata, not payloads (e.g., "API call succeeded" not full response)
   - Sanitize error messages before logging

3. **Production Monitoring**
   - Use centralized logging service (e.g., CloudWatch, Splunk)
   - Implement log retention policies
   - Regular compliance audits
   - Automated PII detection in production logs
