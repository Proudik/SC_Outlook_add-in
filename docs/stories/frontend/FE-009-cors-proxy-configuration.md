# FE-009: CORS & Proxy Configuration

**Story ID:** FE-009
**Story Points:** 5
**Epic Link:** Infrastructure & DevOps
**Status:** Ready for Development

## Description

Configure CORS (Cross-Origin Resource Sharing) and API proxy infrastructure to support multi-tenant workspace architecture where each customer's SingleCase instance runs on a different subdomain (e.g., `customer1.singlecase.ch`, `customer2.singlecase.ch`). The add-in must dynamically route API requests through a proxy that handles CORS and supports customer-specific workspace hosts.

Update development proxy (webpack dev server) and production deployment to support dynamic workspace host routing.

## Acceptance Criteria

1. **Dynamic Workspace Host Routing**
   - Frontend sends API requests to: `/singlecase/{workspaceHost}/publicapi/v1/...`
   - Proxy extracts `workspaceHost` from URL path
   - Proxy forwards request to: `https://{workspaceHost}/publicapi/v1/...`
   - Support any workspace host (customer subdomains)
   - Validate workspace host format (prevent open proxy abuse)

2. **Development Proxy (Webpack)**
   - Webpack dev server middleware proxies `/singlecase/*` requests
   - Extract workspace host from URL path segment
   - Forward to actual workspace host with proper headers
   - Handle HTTPS certificate validation
   - Log all proxy requests for debugging

3. **Production Proxy Configuration**
   - Deploy proxy service (nginx, API Gateway, or serverless function)
   - Handle CORS preflight (OPTIONS) requests
   - Add proper CORS headers to responses
   - Rate limit per workspace host
   - Monitor proxy health and errors

4. **CORS Headers**
   - `Access-Control-Allow-Origin`: Add-in origin (https://localhost:3000 dev, production domain)
   - `Access-Control-Allow-Methods`: GET, POST, PUT, PATCH, DELETE
   - `Access-Control-Allow-Headers`: Authentication, Content-Type
   - `Access-Control-Allow-Credentials`: true (for cookies if needed)
   - Handle preflight OPTIONS requests

5. **Error Handling**
   - 400 Bad Request if workspace host is invalid
   - 502 Bad Gateway if upstream workspace is unreachable
   - 504 Gateway Timeout if upstream is slow
   - Log all proxy errors for monitoring

6. **Security**
   - Whitelist allowed workspace host patterns (*.singlecase.ch)
   - Prevent open proxy abuse (no arbitrary host proxying)
   - Remove sensitive headers (Cookie, Set-Cookie)
   - Add request ID header for tracing

## Technical Requirements

### Development Environment

1. **Update webpack.config.js**
   - Enhance existing `/singlecase` middleware
   - Add workspace host validation
   - Add request/response logging
   - Add timeout handling
   - Add error handling

2. **Proxy Middleware**
   ```javascript
   // In webpack.config.js setupMiddlewares
   devServer.app.use('/singlecase', async (req, res, next) => {
     try {
       // Extract workspace host from URL: /singlecase/{host}/publicapi/v1/...
       const match = req.url.match(/^\/singlecase\/([^/?#]+)(\/.*)?$/);
       if (!match) {
         return res.status(400).send('Invalid proxy URL format');
       }

       const encodedHost = match[1];
       const restPath = match[2] || '';
       const workspaceHost = decodeURIComponent(encodedHost);

       // Validate workspace host (whitelist pattern)
       if (!isValidWorkspaceHost(workspaceHost)) {
         return res.status(400).send('Invalid workspace host');
       }

       const upstreamUrl = `https://${workspaceHost}${restPath}`;
       console.log(`[Proxy] ${req.method} ${upstreamUrl}`);

       // Forward request
       const upstreamResponse = await fetch(upstreamUrl, {
         method: req.method,
         headers: {
           ...sanitizeHeaders(req.headers),
           'X-Forwarded-For': req.ip,
           'X-Request-ID': generateRequestId(),
         },
         body: req.method !== 'GET' && req.method !== 'HEAD'
           ? await readRequestBody(req)
           : undefined,
       });

       // Forward response
       res.status(upstreamResponse.status);
       upstreamResponse.headers.forEach((value, key) => {
         if (key.toLowerCase() !== 'set-cookie') {
           res.setHeader(key, value);
         }
       });

       const responseBody = await upstreamResponse.arrayBuffer();
       res.send(Buffer.from(responseBody));

     } catch (error) {
       console.error('[Proxy] Error:', error);
       res.status(502).send('Proxy error');
     }
   });
   ```

3. **Workspace Host Validation**
   ```javascript
   function isValidWorkspaceHost(host) {
     // Allow only *.singlecase.ch and localhost for dev
     const patterns = [
       /^[\w-]+\.singlecase\.ch$/,
       /^localhost:\d+$/,
       /^127\.0\.0\.1:\d+$/,
     ];
     return patterns.some(pattern => pattern.test(host));
   }
   ```

### Production Environment

1. **Nginx Reverse Proxy** (recommended)
   ```nginx
   # /etc/nginx/sites-available/outlook-addin-proxy

   server {
     listen 443 ssl http2;
     server_name addin.singlecase.com;

     ssl_certificate /path/to/cert.pem;
     ssl_certificate_key /path/to/key.pem;

     # CORS configuration
     add_header 'Access-Control-Allow-Origin' 'https://addin.singlecase.com' always;
     add_header 'Access-Control-Allow-Methods' 'GET, POST, PUT, PATCH, DELETE, OPTIONS' always;
     add_header 'Access-Control-Allow-Headers' 'Authentication, Content-Type, X-Request-ID' always;
     add_header 'Access-Control-Allow-Credentials' 'true' always;

     # Handle preflight requests
     if ($request_method = OPTIONS) {
       return 204;
     }

     # Proxy /singlecase/{host}/publicapi/v1/... to https://{host}/publicapi/v1/...
     location ~ ^/singlecase/([^/]+)(/.*)$ {
       set $workspace_host $1;
       set $rest_path $2;

       # Validate workspace host
       if ($workspace_host !~ ^[\w-]+\.singlecase\.ch$) {
         return 400 "Invalid workspace host";
       }

       # Proxy to upstream
       proxy_pass https://$workspace_host$rest_path;
       proxy_set_header Host $workspace_host;
       proxy_set_header X-Real-IP $remote_addr;
       proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
       proxy_set_header X-Request-ID $request_id;

       # Timeouts
       proxy_connect_timeout 5s;
       proxy_send_timeout 30s;
       proxy_read_timeout 30s;

       # Remove sensitive headers
       proxy_set_header Cookie "";
     }

     # Rate limiting
     limit_req_zone $workspace_host zone=per_workspace:10m rate=10r/s;
     limit_req zone=per_workspace burst=20 nodelay;
   }
   ```

2. **Alternative: Serverless Proxy (AWS Lambda + API Gateway)**
   ```typescript
   // lambda/proxy-handler.ts
   import type { APIGatewayProxyEvent, APIGatewayProxyResult } from 'aws-lambda';

   export async function handler(
     event: APIGatewayProxyEvent
   ): Promise<APIGatewayProxyResult> {
     const match = event.path.match(/^\/singlecase\/([^/]+)(\/.*)?$/);
     if (!match) {
       return {
         statusCode: 400,
         body: JSON.stringify({ error: 'Invalid proxy URL format' }),
       };
     }

     const workspaceHost = decodeURIComponent(match[1]);
     const restPath = match[2] || '';

     if (!isValidWorkspaceHost(workspaceHost)) {
       return {
         statusCode: 400,
         body: JSON.stringify({ error: 'Invalid workspace host' }),
       };
     }

     try {
       const upstreamUrl = `https://${workspaceHost}${restPath}`;
       const response = await fetch(upstreamUrl, {
         method: event.httpMethod,
         headers: sanitizeHeaders(event.headers),
         body: event.body,
       });

       return {
         statusCode: response.status,
         headers: {
           'Access-Control-Allow-Origin': event.headers.origin || '*',
           'Access-Control-Allow-Credentials': 'true',
           'Content-Type': response.headers.get('content-type') || 'application/json',
         },
         body: await response.text(),
       };
     } catch (error) {
       console.error('Proxy error:', error);
       return {
         statusCode: 502,
         body: JSON.stringify({ error: 'Upstream service unavailable' }),
       };
     }
   }
   ```

### Frontend Updates

1. **Update services/singlecase.ts**
   - Ensure all API requests use proxy path format
   - Current implementation already uses `/singlecase/{host}/publicapi/v1/...` format
   - Verify no hardcoded workspace hosts in code

2. **Add Proxy Health Check**
   ```typescript
   // utils/proxyHealthCheck.ts
   export async function checkProxyHealth(workspaceHost: string): Promise<boolean> {
     try {
       const response = await fetch(
         `/singlecase/${encodeURIComponent(workspaceHost)}/publicapi/v1/health`,
         { method: 'GET', timeout: 5000 }
       );
       return response.ok;
     } catch {
       return false;
     }
   }
   ```

### Environment Variables

1. **Development (.env)**
   ```
   # No changes needed - proxy runs in webpack dev server
   ```

2. **Production (.env.production)**
   ```
   # Proxy service URL (if using separate proxy service)
   PROXY_SERVICE_URL=https://proxy.singlecase.com

   # Allowed workspace host patterns
   ALLOWED_WORKSPACE_PATTERNS=*.singlecase.ch,*.singlecase.com
   ```

## API Integration Patterns

1. **Frontend API Request**
   ```typescript
   // In services/singlecase.ts
   async function resolveBaseUrl(): Promise<string> {
     const storedHostRaw = await getStored(STORAGE_KEYS.workspaceHost);
     const host = normalizeHost(storedHostRaw || '');

     if (!host) {
       throw new Error('Workspace host is missing');
     }

     // Proxy path format
     return `/singlecase/${encodeURIComponent(host)}/publicapi/v1`;
   }

   // Example API call
   const baseUrl = await resolveBaseUrl();
   const response = await fetch(`${baseUrl}/cases`, {
     headers: {
       'Authentication': token,
       'Content-Type': 'application/json',
     },
   });
   ```

2. **Proxy Request Flow**
   ```
   Frontend: GET /singlecase/customer1.singlecase.ch/publicapi/v1/cases
        ↓
   Proxy: Extract "customer1.singlecase.ch" from path
        ↓
   Proxy: Forward to https://customer1.singlecase.ch/publicapi/v1/cases
        ↓
   Upstream: Process request, return response
        ↓
   Proxy: Add CORS headers, forward response
        ↓
   Frontend: Receive response
   ```

## Reference Implementation

Review the demo's proxy implementation:

1. **Current Proxy**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/webpack.config.js`
   - Lines 134-198: Existing relay middleware
   - **ENHANCE** with validation, logging, error handling

2. **API Service**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecase.ts`
   - `resolveBaseUrl()` function already uses proxy format
   - No changes needed if proxy format is consistent

## Dependencies

- **Required for**: All API integration stories (FE-001 through FE-004, FE-007)
- **Deployment**: Requires production proxy infrastructure (nginx or serverless)
- **Security**: Requires SSL certificates for production domain

## Notes

1. **Security Considerations**
   - Never proxy to arbitrary hosts (open proxy vulnerability)
   - Whitelist only trusted workspace host patterns
   - Validate host format before proxying
   - Remove authentication cookies (don't leak credentials)
   - Add rate limiting to prevent abuse

2. **Performance**
   - Proxy adds latency (~50-100ms)
   - Use keep-alive connections to upstream
   - Cache DNS lookups for workspace hosts
   - Monitor proxy response times

3. **Monitoring**
   - Log all proxy requests (for debugging)
   - Track proxy error rates per workspace
   - Alert on high error rates (>5%)
   - Monitor upstream availability

4. **Testing Strategy**
   - Test with multiple workspace hosts
   - Test invalid host rejection
   - Test CORS preflight requests
   - Test timeout handling (slow upstream)
   - Test error scenarios (unreachable upstream)
   - Load test proxy with concurrent requests

5. **Production Deployment**
   - Deploy proxy behind load balancer
   - Use health checks for proxy instances
   - Set up auto-scaling for proxy service
   - Configure SSL termination at load balancer
   - Add CloudFlare or CDN for DDoS protection

6. **Future Enhancements**
   - Cache frequently accessed workspace hosts
   - Implement circuit breaker for failing upstreams
   - Add request/response compression
   - Support WebSocket proxying (if needed for real-time features)

## Definition of Done

- [ ] Webpack dev server proxy middleware enhanced
- [ ] Workspace host validation implemented
- [ ] Invalid host requests rejected (400 Bad Request)
- [ ] CORS headers configured correctly
- [ ] Preflight OPTIONS requests handled
- [ ] Production proxy infrastructure documented (nginx or Lambda)
- [ ] Proxy request logging added
- [ ] Error handling for unreachable upstreams
- [ ] Timeout handling (5s connect, 30s read)
- [ ] Rate limiting configured (per workspace)
- [ ] Proxy health check endpoint working
- [ ] Unit tests for host validation
- [ ] Integration tests for proxy flow
- [ ] Load tested with 100+ concurrent requests
- [ ] Documentation updated with proxy architecture
- [ ] Production deployment guide written
