/**
 * Browser-compatible, lightweight MCP SSE Client
 * Uses standard EventSource and fetch to communicate with remote SSE MCP servers.
 * Automatically handles routing requests through a local proxy to bypass CORS.
 */
export class SSEMCPClient {
  private eventSource: EventSource | null = null;
  private postUrl: string | null = null;
  private pendingRequests = new Map<
    number | string,
    { resolve: (val: any) => void; reject: (err: any) => void }
  >();
  private nextId = 1;
  private resolvedUrl: string;

  constructor(private rawSseUrl: string) {
    // If the URL targets the Go India Stocks MCP production server, map it to the local proxy
    // path to bypass CORS. Otherwise, connect directly.
    if (rawSseUrl.startsWith('https://goindia-mcp.fly.dev')) {
      this.resolvedUrl = rawSseUrl.replace('https://goindia-mcp.fly.dev', '/proxy/goindia');
    } else {
      this.resolvedUrl = rawSseUrl;
    }
  }

  /**
   * Connects to the SSE endpoint and awaits the initial handshake
   * which delivers the post/message endpoint.
   */
  connect(timeoutMs = 15000): Promise<void> {
    return new Promise((resolve, reject) => {
      console.log(`[MCP] Connecting to SSE: ${this.resolvedUrl}`);
      
      const timeout = setTimeout(() => {
        this.close();
        reject(new Error(`MCP connection timeout after ${timeoutMs}ms`));
      }, timeoutMs);

      try {
        this.eventSource = new EventSource(this.resolvedUrl);

        // Standard MCP SSE handshake: server sends the target POST endpoint in an 'endpoint' event
        this.eventSource.addEventListener('endpoint', (event: MessageEvent) => {
          clearTimeout(timeout);
          try {
            const rawEndpoint = event.data;
            console.log(`[MCP] Received endpoint event: ${rawEndpoint}`);
            
            // Resolve relative paths against the connection URL
            let resolvedPostUrl = new URL(rawEndpoint, new URL(this.resolvedUrl, window.location.href)).toString();
            
            // If the resolved POST URL points to the production domain, route it through our proxy
            if (resolvedPostUrl.startsWith('https://goindia-mcp.fly.dev')) {
              resolvedPostUrl = resolvedPostUrl.replace('https://goindia-mcp.fly.dev', '/proxy/goindia');
            }
            
            this.postUrl = resolvedPostUrl;
            console.log(`[MCP] Resolved POST endpoint: ${this.postUrl}`);
            resolve();
          } catch (err) {
            this.close();
            reject(new Error(`Failed to resolve message endpoint: ${err instanceof Error ? err.message : String(err)}`));
          }
        });

        // Listen for responses to client requests
        this.eventSource.addEventListener('message', (event: MessageEvent) => {
          try {
            const message = JSON.parse(event.data);
            if (message.id !== undefined) {
              const pending = this.pendingRequests.get(message.id);
              if (pending) {
                this.pendingRequests.delete(message.id);
                if (message.error) {
                  pending.reject(message.error);
                } else {
                  pending.resolve(message.result);
                }
              }
            }
          } catch (err) {
            console.error('[MCP] Error parsing message from server:', err);
          }
        });

        this.eventSource.onerror = (err) => {
          clearTimeout(timeout);
          console.error('[MCP] SSE transport error:', err);
          reject(new Error('MCP SSE connection error. Please verify the URL/token and that the proxy is functioning.'));
        };
      } catch (err) {
        clearTimeout(timeout);
        reject(err);
      }
    });
  }

  /**
   * Sends a JSON-RPC 2.0 request to the message endpoint
   */
  async sendRequest(method: string, params: any = {}): Promise<any> {
    if (!this.postUrl) {
      throw new Error('Not connected to MCP server. Call connect() first.');
    }

    const id = this.nextId++;
    const requestBody = {
      jsonrpc: '2.0',
      method,
      params,
      id,
    };

    const responsePromise = new Promise((resolve, reject) => {
      this.pendingRequests.set(id, { resolve, reject });
    });

    try {
      const response = await fetch(this.postUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        this.pendingRequests.delete(id);
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
    } catch (err) {
      this.pendingRequests.delete(id);
      throw new Error(`Failed to transmit MCP request: ${err instanceof Error ? err.message : String(err)}`);
    }

    return responsePromise;
  }

  /**
   * Returns the list of tools exposed by the MCP server
   */
  async listTools(): Promise<{ tools: Array<{ name: string; description?: string; inputSchema?: any }> }> {
    return this.sendRequest('tools/list');
  }

  /**
   * Calls a tool by name with arguments
   */
  async callTool(name: string, args: any): Promise<{ content: Array<{ type: string; text?: string; [key: string]: any }>; isError?: boolean }> {
    return this.sendRequest('tools/call', { name, arguments: args });
  }

  /**
   * Closes the SSE connection and cancels pending requests
   */
  close() {
    console.log('[MCP] Closing connection');
    if (this.eventSource) {
      this.eventSource.close();
      this.eventSource = null;
    }
    this.postUrl = null;
    this.pendingRequests.forEach((req) => req.reject(new Error('Connection closed by client')));
    this.pendingRequests.clear();
  }
}
