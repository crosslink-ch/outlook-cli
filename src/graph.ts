const GRAPH_ORIGIN = 'https://graph.microsoft.com';
const GRAPH_ROOT = `${GRAPH_ORIGIN}/v1.0`;
const SUMMARY_FIELDS = 'id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments';

export interface MessageSummary {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  receivedDateTime?: string;
  bodyPreview?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
}

interface Page<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export interface MessageDetail extends MessageSummary {
  body?: { contentType: string; content: string };
  toRecipients?: unknown[];
  ccRecipients?: unknown[];
  sentDateTime?: string;
  webLink?: string;
}

export class GraphClient {
  constructor(private readonly token: string, private readonly transport: typeof fetch = fetch) {}

  private async get<T>(url: URL, preferText = false): Promise<T> {
    if (url.origin !== GRAPH_ORIGIN || (url.pathname !== '/v1.0/me' && !url.pathname.startsWith('/v1.0/me/'))) {
      throw new Error('Refusing a Graph URL outside the signed-in mailbox');
    }
    const response = await this.transport(url.toString(), {
      method: 'GET',
      redirect: 'error',
      signal: AbortSignal.timeout(15_000),
      headers: {
        Authorization: `Bearer ${this.token}`,
        ...(preferText ? { Prefer: 'outlook.body-content-type="text"' } : {}),
      },
    });
    if (!response.ok) {
      let code = '';
      try { code = String(((await response.json()) as { error?: { code?: string } }).error?.code ?? ''); } catch { /* no JSON */ }
      throw new Error(`Microsoft Graph returned HTTP ${response.status}${code ? ` (${code})` : ''}`);
    }
    return await response.json() as T;
  }

  async me(): Promise<{ id: string; displayName?: string; mail?: string; userPrincipalName?: string }> {
    const url = new URL(`${GRAPH_ROOT}/me`);
    url.searchParams.set('$select', 'id,displayName,mail,userPrincipalName');
    return this.get(url);
  }

  async read(id: string): Promise<MessageDetail> {
    if (!id.trim() || id.length > 1024) throw new Error('Message ID is required (max 1024 characters)');
    const url = new URL(`${GRAPH_ROOT}/me/messages/${encodeURIComponent(id)}`);
    url.searchParams.set('$select', 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,isRead,hasAttachments,webLink');
    return this.get<MessageDetail>(url, true);
  }

  async search(query: string, limit: number): Promise<MessageSummary[]> {
    if (!query.trim() || query.length > 256) throw new Error('Search query must be 1–256 characters');
    if (!Number.isSafeInteger(limit) || limit < 1 || limit > 100) throw new Error('Limit must be 1–100');
    const url = new URL(`${GRAPH_ROOT}/me/messages`);
    url.searchParams.set('$search', `"${query.trim().replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`);
    url.searchParams.set('$top', String(limit));
    url.searchParams.set('$select', SUMMARY_FIELDS);
    return this.collect(url, limit);
  }

  async list(folder: string, limit: number): Promise<MessageSummary[]> {
    if (!Number.isSafeInteger(limit) || limit < 1 || limit > 100) throw new Error('Limit must be 1–100');
    if (!folder || folder.includes('/') || folder === '.' || folder === '..') throw new Error('Invalid folder');
    const url = new URL(`${GRAPH_ROOT}/me/mailFolders/${encodeURIComponent(folder)}/messages`);
    url.searchParams.set('$top', String(limit));
    url.searchParams.set('$select', SUMMARY_FIELDS);
    return this.collect(url, limit);
  }

  private async collect(url: URL, limit: number): Promise<MessageSummary[]> {
    const results: MessageSummary[] = [];
    const seen = new Set<string>();
    let next: URL | undefined = url;
    while (next && results.length < limit) {
      if (seen.size >= 20 || seen.has(next.href)) throw new Error('Microsoft Graph returned excessive or repeated pages');
      seen.add(next.href);
      const page: Page<MessageSummary> = await this.get(next);
      if (!Array.isArray(page.value)) throw new Error('Invalid Microsoft Graph page');
      results.push(...page.value.slice(0, limit - results.length));
      next = page['@odata.nextLink'] ? new URL(page['@odata.nextLink']) : undefined;
    }
    return results;
  }
}
