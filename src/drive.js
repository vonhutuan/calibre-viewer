const graph = 'https://graph.microsoft.com/v1.0/me/drive';
const downloadProperty = '@microsoft.graph.downloadUrl';
const alternateDownloadProperty = '@content.downloadUrl';

export class DriveError extends Error {
  constructor(status, operation) {
    super(status ? `OneDrive returned HTTP ${status} during ${operation}.` : `OneDrive: ${operation}.`);
    this.status = status;
    this.operation = operation;
  }
}

export function itemEndpoint(folderId, path) {
  const segments = path.split('/');
  if (segments.some(segment => !segment || segment === '.' || segment === '..' || segment.includes('\\'))) {
    throw new Error('Invalid library path.');
  }
  const encoded = segments.map(segment => encodeURIComponent(segment)
    .replace(/[!'()*]/g, char => '%' + char.charCodeAt(0).toString(16).toUpperCase())).join('/');
  const folder = folderId === 'root' ? 'root' : `items/${encodeURIComponent(folderId)}`;
  return `${graph}/${folder}:/${encoded}`;
}

export class OneDrive {
  constructor(getToken, request = (url, options) => globalThis.fetch(url, options)) {
    this.getToken = getToken;
    this.request = request;
  }

  async item(folderId, path, properties, signal) {
    return this.itemMetadata(itemEndpoint(folderId, path), properties, signal);
  }

  async itemMetadata(endpoint, properties, signal) {
    const url = new URL(endpoint);
    if (properties) url.searchParams.set('$select', properties);
    const token = await this.getToken();
    let response;
    try {
      response = await this.request(url.href, {
        headers: { Authorization: `Bearer ${token}` },
        cache: 'no-store',
        signal,
      });
    } catch (error) {
      if (signal?.aborted) throw error;
      throw new DriveError(undefined, 'Graph connection');
    }
    if (!response.ok) throw new DriveError(response.status, 'file lookup');
    return response.json();
  }

  async downloadUrl(folderId, path, signal) {
    const item = await this.item(folderId, path, `id,${downloadProperty}`, signal);
    let link = item[downloadProperty] || item[alternateDownloadProperty];
    if (!link) {
      const endpoint = item.id ? `${graph}/items/${encodeURIComponent(item.id)}` : itemEndpoint(folderId, path);
      const metadata = await this.itemMetadata(endpoint, undefined, signal);
      link = metadata[downloadProperty] || metadata[alternateDownloadProperty];
    }
    if (!link) {
      throw new DriveError(undefined, 'missing download link');
    }
    const url = new URL(link);
    if (url.protocol !== 'https:') throw new Error('Invalid OneDrive download link.');
    return url.href;
  }

  async file(folderId, path, signal) {
    const url = await this.downloadUrl(folderId, path, signal);
    let response;
    try {
      response = await this.request(url, { cache: 'no-store', signal, referrerPolicy: 'no-referrer' });
    } catch (error) {
      if (signal?.aborted) throw error;
      throw new DriveError(undefined, 'download connection');
    }
    if (!response.ok) throw new DriveError(response.status, 'file download');
    return response;
  }
}
