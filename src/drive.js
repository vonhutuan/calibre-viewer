const graph = 'https://graph.microsoft.com/v1.0/me/drive';
const downloadProperty = '@microsoft.graph.downloadUrl';

export class DriveError extends Error {
  constructor(status) {
    super(`OneDrive returned HTTP ${status}.`);
    this.status = status;
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
    const url = new URL(itemEndpoint(folderId, path));
    url.searchParams.set('$select', properties);
    const token = await this.getToken();
    const response = await this.request(url.href, {
      headers: { Authorization: `Bearer ${token}` },
      cache: 'no-store',
      signal,
    });
    if (!response.ok) throw new DriveError(response.status);
    return response.json();
  }

  async downloadUrl(folderId, path, signal) {
    const item = await this.item(folderId, path, `id,${downloadProperty}`, signal);
    const url = new URL(item[downloadProperty]);
    if (url.protocol !== 'https:') throw new Error('Invalid OneDrive download link.');
    return url.href;
  }

  async file(folderId, path, signal) {
    const url = await this.downloadUrl(folderId, path, signal);
    const response = await this.request(url, { cache: 'no-store', signal, referrerPolicy: 'no-referrer' });
    if (!response.ok) throw new DriveError(response.status);
    return response;
  }
}
