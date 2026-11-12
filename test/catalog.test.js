import test from 'node:test';
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readCatalog, matchesBook } from '../src/catalog.js';
import { itemEndpoint, OneDrive, DriveError } from '../src/drive.js';

test('reads Calibre metadata with ordered authors, cover flags and original format names', () => {
  const db = new DatabaseSync(':memory:');
  db.exec(`
    CREATE TABLE books (id, title, sort, path, has_cover, series_index);
    CREATE TABLE authors (id, name);
    CREATE TABLE books_authors_link (id, book, author);
    CREATE TABLE data (book, name, format);
    CREATE TABLE comments (book, text);
    CREATE TABLE series (id, name);
    CREATE TABLE books_series_link (book, series);
    CREATE TABLE tags (id, name);
    CREATE TABLE books_tags_link (book, tag);
    INSERT INTO books VALUES (1,'Café','Cafe','René/Café (1)',1,2),(2,'Empty','Empty','Empty (2)',0,1);
    INSERT INTO comments VALUES (1,'<p>A description.</p>');
    INSERT INTO series VALUES (3,'Coffee Stories');
    INSERT INTO books_series_link VALUES (1,3);
    INSERT INTO tags VALUES (5,'Fiction');
    INSERT INTO books_tags_link VALUES (1,5);
    INSERT INTO authors VALUES (1,'João'),(2,'René');
    INSERT INTO books_authors_link VALUES (2,1,1),(1,1,2);
    INSERT INTO data VALUES (1,'Original #1','EPUB'),(1,'Second','PDF');
  `);
  const [book, empty] = readCatalog({ exec: sql => [{ values: db.prepare(sql).all().map(Object.values) }] });
  assert.deepEqual(book.authors, ['René', 'João']);
  assert.equal(book.hasCover, true);
  assert.equal(book.description, '<p>A description.</p>');
  assert.equal(book.series, 'Coffee Stories');
  assert.equal(book.seriesIndex, 2);
  assert.deepEqual(book.tags, ['Fiction']);
  assert.deepEqual(book.formats, [{ name: 'Original #1', format: 'EPUB' }, { name: 'Second', format: 'PDF' }]);
  assert.deepEqual(empty.formats, []);
  assert.equal(empty.hasCover, false);
  assert.equal(empty.description, null);
  assert.deepEqual(empty.tags, []);
  assert.ok(matchesBook(book, 'cafe joao'));
  assert.equal(matchesBook(book, 'unrelated'), false);
  db.close();
});

test('encodes Graph path segments and rejects paths outside the library', () => {
  assert.match(itemEndpoint('root', 'René/Café #1 (2)/100%.epub'), /Ren%C3%A9\/Caf%C3%A9%20%231%20%282%29\/100%25.epub$/);
  assert.throws(() => itemEndpoint('root', '../outside'), /Invalid/);
});

test('requests fresh links with a token and fetches file bytes without one', async () => {
  const calls = [];
  const drive = new OneDrive(async () => 'token', async (url, options) => {
    calls.push([url, options]);
    if (url.includes('missing')) return new Response(null, { status: 404 });
    if (url.startsWith('https://graph.microsoft.com')) {
      return Response.json({ '@microsoft.graph.downloadUrl': 'https://files.example/book' });
    }
    return new Response('content');
  });
  assert.equal(await (await drive.file('folder', 'metadata.db')).text(), 'content');
  assert.equal(calls[0][1].headers.Authorization, 'Bearer token');
  assert.equal(calls[1][1].headers, undefined);
  await assert.rejects(drive.downloadUrl('folder', 'missing'), error => error instanceof DriveError && error.status === 404);
});

test('calls the default fetch with its global receiver', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = function () {
    assert.ok(this === globalThis, 'Browser fetch requires the global receiver');
    return Promise.resolve(Response.json({ id: 'folder', folder: {} }));
  };
  try {
    const drive = new OneDrive(async () => 'token');
    const folder = await drive.item('root', 'Calibre Library', 'id,folder');
    assert.equal(folder.id, 'folder');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('retries without a field filter when the OneDrive download annotation is omitted', async () => {
  const requests = [];
  const drive = new OneDrive(async () => 'token', async url => {
    requests.push(new URL(url));
    return requests.length === 1
      ? Response.json({ id: 'file!123' })
      : Response.json({ '@content.downloadUrl': 'https://files.example/metadata.db' });
  });
  assert.equal(await drive.downloadUrl('folder', 'metadata.db'), 'https://files.example/metadata.db');
  assert.equal(requests.length, 2);
  assert.match(requests[1].pathname, /\/items\/file!123$/);
  assert.equal(requests[1].search, '');
});
