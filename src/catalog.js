function rows(db, sql) {
  return db.exec(sql)[0]?.values ?? [];
}

function normalize(text) {
  return text.normalize('NFKD').replace(/\p{M}/gu, '').toLowerCase();
}

export function readCatalog(db) {
  const books = new Map();
  for (const [id, title, path, hasCover] of rows(db, 'SELECT id, title, path, has_cover FROM books ORDER BY sort COLLATE NOCASE, id')) {
    books.set(id, { title, path, hasCover: Boolean(hasCover), authors: [], formats: [] });
  }
  for (const [id, name] of rows(db, 'SELECT link.book, authors.name FROM books_authors_link AS link JOIN authors ON authors.id = link.author ORDER BY link.id')) {
    books.get(id)?.authors.push(name);
  }
  for (const [id, name, format] of rows(db, 'SELECT book, name, format FROM data ORDER BY format')) {
    books.get(id)?.formats.push({ name, format });
  }
  return [...books.values()];
}

export function matchesBook(book, query) {
  const haystack = normalize([book.title, ...book.authors].join(' '));
  return normalize(query).trim().split(/\s+/).every(word => haystack.includes(word));
}
