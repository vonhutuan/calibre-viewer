function rows(db, sql) {
  return db.exec(sql)[0]?.values ?? [];
}

function normalize(text) {
  return text.normalize('NFKD').replace(/\p{M}/gu, '').toLowerCase();
}

const authorCollator = new Intl.Collator(undefined, { sensitivity: 'base' });

export function sortBooks(books, order) {
  if (order === 'author') {
    return [...books].sort((a, b) => authorCollator.compare(
      a.authorSort || a.authors[0] || '',
      b.authorSort || b.authors[0] || '',
    ));
  }
  if (order === 'added' || order === 'published') {
    return [...books].sort((a, b) => (b[order] ?? -Infinity) - (a[order] ?? -Infinity));
  }
  return books;
}

export function readCatalog(db) {
  const books = new Map();
  for (const [id, title, path, hasCover, description, series, seriesIndex, authorSort, added, published] of rows(db, `
    SELECT books.id, title, path, has_cover, comments.text, series.name, series_index,
           author_sort, CAST(strftime('%s', timestamp) AS INTEGER),
           CASE WHEN date(pubdate) != '0101-01-01' THEN CAST(strftime('%s', pubdate) AS INTEGER) END
    FROM books
    LEFT JOIN comments ON comments.book = books.id
    LEFT JOIN books_series_link AS link ON link.book = books.id
    LEFT JOIN series ON series.id = link.series
    ORDER BY books.sort COLLATE NOCASE, books.id
  `)) {
    books.set(id, { title, path, hasCover: Boolean(hasCover), description, series, seriesIndex, authorSort, added, published, authors: [], formats: [], tags: [] });
  }
  for (const [id, name] of rows(db, 'SELECT link.book, tags.name FROM books_tags_link AS link JOIN tags ON tags.id = link.tag ORDER BY tags.name COLLATE NOCASE')) {
    books.get(id)?.tags.push(name);
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
