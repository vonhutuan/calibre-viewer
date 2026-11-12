import { BrowserCacheLocation, InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser';
import initSqlJs from 'sql.js';
import wasmUrl from 'sql.js/dist/sql-wasm.wasm?url';
import { readCatalog, matchesBook } from './catalog.js';
import { DriveError, OneDrive } from './drive.js';
import './style.css';

const elements = Object.fromEntries(
  [...document.querySelectorAll('[id]')].map(element => [element.id, element]),
);
const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID?.trim();
const folderPath = import.meta.env.VITE_CALIBRE_FOLDER?.trim() || 'Calibre Library';
const redirectUri = new URL(location.pathname, location.origin).href;
const scopes = ['Files.Read'];
const objectUrls = new Map();
let auth;
let account;
let sqlite;
let controller;
let coverObserver;
let folderId;
let books = [];
let currentBook;

function showError(error, target, fallback) {
  if (error.name === 'AbortError') return;
  const needsSignIn = error instanceof InteractionRequiredAuthError || error.status === 401;
  elements['retry-sign-in'].hidden = !needsSignIn;
  target.textContent = needsSignIn ? 'Your OneDrive session needs a new sign-in.' : fallback;
}

function clearLibrary() {
  controller?.abort();
  coverObserver?.disconnect();
  for (const url of objectUrls.values()) URL.revokeObjectURL(url);
  objectUrls.clear();
  books = [];
  folderId = undefined;
  currentBook = undefined;
  elements.details.close();
  elements['book-title'].textContent = '';
  elements['book-authors'].textContent = '';
  elements['download-status'].textContent = '';
  elements.formats.replaceChildren();
  elements.books.replaceChildren();
  elements.count.textContent = '';
  elements.library.hidden = true;
}

const drive = new OneDrive(async () => {
  const result = await auth.acquireTokenSilent({ account, scopes });
  return result.accessToken;
});

async function loadLibrary() {
  clearLibrary();
  controller = new AbortController();
  const { signal } = controller;
  elements.status.textContent = 'Loading your library…';
  elements['retry-sign-in'].hidden = true;
  elements.refresh.disabled = true;
  try {
    const folder = await drive.item('root', folderPath, 'id,folder', signal);
    if (!folder.folder || !folder.id) throw new Error('Expected a library folder.');
    const response = await drive.file(folder.id, 'metadata.db', signal);
    const bytes = new Uint8Array(await response.arrayBuffer());
    sqlite ??= await initSqlJs({ locateFile: () => wasmUrl });
    if (signal.aborted) return;
    const db = new sqlite.Database(bytes);
    try {
      books = readCatalog(db);
    } finally {
      db.close();
    }
    folderId = folder.id;
    elements.library.hidden = false;
    elements.status.textContent = '';
    renderBooks();
  } catch (error) {
    if (!signal.aborted) {
      const message = error.status === 404
        ? `Could not find ${folderPath}/metadata.db in this OneDrive.`
        : 'Could not load your library. Try Refresh.';
      showError(error, elements.status, message);
    }
  } finally {
    if (!signal.aborted) elements.refresh.disabled = false;
  }
}

async function showCover(book, image, signal) {
  try {
    let url = objectUrls.get(book.path);
    if (!url) {
      const response = await drive.file(folderId, `${book.path}/cover.jpg`, signal);
      const blob = await response.blob();
      if (signal.aborted) return;
      url = objectUrls.get(book.path);
      if (!url) {
        url = URL.createObjectURL(blob);
        objectUrls.set(book.path, url);
      }
    }
    image.src = url;
    image.hidden = false;
  } catch {
    // The cover title remains visible if the image is unavailable.
  }
}

function renderBooks() {
  coverObserver?.disconnect();
  const signal = controller.signal;
  const visible = books.filter(book => matchesBook(book, elements.search.value));
  const covers = new Map();
  coverObserver = new IntersectionObserver((entries, observer) => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      observer.unobserve(entry.target);
      void showCover(covers.get(entry.target), entry.target.querySelector('img'), signal);
    }
  }, { rootMargin: '160px' });

  const items = visible.map(book => {
    const item = elements['book-card'].content.firstElementChild.cloneNode(true);
    item.querySelector('.title').textContent = book.title;
    item.querySelector('.cover-title').textContent = book.title;
    item.querySelector('.authors').textContent = book.authors.join(', ');
    item.querySelector('button').addEventListener('click', () => showBook(book));
    if (book.hasCover) covers.set(item, book);
    return item;
  });
  elements.books.replaceChildren(...items);
  for (const item of covers.keys()) coverObserver.observe(item);
  elements.count.textContent = visible.length
    ? `${visible.length} of ${books.length} books`
    : (books.length ? 'No matching books' : 'No books in this library');
}

function showBook(book) {
  currentBook = book;
  elements['book-title'].textContent = book.title;
  elements['book-authors'].textContent = book.authors.join(', ');
  elements['download-status'].textContent = book.formats.length ? '' : 'No ebook files are listed for this book.';
  elements.formats.replaceChildren(...book.formats.map(format => {
    const button = document.createElement('button');
    button.className = 'primary';
    button.type = 'button';
    button.textContent = `Download ${format.format}`;
    button.addEventListener('click', () => downloadBook(book, format, button));
    return button;
  }));
  elements.details.showModal();
}

async function downloadBook(book, format, button) {
  const { signal } = controller;
  const filename = `${format.name}.${format.format.toLowerCase()}`;
  const status = elements['download-status'];
  status.textContent = `Preparing ${format.format}…`;
  button.disabled = true;
  try {
    const response = await drive.file(folderId, `${book.path}/${filename}`, signal);
    const blob = await response.blob();
    if (signal.aborted) return;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30_000);
    if (currentBook === book && elements.details.open) status.textContent = 'Download requested.';
  } catch (error) {
    if (!signal.aborted && currentBook === book && elements.details.open) {
      const message = error.status === 404
        ? `${format.format} is unavailable in OneDrive.`
        : 'Download failed. Try again.';
      showError(error, status, message);
      if (!elements['retry-sign-in'].hidden) {
        elements.details.close();
        elements.status.textContent = status.textContent;
      }
    }
  } finally {
    button.disabled = false;
  }
}

async function signIn() {
  try {
    await auth.loginRedirect({ scopes });
  } catch (error) {
    showError(error, elements.status, 'Could not start Microsoft sign-in. Try again.');
  }
}

async function start() {
  if (!clientId) {
    elements.status.textContent = 'Microsoft sign-in needs to be configured for this deployment.';
    return;
  }
  auth = new PublicClientApplication({
    auth: { clientId, authority: 'https://login.microsoftonline.com/consumers', redirectUri },
    cache: { cacheLocation: BrowserCacheLocation.SessionStorage },
  });
  try {
    await auth.initialize();
    const result = await auth.handleRedirectPromise();
    const accounts = auth.getAllAccounts();
    account = result?.account ?? auth.getActiveAccount() ?? (accounts.length === 1 ? accounts[0] : null);
    if (account) {
      auth.setActiveAccount(account);
      elements.welcome.hidden = true;
      elements.actions.hidden = false;
      await loadLibrary();
    }
  } catch {
    elements.status.textContent = 'Could not complete Microsoft sign-in. Try again.';
  } finally {
    elements['sign-in'].disabled = false;
  }
}

elements['sign-in'].addEventListener('click', signIn);
elements['retry-sign-in'].addEventListener('click', signIn);
elements.refresh.addEventListener('click', loadLibrary);
elements.search.addEventListener('input', renderBooks);
elements['sign-out'].addEventListener('click', async () => {
  clearLibrary();
  elements.search.value = '';
  elements.actions.hidden = true;
  elements.welcome.hidden = false;
  elements['retry-sign-in'].hidden = true;
  elements.status.textContent = '';
  try {
    await auth.logoutRedirect({ account, postLogoutRedirectUri: redirectUri });
  } catch {
    elements.status.textContent = 'Microsoft sign-out could not finish. Close this tab to clear the viewer session.';
  }
});
void start();
