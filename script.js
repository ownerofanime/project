const intro = document.getElementById('intro');
const galleryGrid = document.getElementById('galleryGrid');

const lockedPane = document.getElementById('lockedPane');
const uploadPane = document.getElementById('uploadPane');
const passwordInput = document.getElementById('password');
const unlockBtn = document.getElementById('unlockBtn');
const authMessage = document.getElementById('authMessage');

const photoTitle = document.getElementById('photoTitle');
const photoSize = document.getElementById('photoSize');
const photoFile = document.getElementById('photoFile');
const addPhotoBtn = document.getElementById('addPhotoBtn');
const uploadMessage = document.getElementById('uploadMessage');

const STORAGE_KEY = 'aria_portfolio_photos_v1';
const SALT = 'aria-lens-enterprise-salt-v1';
// Demo hash for password: OwnerOnly#2026
const OWNER_PASSWORD_HASH_HEX = '555bd63502fe7e2d77901a3f90d80519a69b36c579e6b9642dcd6fa9f93a1929';

const defaultPhotos = [
  { title: 'Shoreline Calm', size: 'large', src: 'https://images.unsplash.com/photo-1507525428034-b723cf961d3e?auto=format&fit=crop&w=1200&q=80' },
  { title: 'White Horizon', size: 'medium', src: 'https://images.unsplash.com/photo-1493244040629-496f6d136cc3?auto=format&fit=crop&w=900&q=80' },
  { title: 'Urban Silence', size: 'small', src: 'https://images.unsplash.com/photo-1495422964407-28ecb5f68cc8?auto=format&fit=crop&w=900&q=80' },
  { title: 'Motion', size: 'wide', src: 'https://images.unsplash.com/photo-1449824913935-59a10b8d2000?auto=format&fit=crop&w=1200&q=80' },
  { title: 'Portrait Frame', size: 'tall', src: 'https://images.unsplash.com/photo-1506794778202-cad84cf45f1d?auto=format&fit=crop&w=1000&q=80' }
];

let ownerUnlocked = false;
let photos = loadPhotos();

window.addEventListener('load', () => {
  setTimeout(() => intro.classList.add('hidden'), 1200);
  renderGallery();
});

unlockBtn.addEventListener('click', async () => {
  const candidate = passwordInput.value || '';
  const hashed = await sha256Hex(`${SALT}:${candidate}`);

  if (hashed === OWNER_PASSWORD_HASH_HEX) {
    ownerUnlocked = true;
    lockedPane.classList.add('hidden');
    uploadPane.classList.remove('hidden');
    uploadPane.setAttribute('aria-hidden', 'false');
    authMessage.textContent = '';
  } else {
    ownerUnlocked = false;
    authMessage.textContent = 'Access denied: password is incorrect.';
  }
});

addPhotoBtn.addEventListener('click', async () => {
  if (!ownerUnlocked) {
    uploadMessage.textContent = 'Owner mode is locked.';
    return;
  }

  const file = photoFile.files?.[0];
  const title = (photoTitle.value || '').trim();
  const size = photoSize.value;

  if (!file || !title) {
    uploadMessage.textContent = 'Please provide both a title and an image file.';
    return;
  }

  if (!file.type.startsWith('image/')) {
    uploadMessage.textContent = 'Only image files are allowed.';
    return;
  }

  const dataUrl = await fileToDataUrl(file);
  photos.unshift({ title, size, src: dataUrl });
  savePhotos(photos);
  renderGallery();

  photoTitle.value = '';
  photoFile.value = '';
  photoSize.value = 'medium';
  uploadMessage.textContent = 'Photo added successfully.';
});

function renderGallery() {
  galleryGrid.innerHTML = '';
  photos.forEach((photo) => {
    const tile = document.createElement('article');
    tile.className = `tile size-${photo.size}`;

    const image = document.createElement('img');
    image.src = photo.src;
    image.alt = photo.title;
    image.loading = 'lazy';

    const caption = document.createElement('p');
    caption.textContent = photo.title;

    tile.append(image, caption);
    galleryGrid.appendChild(tile);
  });
}

function loadPhotos() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [...defaultPhotos];

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    return [...defaultPhotos];
  } catch {
    return [...defaultPhotos];
  }
}

function savePhotos(collection) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(collection));
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function sha256Hex(text) {
  const encoder = new TextEncoder();
  const data = encoder.encode(text);
  const digest = await crypto.subtle.digest('SHA-256', data);
  const bytes = [...new Uint8Array(digest)];
  return bytes.map((b) => b.toString(16).padStart(2, '0')).join('');
}
