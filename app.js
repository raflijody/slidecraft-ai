const form = document.getElementById('generator-form');
const statusEl = document.getElementById('form-status');
const submitBtn = document.getElementById('submit-btn');

function showFile(inputId, zoneId, defaultTitle) {
  const input = document.getElementById(inputId);
  const zone = document.getElementById(zoneId);
  input.addEventListener('change', () => {
    const file = input.files[0];
    const title = zone.querySelector('.upload-copy b');
    const sub = zone.querySelector('.upload-copy small');
    if (file) {
      title.textContent = file.name;
      sub.textContent = `${(file.size / 1024 / 1024).toFixed(2)} MB · siap digunakan`;
      zone.style.borderColor = '#73917e';
    } else {
      title.textContent = defaultTitle;
      zone.style.borderColor = '';
    }
  });
}
showFile('source_file', 'source-drop', 'Unggah materi kuliah atau catatan');
showFile('template_file', 'template-drop', 'Gunakan desain template kamu');

document.querySelectorAll('.style-card').forEach(card => {
  card.addEventListener('click', () => {
    document.querySelectorAll('.style-card').forEach(c => c.classList.remove('selected'));
    card.classList.add('selected');
    const radio = card.querySelector('input');
    radio.checked = true;
  });
});

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const topic = document.getElementById('topic').value.trim();
  if (!topic) return;
  submitBtn.disabled = true;
  submitBtn.querySelector('span:first-child').textContent = 'Menyusun presentasi…';
  statusEl.textContent = 'Sedang menyusun alur dan menyiapkan file. Jangan tutup halaman ini.';
  try {
    const data = new FormData(form);
    const response = await fetch('/api/generate', { method: 'POST', body: data });
    if (!response.ok) {
      let message = 'Terjadi kesalahan saat membuat presentasi.';
      try { const payload = await response.json(); message = payload.detail || message; } catch {}
      throw new Error(message);
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'slidecraft-presentasi.pptx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    const mode = response.headers.get('X-Generation-Mode');
    const warning = response.headers.get('X-Template-Warning');
    statusEl.textContent = (mode === 'fallback' ? 'File dibuat dalam mode draf tanpa API AI. ' : 'File PPTX berhasil dibuat. ') + (warning || 'Periksa hasil sebelum digunakan.');
  } catch (error) {
    statusEl.textContent = `Belum berhasil: ${error.message}`;
  } finally {
    submitBtn.disabled = false;
    submitBtn.querySelector('span:first-child').textContent = 'Susun presentasi';
  }
});
