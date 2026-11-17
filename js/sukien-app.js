/**
 * SỰ KIỆN PLAY v2.4.0 - Bàn Điều Hành Âm Thanh Hội Nghị, Đại Hội & Lễ Chào Cờ Offline
 * Tích hợp Web Audio API Synthesizer + HTML5 Audio + Lưu trữ IndexedDB Ngoại Tuyến 100%
 */

const DEFAULT_SOUND_TRACKS = [
  {
    id: 'quoc_ca',
    key: '1',
    title: 'Quốc Ca Việt Nam',
    subtitle: 'Nghi lễ Chào cờ · Trang trọng',
    color: '#dc2626',
    duration: 68,
    loop: false,
    volume: 0.9,
    synthType: 'anthem'
  },
  {
    id: 'quoc_te_ca',
    key: '2',
    title: 'Quốc Tế Ca (Nghi Lễ Đảng)',
    subtitle: 'Đại hội Đảng bộ · Sinh hoạt Đảng',
    color: '#eab308',
    duration: 64,
    loop: false,
    volume: 0.9,
    synthType: 'internationale'
  },
  {
    id: 'vo_tay',
    key: '3',
    title: 'Tiếng Vỗ Tay Đại Hội',
    subtitle: 'Hiệu ứng khán giả · Chúc mừng',
    color: '#84cc16',
    duration: 8,
    loop: false,
    volume: 0.85,
    synthType: 'applause'
  },
  {
    id: 'trao_thuong_1',
    key: '4',
    title: 'Nhạc Trao Thưởng 1',
    subtitle: 'Victory · Tuyên dương & Khen thưởng',
    color: '#f59e0b',
    duration: 45,
    loop: true,
    volume: 0.8,
    synthType: 'fanfare1'
  },
  {
    id: 'trao_thuong_2',
    key: '5',
    title: 'Nhạc Trao Thưởng 2',
    subtitle: 'Vinh danh · Trao Huy hiệu Đảng & Giấy khen',
    color: '#f97316',
    duration: 50,
    loop: true,
    volume: 0.8,
    synthType: 'fanfare2'
  },
  {
    id: 'dai_bieu',
    key: '6',
    title: 'Chào Mừng Đại Biểu',
    subtitle: 'Nhạc nền đón đại biểu lên sân khấu',
    color: '#a855f7',
    duration: 30,
    loop: true,
    volume: 0.8,
    synthType: 'welcome'
  }
];

let soundTracks = [...DEFAULT_SOUND_TRACKS];
let activePlayers = {};
let audioCtx = null;
let isMcDucking = false;
let isAutoFade = true;

const DB_NAME = 'SuKienPlayOfflineDB';
const STORE_NAME = 'custom_tracks';

function openAudioDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = (e) => reject(e.target.error);
  });
}

async function loadSavedCustomTracks() {
  try {
    const db = await openAudioDb();
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => resolve([]);
    });
  } catch (e) {
    return [];
  }
}

async function saveCustomTrackToDb(trackObj) {
  try {
    const db = await openAudioDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put(trackObj);
      tx.oncomplete = () => resolve(true);
      tx.onerror = (e) => reject(e);
    });
  } catch (e) {
    console.warn('IndexedDB error:', e);
  }
}

function getAudioContext() {
  if (!audioCtx) {
    const AudioContextClass = window.AudioContext || window.webkitAudioContext;
    audioCtx = new AudioContextClass();
  }
  if (audioCtx.state === 'suspended') {
    audioCtx.resume();
  }
  return audioCtx;
}

function formatDuration(sec) {
  const s = Math.max(0, Math.floor(sec || 0));
  const m = Math.floor(s / 60);
  const rem = s % 60;
  return `${m}:${rem < 10 ? '0' : ''}${rem}`;
}

function renderSoundboard() {
  const grid = document.getElementById('soundboardGrid');
  if (!grid) return;

  grid.innerHTML = soundTracks.map((t, idx) => {
    const p = activePlayers[t.id];
    const isPlaying = p && p.playing;
    const curTime = p ? p.currentTime : 0;
    return `
      <div class="sound-card ${isPlaying ? 'playing' : ''}" id="card_${t.id}">
        <div style="display:flex; align-items:center; justify-content:space-between; gap:10px;">
          <div>
            <div style="display:flex; align-items:center; gap:8px;">
              <span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:${t.color}; box-shadow:0 0 8px ${t.color};"></span>
              <strong style="font-size:1.08rem; color:var(--text-main);">${escapeHtml(t.title)}</strong>
            </div>
            <div style="font-size:0.8rem; color:var(--text-secondary); margin-top:2px;">${escapeHtml(t.subtitle)}</div>
          </div>
          <span class="card-version" title="Phím tắt bàn phím">Phím ${t.key || (idx + 1)}</span>
        </div>

        <div style="display:flex; align-items:center; gap:14px; margin:6px 0;">
          <button type="button" class="sound-play-btn"
            onmousedown="handlePlayButtonMouseDown(event, '${t.id}')"
            onclick="toggleTrack('${t.id}')"
            title="Nhấn để Phát/Tạm dừng, hoặc Giữ & Kéo sang trái để đưa nhạc về đầu">
            ${isPlaying ? '⏸' : '▶'}
          </button>
          <div style="flex:1;">
            <div style="display:flex; justify-content:space-between; font-family:var(--font-mono); font-size:0.82rem; color:var(--accent-color); margin-bottom:4px;">
              <span id="time_${t.id}">${formatDuration(curTime)} · ${formatDuration(t.duration)}</span>
              <span id="status_${t.id}" style="font-weight:700;">${isPlaying ? 'ĐANG PHÁT' : 'SẴN SÀNG'}</span>
            </div>
            <input type="range" id="seek_${t.id}" min="0" max="${t.duration}" step="0.1" value="${curTime}"
              oninput="seekTrack('${t.id}', this.value)"
              style="width:100%; accent-color:#facc15; cursor:pointer;">
          </div>
        </div>

        <div style="display:flex; align-items:center; justify-content:space-between; gap:12px; padding-top:8px; border-top:1px solid rgba(250,204,21,0.2);">
          <div style="display:flex; align-items:center; gap:6px; flex:1;">
            <span style="font-size:0.85rem;">🔊</span>
            <input type="range" min="0" max="1" step="0.01" value="${t.volume}"
              oninput="setTrackVolume('${t.id}', this.value)"
              style="width:100%; accent-color:#facc15; cursor:pointer;">
          </div>
          <label style="font-size:0.8rem; color:var(--text-secondary); display:flex; align-items:center; gap:5px; cursor:pointer;">
            <input type="checkbox" ${t.loop ? 'checked' : ''} onchange="setTrackLoop('${t.id}', this.checked)"> Lặp lại
          </label>
          <button type="button" class="btn-tool" style="padding:4px 8px; font-size:0.75rem;" onclick="rewindTrack('${t.id}')" title="Đưa nhạc về đầu 0:00">⏮ Đầu</button>
        </div>
      </div>
    `;
  }).join('');
}

function handlePlayButtonMouseDown(e, trackId) {
  const startX = e.clientX;
  const onMove = (moveEvt) => {
    if (startX - moveEvt.clientX > 28) {
      rewindTrack(trackId);
      if (window.showToast) window.showToast('Đã kéo đưa nhạc về đầu (0:00)!');
      window.removeEventListener('mousemove', onMove);
    }
  };
  const onUp = () => window.removeEventListener('mousemove', onMove);
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp, { once: true });
}

function startSynthEngine(track) {
  const ctx = getAudioContext();
  const masterGain = ctx.createGain();
  const targetVol = isMcDucking ? track.volume * 0.15 : track.volume;
  if (isAutoFade) {
    masterGain.gain.setValueAtTime(0.01, ctx.currentTime);
    masterGain.gain.linearRampToValueAtTime(targetVol, ctx.currentTime + 1.2);
  } else {
    masterGain.gain.setValueAtTime(targetVol, ctx.currentTime);
  }
  masterGain.connect(ctx.destination);

  let noteTimer = null;
  let stepIdx = 0;

  // Giai điệu trang trọng / sôi động cho từng thẻ sự kiện
  const melodies = {
    anthem: [261.63, 293.66, 329.63, 349.23, 392.00, 440.00, 392.00, 329.63, 261.63, 392.00, 523.25],
    internationale: [293.66, 392.00, 370.00, 440.00, 392.00, 329.63, 293.66, 246.94, 293.66, 392.00],
    fanfare1: [392.00, 392.00, 392.00, 523.25, 659.25, 587.33, 523.25, 659.25, 783.99],
    fanfare2: [329.63, 392.00, 523.25, 659.25, 587.33, 659.25, 783.99, 1046.50],
    welcome: [261.63, 329.63, 392.00, 523.25, 392.00, 329.63, 293.66, 349.23, 440.00, 587.33]
  };

  if (track.synthType === 'applause') {
    const playClapBurst = () => {
      const bufferSize = ctx.sampleRate * 0.12;
      const buffer = ctx.createBuffer(1, bufferSize, ctx.sampleRate);
      const data = buffer.getChannelData(0);
      for (let i = 0; i < bufferSize; i++) {
        data[i] = (Math.random() * 2 - 1) * Math.exp(-i / (ctx.sampleRate * 0.04));
      }
      const noise = ctx.createBufferSource();
      noise.buffer = buffer;
      const filter = ctx.createBiquadFilter();
      filter.type = 'bandpass';
      filter.frequency.value = 1100 + Math.random() * 600;
      noise.connect(filter);
      filter.connect(masterGain);
      noise.start();
    };
    noteTimer = setInterval(playClapBurst, 95);
  } else {
    const notes = melodies[track.synthType] || melodies.fanfare1;
    const playNextNote = () => {
      const freq = notes[stepIdx % notes.length];
      stepIdx++;
      const osc = ctx.createOscillator();
      const noteGain = ctx.createGain();
      osc.type = 'triangle';
      osc.frequency.setValueAtTime(freq, ctx.currentTime);
      noteGain.gain.setValueAtTime(0.35, ctx.currentTime);
      noteGain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.42);
      osc.connect(noteGain);
      noteGain.connect(masterGain);
      osc.start();
      osc.stop(ctx.currentTime + 0.45);
    };
    playNextNote();
    noteTimer = setInterval(playNextNote, 460);
  }

  return {
    masterGain,
    stop: () => {
      if (noteTimer) clearInterval(noteTimer);
      try {
        masterGain.gain.linearRampToValueAtTime(0.0001, ctx.currentTime + 0.25);
      } catch (e) {}
    }
  };
}

function toggleTrack(trackId) {
  const track = soundTracks.find(t => t.id === trackId);
  if (!track) return;

  let player = activePlayers[trackId];
  if (player && player.playing) {
    pauseTrack(trackId);
    return;
  }

  if (!player) {
    player = {
      playing: false,
      currentTime: 0,
      audioEl: track.audioDataUrl ? new Audio(track.audioDataUrl) : null,
      synth: null,
      timer: null
    };
    activePlayers[trackId] = player;
  }

  player.playing = true;
  if (player.audioEl) {
    player.audioEl.volume = isMcDucking ? track.volume * 0.15 : track.volume;
    player.audioEl.loop = track.loop;
    player.audioEl.currentTime = player.currentTime;
    player.audioEl.play().catch(() => {});
  } else {
    player.synth = startSynthEngine(track);
  }

  player.timer = setInterval(() => {
    if (player.audioEl) {
      player.currentTime = player.audioEl.currentTime;
      if (player.audioEl.ended && !track.loop) {
        pauseTrack(trackId, true);
        return;
      }
    } else {
      player.currentTime += 0.2;
      if (player.currentTime >= track.duration) {
        if (track.loop) {
          player.currentTime = 0;
        } else {
          pauseTrack(trackId, true);
          return;
        }
      }
    }
    const timeEl = document.getElementById(`time_${trackId}`);
    const seekEl = document.getElementById(`seek_${trackId}`);
    if (timeEl) timeEl.textContent = `${formatDuration(player.currentTime)} · ${formatDuration(track.duration)}`;
    if (seekEl) seekEl.value = player.currentTime;
  }, 200);

  renderSoundboard();
}

function pauseTrack(trackId, resetToZero = false) {
  const player = activePlayers[trackId];
  if (!player) return;
  player.playing = false;
  if (player.timer) clearInterval(player.timer);
  if (player.audioEl) {
    player.audioEl.pause();
    if (resetToZero) player.audioEl.currentTime = 0;
  }
  if (player.synth) {
    player.synth.stop();
    player.synth = null;
  }
  if (resetToZero) player.currentTime = 0;
  renderSoundboard();
}

function rewindTrack(trackId) {
  const player = activePlayers[trackId];
  if (player) {
    player.currentTime = 0;
    if (player.audioEl) player.audioEl.currentTime = 0;
  }
  renderSoundboard();
}

function seekTrack(trackId, val) {
  const sec = parseFloat(val) || 0;
  if (!activePlayers[trackId]) {
    activePlayers[trackId] = { playing: false, currentTime: sec, audioEl: null, synth: null, timer: null };
  } else {
    activePlayers[trackId].currentTime = sec;
    if (activePlayers[trackId].audioEl) {
      activePlayers[trackId].audioEl.currentTime = sec;
    }
  }
  const track = soundTracks.find(t => t.id === trackId);
  const timeEl = document.getElementById(`time_${trackId}`);
  if (timeEl && track) {
    timeEl.textContent = `${formatDuration(sec)} · ${formatDuration(track.duration)}`;
  }
}

function setTrackVolume(trackId, val) {
  const v = parseFloat(val);
  const track = soundTracks.find(t => t.id === trackId);
  if (track) track.volume = v;
  const player = activePlayers[trackId];
  if (player) {
    const effective = isMcDucking ? v * 0.15 : v;
    if (player.audioEl) player.audioEl.volume = effective;
    if (player.synth && player.synth.masterGain && audioCtx) {
      player.synth.masterGain.gain.setValueAtTime(effective, audioCtx.currentTime);
    }
  }
}

function setTrackLoop(trackId, checked) {
  const track = soundTracks.find(t => t.id === trackId);
  if (track) track.loop = checked;
  if (activePlayers[trackId]?.audioEl) {
    activePlayers[trackId].audioEl.loop = checked;
  }
}

function stopAllTracks() {
  Object.keys(activePlayers).forEach(id => pauseTrack(id, true));
  if (window.showToast) window.showToast('Đã dừng toàn bộ âm thanh sự kiện!');
}

function toggleMcDucking() {
  isMcDucking = !isMcDucking;
  const btn = document.getElementById('btnMcDucking');
  if (btn) {
    btn.textContent = isMcDucking ? '🎙 Đang Hạ Âm Nền (15% Cho MC)' : '🎙 Hạ Âm Nền Khi MC Phát Biểu';
  }
  soundTracks.forEach(t => setTrackVolume(t.id, t.volume));
  if (window.showToast) {
    window.showToast(isMcDucking ? 'Đã hạ âm lượng nhạc nền xuống 15% để đại biểu/MC phát biểu!' : 'Đã khôi phục âm lượng chuẩn 100%!');
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

document.addEventListener('DOMContentLoaded', async () => {
  if (!document.getElementById('soundboardGrid')) return;
  const saved = await loadSavedCustomTracks();
  if (saved.length > 0) {
    soundTracks = [...DEFAULT_SOUND_TRACKS, ...saved];
  }
  renderSoundboard();

  const fileInput = document.getElementById('customAudioInput');
  if (fileInput) {
    fileInput.addEventListener('change', async (e) => {
      const files = Array.from(e.target.files || []);
      for (const file of files) {
        const reader = new FileReader();
        const dataUrl = await new Promise((res) => {
          reader.onload = () => res(reader.result);
          reader.readAsDataURL(file);
        });
        const tempAudio = new Audio(dataUrl);
        const dur = await new Promise((res) => {
          tempAudio.onloadedmetadata = () => res(tempAudio.duration || 60);
          tempAudio.onerror = () => res(60);
        });
        const newTrack = {
          id: 'custom_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
          key: String(soundTracks.length + 1),
          title: file.name.replace(/\.[^.]+$/, ''),
          subtitle: 'Nhạc sự kiện tùy chọn · Lưu Offline IndexedDB',
          color: '#facc15',
          duration: Math.round(dur),
          loop: false,
          volume: 0.85,
          audioDataUrl: dataUrl
        };
        soundTracks.push(newTrack);
        await saveCustomTrackToDb(newTrack);
      }
      fileInput.value = '';
      renderSoundboard();
      if (window.showToast) window.showToast(`Đã thêm ${files.length} bản nhạc sự kiện vào bộ nhớ Offline!`);
    });
  }

  window.addEventListener('keydown', (e) => {
    if (['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName)) return;
    if (e.code === 'Space') {
      e.preventDefault();
      stopAllTracks();
      return;
    }
    const idx = parseInt(e.key, 10);
    if (idx >= 1 && idx <= soundTracks.length) {
      e.preventDefault();
      toggleTrack(soundTracks[idx - 1].id);
    }
  });
});

window.toggleTrack = toggleTrack;
window.rewindTrack = rewindTrack;
window.seekTrack = seekTrack;
window.setTrackVolume = setTrackVolume;
window.setTrackLoop = setTrackLoop;
window.stopAllTracks = stopAllTracks;
window.toggleMcDucking = toggleMcDucking;
window.handlePlayButtonMouseDown = handlePlayButtonMouseDown;
