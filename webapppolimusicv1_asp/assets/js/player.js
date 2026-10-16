/* =====================================================================
   PoliMusic - Reproductor de audio
   Cola de reproduccion, controles, barra de progreso, volumen,
   aleatorio, repeticion y contador de reproducciones.
   ===================================================================== */
(function () {
    'use strict';

    var PM = window.PoliMusic || (window.PoliMusic = {});
    var BASE = window.PM_BASE || '/';

    var audio = document.getElementById('pmAudio');
    if (!audio) { return; }

    /* ------------------------- Referencias DOM ------------------------- */
    var el = {
        cover:   document.getElementById('pCover'),
        title:   document.getElementById('pTitle'),
        artist:  document.getElementById('pArtist'),
        play:    document.getElementById('pPlay'),
        prev:    document.getElementById('pPrev'),
        next:    document.getElementById('pNext'),
        shuffle: document.getElementById('pShuffle'),
        repeat:  document.getElementById('pRepeat'),
        bar:     document.getElementById('pBar'),
        fill:    document.getElementById('pFill'),
        cur:     document.getElementById('pCur'),
        dur:     document.getElementById('pDur'),
        vol:     document.getElementById('pVol'),
        volFill: document.getElementById('pVolFill'),
        mute:    document.getElementById('pMute')
    };

    /* ---------------------------- Estado ------------------------------- */
    var queue = [];
    var index = -1;
    var shuffle = false;
    var repeat = 'off';          // off | all | one
    var lastPlayId = null;
    var seeking = false;

    /* --------------------------- Utilidades ---------------------------- */
    function ls(key, def) {
        try {
            var v = window.localStorage.getItem('pm_' + key);
            return v === null ? def : v;
        } catch (e) { return def; }
    }
    function lsSet(key, value) {
        try { window.localStorage.setItem('pm_' + key, value); } catch (e) { /* ignore */ }
    }
    function fmt(sec) {
        if (!isFinite(sec) || sec < 0) { sec = 0; }
        var m = Math.floor(sec / 60);
        var s = Math.floor(sec % 60);
        return m + ':' + (s < 10 ? '0' : '') + s;
    }
    function setIcon(node, path) {
        if (!node) { return; }
        var svg = node.querySelector('svg');
        if (svg) { svg.innerHTML = path; }
    }
    function toast(msg, isErr) {
        if (PM.toast) { PM.toast(msg, isErr); return; }
        window.alert(msg);
    }

    var ICONS = {
        play:  '<polygon points="6 4 20 12 6 20 6 4"></polygon>',
        pause: '<rect x="6" y="4" width="4" height="16" rx="1"></rect><rect x="14" y="4" width="4" height="16" rx="1"></rect>'
    };

    /* ------------------------- Contador de plays ----------------------- */
    function registerPlay(track) {
        if (!track || !track.id || lastPlayId === track.id) { return; }
        lastPlayId = track.id;
        try {
            var body = 'id=' + encodeURIComponent(track.id);
            if (window.fetch) {
                fetch(BASE + 'api_play.asp', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: body,
                    credentials: 'same-origin'
                }).then(function (r) { return r.json(); }).then(function (data) {
                    if (data && data.ok) {
                        var cells = document.querySelectorAll('[data-plays-for="' + track.id + '"]');
                        for (var i = 0; i < cells.length; i++) {
                            cells[i].textContent = data.plays;
                        }
                    }
                }).catch(function () { /* silencioso */ });
            }
        } catch (e) { /* silencioso */ }
    }

    /* --------------------------- Reproduccion -------------------------- */
    function load(i, autoplay) {
        if (i < 0 || i >= queue.length) { return; }
        index = i;
        var t = queue[i];

        if (!t.src) {
            toast('El archivo de audio de "' + t.title + '" no esta disponible en el servidor.', true);
            return;
        }

        audio.src = t.src;
        el.title.textContent = t.title;
        el.artist.textContent = t.artist || 'Artista desconocido';
        if (el.cover) {
            el.cover.src = t.cover || '';
            el.cover.alt = t.title;
        }
        document.title = t.title + ' - PoliMusic';

        if (autoplay !== false) {
            var p = audio.play();
            if (p && p.catch) {
                p.catch(function () {
                    toast('El navegador bloqueo la reproduccion automatica. Pulse play.', true);
                });
            }
        }
        highlight();
        updateMediaSession(t);
    }

    function playQueue(list, startAt) {
        if (!list || !list.length) {
            toast('No hay canciones para reproducir.', true);
            return;
        }
        queue = list.slice(0);
        load(startAt || 0, true);
    }

    function toggle() {
        if (index < 0) {
            if (queue.length) { load(0, true); }
            return;
        }
        if (audio.paused) {
            var p = audio.play();
            if (p && p.catch) { p.catch(function () { }); }
        } else {
            audio.pause();
        }
    }

    function nextIndex(step) {
        if (!queue.length) { return -1; }
        if (shuffle && queue.length > 1) {
            var r = index;
            while (r === index) { r = Math.floor(Math.random() * queue.length); }
            return r;
        }
        var n = index + step;
        if (n >= queue.length) { n = 0; }
        if (n < 0) { n = queue.length - 1; }
        return n;
    }

    function next(auto) {
        if (!queue.length) { return; }
        if (auto && repeat === 'one') {
            audio.currentTime = 0;
            audio.play();
            return;
        }
        if (auto && repeat === 'off' && !shuffle && index === queue.length - 1) {
            audio.pause();
            audio.currentTime = 0;
            return;
        }
        load(nextIndex(1), true);
    }

    function prev() {
        if (!queue.length) { return; }
        if (audio.currentTime > 3) {
            audio.currentTime = 0;
            return;
        }
        load(nextIndex(-1), true);
    }

    /* ----------------------------- Resaltado --------------------------- */
    function highlight() {
        var rows = document.querySelectorAll('.track');
        var current = (index >= 0 && queue[index]) ? String(queue[index].id) : '';
        for (var i = 0; i < rows.length; i++) {
            var id = rows[i].getAttribute('data-id');
            if (id && id === current) {
                rows[i].classList.add('playing');
            } else {
                rows[i].classList.remove('playing');
            }
        }
    }

    /* --------------------------- Media Session ------------------------- */
    function updateMediaSession(t) {
        if (!('mediaSession' in navigator)) { return; }
        try {
            navigator.mediaSession.metadata = new window.MediaMetadata({
                title: t.title,
                artist: t.artist || 'PoliMusic',
                album: t.album || 'PoliMusic',
                artwork: t.cover ? [{ src: t.cover, sizes: '512x512', type: 'image/png' }] : []
            });
            navigator.mediaSession.setActionHandler('play', function () { audio.play(); });
            navigator.mediaSession.setActionHandler('pause', function () { audio.pause(); });
            navigator.mediaSession.setActionHandler('previoustrack', prev);
            navigator.mediaSession.setActionHandler('nexttrack', function () { next(false); });
        } catch (e) { /* navegadores sin soporte completo */ }
    }

    /* ------------------------------ Eventos ---------------------------- */
    audio.addEventListener('play', function () {
        if (el.play) { el.play.classList.add('playing'); }
        setIcon(el.play, ICONS.pause);
        var t = queue[index];
        if (t) { registerPlay(t); }
    });
    audio.addEventListener('pause', function () {
        if (el.play) { el.play.classList.remove('playing'); }
        setIcon(el.play, ICONS.play);
    });
    audio.addEventListener('ended', function () { next(true); });
    audio.addEventListener('error', function () {
        if (audio.src) {
            var t = queue[index];
            toast('No se pudo cargar el audio' + (t ? ' de "' + t.title + '"' : '') + '.', true);
        }
    });
    audio.addEventListener('timeupdate', function () {
        if (seeking) { return; }
        var d = audio.duration;
        if (!isFinite(d) || d <= 0) {
            el.cur.textContent = '0:00';
            el.fill.style.width = '0%';
            return;
        }
        var pct = (audio.currentTime / d) * 100;
        el.fill.style.width = pct + '%';
        el.cur.textContent = fmt(audio.currentTime);
        el.dur.textContent = fmt(d);
    });
    audio.addEventListener('loadedmetadata', function () {
        el.dur.textContent = fmt(audio.duration);
    });
    audio.addEventListener('volumechange', function () {
        var v = audio.muted ? 0 : audio.volume;
        el.volFill.style.width = (v * 100) + '%';
        if (el.mute) { el.mute.classList.toggle('on', audio.muted || audio.volume === 0); }
    });

    if (el.play) { el.play.addEventListener('click', toggle); }
    if (el.next) { el.next.addEventListener('click', function () { next(false); }); }
    if (el.prev) { el.prev.addEventListener('click', prev); }

    if (el.shuffle) {
        el.shuffle.addEventListener('click', function () {
            shuffle = !shuffle;
            el.shuffle.classList.toggle('on', shuffle);
            lsSet('shuffle', shuffle ? '1' : '0');
        });
    }
    if (el.repeat) {
        el.repeat.addEventListener('click', function () {
            repeat = repeat === 'off' ? 'all' : (repeat === 'all' ? 'one' : 'off');
            el.repeat.classList.toggle('on', repeat !== 'off');
            el.repeat.setAttribute('data-mode', repeat);
            el.repeat.title = repeat === 'off' ? 'Repetir: desactivado'
                : (repeat === 'all' ? 'Repetir: todas' : 'Repetir: una');
            lsSet('repeat', repeat);
        });
    }

    function seekFromEvent(ev) {
        var rect = el.bar.getBoundingClientRect();
        var x = (ev.touches ? ev.touches[0].clientX : ev.clientX) - rect.left;
        var ratio = Math.min(Math.max(x / rect.width, 0), 1);
        if (isFinite(audio.duration) && audio.duration > 0) {
            audio.currentTime = ratio * audio.duration;
            el.fill.style.width = (ratio * 100) + '%';
            el.cur.textContent = fmt(audio.currentTime);
        }
    }
    if (el.bar) {
        el.bar.addEventListener('mousedown', function (ev) {
            seeking = true;
            seekFromEvent(ev);
            document.addEventListener('mousemove', onDrag);
            document.addEventListener('mouseup', stopDrag);
        });
        el.bar.addEventListener('touchstart', function (ev) {
            seeking = true;
            seekFromEvent(ev);
            document.addEventListener('touchmove', onDrag);
            document.addEventListener('touchend', stopDrag);
        });
    }
    function onDrag(ev) { seekFromEvent(ev); }
    function stopDrag() {
        seeking = false;
        document.removeEventListener('mousemove', onDrag);
        document.removeEventListener('mouseup', stopDrag);
        document.removeEventListener('touchmove', onDrag);
        document.removeEventListener('touchend', stopDrag);
    }

    function volumeFromEvent(ev) {
        var rect = el.vol.getBoundingClientRect();
        var x = (ev.touches ? ev.touches[0].clientX : ev.clientX) - rect.left;
        var v = Math.min(Math.max(x / rect.width, 0), 1);
        audio.volume = v;
        audio.muted = false;
        lsSet('volume', v);
    }
    if (el.vol) {
        el.vol.addEventListener('mousedown', function (ev) {
            volumeFromEvent(ev);
            document.addEventListener('mousemove', onVolDrag);
            document.addEventListener('mouseup', stopVolDrag);
        });
        el.vol.addEventListener('touchstart', function (ev) {
            volumeFromEvent(ev);
        });
    }
    function onVolDrag(ev) { volumeFromEvent(ev); }
    function stopVolDrag() { document.removeEventListener('mousemove', onVolDrag); }

    if (el.mute) {
        el.mute.addEventListener('click', function () {
            audio.muted = !audio.muted;
        });
    }

    /* --------------------- Atajos de teclado --------------------------- */
    document.addEventListener('keydown', function (ev) {
        var tag = (ev.target && ev.target.tagName) || '';
        if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') { return; }
        if (ev.code === 'Space') { ev.preventDefault(); toggle(); }
        else if (ev.code === 'ArrowRight' && ev.shiftKey) { next(false); }
        else if (ev.code === 'ArrowLeft' && ev.shiftKey) { prev(); }
    });

    /* --------------------- Delegacion de clics ------------------------- */
    document.addEventListener('click', function (ev) {
        if (!ev.target.closest) { return; }

        var explicit = ev.target.closest('[data-play-id]');
        var row = ev.target.closest('.track');
        if (!explicit && !row) { return; }

        // Un clic sobre otro control (anadir a playlist, enlace, etc.)
        // no debe iniciar la reproduccion.
        if (!explicit && ev.target.closest('button, a, input, select, textarea, label')) { return; }

        var targetId = explicit ? explicit.getAttribute('data-play-id')
                                : row.getAttribute('data-id');
        var scope = explicit || row;
        var container = scope.closest('[data-tracklist]');
        var rows = container ? container.querySelectorAll('.track') : [row];

        var list = [];
        var start = 0;
        for (var i = 0; i < rows.length; i++) {
            var raw = rows[i].getAttribute('data-track');
            if (!raw) { continue; }
            try {
                var t = JSON.parse(raw);
                if (String(t.id) === String(targetId)) { start = list.length; }
                list.push(t);
            } catch (e) { /* atributo invalido */ }
        }
        if (!list.length) { return; }
        ev.preventDefault();
        playQueue(list, start);
    });

    /* --------------------------- Inicializacion ------------------------ */
    (function init() {
        var v = parseFloat(ls('volume', '0.8'));
        if (isNaN(v)) { v = 0.8; }
        audio.volume = Math.min(Math.max(v, 0), 1);

        shuffle = ls('shuffle', '0') === '1';
        if (el.shuffle) { el.shuffle.classList.toggle('on', shuffle); }

        repeat = ls('repeat', 'off');
        if (el.repeat) {
            el.repeat.classList.toggle('on', repeat !== 'off');
            el.repeat.setAttribute('data-mode', repeat);
        }
        setIcon(el.play, ICONS.play);
        if (el.volFill) { el.volFill.style.width = (audio.volume * 100) + '%'; }
    })();

    /* ------------------------------ API -------------------------------- */
    PM.player = {
        play: playQueue,
        toggle: toggle,
        next: function () { next(false); },
        prev: prev,
        current: function () { return index >= 0 ? queue[index] : null; },
        queue: function () { return queue.slice(0); },
        isShuffle: function () { return shuffle; },
        repeatMode: function () { return repeat; }
    };
})();
