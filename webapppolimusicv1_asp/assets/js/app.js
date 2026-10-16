/* =====================================================================
   PoliMusic - Logica de interfaz
   Navegacion sin recarga (la musica no se interrumpe), busqueda,
   notificaciones, modales y utilidades de formularios.
   ===================================================================== */
(function () {
    'use strict';

    var PM = window.PoliMusic || (window.PoliMusic = {});
    var BASE = window.PM_BASE || '/';

    /* =================================================================
       Notificaciones flotantes
       ================================================================= */
    var toastHost = null;
    PM.toast = function (msg, isErr) {
        if (!toastHost) {
            toastHost = document.createElement('div');
            toastHost.className = 'toast-host';
            document.body.appendChild(toastHost);
        }
        var t = document.createElement('div');
        t.className = 'toast' + (isErr ? ' err' : '');
        t.textContent = msg;
        toastHost.appendChild(t);
        setTimeout(function () {
            t.style.transition = 'opacity .3s ease, transform .3s ease';
            t.style.opacity = '0';
            t.style.transform = 'translateX(20px)';
            setTimeout(function () { if (t.parentNode) { t.parentNode.removeChild(t); } }, 320);
        }, 3400);
    };

    /* =================================================================
       Modal generico
       ================================================================= */
    var modalEl = document.getElementById('pmModal');

    PM.modal = {
        open: function (title, bodyHtml, footHtml) {
            if (!modalEl) { return; }
            modalEl.querySelector('.modal-head h3').innerHTML = title;
            modalEl.querySelector('.modal-body').innerHTML = bodyHtml;
            var foot = modalEl.querySelector('.modal-foot');
            foot.innerHTML = footHtml || '';
            modalEl.classList.add('open');
            document.body.classList.add('overlay-open');
        },
        close: function () {
            if (!modalEl) { return; }
            modalEl.classList.remove('open');
            document.body.classList.remove('overlay-open');
        },
        isOpen: function () { return modalEl && modalEl.classList.contains('open'); }
    };

    if (modalEl) {
        modalEl.addEventListener('click', function (ev) {
            if (ev.target === modalEl || ev.target.closest('[data-modal-close]')) {
                PM.modal.close();
            }
        });
        document.addEventListener('keydown', function (ev) {
            if (ev.key === 'Escape') { PM.modal.close(); }
        });
    }

    /* =================================================================
       Navegacion por vistas (fetch + pushState)
       ================================================================= */
    var viewRoot = document.getElementById('viewRoot');

    function fragmentUrl(href) {
        var q = href.indexOf('?');
        var qs = q >= 0 ? href.substring(q + 1) : 'view=home';
        return BASE + 'api_view.asp?' + qs;
    }

    function viewOf(href) {
        var m = /[?&]view=([^&]*)/.exec(href);
        return m ? decodeURIComponent(m[1]) : 'home';
    }

    function syncNav(href) {
        var v = viewOf(href);
        var items = document.querySelectorAll('.nav-item[data-nav]');
        for (var i = 0; i < items.length; i++) {
            var n = items[i].getAttribute('data-nav');
            var on = (n === v) || (v === 'album' && n === 'albums') ||
                     (v === 'artist' && n === 'artists') ||
                     (v === 'playlist' && n === 'playlists') ||
                     (v === 'search' && n === 'home');
            items[i].classList.toggle('active', on);
        }
    }

    function afterLoad() {
        bindFilters();
        if (window.PoliMusic && window.PoliMusic.player) { /* cola intacta */ }
        var q = document.getElementById('pmSearch');
        if (q) {
            var m = /[?&]q=([^&]*)/.exec(window.location.search);
            if (m) { q.value = decodeURIComponent(m[1].replace(/\+/g, ' ')); }
        }
    }

    function navigate(href, push) {
        if (!viewRoot) { window.location.href = href; return; }
        viewRoot.style.opacity = '.4';
        fetch(fragmentUrl(href), { credentials: 'same-origin' })
            .then(function (r) {
                if (!r.ok) { throw new Error('HTTP ' + r.status); }
                return r.text();
            })
            .then(function (html) {
                viewRoot.innerHTML = html;
                viewRoot.style.opacity = '1';
                if (push) { history.pushState({ href: href }, '', href); }
                syncNav(href);
                window.scrollTo(0, 0);
                afterLoad();
            })
            .catch(function () { window.location.href = href; });
    }
    PM.navigate = navigate;

    document.addEventListener('click', function (ev) {
        var a = ev.target.closest ? ev.target.closest('a[data-nav]') : null;
        if (!a) { return; }
        if (ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.button !== 0) { return; }
        var href = a.getAttribute('href');
        if (!href) { return; }
        ev.preventDefault();
        navigate(href, true);
    });

    window.addEventListener('popstate', function (ev) {
        var href = (ev.state && ev.state.href) || (location.pathname.split('/').pop() + location.search);
        navigate(href, false);
    });

    /* =================================================================
       Buscador superior
       ================================================================= */
    var search = document.getElementById('pmSearch');
    if (search) {
        var timer = null;
        search.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                var q = search.value.trim();
                if (q.length) { navigate('home.asp?view=search&q=' + encodeURIComponent(q), true); }
            }
        });
        search.addEventListener('input', function () {
            clearTimeout(timer);
            var q = search.value.trim();
            if (q.length < 2) { return; }
            timer = setTimeout(function () {
                navigate('home.asp?view=search&q=' + encodeURIComponent(q), true);
            }, 620);
        });
    }

    /* =================================================================
       Filtro local de listas (canciones, tablas)
       ================================================================= */
    function bindFilters() {
        var inputs = document.querySelectorAll('[data-filter-target]');
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].getAttribute('data-bound') === '1') { continue; }
            inputs[i].setAttribute('data-bound', '1');
            inputs[i].addEventListener('input', function () {
                var sel = this.getAttribute('data-filter-target');
                var term = this.value.toLowerCase().trim();
                var rows = document.querySelectorAll(sel);
                var visible = 0;
                for (var j = 0; j < rows.length; j++) {
                    var hay = (rows[j].textContent || '').toLowerCase();
                    var show = term === '' || hay.indexOf(term) >= 0;
                    rows[j].style.display = show ? '' : 'none';
                    if (show) { visible++; }
                }
                var counter = document.querySelector(this.getAttribute('data-filter-count') || '');
                if (counter) { counter.textContent = visible; }
            });
        }
    }

    /* =================================================================
       Anadir cancion a una playlist
       ================================================================= */
    document.addEventListener('click', function (ev) {
        var btn = ev.target.closest ? ev.target.closest('[data-add-song]') : null;
        if (!btn) { return; }
        ev.preventDefault();
        ev.stopPropagation();
        var songId = btn.getAttribute('data-add-song');
        PM.modal.open('Anadir a una playlist', '<p class="muted small">Cargando...</p>', '');
        fetch(BASE + 'api_playlist_pick.asp?song=' + encodeURIComponent(songId),
              { credentials: 'same-origin' })
            .then(function (r) { return r.text(); })
            .then(function (html) {
                PM.modal.open('Anadir a una playlist', html,
                    '<button type="button" class="btn btn-ghost" data-modal-close>Cancelar</button>');
            })
            .catch(function () { PM.modal.close(); PM.toast('No se pudo cargar sus playlists.', true); });
    });

    document.addEventListener('click', function (ev) {
        var pick = ev.target.closest ? ev.target.closest('[data-pick-playlist]') : null;
        if (!pick) { return; }
        ev.preventDefault();
        var pid = pick.getAttribute('data-pick-playlist');
        var sid = pick.getAttribute('data-song');
        var body = 'playlist=' + encodeURIComponent(pid) + '&song=' + encodeURIComponent(sid);
        fetch(BASE + 'api_playlist_add.asp', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: body,
            credentials: 'same-origin'
        })
            .then(function (r) { return r.json(); })
            .then(function (d) {
                PM.modal.close();
                PM.toast(d && d.msg ? d.msg : 'Operacion completada.', !(d && d.ok));
            })
            .catch(function () { PM.toast('No se pudo completar la operacion.', true); });
    });

    /* =================================================================
       Confirmacion de acciones destructivas
       ================================================================= */
    document.addEventListener('click', function (ev) {
        var el = ev.target.closest ? ev.target.closest('[data-confirm]') : null;
        if (!el) { return; }
        var msg = el.getAttribute('data-confirm') || '\u00bfConfirma la operacion?';
        if (!window.confirm(msg)) {
            ev.preventDefault();
            ev.stopPropagation();
        }
    });
    document.addEventListener('submit', function (ev) {
        var form = ev.target;
        if (form.getAttribute && form.getAttribute('data-confirm')) {
            if (!window.confirm(form.getAttribute('data-confirm'))) { ev.preventDefault(); }
        }
    });

    /* =================================================================
       Menu de usuario y barra lateral movil
       ================================================================= */
    var chip = document.getElementById('pmUserChip');
    var menu = document.getElementById('pmUserMenu');
    if (chip && menu) {
        chip.addEventListener('click', function (ev) {
            ev.stopPropagation();
            menu.classList.toggle('open');
        });
        document.addEventListener('click', function () { menu.classList.remove('open'); });
    }

    var toggle = document.getElementById('pmSidebarToggle');
    var sidebar = document.getElementById('pmSidebar');
    if (toggle && sidebar) {
        toggle.addEventListener('click', function (ev) {
            ev.stopPropagation();
            sidebar.classList.toggle('open');
        });
        document.addEventListener('click', function (ev) {
            if (window.innerWidth <= 900 && sidebar.classList.contains('open') &&
                !sidebar.contains(ev.target) && ev.target !== toggle) {
                sidebar.classList.remove('open');
            }
        });
    }

    /* =================================================================
       Reproducir toda una lista
       ================================================================= */
    document.addEventListener('click', function (ev) {
        var btn = ev.target.closest ? ev.target.closest('[data-play-all]') : null;
        if (!btn) { return; }
        ev.preventDefault();
        var container = document.querySelector(btn.getAttribute('data-play-all'));
        if (!container) { return; }
        var rows = container.querySelectorAll('.track');
        var list = [];
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].style.display === 'none') { continue; }
            try { list.push(JSON.parse(rows[i].getAttribute('data-track'))); } catch (e) { /* ignorar */ }
        }
        if (!list.length) { PM.toast('No hay canciones para reproducir.', true); return; }
        if (btn.hasAttribute('data-play-shuffle')) {
            for (var k = list.length - 1; k > 0; k--) {
                var j = Math.floor(Math.random() * (k + 1));
                var tmp = list[k]; list[k] = list[j]; list[j] = tmp;
            }
        }
        if (PM.player) { PM.player.play(list, 0); }
    });

    /* =================================================================
       Inicializacion
       ================================================================= */
    afterLoad();
    if (viewRoot && viewRoot.getAttribute('data-view')) {
        syncNav('home.asp?view=' + viewRoot.getAttribute('data-view'));
    }
})();
