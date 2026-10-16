<%
' =====================================================================
'  PoliMusic - Pie de pagina: reproductor, modal y scripts
' =====================================================================
%>        </main>
    </div>
</div>

<!-- ============================ REPRODUCTOR ========================= -->
<footer class="player" id="pmPlayer" role="contentinfo">
    <div class="player-now">
        <img class="player-cover" id="pCover" alt="Caratula"
             src="<%=H(PlaceholderUrl("album", "player", "PM"))%>">
        <div class="player-meta">
            <p class="player-title" id="pTitle">Sin reproduccion</p>
            <p class="player-artist" id="pArtist">Seleccione una cancion de la lista</p>
        </div>
    </div>

    <div class="player-center">
        <div class="player-controls">
            <button type="button" class="pc-btn" id="pShuffle" title="Aleatorio" aria-label="Aleatorio">
                <%=Icon("shuffle")%>
            </button>
            <button type="button" class="pc-btn" id="pPrev" title="Anterior" aria-label="Anterior">
                <%=Icon("skip-back")%>
            </button>
            <button type="button" class="pc-play" id="pPlay" title="Reproducir / Pausar" aria-label="Reproducir">
                <svg class="i-play" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                    <path d="M7 4.5 19.5 12 7 19.5z"></path>
                </svg>
                <svg class="i-pause" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                    <rect x="6.5" y="4.5" width="4" height="15" rx="1.2"></rect>
                    <rect x="13.5" y="4.5" width="4" height="15" rx="1.2"></rect>
                </svg>
            </button>
            <button type="button" class="pc-btn" id="pNext" title="Siguiente" aria-label="Siguiente">
                <%=Icon("skip-forward")%>
            </button>
            <button type="button" class="pc-btn" id="pRepeat" title="Repetir" aria-label="Repetir">
                <%=Icon("repeat")%>
            </button>
        </div>
        <div class="player-progress">
            <span class="p-time" id="pCur">0:00</span>
            <div class="p-bar" id="pBar" role="slider" aria-label="Progreso">
                <div class="p-fill" id="pFill"></div>
            </div>
            <span class="p-time" id="pDur">0:00</span>
        </div>
    </div>

    <div class="player-right">
        <button type="button" class="pc-btn" id="pMute" title="Silenciar" aria-label="Silenciar">
            <%=Icon("volume")%>
        </button>
        <div class="vol-bar" id="pVol" role="slider" aria-label="Volumen">
            <div class="vol-fill" id="pVolFill"></div>
        </div>
    </div>
</footer>

<audio id="pmAudio" preload="metadata"></audio>

<!-- ============================== MODAL ============================= -->
<div class="modal" id="pmModal" role="dialog" aria-modal="true">
    <div class="modal-card">
        <div class="modal-head">
            <h3>Titulo</h3>
            <button type="button" class="btn-icon" data-modal-close aria-label="Cerrar">
                <%=Icon("x")%>
            </button>
        </div>
        <div class="modal-body"></div>
        <div class="modal-foot"></div>
    </div>
</div>

<script type="text/javascript">
    window.PM_BASE = "<%=H(AppBase())%>";
</script>
<script src="<%=U("assets/js/player.js")%>?v=<%=H(APP_VERSION)%>"></script>
<script src="<%=U("assets/js/app.js")%>?v=<%=H(APP_VERSION)%>"></script>
</body>
</html>
