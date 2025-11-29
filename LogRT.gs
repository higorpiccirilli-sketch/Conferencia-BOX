/**
 * ===================================================================================
 * LogRT – Logger em tempo real (Minimalist Edition)
 * Arquivo: LogRT.gs
 * -----------------------------------------------------------------------------------
 * Cache por usuário (UserCache): chave lanc_status_<execId>, TTL 6h
 * Buffer: até 1000 linhas
 * API:
 * RT.push(execId, text, level)
 * RT.fetchLast(execId)
 * RT.clear(execId)
 * ===================================================================================
 */

var RT = (function(){
  var KEY_PREFIX   = 'lanc_status_';
  var TTL_SECONDS  = 21600; // 6h
  var MAX_LINES    = 1000;
  var TZ           = 'America/Sao_Paulo';

  function _key(execId){ return KEY_PREFIX + String(execId || 'default'); }
  function _now() { return Utilities.formatDate(new Date(), TZ, "HH:mm:ss"); } // Simplificado para apenas Hora

  function _load(execId) {
    var raw = CacheService.getUserCache().get(_key(execId));
    if (!raw) return [];
    try { return JSON.parse(raw) || []; } catch (e) { return []; }
  }

  function _save(execId, arr) {
    var buf = (arr || []).slice(-MAX_LINES);
    CacheService.getUserCache().put(_key(execId), JSON.stringify(buf), TTL_SECONDS);
  }

  function push(execId, text, level) {
    var arr = _load(execId);
    arr.push({ ts: _now(), level: String(level || 'info'), text: String(text) });
    _save(execId, arr);
    return true;
  }

  function fetchLast(execId) {
    var arr = _load(execId);
    return { line: arr.length ? arr[arr.length - 1] : null, count: arr.length };
  }

  function clear(execId) {
    CacheService.getUserCache().remove(_key(execId));
    return true;
  }

  // Wrappers simples
  return {
    iniciando: function(id, m) { return push(id, '[INICIANDO] ' + m, 'info'); },
    andamento: function(id, m) { return push(id, '... ' + m, 'info'); },
    ok:        function(id, m) { return push(id, '[OK] ' + m, 'ok'); },
    aviso:     function(id, m) { return push(id, '[AVISO] ' + m, 'warn'); },
    erro:      function(id, m) { return push(id, '[ERRO] ' + m, 'error'); },
    erroFatal: function(id, m) { return push(id, '[FATAL] ' + m, 'error'); },
    final:     function(id, m) { return push(id, '✅ ' + m, 'final'); },
    
    // Funções internas expostas
    push: push,
    fetchLast: fetchLast,
    clear: clear
  };
})();

function RT_fetchLast(execId){ return RT.fetchLast(execId); }
function RT_clear(execId){ return RT.clear(execId); }
