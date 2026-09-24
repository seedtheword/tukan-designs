/* ============================================================
   TuKan Designs — Newsletter + Consultation request
   - Consultation: a request form. POSTs { action:'consultation', ... }
     to the order handler (cfg.orderHandlerUrl), which emails Brandon.
     The visitor sees a "we'll get back to you" confirmation. Falls back
     to a mailto draft if no backend is configured.
   - Newsletter: POSTs { action:'newsletter', email } to the same
     handler. Falls back to a mailto if no backend is configured.
   ============================================================ */
(function () {
  'use strict';

  var CONFIG = null;
  async function loadConfig() {
    if (CONFIG) return CONFIG;
    try {
      CONFIG = await fetch('assets/data/site-config.json?t=' + Date.now(), { cache: 'no-store' })
        .then(function (r) { return r.json(); });
    } catch (_) { CONFIG = {}; }
    return CONFIG;
  }

  // POST to the Google Apps Script Web App (cfg.orderHandlerUrl), same pattern
  // as the Seed the Word site: JSON as text/plain (no CORS preflight); the
  // script returns { ok: true } and emails Brandon. Returns true on success,
  // false if not configured, throws on error.
  async function sendToBackend(cfg, payload) {
    if (cfg && cfg.orderHandlerUrl) {
      var res = await fetch(cfg.orderHandlerUrl, {
        method: 'POST',
        mode: 'cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      }).then(function (r) { return r.json(); });
      if (res && res.ok) return true;
      throw new Error((res && res.error) || 'failed');
    }
    return false; // not configured -> caller uses mailto fallback
  }

  // ── Consultation request form ────────────────────────────────
  function initConsult() {
    var form = document.getElementById('consult-form');
    if (!form) return;
    var btn = document.getElementById('cs-submit');
    var statusEl = document.getElementById('cs-status');

    function setStatus(msg, kind) {
      statusEl.textContent = msg || '';
      statusEl.className = 'consult-form__status' + (kind ? ' consult-form__status--' + kind : '');
    }
    function val(id) { var el = document.getElementById(id); return el ? el.value.trim() : ''; }

    form.addEventListener('submit', async function (e) {
      e.preventDefault();
      var name = val('cs-name');
      var email = val('cs-email');
      if (!name || !email || email.indexOf('@') === -1) {
        setStatus('Please enter your name and a valid email.', 'err');
        return;
      }

      var payload = {
        action: 'consultation',
        _subject: 'Consultation Request — ' + name,
        name: name,
        email: email,
        phone: val('cs-phone'),
        method: val('cs-method'),
        availability: val('cs-availability'),
        notes: val('cs-notes'),
        submittedAt: new Date().toISOString()
      };

      btn.disabled = true; btn.textContent = 'Sending…';
      setStatus('');

      var cfg = await loadConfig();
      try {
        var sent = await sendToBackend(cfg, payload);
        if (sent) {
          setStatus('Thanks, ' + name.split(' ')[0] + '! Your request is in — we\'ll get back to you by email to set up a time.', 'ok');
          form.reset();
          btn.disabled = false; btn.textContent = 'Request My Consultation →';
          return;
        }
        // No backend — open an email draft to Brandon so nothing is lost.
        var body = encodeURIComponent(
          'Consultation request from the website:\n\n' +
          'Name: ' + name + '\nEmail: ' + email + '\nPhone: ' + payload.phone +
          '\nPreferred contact: ' + payload.method +
          '\nAvailability: ' + payload.availability +
          '\n\n' + (payload.notes || '(no notes)')
        );
        window.location.href = 'mailto:hello@tukandesigns.com?subject=' +
          encodeURIComponent('Consultation Request — ' + name) + '&body=' + body;
        setStatus('Opening your email app to send your request…', 'ok');
      } catch (err) {
        setStatus('Something went wrong. Please try again, or call/email us directly.', 'err');
      } finally {
        btn.disabled = false; btn.textContent = 'Request My Consultation →';
      }
    });
  }

  // ── Newsletter ───────────────────────────────────────────────
  function initNewsletter() {
    var form = document.getElementById('newsletter-form');
    if (!form) return;
    var emailInput = document.getElementById('nl-email');
    var btn = document.getElementById('nl-submit');
    var statusEl = document.getElementById('nl-status');

    function setStatus(msg, kind) {
      statusEl.textContent = msg || '';
      statusEl.className = 'newsletter__status' + (kind ? ' newsletter__status--' + kind : '');
    }

    form.addEventListener('submit', async function (e) {
      e.preventDefault();
      var email = emailInput.value.trim();
      if (!email || email.indexOf('@') === -1) { setStatus('Please enter a valid email.', 'err'); return; }

      btn.disabled = true; btn.textContent = '…';
      setStatus('');

      var cfg = await loadConfig();
      try {
        var sent = await sendToBackend(cfg, { action: 'newsletter', _subject: 'Newsletter signup — TuKan Designs', email: email, source: 'tukan-website' });
        if (sent) {
          setStatus('You\'re in! Watch your inbox for new pieces and sales.', 'ok');
          form.reset();
          btn.disabled = false; btn.textContent = 'Subscribe';
          return;
        }
        // No backend — open an email to subscribe manually.
        window.location.href = 'mailto:hello@tukandesigns.com?subject=' +
          encodeURIComponent('Newsletter signup') + '&body=' +
          encodeURIComponent('Please add me to the TuKan Designs list: ' + email);
        setStatus('Opening your email app to confirm…', 'ok');
      } catch (err) {
        setStatus('Something went wrong. Please try again or email us.', 'err');
      } finally {
        btn.disabled = false; btn.textContent = 'Subscribe';
      }
    });
  }

  function init() { initConsult(); initNewsletter(); }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else { init(); }
})();
