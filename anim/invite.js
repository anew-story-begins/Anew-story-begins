/* Animation layer for the invitation pages (GSAP + ScrollTrigger + Motion).
   Adds motion only: it never changes RSVP, notes, the Google Sheet connection or the music. */
(function () {
  if (!window.gsap) return;
  var reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  var M = window.Motion || null;
  if (window.ScrollTrigger) gsap.registerPlugin(ScrollTrigger);

  var css = document.createElement('style');
  css.textContent =
    /* individual transform properties compose with the existing fit() transform on .ph */
    '.ph{scale:var(--s,1);translate:0 var(--ty,0px)}' +
    '.bd{translate:0 var(--py,0px)}' +
    '.shine{position:absolute;pointer-events:none;z-index:6;-webkit-mask-repeat:no-repeat;mask-repeat:no-repeat;-webkit-mask-size:100% 100%;mask-size:100% 100%;' +
    'background:linear-gradient(105deg,rgba(255,244,200,0) 40%,rgba(255,246,214,.85) 50%,rgba(255,244,200,0) 60%);background-size:320% 100%;background-position:160% 0;mix-blend-mode:screen;opacity:.9}' +
    '#rSubmit{overflow:hidden}#rSubmit::after{content:"";position:absolute;inset:0;border-radius:inherit;background:linear-gradient(105deg,transparent 30%,rgba(255,250,225,.55) 50%,transparent 70%);background-size:300% 100%;background-position:150% 0;animation:btnShine 4.2s ease-in-out 1.5s infinite}' +
    '@keyframes btnShine{0%,60%{background-position:150% 0}100%{background-position:-50% 0}}' +
    '.cd{perspective:240px}' +
    '@media(prefers-reduced-motion:reduce){#rSubmit::after{animation:none}}';
  document.head.appendChild(css);

  var secs = [].slice.call(document.querySelectorAll('.scr'));

  /* 1. Each page eases in as it arrives, then its text and controls follow one after another */
  secs.forEach(function (sec, i) {
    var ph = sec.querySelector('.ph');
    if (!ph) return;
    var items = [].slice.call(ph.querySelectorAll('.tx, .fld, .val, .cdl, .cd, .sw, .note5, .msg4, .msgtext > *'));
    var tl = gsap.timeline({ paused: true, defaults: { ease: 'power3.out' } });
    tl.fromTo(ph, { '--s': 0.94, '--ty': '22px', opacity: 0.25 }, { '--s': 1, '--ty': '0px', opacity: 1, duration: 0.9 });
    if (items.length) tl.fromTo(items, { opacity: 0, y: 12 }, { opacity: 1, y: 0, duration: 0.55, stagger: 0.06, clearProps: 'transform' }, 0.25);
    if (reduce) { tl.progress(1); return; }
    if (i === 0) { gsap.delayedCall(0.35, function () { tl.restart(); }); }
    if (window.ScrollTrigger) ScrollTrigger.create({
      trigger: sec, start: 'top 65%', end: 'bottom 35%',
      onEnter: function () { if (i) tl.restart(); }, onEnterBack: function () { tl.restart(); }
    });
    /* soft depth: the blurred backdrop drifts slower than the page */
    var bd = sec.querySelector('.bd');
    if (bd && window.ScrollTrigger) gsap.fromTo(bd, { '--py': '-36px' }, { '--py': '36px', ease: 'none',
      scrollTrigger: { trigger: sec, start: 'top bottom', end: 'bottom top', scrub: true } });
  });

  /* 3. The "next page" arrows breathe gently so guests know to scroll */
  if (!reduce) gsap.to('.nx', { y: 6, duration: 1.1, ease: 'sine.inOut', yoyo: true, repeat: -1 });

  /* 4. Countdown digits flip when they change */
  if (M && !reduce) [].slice.call(document.querySelectorAll('.cd')).forEach(function (el) {
    var last = el.textContent;
    new MutationObserver(function () {
      if (el.textContent === last) return; last = el.textContent;
      M.animate(el, { transform: ['rotateX(-80deg)', 'rotateX(0deg)'], opacity: [0.2, 1] }, { duration: 0.4, easing: [0.2, 0.8, 0.2, 1] });
    }).observe(el, { childList: true, characterData: true, subtree: true });
  });

  /* 5. Springy feedback: reminder switches, the toast message and the calendar pop-up */
  if (M && !reduce) {
    [].slice.call(document.querySelectorAll('.sw')).forEach(function (b) {
      b.addEventListener('click', function () { M.animate(b, { transform: ['scale(1)', 'scale(1.18)', 'scale(1)'] }, { duration: 0.35 }); });
    });
    var toast = document.getElementById('toast');
    if (toast) new MutationObserver(function () {
      if (toast.classList.contains('on')) M.animate(toast, { transform: ['translateX(-50%) translateY(26px) scale(.92)', 'translateX(-50%) translateY(-4px) scale(1.02)', 'translateX(-50%) translateY(0) scale(1)'] }, { duration: 0.55, easing: [0.22, 1, 0.36, 1] });
    }).observe(toast, { attributes: true, attributeFilter: ['class'] });
    var cal = document.getElementById('cal');
    if (cal) new MutationObserver(function () {
      var bx = cal.querySelector('.bx');
      if (cal.classList.contains('open') && bx) M.animate(bx, { transform: ['scale(.86)', 'scale(1.03)', 'scale(1)'], opacity: [0, 1] }, { duration: 0.5, easing: [0.22, 1, 0.36, 1] });
    }).observe(cal, { attributes: true, attributeFilter: ['class'] });
    var menu = document.getElementById('menu');
    if (menu) new MutationObserver(function () {
      var panel = menu.firstElementChild && menu.firstElementChild.nextElementSibling || menu.lastElementChild;
      if (menu.classList.contains('open') && panel) M.animate(panel, { transform: ['translateY(18px)', 'translateY(0)'], opacity: [0, 1] }, { duration: 0.4, easing: [0.22, 1, 0.36, 1] });
    }).observe(menu, { attributes: true, attributeFilter: ['class'] });
    /* RSVP sent: a soft gold glow around the form */
    var msg = document.getElementById('rMsg');
    if (msg) new MutationObserver(function () {
      if (/Thank you/.test(msg.textContent)) M.animate('#p4 .fld', { boxShadow: ['0 0 0 0 rgba(240,213,143,0)', '0 0 0 3px rgba(240,213,143,.75)', '0 0 0 0 rgba(240,213,143,0)'] }, { duration: 1.2 });
    }).observe(msg, { childList: true, characterData: true, subtree: true });
  }
})();
