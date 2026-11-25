/* Animation layer for the key page (GSAP). Keeps the original key-turning animation and sounds;
   adds an entrance, floating gold motes, a gentle push toward the keyhole, a gold burst and a soft handoff. */
(function () {
  if (!window.gsap) { document.documentElement.classList.add('ready'); return; }
  var reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  var st = document.getElementById('stage');
  if (!st) return;
  var css = document.createElement('style');
  css.textContent = '#stage{scale:var(--s,1);translate:var(--tx,0px) var(--ty,0px)}' +
    '.mote{position:absolute;width:5px;height:5px;margin:-2.5px;border-radius:50%;pointer-events:none;z-index:8;background:radial-gradient(circle,#fff6d8 0%,#f3cf7a 45%,rgba(243,207,122,0) 72%)}' +
    '.spark{position:absolute;width:7px;height:7px;margin:-3.5px;border-radius:50%;pointer-events:none;z-index:18;background:radial-gradient(circle,#fffdf2 0%,#ffe08a 40%,rgba(255,224,138,0) 70%)}';
  document.head.appendChild(css);
  var HX = 290, HY = 250; // keyhole, in stage coordinates

  /* 1. Entrance: the door comes out of the dark with a slow settle */
  if (!reduce) {
    gsap.fromTo('#bd1', { opacity: 0 }, { opacity: 1, duration: 1.2, ease: 'power2.out' });
    gsap.fromTo(st, { opacity: 0, '--s': 1.04 }, { opacity: 1, '--s': 1, duration: 1.2, ease: 'power2.out' });
  } else document.documentElement.classList.add('ready');

  /* 2. Gold motes drift around the keyhole while the guest decides to tap */
  var motes = [];
  if (!reduce) for (var i = 0; i < 12; i++) {
    var m = document.createElement('i'); m.className = 'mote'; st.appendChild(m); motes.push(m);
    (function drift(el) {
      var a = Math.random() * 6.283, r = 28 + Math.random() * 70;
      gsap.set(el, { left: HX + Math.cos(a) * r * 0.6, top: HY + Math.sin(a) * r, opacity: 0, scale: 0.4 + Math.random() * 0.8 });
      gsap.to(el, { x: (Math.random() - 0.5) * 30, y: -20 - Math.random() * 40, opacity: 0.9, duration: 1.4 + Math.random() * 1.2, ease: 'sine.out', delay: Math.random() * 2,
        onComplete: function () { gsap.to(el, { opacity: 0, duration: 0.9, onComplete: function () { gsap.set(el, { x: 0, y: 0 }); drift(el); } }); } });
    })(m);
  }

  /* 3. On the tap: a burst of gold as the lock opens (no zoom) */
  var tap = document.getElementById('tap');
  if (tap) tap.addEventListener('click', function () {
    if (!document.body.classList.contains('go')) return;
    motes.forEach(function (el) { gsap.killTweensOf(el); gsap.to(el, { opacity: 0, duration: 0.4 }); });
    if (reduce) return;
    gsap.delayedCall(4.05, function () {
      for (var j = 0; j < 28; j++) {
        var s = document.createElement('i'); s.className = 'spark'; st.appendChild(s);
        var a = Math.random() * 6.283, d = 60 + Math.random() * 150;
        gsap.set(s, { left: HX, top: HY, scale: 0.5 + Math.random() });
        gsap.to(s, { x: Math.cos(a) * d, y: Math.sin(a) * d, opacity: 0, duration: 0.9 + Math.random() * 0.6, ease: 'power2.out', onComplete: function () { this.targets()[0].remove(); } });
      }
    });
  });

  /* 4. Handoff: the invitation's own gold layer fades out (0.4s), so no extra fade here */
})();
