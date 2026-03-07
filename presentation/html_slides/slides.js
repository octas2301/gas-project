/**
 * スライドショー用：キーボード左右で前後移動（Zenn / slide_maker 方式を参考）
 * フルスクリーンは F11 またはブラウザのメニューから
 * ※ GAS では .gs として読み込まれるため、サーバー側では document が未定義。ブラウザ時のみ実行する。
 */
(function () {
  if (typeof document === 'undefined' || typeof window === 'undefined') return;
  var slides = document.querySelectorAll('.slide');
  var current = 0;

  function show(i) {
    if (i < 0) i = 0;
    if (i >= slides.length) i = slides.length - 1;
    current = i;
    slides.forEach(function (s, j) {
      s.style.display = j === current ? 'flex' : 'none';
    });
    if (window.location.hash !== '#full') {
      window.location.hash = 'slide-' + (current + 1);
    }
  }

  document.addEventListener('keydown', function (e) {
    if (e.key === 'ArrowRight' || e.key === ' ') {
      e.preventDefault();
      show(current + 1);
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      show(current - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      show(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      show(slides.length - 1);
    }
  });

  // スライドショーモード時のみ 1 枚表示。通常は全スライド表示（印刷・PDF用）
  var hash = (window.location.hash || '').toLowerCase();
  if (hash === '#present' || hash === '#full') {
    show(0);
  }
})();
