/* 点薪官网交互 */
(function () {
    'use strict';

    document.documentElement.classList.add('site-js');

    /* 顶栏滚动阴影 */
    var topbar = document.getElementById('topbar');
    window.addEventListener('scroll', function () {
        topbar.classList.toggle('scrolled', window.scrollY > 8);
    }, { passive: true });

    /* 手机端折叠菜单 */
    var burger = document.getElementById('navBurger');
    var drop = document.getElementById('navDrop');
    if (burger && drop) {
        burger.addEventListener('click', function () {
            drop.classList.toggle('open');
        });
        drop.addEventListener('click', function (e) {
            if (e.target.tagName === 'A') drop.classList.remove('open');
        });
    }

    /* 首屏手机框截图自动轮播（宽屏启用） */
    var heroImg = document.getElementById('heroImg');
    var heroDots = document.getElementById('heroDots');
    var SHOTS = [
        { src: 'img/calendar.png', alt: '工时日历' },
        { src: 'img/motivate.png', alt: '上班鼓励器' },
        { src: 'img/summary.png', alt: '月度总结' },
        { src: 'img/sheet.png', alt: '记工时' },
        { src: 'img/settings.png', alt: '工资设置' }
    ];
    if (heroImg && heroDots) {
        var idx = 0;
        var timer = 0;
        heroDots.innerHTML = SHOTS.map(function (s, i) {
            return '<i data-i="' + i + '"' + (i === 0 ? ' class="on"' : '') + ' title="' + s.alt + '"></i>';
        }).join('');

        function show(i) {
            idx = (i + SHOTS.length) % SHOTS.length;
            var dotList = heroDots.children;
            for (var k = 0; k < dotList.length; k++) {
                dotList[k].className = k === idx ? 'on' : '';
            }
            heroImg.style.opacity = '0';
            setTimeout(function () {
                heroImg.src = SHOTS[idx].src;
                heroImg.alt = SHOTS[idx].alt;
                heroImg.style.opacity = '1';
            }, 180);
        }
        function restart() {
            if (timer) clearInterval(timer);
            timer = setInterval(function () { show(idx + 1); }, 3600);
        }
        heroDots.addEventListener('click', function (e) {
            var i = e.target.getAttribute('data-i');
            if (i !== null) { show(Number(i)); restart(); }
        });
        restart();
    }

    /* 滚动显现 */
    if ('IntersectionObserver' in window) {
        var targets = document.querySelectorAll('.feat, .detail, .dl-card, .changelog, .faq');
        var io = new IntersectionObserver(function (entries) {
            entries.forEach(function (en) {
                if (en.isIntersecting) {
                    en.target.classList.add('in');
                    io.unobserve(en.target);
                }
            });
        }, { threshold: 0.12 });
        targets.forEach(function (el) {
            el.classList.add('reveal');
            io.observe(el);
        });
        /* 兜底：任何原因未触发时，3 秒后全部直接显示（无过渡） */
        setTimeout(function () {
            targets.forEach(function (el) {
                el.classList.add('now');
            });
            io.disconnect();
        }, 3000);
    }

    /* FAQ 单开互斥 */
    var faqs = document.querySelectorAll('details.faq');
    faqs.forEach(function (d) {
        d.addEventListener('toggle', function () {
            if (d.open) {
                faqs.forEach(function (o) { if (o !== d) o.open = false; });
            }
        });
    });

    /* 反馈 iframe 高度自适应（同源） */
    var frame = document.getElementById('fbFrame');
    if (frame) {
        frame.addEventListener('load', function () {
            try {
                var doc = frame.contentDocument;
                if (doc && doc.body) {
                    frame.style.height = Math.max(560, doc.body.scrollHeight + 24) + 'px';
                }
            } catch (e) { /* 跨域时保持默认高度 */ }
        });
    }
})();
