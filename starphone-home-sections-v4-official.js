
(() => {
  'use strict';

  /* =========================================================
     STARPHONE HOME SECTIONS V4.0-OFFICIAL
     - Inserta recomendaciones personalizadas + Pabellón de Marcas
     - Se coloca antes de #source-tabs
     - No modifica tarjetas de producto, carrito, favoritos ni WhatsApp
     ========================================================= */

  const CONFIG = {
    insertBeforeId: 'source-tabs',
    productsGridId: 'grid',
    activeFiltersRowId: 'active-filters-row',
    newKeyword: 'nuevo',
    sectionMaxWidth: '80rem',
    recommendationLimit: 12,
    recommendationBrandWeight: 100,
    recommendationCategoryWeight: 35,
    recommendationTypeWeight: 15,

    recommendations: {
      title: 'Productos nuevos para ti',
      personalizedSubtitle: 'Basado en tus compras anteriores',
      genericSubtitle: 'Descubre las novedades de nuestro catálogo',
      button: 'Ver todos'
    },

    brands: [
      {
        filter: 'XIAOMI',
        title: 'XIAOMI',
        description: 'Tecnología para una vida inteligente',
        image:
          'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT6MuRn8XB-XLybgpp-zyQsPr8iZNzn3UVy6nc1G6Vhe7D0U_Mx_oOZdEY&s=10'
      },
      {
        filter: 'DJI',
        title: 'DJI',
        description: 'Drones · Cámaras · Audio',
        image:
          'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcR-1ad3nktSDVo1978cnsU-5cP_MGq3Rh0zKu9zczpxXg&s=10'
      },
      {
        filter: 'HUAWEI',
        title: 'HUAWEI',
        description: 'Wearables · Audio · Tecnología',
        image:
          'https://consumer.huawei.com/content/dam/huawei-cbg-site/cn/mkt/pdp/wearables/watch-fit4/images/kv/huawei-watch-fit4-kv.jpg'
      },
      {
        filter: 'Tronsmart',
        title: 'Tronsmart',
        description: 'Potencia para tu música',
        image:
          'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRjckY1tYpLGpaPn4Q4Ru04nHU7PZKRdSTSYy60OZF2dpQaUCQoFetw_3Rq&s=10'
      },
      {
        filter: 'GAMESIR',
        title: 'GAMESIR',
        description: 'Accesorios para gaming',
        image:
          'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQER1yZcReVco--0LA8RYmRvnm7M0laq3vQ7THyvDAAMEMkIhUdT0Gh478&s=10'
      },
      {
        filter: 'SOLOHOUR',
        title: 'SOLOHOUR',
        description: 'Audio & Charging',
        image:
          'https://www.orrohome.cl/cdn/shop/files/COMBO_3_v2_1.png?v=1782408944'
      }
    ]
  };

  const IDS = {
    style: 'sp-home-v3-style',
    root: 'sp-home-v3',
    status: 'sp-home-v3-status'
  };

  const CSS = `
    /* 全站图片保护 */
    img {
      -webkit-user-drag: none;
      -webkit-touch-callout: none;
      user-select: none;
      -webkit-user-select: none;
    }

    #${IDS.root} {
      width: 100%;
      max-width: ${CONFIG.sectionMaxWidth};
      margin: 0 auto;
      padding: 1rem 1rem 0;
    }

    #${IDS.root},
    #${IDS.root} * {
      box-sizing: border-box;
    }

    .spv3-new {
      position: relative;
      width: 100%;
      max-width: none;
      min-height: 230px;
      overflow: hidden;
      display: flex;
      align-items: center;
      border: 0;
      border-radius: 28px;
      padding: 34px 42px;
      color: #fff;
      text-align: left;
      cursor: pointer;
      background:
        linear-gradient(90deg,
          rgba(13, 27, 54, .97) 0%,
          rgba(26, 52, 92, .80) 48%,
          rgba(77, 105, 148, .42) 100%
        ),
        var(--spv3-new-image) center / cover no-repeat;
      box-shadow: 0 18px 42px rgba(15, 27, 53, .14);
      transition: transform .25s ease, box-shadow .25s ease;
      isolation: isolate;
      touch-action: pan-y;
    }

    .spv3-new::after {
      content: "";
      position: absolute;
      inset: 0;
      z-index: 0;
      pointer-events: none;
      background:
        linear-gradient(90deg,
          rgba(8, 20, 46, .98) 0%,
          rgba(17, 39, 76, .90) 42%,
          rgba(17, 39, 76, .28) 72%,
          rgba(17, 39, 76, .08) 100%
        );
    }

    .spv3-new-copy {
      position: relative;
      z-index: 3;
      width: min(58%, 650px);
    }

    .spv3-new-product {
      position: absolute;
      z-index: 2;
      top: 10%;
      right: 4%;
      bottom: 10%;
      width: 38%;
      display: flex;
      align-items: center;
      justify-content: center;
      pointer-events: none;
    }

    .spv3-new-product img {
      width: 100%;
      height: 100%;
      object-fit: contain;
      filter: drop-shadow(0 18px 24px rgba(0,0,0,.25));
      opacity: 0;
      transform: translateX(16px) scale(.97);
      transition: opacity .45s ease, transform .45s ease;
    }

    .spv3-new-product img.is-visible {
      opacity: 1;
      transform: translateX(0) scale(1);
    }

    .spv3-new-brand {
      display: block;
      margin-bottom: 5px;
      color: #93c5fd;
      font-size: 12px;
      font-weight: 900;
      letter-spacing: .12em;
      text-transform: uppercase;
    }

    .spv3-new-name {
      display: block;
      max-width: 620px;
      margin: 0 0 10px;
      color: #fff;
      font-size: clamp(25px, 3.2vw, 43px);
      line-height: 1.08;
      font-weight: 950;
      letter-spacing: -.035em;
    }

    .spv3-new-meta {
      display: block;
      margin-bottom: 18px;
      color: rgba(255,255,255,.74);
      font-size: 15px;
      font-weight: 700;
    }

    .spv3-dots {
      position: absolute;
      z-index: 6;
      left: 42px;
      bottom: 18px;
      display: flex;
      gap: 7px;
    }

    .spv3-dot {
      width: 8px;
      height: 8px;
      border: 0;
      border-radius: 999px;
      padding: 0;
      background: rgba(255,255,255,.38);
      cursor: pointer;
      pointer-events: auto;
      transition: width .25s ease, background .25s ease, transform .2s ease;
    }

    .spv3-dot:hover {
      transform: scale(1.16);
      background: rgba(255,255,255,.78);
    }

    .spv3-dot.is-active {
      width: 22px;
      background: #fff;
    }


    .spv3-carousel-arrow {
      position: absolute;
      z-index: 7;
      top: 50%;
      width: 42px;
      height: 42px;
      border: 1px solid rgba(255,255,255,.25);
      border-radius: 999px;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 0;
      color: #fff;
      background: rgba(15,23,42,.36);
      box-shadow: 0 8px 22px rgba(0,0,0,.18);
      backdrop-filter: blur(10px);
      cursor: pointer;
      opacity: 0;
      transform: translateY(-50%) scale(.94);
      transition:
        opacity .22s ease,
        transform .22s ease,
        background .22s ease;
    }

    .spv3-new:hover .spv3-carousel-arrow,
    .spv3-carousel-arrow:focus-visible {
      opacity: 1;
      transform: translateY(-50%) scale(1);
    }

    .spv3-carousel-arrow:hover {
      background: rgba(37,99,235,.88);
    }

    .spv3-carousel-prev { left: 14px; }
    .spv3-carousel-next { right: 14px; }

    .spv3-carousel-arrow svg {
      width: 20px;
      height: 20px;
      pointer-events: none;
    }

    .spv3-new:hover {
      transform: translateY(-3px);
      box-shadow: 0 22px 48px rgba(15, 27, 53, .18);
    }

    .spv4-product-card:focus-visible,
    .spv4-view-all:focus-visible,
    .spv3-brand-card:focus-visible {
      outline: 3px solid #60a5fa;
      outline-offset: 3px;
    }

    .spv3-kicker {
      display: block;
      margin: 0 0 8px;
      color: rgba(255,255,255,.76);
      font-size: 12px;
      line-height: 1;
      font-weight: 900;
      letter-spacing: .16em;
      text-transform: uppercase;
    }

    .spv3-new-title {
      display: block;
      margin: 0 0 10px;
      font-size: clamp(32px, 4vw, 50px);
      line-height: 1.04;
      font-weight: 950;
      letter-spacing: -.045em;
    }

    .spv3-new-subtitle {
      display: block;
      margin: 0 0 18px;
      color: rgba(255,255,255,.80);
      font-size: 18px;
    }

    .spv3-link {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      color: #fff;
      font-size: 16px;
      font-weight: 900;
    }

    /* =========================================================
       V4 · PRODUCTOS NUEVOS PARA TI
       ========================================================= */
    .spv4-recommendations {
      position: relative;
      width: 100%;
      overflow: hidden;
      border: 1px solid rgba(226,232,240,.92);
      border-radius: 28px;
      padding: 24px 24px 20px;
      background:
        radial-gradient(circle at 8% 0%, rgba(37,99,235,.08), transparent 30%),
        #fff;
      box-shadow: 0 16px 42px rgba(15,23,42,.08);
    }

    .spv4-recommendations-head {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 20px;
      margin-bottom: 18px;
    }

    .spv4-recommendations-copy {
      min-width: 0;
    }

    .spv4-recommendations-title {
      display: flex;
      align-items: center;
      gap: 10px;
      margin: 0;
      color: #0f172a;
      font-size: clamp(24px, 2.6vw, 36px);
      line-height: 1.08;
      font-weight: 950;
      letter-spacing: -.035em;
    }

    .spv4-recommendations-star {
      width: 38px;
      height: 38px;
      flex: 0 0 auto;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 13px;
      background: linear-gradient(145deg, #eff6ff, #dbeafe);
      box-shadow: inset 0 0 0 1px rgba(37,99,235,.10);
      font-size: 19px;
    }

    .spv4-recommendations-subtitle {
      margin: 7px 0 0 48px;
      color: #94a3b8;
      font-size: 14px;
      font-weight: 700;
    }

    .spv4-profile-tags {
      display: flex;
      flex-wrap: wrap;
      gap: 7px;
      margin: 12px 0 0 48px;
    }

    .spv4-profile-tag {
      display: inline-flex;
      align-items: center;
      min-height: 26px;
      border: 1px solid #dbeafe;
      border-radius: 999px;
      padding: 5px 10px;
      background: #eff6ff;
      color: #2563eb;
      font-size: 10px;
      line-height: 1;
      font-weight: 900;
      letter-spacing: .04em;
      text-transform: uppercase;
    }

    .spv4-view-all {
      flex: 0 0 auto;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      border: 0;
      border-radius: 999px;
      padding: 10px 13px;
      background: transparent;
      color: #2563eb;
      cursor: pointer;
      font-size: 12px;
      font-weight: 950;
      text-transform: uppercase;
      transition: background .2s ease, transform .2s ease;
    }

    .spv4-view-all:hover {
      background: #eff6ff;
      transform: translateX(2px);
    }

    .spv4-recommendations-viewport {
      position: relative;
    }

    .spv4-recommendations-list {
      display: grid;
      grid-auto-flow: column;
      grid-auto-columns: minmax(155px, 1fr);
      grid-template-rows: 1fr;
      gap: 14px;
      overflow-x: auto;
      overscroll-behavior-inline: contain;
      scroll-snap-type: x proximity;
      scroll-behavior: smooth;
      scrollbar-width: none;
      padding: 3px 2px 8px;
    }

    .spv4-recommendations-list::-webkit-scrollbar {
      display: none;
    }

    .spv4-product-card {
      position: relative;
      min-width: 0;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      scroll-snap-align: start;
      border: 1px solid #e2e8f0;
      border-radius: 20px;
      padding: 10px;
      background: rgba(255,255,255,.96);
      color: #0f172a;
      cursor: pointer;
      text-align: left;
      box-shadow: 0 8px 22px rgba(15,23,42,.055);
      transition: transform .22s ease, box-shadow .22s ease, border-color .22s ease;
    }

    .spv4-product-card:hover {
      transform: translateY(-3px);
      border-color: #bfdbfe;
      box-shadow: 0 14px 30px rgba(15,23,42,.11);
    }

    .spv4-product-image-wrap {
      position: relative;
      width: 100%;
      aspect-ratio: 1 / 1;
      overflow: hidden;
      display: flex;
      align-items: center;
      justify-content: center;
      border-radius: 15px;
      background: linear-gradient(145deg, #f8fafc, #f1f5f9);
    }

    .spv4-product-image-wrap img {
      width: 88%;
      height: 88%;
      object-fit: contain;
      transition: transform .28s ease;
    }

    .spv4-product-card:hover .spv4-product-image-wrap img {
      transform: scale(1.045);
    }

    .spv4-new-badge {
      position: absolute;
      z-index: 2;
      top: 8px;
      left: 8px;
      border-radius: 999px;
      padding: 5px 8px;
      background: #16a34a;
      color: #fff;
      box-shadow: 0 5px 12px rgba(22,163,74,.24);
      font-size: 8px;
      line-height: 1;
      font-weight: 950;
      letter-spacing: .08em;
    }

    .spv4-product-brand {
      margin-top: 11px;
      color: #94a3b8;
      font-size: 9px;
      line-height: 1.2;
      font-weight: 950;
      letter-spacing: .08em;
      text-transform: uppercase;
    }

    .spv4-product-name {
      display: -webkit-box;
      min-height: 34px;
      overflow: hidden;
      margin-top: 4px;
      color: #0f172a;
      font-size: 12px;
      line-height: 1.35;
      font-weight: 900;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
    }

    .spv4-product-footer {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      margin-top: auto;
      padding-top: 11px;
    }

    .spv4-product-price {
      color: #2563eb;
      font-size: 14px;
      line-height: 1;
      font-weight: 950;
    }

    .spv4-product-open {
      width: 28px;
      height: 28px;
      flex: 0 0 auto;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 999px;
      background: #eff6ff;
      color: #2563eb;
      font-size: 17px;
      font-weight: 900;
    }

    .spv4-scroll-arrow {
      position: absolute;
      z-index: 6;
      top: 50%;
      width: 38px;
      height: 38px;
      border: 1px solid #e2e8f0;
      border-radius: 999px;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 0;
      background: rgba(255,255,255,.96);
      color: #0f172a;
      box-shadow: 0 8px 22px rgba(15,23,42,.13);
      cursor: pointer;
      opacity: 0;
      transform: translateY(-50%) scale(.94);
      transition: opacity .2s ease, transform .2s ease, background .2s ease;
    }

    .spv4-recommendations:hover .spv4-scroll-arrow,
    .spv4-scroll-arrow:focus-visible {
      opacity: 1;
      transform: translateY(-50%) scale(1);
    }

    .spv4-scroll-arrow:hover {
      background: #2563eb;
      color: #fff;
    }

    .spv4-scroll-prev { left: -8px; }
    .spv4-scroll-next { right: -8px; }

    .spv4-empty {
      min-height: 190px;
      display: flex;
      align-items: center;
      justify-content: center;
      border: 1px dashed #cbd5e1;
      border-radius: 20px;
      color: #94a3b8;
      text-align: center;
      font-size: 13px;
      font-weight: 800;
    }

    .dark .spv4-recommendations,
    [data-theme="dark"] .spv4-recommendations {
      border-color: #334155;
      background:
        radial-gradient(circle at 8% 0%, rgba(59,130,246,.13), transparent 30%),
        #1e293b;
    }

    .dark .spv4-recommendations-title,
    [data-theme="dark"] .spv4-recommendations-title,
    .dark .spv4-product-name,
    [data-theme="dark"] .spv4-product-name {
      color: #f8fafc;
    }

    .dark .spv4-product-card,
    [data-theme="dark"] .spv4-product-card {
      border-color: #334155;
      background: #0f172a;
    }

    .dark .spv4-product-image-wrap,
    [data-theme="dark"] .spv4-product-image-wrap {
      background: linear-gradient(145deg, #1e293b, #111827);
    }

    @media (min-width: 1180px) {
      .spv4-recommendations-list {
        grid-auto-columns: calc((100% - 84px) / 7);
      }
    }

    @media (min-width: 900px) and (max-width: 1179px) {
      .spv4-recommendations-list {
        grid-auto-columns: calc((100% - 56px) / 5);
      }
    }

    .spv3-heading {
      margin: 46px 2px 20px;
    }

    .spv3-heading h2 {
      margin: 0 0 6px;
      color: #0f172a;
      font-size: clamp(31px, 3.2vw, 44px);
      line-height: 1.05;
      font-weight: 950;
      letter-spacing: -.045em;
    }

    .spv3-heading p {
      margin: 0;
      color: #94a3b8;
      font-size: 17px;
      font-weight: 500;
    }

    .spv3-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 18px;
      margin-bottom: 34px;
    }

    .spv3-brand-card {
      --spv3-x: 50%;
      --spv3-y: 50%;
      position: relative;
      min-height: 340px;
      overflow: hidden;
      border: 0;
      border-radius: 26px;
      padding: 0;
      background: #111827;
      cursor: pointer;
      box-shadow: 0 15px 36px rgba(15, 23, 42, .12);
      text-align: left;
      isolation: isolate;
      transition: transform .24s ease, box-shadow .24s ease;
    }

    .spv3-brand-card:hover {
      transform: translateY(-3px) scale(1.008);
      box-shadow: 0 20px 45px rgba(15, 23, 42, .18);
    }

    .spv3-brand-card img {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      object-fit: cover;
      display: block;
      filter: saturate(.92) contrast(1.03);
      transition: transform .40s ease, filter .30s ease;
    }

    .spv3-brand-card:hover img {
      transform: scale(1.025);
      filter: saturate(1) contrast(1.05);
    }


    .spv3-brand-card::before {
      content: "";
      position: absolute;
      inset: 0;
      z-index: 2;
      opacity: 0;
      pointer-events: none;
      background:
        radial-gradient(
          420px circle at var(--spv3-x) var(--spv3-y),
          rgba(255,255,255,.18),
          transparent 44%
        );
      transition: opacity .22s ease;
    }

    .spv3-brand-card:hover::before {
      opacity: 1;
    }



    .spv3-brand-card[data-brand="XIAOMI"] img { object-position: 50% 48%; }
    .spv3-brand-card[data-brand="DJI"] img { object-position: 50% 45%; }
    .spv3-brand-card[data-brand="HUAWEI"] img { object-position: 50% 48%; }
    .spv3-brand-card[data-brand="QCY"] img { object-position: 50% 38%; }
    .spv3-brand-card[data-brand="GAMESIR"] img { object-position: 50% 52%; }
    .spv3-brand-card[data-brand="SOLOHOUR"] img { object-position: 50% 48%; }

    .spv3-brand-card::after {
      content: "";
      position: absolute;
      inset: 0;
      z-index: 1;
      background:
        linear-gradient(180deg,
          rgba(0,0,0,.02) 28%,
          rgba(0,0,0,.82) 100%
        );
    }

    .spv3-brand-copy {
      position: absolute;
      z-index: 2;
      left: 24px;
      right: 24px;
      bottom: 22px;
      color: #fff;
    }

    .spv3-brand-name {
      display: block;
      margin-bottom: 5px;
      font-size: 31px;
      line-height: 1;
      font-weight: 950;
      letter-spacing: -.035em;
    }

    .spv3-brand-desc {
      display: block;
      margin-bottom: 15px;
      color: rgba(255,255,255,.78);
      font-size: 16px;
      line-height: 1.35;
    }

    #${IDS.status} {
      position: sticky;
      top: 74px;
      z-index: 25;
      display: none;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin: 0 0 18px;
      padding: 12px 14px;
      border: 1px solid rgba(47, 111, 237, .18);
      border-radius: 14px;
      background: rgba(239, 246, 255, .96);
      color: #1e3a8a;
      box-shadow: 0 8px 22px rgba(37, 99, 235, .08);
      backdrop-filter: blur(10px);
      font-size: 14px;
      font-weight: 800;
    }

    #${IDS.status}.is-visible {
      display: flex;
    }

    .spv3-clear {
      flex: 0 0 auto;
      border: 0;
      border-radius: 999px;
      padding: 7px 12px;
      background: #2563eb;
      color: #fff;
      cursor: pointer;
      font-weight: 900;
    }

    .dark .spv3-heading h2,
    [data-theme="dark"] .spv3-heading h2 {
      color: #fff;
    }

    .dark #${IDS.status},
    [data-theme="dark"] #${IDS.status} {
      border-color: rgba(96,165,250,.25);
      background: rgba(15,23,42,.94);
      color: #dbeafe;
    }

    @media (prefers-reduced-motion: reduce) {
      .spv3-new-product img {
        transition: none !important;
      }
    }

    @media (max-width: 1050px) and (min-width: 768px) {
      .spv3-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
    }

    @media (max-width: 767px) {
      #${IDS.root} {
        padding: 14px 14px 0;
      }

      .spv3-new {
        min-height: 190px;
        border-radius: 22px;
        padding: 22px;
      }

      .spv3-new-copy {
        width: 64%;
      }

      .spv3-new-product {
        top: 14%;
        right: 1%;
        bottom: 14%;
        width: 38%;
      }

      .spv3-new-name {
        max-width: 100%;
        font-size: 23px;
        line-height: 1.08;
      }

      .spv3-new-meta {
        margin-bottom: 11px;
        font-size: 12px;
      }

      .spv3-new-brand {
        font-size: 9px;
      }

      .spv3-dots {
        left: 22px;
        bottom: 12px;
      }

      .spv3-carousel-arrow {
        width: 34px;
        height: 34px;
        opacity: .86;
        transform: translateY(-50%) scale(1);
      }

      .spv3-carousel-prev { left: 8px; }
      .spv3-carousel-next { right: 8px; }

      .spv3-kicker {
        font-size: 10px;
      }

      .spv3-new-title {
        max-width: 88%;
        margin-bottom: 7px;
        font-size: 28px;
      }

      .spv3-new-subtitle {
        margin-bottom: 12px;
        font-size: 14px;
      }

      .spv3-link {
        font-size: 14px;
      }

      .spv4-recommendations {
        border-radius: 22px;
        padding: 18px 14px 14px;
      }

      .spv4-recommendations-head {
        gap: 8px;
        margin-bottom: 14px;
      }

      .spv4-recommendations-title {
        gap: 8px;
        font-size: 21px;
      }

      .spv4-recommendations-star {
        width: 32px;
        height: 32px;
        border-radius: 11px;
        font-size: 16px;
      }

      .spv4-recommendations-subtitle {
        margin: 6px 0 0 40px;
        font-size: 11px;
      }

      .spv4-profile-tags {
        margin: 9px 0 0 40px;
        gap: 5px;
      }

      .spv4-profile-tag {
        min-height: 23px;
        padding: 4px 8px;
        font-size: 8px;
      }

      .spv4-view-all {
        padding: 7px 5px;
        font-size: 9px;
      }

      .spv4-recommendations-list {
        grid-auto-columns: calc((100% - 12px) / 2.22);
        gap: 12px;
        margin-right: -14px;
        padding-right: 14px;
      }

      .spv4-product-card {
        border-radius: 17px;
        padding: 8px;
      }

      .spv4-product-image-wrap {
        border-radius: 13px;
      }

      .spv4-product-name {
        min-height: 32px;
        font-size: 11px;
      }

      .spv4-product-price {
        font-size: 13px;
      }

      .spv4-scroll-arrow {
        display: none;
      }

      .spv3-heading {
        margin: 31px 2px 16px;
      }

      .spv3-heading h2 {
        font-size: 30px;
      }

      .spv3-heading p {
        font-size: 15px;
      }

      .spv3-grid {
        grid-template-columns: 1fr;
        gap: 12px;
        margin-bottom: 26px;
      }

      .spv3-brand-card {
        min-height: 168px;
        border-radius: 21px;
      }

      .spv3-brand-copy {
        left: 19px;
        right: 19px;
        bottom: 17px;
      }

      .spv3-brand-name {
        font-size: 24px;
      }

      .spv3-brand-desc {
        margin-bottom: 10px;
        font-size: 14px;
      }

      #${IDS.status} {
        top: 66px;
      }
    }
  `;

  function addStyles() {
    if (document.getElementById(IDS.style)) return;

    const style = document.createElement('style');
    style.id = IDS.style;
    style.textContent = CSS;
    document.head.appendChild(style);
  }

  function safeCall(name, ...args) {
    const fn = window[name];
    if (typeof fn !== 'function') return false;

    try {
      fn(...args);
      return true;
    } catch (error) {
      console.warn(`[Starphone Home V4.0] ${name} falló:`, error);
      return false;
    }
  }

  /*
   * index.html declares products, activeFilters and isFavoriteMode with let.
   * They are available by identifier in this classic script, but not as
   * window.products / window.activeFilters properties.
   */
  function getProducts() {
    try {
      return Array.isArray(products) ? products : [];
    } catch (_) {
      return [];
    }
  }

  function clearNativeInputs() {
    const input = document.getElementById('search');
    if (input) input.value = '';
  }

  function resetFilters() {
    /* Use the site's own reset so all native UI state stays consistent. */
    safeCall('resetFilters', false);

    try {
      activeFilters.keyword = '';
      activeFilters.MARCA = null;
      activeFilters.CATEGORIA = null;
      activeFilters.TIPO = null;
      activeFilters.priceRange = null;
      isFavoriteMode = false;
    } catch (error) {
      console.warn('[Starphone Home V4.0] Error limpiando filtros:', error);
    }

    clearNativeInputs();
  }

  function normalizeBrandName(name) {
    const wanted = String(name || '').trim().toUpperCase();
    const found = getProducts().find((product) =>
      String(product?.MARCA || '').trim().toUpperCase() === wanted
    );
    return found ? String(found.MARCA).trim() : name;
  }

  function rerender() {
    /* Refresh the same pieces the native filters refresh. */
    safeCall('renderFilters');
    safeCall('renderSourceTabs');
    safeCall('render');
    safeCall('updateActiveFiltersPills');
    safeCall('updateFilterIconStatus');
    syncHomeStatusFromNativeFilters();

    window.dispatchEvent(
      new CustomEvent('starphone-home-filter-change', {
        detail: {
          activeFilters: (() => {
            try { return { ...activeFilters }; }
            catch (_) { return {}; }
          })()
        }
      })
    );
  }

  function getProductsTarget() {
    return (
      document.getElementById(CONFIG.insertBeforeId) ||
      document.getElementById(CONFIG.activeFiltersRowId) ||
      document.getElementById(CONFIG.productsGridId) ||
      document.querySelector('main')
    );
  }

  function scrollToProducts() {
    const target = getProductsTarget();
    if (!target) return;

    const stickyHeader = document.querySelector('header.glass-header');
    const headerHeight = stickyHeader ? stickyHeader.getBoundingClientRect().height : 0;
    const top = target.getBoundingClientRect().top + window.scrollY - headerHeight - 14;

    window.scrollTo({
      top: Math.max(0, top),
      behavior: 'smooth'
    });
  }

  function showStatus(text) {
    const status = document.getElementById(IDS.status);
    if (!status) return;

    const label = status.querySelector('[data-spv3-status-label]');
    if (label) label.textContent = text;

    status.classList.add('is-visible');
  }

  function hideStatus() {
    document.getElementById(IDS.status)?.classList.remove('is-visible');
  }

  async function applyNewFilter() {
    let defaultSourceKey = null;

    try {
      if (Array.isArray(sources) && sources.length > 0) {
        defaultSourceKey = sources[0].key;
      }
    } catch (_) {}

    /*
     * 点击新品横幅：
     * 1. 退出 FAVORITOS
     * 2. 回到第一个数据源（LISTA ACTUAL）
     * 3. 等产品加载完成
     * 4. 筛选 NUEVO 并滚动到商品区
     */
    try {
      isFavoriteMode = false;
    } catch (_) {}

    if (defaultSourceKey) {
      let needSwitch = false;

      try {
        needSwitch = currentSourceKey !== defaultSourceKey;
      } catch (_) {}

      if (needSwitch) {
        safeCall('setSource', defaultSourceKey);

        const startedAt = Date.now();
        await new Promise((resolve) => {
          const timer = window.setInterval(() => {
            let ready = false;

            try {
              ready =
                currentSourceKey === defaultSourceKey &&
                Array.isArray(products) &&
                products.length > 0;
            } catch (_) {}

            if (ready || Date.now() - startedAt > 5000) {
              window.clearInterval(timer);
              resolve();
            }
          }, 80);
        });
      }
    }

    resetFilters();

    try {
      activeFilters.keyword = CONFIG.newKeyword;
      isFavoriteMode = false;
    } catch (error) {
      console.warn('[Starphone Home V4.0] No se pudo aplicar Nuevo:', error);
    }

    rerender();
    showStatus('Filtro activo: Novedades');

    window.setTimeout(scrollToProducts, 100);
  }

  async function applyBrandFilter(name) {
    let defaultSourceKey = null;

    try {
      if (Array.isArray(sources) && sources.length > 0) {
        defaultSourceKey = sources[0].key;
      }
    } catch (_) {}

    // 先退出 FAVORITOS，再回到第一个数据源 LISTA ACTUAL
    try {
      isFavoriteMode = false;
    } catch (_) {}

    if (defaultSourceKey) {
      let needSwitch = false;

      try {
        needSwitch = currentSourceKey !== defaultSourceKey;
      } catch (_) {}

      if (needSwitch) {
        safeCall('setSource', defaultSourceKey);

        const startedAt = Date.now();

        await new Promise((resolve) => {
          const timer = window.setInterval(() => {
            let ready = false;

            try {
              ready =
                currentSourceKey === defaultSourceKey &&
                Array.isArray(products) &&
                products.length > 0;
            } catch (_) {}

            if (ready || Date.now() - startedAt > 5000) {
              window.clearInterval(timer);
              resolve();
            }
          }, 80);
        });
      }
    }

    resetFilters();

    const normalized = normalizeBrandName(name);

    try {
      activeFilters.MARCA = normalized;
      isFavoriteMode = false;
    } catch (error) {
      console.warn('[Starphone Home V4.0] No se pudo aplicar la marca:', error);
    }

    rerender();
    showStatus(`Marca seleccionada: ${normalized}`);

    window.setTimeout(scrollToProducts, 100);
  }

  function clearHomeFilter() {
    resetFilters();
    rerender();
    hideStatus();

    window.setTimeout(scrollToProducts, 80);
  }


  function syncHomeStatusFromNativeFilters() {
    const status = document.getElementById(IDS.status);
    if (!status) return;

    let keyword = '';
    let brand = null;

    try {
      keyword = String(activeFilters?.keyword || '').trim();
      brand = activeFilters?.MARCA || null;
    } catch (_) {}

    const isNewFilter =
      keyword.toUpperCase() === String(CONFIG.newKeyword || '').toUpperCase();

    if (brand) {
      showStatus(`Marca seleccionada: ${brand}`);
      return;
    }

    if (isNewFilter) {
      showStatus('Filtro activo: Novedades');
      return;
    }

    hideStatus();
  }

  function setupNativeFilterSync() {
    if (document.documentElement.dataset.spHomeFilterSync === '1') return;
    document.documentElement.dataset.spHomeFilterSync = '1';

    const sync = () => {
      window.setTimeout(syncHomeStatusFromNativeFilters, 0);
    };

    // The native filter pill row changes whenever a pill is removed.
    const pillRow = document.getElementById(CONFIG.activeFiltersRowId);
    if (pillRow) {
      const observer = new MutationObserver(sync);
      observer.observe(pillRow, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['style', 'class']
      });
    }

    // Capture clicks on native filter controls and clear buttons.
    document.addEventListener('click', (event) => {
      const target = event.target.closest(
        '.filter-pill-remove, [onclick*="removeFilter"], [onclick*="resetFilters"], [onclick*="Limpiar"], #filter-content button'
      );
      if (target) sync();
    }, true);

    // Keep the custom status synchronized after native rendering.
    window.addEventListener('starphone-home-filter-change', sync);

    sync();
  }

  function createStatus() {
    const status = document.createElement('div');
    status.id = IDS.status;

    status.innerHTML = `
      <span data-spv3-status-label></span>
      <button class="spv3-clear" type="button">Ver todos</button>
    `;

    status.querySelector('.spv3-clear').addEventListener('click', clearHomeFilter);

    return status;
  }



  const recommendationState = {
    allProducts: [],
    newProducts: [],
    recommendedProducts: [],
    preferredBrands: [],
    preferredCategories: [],
    personalized: false,
    loaded: false,
    loading: false,
    sourceSignature: ''
  };

  function isOfferSource(source) {
    const label = String(source?.label || '').trim().toUpperCase();
    const key = String(source?.key || '').trim().toUpperCase();
    return label.includes('OFERTA') || key.includes('OFERTA');
  }

  function normalizeKey(value) {
    return String(value || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .replace(/\s+/g, ' ')
      .trim();
  }

  function productIdentity(product) {
    return normalizeKey(
      product?.id ||
      product?.PRODUCTO ||
      `${product?.MARCA || ''}|${product?.Imagen_Path || ''}`
    );
  }

  function isNewProduct(product) {
    const newFlag = normalizeKey(product?.['新到货']);
    const type = normalizeKey(product?.TIPO);
    const name = normalizeKey(product?.PRODUCTO);

    return (
      newFlag === 'NUEVO' ||
      type === 'NUEVO' ||
      name.includes('NUEVO')
    );
  }

  function dedupeProducts(list) {
    const map = new Map();

    list.forEach((product) => {
      const key = productIdentity(product);
      if (key && !map.has(key)) map.set(key, product);
    });

    return [...map.values()];
  }

  function productImage(product) {
    const raw = String(
      product?.Imagen_Path ||
      product?.IMAGEN ||
      product?.image ||
      ''
    ).trim();

    return raw ? encodeURI(raw.replace(/\\/g, '/')) : '';
  }

  function productPrice(product) {
    const raw =
      product?.['Precio ( USD )'] ??
      product?.['Precio USD'] ??
      product?.PRECIO ??
      '';

    const value = parseFloat(raw);

    return Number.isFinite(value)
      ? `$${value.toLocaleString('en-US', {
          minimumFractionDigits: value % 1 ? 2 : 0,
          maximumFractionDigits: 2
        })}`
      : '';
  }

  function readOrderHistory() {
    try {
      const history = JSON.parse(
        window.localStorage.getItem('order_history') || '[]'
      );
      return Array.isArray(history) ? history : [];
    } catch (error) {
      console.warn('[Starphone Home V4.0] No se pudo leer el historial:', error);
      return [];
    }
  }

  function findHistoricalProduct(item, allProducts) {
    const wantedId = normalizeKey(item?.id);
    const wantedName = normalizeKey(item?.name);

    return allProducts.find((product) => {
      const productId = normalizeKey(product?.id);
      const productName = normalizeKey(product?.PRODUCTO);

      return (
        (wantedId && (productId === wantedId || productName === wantedId)) ||
        (wantedName && productName === wantedName)
      );
    }) || null;
  }

  function incrementScore(map, key, amount) {
    const normalized = normalizeKey(key);
    if (!normalized) return;
    map.set(normalized, (map.get(normalized) || 0) + amount);
  }

  function rankMap(map) {
    return [...map.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([name, score]) => ({ name, score }));
  }

  function getFallbackPopularity(product) {
    const possibleValues = [
      product?.HOT,
      product?.POPULAR,
      product?.['热门'],
      product?.['销量'],
      product?.VENTAS,
      product?.SALES
    ];

    for (const value of possibleValues) {
      const numeric = Number(value);
      if (Number.isFinite(numeric)) return numeric;

      const text = normalizeKey(value);
      if (['HOT', 'SI', 'YES', 'POPULAR', 'MAS VENDIDO'].includes(text)) {
        return 1;
      }
    }

    return 0;
  }

  function buildPersonalizedRecommendations(allProducts, newProducts) {
    const history = readOrderHistory();
    const brandScores = new Map();
    const categoryScores = new Map();
    const typeScores = new Map();

    history.forEach((order) => {
      const items = Array.isArray(order?.items) ? order.items : [];

      items.forEach((item) => {
        const product = findHistoricalProduct(item, allProducts);
        if (!product) return;

        const quantity = Math.max(1, Number(item?.qty) || 1);
        incrementScore(brandScores, product?.MARCA, quantity);
        incrementScore(categoryScores, product?.CATEGORIA, quantity);
        incrementScore(typeScores, product?.TIPO, quantity);
      });
    });

    const rankedBrands = rankMap(brandScores);
    const rankedCategories = rankMap(categoryScores);
    const rankedTypes = rankMap(typeScores);

    const scoreProduct = (product, originalIndex) => {
      const brand = normalizeKey(product?.MARCA);
      const category = normalizeKey(product?.CATEGORIA);
      const type = normalizeKey(product?.TIPO);

      const brandRank = rankedBrands.findIndex((entry) => entry.name === brand);
      const categoryRank = rankedCategories.findIndex((entry) => entry.name === category);
      const typeRank = rankedTypes.findIndex((entry) => entry.name === type);

      let score = 0;
      let tier = 3;

      if (brandRank >= 0) {
        tier = 1;
        score +=
          CONFIG.recommendationBrandWeight * rankedBrands[brandRank].score -
          brandRank;
      } else if (categoryRank >= 0) {
        tier = 2;
        score +=
          CONFIG.recommendationCategoryWeight * rankedCategories[categoryRank].score -
          categoryRank;
      } else if (typeRank >= 0) {
        tier = 2;
        score +=
          CONFIG.recommendationTypeWeight * rankedTypes[typeRank].score -
          typeRank;
      }

      score += getFallbackPopularity(product) * 5;

      return {
        product,
        tier,
        score,
        originalIndex
      };
    };

    const rankedProducts = newProducts
      .map(scoreProduct)
      .sort((a, b) =>
        a.tier - b.tier ||
        b.score - a.score ||
        a.originalIndex - b.originalIndex
      )
      .slice(0, CONFIG.recommendationLimit)
      .map((entry) => entry.product);

    return {
      products: rankedProducts,
      personalized: rankedBrands.length > 0 || rankedCategories.length > 0,
      preferredBrands: rankedBrands.slice(0, 3).map((entry) => entry.name),
      preferredCategories: rankedCategories.slice(0, 2).map((entry) => entry.name)
    };
  }

  async function loadRecommendationPool(force = false) {
    if (recommendationState.loading) return;

    let sourceList = [];
    try {
      sourceList = Array.isArray(sources) ? sources : [];
    } catch (_) {
      sourceList = [];
    }

    if (!sourceList.length) return;

    const eligibleSources = sourceList.filter((source) => !isOfferSource(source));
    const sourceSignature = eligibleSources
      .map((source) => `${source?.key || ''}|${source?.file || ''}`)
      .join('::');

    if (
      !force &&
      recommendationState.loaded &&
      recommendationState.sourceSignature === sourceSignature
    ) {
      renderRecommendations();
      return;
    }

    recommendationState.loading = true;

    try {
      const results = await Promise.all(
        eligibleSources.map(async (source) => {
          try {
            const response = await fetch(
              `${source.file}${String(source.file).includes('?') ? '&' : '?'}v=${Date.now()}`
            );

            if (!response.ok) return [];

            const raw = await response.json();
            if (!Array.isArray(raw)) return [];

            return raw.map((product) => ({
              ...product,
              id: product?.id || product?.PRODUCTO,
              __sourceKey: source?.key || '',
              __sourceLabel: source?.label || ''
            }));
          } catch (error) {
            console.warn(
              `[Starphone Home V4.0] No se pudo cargar ${source?.label || source?.key || 'fuente'}:`,
              error
            );
            return [];
          }
        })
      );

      const allProducts = dedupeProducts(results.flat());
      const newProducts = allProducts.filter(isNewProduct);
      const recommendation = buildPersonalizedRecommendations(
        allProducts,
        newProducts
      );

      recommendationState.allProducts = allProducts;
      recommendationState.newProducts = newProducts;
      recommendationState.recommendedProducts = recommendation.products;
      recommendationState.personalized = recommendation.personalized;
      recommendationState.preferredBrands = recommendation.preferredBrands;
      recommendationState.preferredCategories = recommendation.preferredCategories;
      recommendationState.sourceSignature = sourceSignature;
      recommendationState.loaded = true;

      renderRecommendations();
    } finally {
      recommendationState.loading = false;
    }
  }

  function recommendationCardHtml(product, index) {
    const image = productImage(product);
    const name = String(product?.PRODUCTO || 'Producto nuevo');
    const brand = String(product?.MARCA || 'NUEVO');
    const price = productPrice(product);

    return `
      <button
        class="spv4-product-card"
        type="button"
        data-spv4-product-index="${index}"
        aria-label="Ver ${name.replace(/"/g, '&quot;')}"
      >
        <span class="spv4-product-image-wrap">
          <span class="spv4-new-badge">NUEVO</span>
          ${image
            ? `<img src="${image}" alt="" loading="lazy" decoding="async">`
            : `<span aria-hidden="true" style="font-size:34px">📦</span>`
          }
        </span>

        <span class="spv4-product-brand">${brand}</span>
        <strong class="spv4-product-name">${name}</strong>

        <span class="spv4-product-footer">
          <span class="spv4-product-price">${price || 'Consultar'}</span>
          <span class="spv4-product-open" aria-hidden="true">→</span>
        </span>
      </button>
    `;
  }

  function renderRecommendations() {
    const root = document.getElementById(IDS.root);
    if (!root) return;

    const list = root.querySelector('[data-spv4-recommendations-list]');
    const subtitle = root.querySelector('[data-spv4-recommendations-subtitle]');
    const tags = root.querySelector('[data-spv4-profile-tags]');

    if (!list || !subtitle || !tags) return;

    const products = recommendationState.recommendedProducts;

    subtitle.textContent = recommendationState.personalized
      ? CONFIG.recommendations.personalizedSubtitle
      : CONFIG.recommendations.genericSubtitle;

    const tagNames = recommendationState.preferredBrands.length
      ? recommendationState.preferredBrands
      : recommendationState.preferredCategories;

    tags.innerHTML = tagNames
      .map((name) => `<span class="spv4-profile-tag">${name}</span>`)
      .join('');
    tags.hidden = tagNames.length === 0;

    if (!products.length) {
      list.innerHTML = `
        <div class="spv4-empty">
          No hay novedades disponibles en este momento.
        </div>
      `;
      return;
    }

    list.innerHTML = products
      .map(recommendationCardHtml)
      .join('');
  }

  async function openRecommendedProduct(product) {
    if (!product) return;

    const sourceKey = product.__sourceKey;

    try {
      isFavoriteMode = false;
    } catch (_) {}

    if (sourceKey) {
      let shouldSwitch = false;

      try {
        shouldSwitch = currentSourceKey !== sourceKey;
      } catch (_) {}

      if (shouldSwitch) {
        safeCall('setSource', sourceKey);

        const startedAt = Date.now();
        await new Promise((resolve) => {
          const timer = window.setInterval(() => {
            let ready = false;

            try {
              ready =
                currentSourceKey === sourceKey &&
                Array.isArray(products) &&
                products.length > 0;
            } catch (_) {}

            if (ready || Date.now() - startedAt > 5000) {
              window.clearInterval(timer);
              resolve();
            }
          }, 80);
        });
      }
    }

    resetFilters();

    const productName = String(product?.PRODUCTO || '').trim();

    try {
      activeFilters.keyword = productName;
      isFavoriteMode = false;
    } catch (error) {
      console.warn('[Starphone Home V4.0] No se pudo abrir el producto:', error);
    }

    const searchInput = document.getElementById('search');
    if (searchInput) searchInput.value = productName;

    rerender();
    showStatus(`Producto seleccionado: ${productName}`);
    window.setTimeout(scrollToProducts, 100);
  }

  function scrollRecommendationList(direction) {
    const root = document.getElementById(IDS.root);
    const list = root?.querySelector('[data-spv4-recommendations-list]');
    if (!list) return;

    const amount = Math.max(260, list.clientWidth * .82);
    list.scrollBy({
      left: direction === 'prev' ? -amount : amount,
      behavior: 'smooth'
    });
  }

  function setupRecommendations(root) {
    if (!root) return;

    const list = root.querySelector('[data-spv4-recommendations-list]');
    const viewAll = root.querySelector('[data-spv4-view-all]');
    const previous = root.querySelector('[data-spv4-scroll-prev]');
    const next = root.querySelector('[data-spv4-scroll-next]');

    viewAll?.addEventListener('click', applyNewFilter);
    previous?.addEventListener('click', () => scrollRecommendationList('prev'));
    next?.addEventListener('click', () => scrollRecommendationList('next'));

    list?.addEventListener('click', (event) => {
      const card = event.target.closest('[data-spv4-product-index]');
      if (!card) return;

      const index = Number(card.dataset.spv4ProductIndex);
      openRecommendedProduct(recommendationState.recommendedProducts[index]);
    });

    let tries = 0;
    const waitForSources = window.setInterval(() => {
      tries += 1;

      let hasSources = false;
      try {
        hasSources = Array.isArray(sources) && sources.length > 0;
      } catch (_) {}

      if (hasSources) {
        window.clearInterval(waitForSources);
        loadRecommendationPool(true);
      } else if (tries >= 40) {
        window.clearInterval(waitForSources);
      }
    }, 300);

    window.addEventListener('storage', (event) => {
      if (event.key === 'order_history') {
        loadRecommendationPool(true);
      }
    });
  }

  function setupPointerGlow(root) {
    if (!root || window.matchMedia('(pointer: coarse)').matches) return;

    root.querySelectorAll('.spv3-brand-card').forEach((card) => {
      card.addEventListener('pointermove', (event) => {
        const rect = card.getBoundingClientRect();
        const x = ((event.clientX - rect.left) / rect.width) * 100;
        const y = ((event.clientY - rect.top) / rect.height) * 100;
        card.style.setProperty('--spv3-x', `${x}%`);
        card.style.setProperty('--spv3-y', `${y}%`);
      });

      card.addEventListener('pointerleave', () => {
        card.style.setProperty('--spv3-x', '50%');
        card.style.setProperty('--spv3-y', '50%');
      });
    });
  }

  function createSection() {
    const root = document.createElement('section');
    root.id = IDS.root;
    root.setAttribute('aria-label', 'Recomendaciones y pabellón de marcas');

    const brandsHtml = CONFIG.brands.map((brand) => `
      <button
        class="spv3-brand-card"
        type="button"
        data-brand="${brand.filter}"
        aria-label="Ver productos ${brand.title}"
      >
        <img
          src="${brand.image}"
          alt=""
          loading="lazy"
          decoding="async"
        >
        <span class="spv3-brand-copy">
          <strong class="spv3-brand-name">${brand.title}</strong>
          <span class="spv3-brand-desc">${brand.description}</span>
          <span class="spv3-link">
            Ver productos
            <span aria-hidden="true">→</span>
          </span>
        </span>
      </button>
    `).join('');

    root.innerHTML = `
      <section class="spv4-recommendations" aria-labelledby="spv4-recommendations-title">
        <div class="spv4-recommendations-head">
          <div class="spv4-recommendations-copy">
            <h2 class="spv4-recommendations-title" id="spv4-recommendations-title">
              <span class="spv4-recommendations-star" aria-hidden="true">⭐</span>
              <span>${CONFIG.recommendations.title}</span>
            </h2>

            <p
              class="spv4-recommendations-subtitle"
              data-spv4-recommendations-subtitle
            >
              ${CONFIG.recommendations.genericSubtitle}
            </p>

            <div
              class="spv4-profile-tags"
              data-spv4-profile-tags
              hidden
            ></div>
          </div>

          <button
            class="spv4-view-all"
            type="button"
            data-spv4-view-all
          >
            ${CONFIG.recommendations.button}
            <span aria-hidden="true">→</span>
          </button>
        </div>

        <div class="spv4-recommendations-viewport">
          <button
            class="spv4-scroll-arrow spv4-scroll-prev"
            type="button"
            data-spv4-scroll-prev
            aria-label="Ver productos anteriores"
          >
            ‹
          </button>

          <div
            class="spv4-recommendations-list"
            data-spv4-recommendations-list
          >
            <div class="spv4-empty">Cargando novedades...</div>
          </div>

          <button
            class="spv4-scroll-arrow spv4-scroll-next"
            type="button"
            data-spv4-scroll-next
            aria-label="Ver más productos"
          >
            ›
          </button>
        </div>
      </section>

      <header class="spv3-heading">
        <h2>Pabellón de Marcas</h2>
        <p>Explora nuestras marcas destacadas</p>
      </header>

      <div class="spv3-grid">
        ${brandsHtml}
      </div>
    `;

    root.querySelectorAll('.spv3-brand-card').forEach((card) => {
      card.addEventListener('click', () => {
        applyBrandFilter(card.dataset.brand);
      });
    });

    setupPointerGlow(root);
    setupRecommendations(root);

    return root;
  }

  function insertModule() {
    if (document.getElementById(IDS.root)) return true;

    const insertBefore = document.getElementById(CONFIG.insertBeforeId);
    if (!insertBefore) return false;

    const parent = insertBefore.parentElement;
    if (!parent) return false;

    const section = createSection();
    const status = createStatus();

    parent.insertBefore(section, insertBefore);
    parent.insertBefore(status, insertBefore);

    setupNativeFilterSync();
    syncHomeStatusFromNativeFilters();

    return true;
  }


  function enableProtection() {
    if (document.documentElement.dataset.spProtectionEnabled === '1') return;
    document.documentElement.dataset.spProtectionEnabled = '1';

    // 全站禁止右键菜单
    document.addEventListener('contextmenu', (event) => {
      event.preventDefault();
    }, { capture: true });

    // 禁止所有图片拖拽
    document.addEventListener('dragstart', (event) => {
      if (event.target instanceof HTMLImageElement) {
        event.preventDefault();
      }
    }, { capture: true });

    // 处理当前页面以及后续动态生成的图片
    const protectImages = (root = document) => {
      root.querySelectorAll?.('img').forEach((img) => {
        img.draggable = false;
        img.setAttribute('draggable', 'false');
        img.style.webkitUserDrag = 'none';
        img.style.userSelect = 'none';
      });
    };

    protectImages();

    const imageObserver = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        mutation.addedNodes.forEach((node) => {
          if (!(node instanceof Element)) return;

          if (node.matches('img')) {
            node.draggable = false;
            node.setAttribute('draggable', 'false');
            node.style.webkitUserDrag = 'none';
            node.style.userSelect = 'none';
          }

          protectImages(node);
        });
      });
    });

    imageObserver.observe(document.documentElement, {
      childList: true,
      subtree: true
    });
  }

  function init() {
    enableProtection();
    addStyles();

    if (insertModule()) return;

    const observer = new MutationObserver(() => {
      if (insertModule()) observer.disconnect();
    });

    observer.observe(document.documentElement, {
      childList: true,
      subtree: true
    });

    window.setTimeout(() => observer.disconnect(), 15000);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
