
(() => {
  'use strict';

  /* =========================================================
     STARPHONE HOME SECTIONS V3.6-SYNC
     - Inserta Novedades + Pabellón de Marcas
     - Se coloca antes de #source-tabs
     - No modifica tarjetas de producto, carrito, favoritos ni WhatsApp
     ========================================================= */

  const CONFIG = {
    insertBeforeId: 'source-tabs',
    productsGridId: 'grid',
    activeFiltersRowId: 'active-filters-row',
    newKeyword: 'nuevo',
    sectionMaxWidth: '80rem',
    carouselInterval: 4200,
    carouselLimit: 10,

    newArrivals: {
      kicker: 'Recién llegados',
      title: '🔥 Novedades de la Semana',
      subtitle: 'Nuevos productos disponibles.',
      button: 'Ver novedades',
      image:
        'https://images.unsplash.com/photo-1517336714731-489689fd1ca8?auto=format&fit=crop&w=1800&q=90'
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
      z-index: 4;
      left: 42px;
      bottom: 18px;
      display: flex;
      gap: 7px;
    }

    .spv3-dot {
      width: 7px;
      height: 7px;
      border: 0;
      border-radius: 999px;
      padding: 0;
      background: rgba(255,255,255,.35);
      pointer-events: none;
      transition: width .25s ease, background .25s ease;
    }

    .spv3-dot.is-active {
      width: 22px;
      background: #fff;
    }

    .spv3-new:hover {
      transform: translateY(-3px);
      box-shadow: 0 22px 48px rgba(15, 27, 53, .18);
    }

    .spv3-new:focus-visible,
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
      console.warn(`[Starphone Home V3.6] ${name} falló:`, error);
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
      console.warn('[Starphone Home V3.6] Error limpiando filtros:', error);
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

  function applyNewFilter() {
    resetFilters();

    try {
      activeFilters.keyword = CONFIG.newKeyword;
    } catch (error) {
      console.warn('[Starphone Home V3.6] No se pudo aplicar Nuevo:', error);
    }

    rerender();
    showStatus('Filtro activo: Novedades');

    window.setTimeout(scrollToProducts, 80);
  }

  function applyBrandFilter(name) {
    resetFilters();

    const normalized = normalizeBrandName(name);
    try {
      activeFilters.MARCA = normalized;
    } catch (error) {
      console.warn('[Starphone Home V3.6] No se pudo aplicar la marca:', error);
    }

    rerender();
    showStatus(`Marca seleccionada: ${normalized}`);

    window.setTimeout(scrollToProducts, 80);
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



  const carouselState = {
    items: [],
    index: 0,
    timer: null,
    signature: ''
  };

  function getNewProducts() {
    try {
      const list = Array.isArray(products) ? products : [];

      return list.filter((product) => {
        const newFlag = String(product?.['新到货'] || '').trim().toUpperCase();
        const type = String(product?.TIPO || '').trim().toUpperCase();
        const name = String(product?.PRODUCTO || '').trim().toUpperCase();

        return (
          newFlag === 'NUEVO' ||
          type === 'NUEVO' ||
          name.includes('NUEVO')
        );
      }).slice(0, CONFIG.carouselLimit);
    } catch (_) {
      return [];
    }
  }

  function productImage(product) {
    const raw = String(product?.Imagen_Path || product?.IMAGEN || product?.image || '').trim();
    return raw ? encodeURI(raw.replace(/\\/g, '/')) : '';
  }

  function productPrice(product) {
    const raw =
      product?.['Precio ( USD )'] ??
      product?.['Precio USD'] ??
      product?.PRECIO ??
      '';

    const value = parseFloat(raw);
    return Number.isFinite(value) ? `$${value.toFixed(value % 1 ? 2 : 0)} USD` : '';
  }

  function carouselSignature(items) {
    return items.map((item) =>
      `${item?.PRODUCTO || ''}|${item?.MARCA || ''}|${productImage(item)}`
    ).join('::');
  }

  function updateCarouselContent(force = false) {
    const root = document.getElementById(IDS.root);
    if (!root) return;

    const items = getNewProducts();
    const signature = carouselSignature(items);

    if (!force && signature === carouselState.signature) return;

    carouselState.items = items;
    carouselState.signature = signature;
    carouselState.index = 0;

    renderCarouselSlide(true);
    restartCarousel();
  }

  function renderCarouselSlide(immediate = false) {
    const root = document.getElementById(IDS.root);
    if (!root) return;

    const brandEl = root.querySelector('[data-spv3-new-brand]');
    const nameEl = root.querySelector('[data-spv3-new-name]');
    const metaEl = root.querySelector('[data-spv3-new-meta]');
    const imageEl = root.querySelector('[data-spv3-new-image]');
    const dotsEl = root.querySelector('[data-spv3-dots]');

    if (!brandEl || !nameEl || !metaEl || !imageEl || !dotsEl) return;

    const item = carouselState.items[carouselState.index];

    if (!item) {
      brandEl.textContent = CONFIG.newArrivals.kicker;
      nameEl.textContent = CONFIG.newArrivals.title.replace(/^🔥\s*/, '');
      metaEl.textContent = CONFIG.newArrivals.subtitle;
      imageEl.removeAttribute('src');
      imageEl.alt = '';
      imageEl.classList.remove('is-visible');
      dotsEl.innerHTML = '';
      return;
    }

    const apply = () => {
      brandEl.textContent = item.MARCA || 'NUEVO';
      nameEl.textContent = item.PRODUCTO || CONFIG.newArrivals.title;
      metaEl.textContent = productPrice(item) || CONFIG.newArrivals.subtitle;

      const image = productImage(item);
      if (image) {
        imageEl.src = image;
        imageEl.alt = item.PRODUCTO || '';
        imageEl.onload = () => imageEl.classList.add('is-visible');
        imageEl.onerror = () => imageEl.classList.remove('is-visible');
      } else {
        imageEl.removeAttribute('src');
        imageEl.alt = '';
      }

      dotsEl.innerHTML = carouselState.items.map((_, index) =>
        `<span class="spv3-dot ${index === carouselState.index ? 'is-active' : ''}"></span>`
      ).join('');
    };

    if (immediate) {
      imageEl.classList.remove('is-visible');
      apply();
      return;
    }

    imageEl.classList.remove('is-visible');
    window.setTimeout(apply, 180);
  }

  function nextCarouselSlide() {
    if (carouselState.items.length <= 1) return;

    carouselState.index =
      (carouselState.index + 1) % carouselState.items.length;

    renderCarouselSlide(false);
  }

  function restartCarousel() {
    if (carouselState.timer) {
      window.clearInterval(carouselState.timer);
      carouselState.timer = null;
    }

    if (carouselState.items.length > 1) {
      carouselState.timer = window.setInterval(
        nextCarouselSlide,
        CONFIG.carouselInterval
      );
    }
  }

  function setupCarousel(root) {
    if (!root) return;

    const banner = root.querySelector('.spv3-new');
    if (!banner) return;

    banner.addEventListener('mouseenter', () => {
      if (carouselState.timer) {
        window.clearInterval(carouselState.timer);
        carouselState.timer = null;
      }
    });

    banner.addEventListener('mouseleave', restartCarousel);

    updateCarouselContent(true);

    /*
     * Products are loaded asynchronously and also change when switching tabs.
     * Watch the product grid so the banner automatically rebuilds itself.
     */
    const grid = document.getElementById(CONFIG.productsGridId);
    if (grid && grid.dataset.spv3CarouselObserved !== '1') {
      grid.dataset.spv3CarouselObserved = '1';

      let refreshTimer = 0;
      const observer = new MutationObserver(() => {
        window.clearTimeout(refreshTimer);
        refreshTimer = window.setTimeout(updateCarouselContent, 80);
      });

      observer.observe(grid, {
        childList: true,
        subtree: false
      });
    }

    let tries = 0;
    const waitForProducts = window.setInterval(() => {
      tries += 1;
      updateCarouselContent();

      if (getNewProducts().length || tries >= 30) {
        window.clearInterval(waitForProducts);
      }
    }, 400);
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
    root.setAttribute('aria-label', 'Novedades y pabellón de marcas');

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
      <button
        class="spv3-new"
        type="button"
        aria-label="${CONFIG.newArrivals.button}"
        style="--spv3-new-image: url('${CONFIG.newArrivals.image}')"
      >
        <span class="spv3-new-copy">
          <span class="spv3-kicker">${CONFIG.newArrivals.kicker}</span>
          <span class="spv3-new-brand" data-spv3-new-brand>${CONFIG.newArrivals.kicker}</span>
          <span class="spv3-new-name" data-spv3-new-name>Novedades de la Semana</span>
          <span class="spv3-new-meta" data-spv3-new-meta>${CONFIG.newArrivals.subtitle}</span>
          <span class="spv3-link">
            ${CONFIG.newArrivals.button}
            <span aria-hidden="true">→</span>
          </span>
        </span>

        <span class="spv3-new-product" aria-hidden="true">
          <img data-spv3-new-image alt="">
        </span>

        <span class="spv3-dots" data-spv3-dots aria-hidden="true"></span>
      </button>

      <header class="spv3-heading">
        <h2>Pabellón de Marcas</h2>
        <p>Explora nuestras marcas destacadas</p>
      </header>

      <div class="spv3-grid">
        ${brandsHtml}
      </div>
    `;

    root.querySelector('.spv3-new').addEventListener('click', applyNewFilter);

    root.querySelectorAll('.spv3-brand-card').forEach((card) => {
      card.addEventListener('click', () => {
        applyBrandFilter(card.dataset.brand);
      });
    });

    setupPointerGlow(root);
    setupCarousel(root);

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
