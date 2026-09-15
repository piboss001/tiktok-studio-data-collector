// ==UserScript==
// @name         TikTok Studio 数据采集器
// @namespace    qualitell.tiktok.collector
// @version      0.3.3
// @description  TikTok Studio 自动显示核心视频指标，支持周报分析、美化 XLSX、视频封面、可点击视频链接。
// @author       Qualitell
// @homepageURL  https://github.com/piboss001/tiktok-studio-data-collector
// @supportURL   https://github.com/piboss001/tiktok-studio-data-collector/issues
// @match        https://www.tiktok.com/tiktokstudio
// @match        https://www.tiktok.com/tiktokstudio/*
// @run-at       document-start
// @grant        unsafeWindow
// @grant        GM_xmlhttpRequest
// @connect      *
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @updateURL    https://raw.githubusercontent.com/piboss001/tiktok-studio-data-collector/main/TikTok-Studio-Data-Collector.user.js
// @downloadURL  https://raw.githubusercontent.com/piboss001/tiktok-studio-data-collector/main/TikTok-Studio-Data-Collector.user.js
// ==/UserScript==

(() => {
  'use strict';

  const page =
    typeof unsafeWindow !== 'undefined'
      ? unsafeWindow
      : window;

  const VERSION = '0.3.3';

  const items = new Map();
  const detailRows = new Map();
  const privateIds = new Set();
  const unknownIds = new Set();

  const autoMetricCache = new Map();
  const autoMetricLoading = new Set();
  const coverCache = new Map();

  let failedCount = 0;
  let busy = false;
  let minimized = false;

  let autoTimer = null;
  let observer = null;

  const originalFetch =
    page.fetch.bind(page);

  const OriginalXHR =
    page.XMLHttpRequest;

  const sleep = ms =>
    new Promise(resolve =>
      setTimeout(resolve, ms)
    );

  const $ = id =>
    document.getElementById(id);

  const COLORS = {
    navy: 'FF0F172A',
    navy2: 'FF1E293B',
    blue: 'FF2563EB',
    blueLight: 'FFDBEAFE',
    cyan: 'FF06B6D4',
    green: 'FF16A34A',
    greenLight: 'FFDCFCE7',
    amber: 'FFD97706',
    amberLight: 'FFFEF3C7',
    red: 'FFDC2626',
    redLight: 'FFFEE2E2',
    slate50: 'FFF8FAFC',
    slate100: 'FFF1F5F9',
    slate200: 'FFE2E8F0',
    slate500: 'FF64748B',
    slate700: 'FF334155',
    white: 'FFFFFFFF'
  };

  /* =========================
     基础
  ========================= */

  function num(value) {
    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {
      return '';
    }

    const n = Number(value);

    return Number.isNaN(n)
      ? ''
      : n;
  }

  function pct(value, digits = 1) {
    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {
      return '-';
    }

    const n = Number(value);

    if (!Number.isFinite(n)) {
      return '-';
    }

    return `${(n * 100).toFixed(digits)}%`;
  }

  function formatDate(timestamp) {
    if (!timestamp) {
      return '';
    }

    return new Date(
      Number(timestamp) * 1000
    ).toLocaleString('zh-CN');
  }

  function safeClone(value) {
    try {
      return JSON.parse(
        JSON.stringify(value)
      );
    } catch {
      return null;
    }
  }

  function csvEscape(value) {
    if (
      value === null ||
      value === undefined
    ) {
      return '';
    }

    const str =
      String(value);

    return /[",\n\r]/.test(str)
      ? `"${str.replace(/"/g, '""')}"`
      : str;
  }

  function mean(values) {
    const arr =
      values.filter(
        v =>
          typeof v === 'number' &&
          Number.isFinite(v)
      );

    if (!arr.length) {
      return '';
    }

    return (
      arr.reduce(
        (a, b) => a + b,
        0
      ) / arr.length
    );
  }

  function median(values) {
    const arr =
      values
        .filter(
          v =>
            typeof v === 'number' &&
            Number.isFinite(v)
        )
        .sort(
          (a, b) => a - b
        );

    if (!arr.length) {
      return '';
    }

    const mid =
      Math.floor(
        arr.length / 2
      );

    return (
      arr.length % 2
        ? arr[mid]
        : (
            arr[mid - 1] +
            arr[mid]
          ) / 2
    );
  }

  function quantile(values, q) {
    const arr =
      values
        .filter(
          v =>
            typeof v === 'number' &&
            Number.isFinite(v)
        )
        .sort(
          (a, b) => a - b
        );

    if (!arr.length) {
      return '';
    }

    if (arr.length === 1) {
      return arr[0];
    }

    const pos =
      (arr.length - 1) * q;

    const base =
      Math.floor(pos);

    const rest =
      pos - base;

    return (
      arr[base + 1] !== undefined
        ? (
            arr[base] +
            rest *
            (
              arr[base + 1] -
              arr[base]
            )
          )
        : arr[base]
    );
  }

  function normalizeText(str) {
    return String(str || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  /* =========================
     预约判断
  ========================= */

  function isScheduled(item) {
    const now =
      Math.floor(
        Date.now() / 1000
      );

    const postTime =
      Number(
        item?.post_time || 0
      );

    const scheduleTime =
      Number(
        item?.schedule_time || 0
      );

    return (
      postTime > now + 60 ||
      scheduleTime > now + 60
    );
  }

  function candidateItems() {
    return [
      ...items.values()
    ].filter(
      item =>
        !isScheduled(item)
    );
  }

  function scheduledItems() {
    return [
      ...items.values()
    ].filter(
      isScheduled
    );
  }

  /* =========================
     捕获作品列表
  ========================= */

  function captureItemList(url, data) {
    if (
      !url ||
      !String(url).includes(
        '/tiktok/creator/manage/item_list/v1/'
      )
    ) {
      return;
    }

    if (
      !data ||
      !Array.isArray(data.item_list)
    ) {
      return;
    }

    for (
      const item of data.item_list
    ) {
      if (!item?.item_id) {
        continue;
      }

      items.set(
        String(item.item_id),
        safeClone(item)
      );
    }

    updateUI(
      data.has_more
        ? `已发现 ${items.size} 条，继续加载…`
        : `已发现 ${items.size} 条`
    );

    scheduleAutoMetrics();
  }

  /* =========================
     Hook fetch
  ========================= */

  page.fetch =
    async function (...args) {
      const response =
        await originalFetch(
          ...args
        );

      try {
        const input =
          args[0];

        const url =
          typeof input === 'string'
            ? input
            : input?.url;

        if (
          url &&
          String(url).includes(
            '/tiktok/creator/manage/item_list/v1/'
          )
        ) {
          response
            .clone()
            .json()
            .then(
              data =>
                captureItemList(
                  url,
                  data
                )
            )
            .catch(() => {});
        }

      } catch {}

      return response;
    };

  /* =========================
     Hook XHR
  ========================= */

  if (
    OriginalXHR?.prototype
  ) {
    const originalOpen =
      OriginalXHR.prototype.open;

    const originalSend =
      OriginalXHR.prototype.send;

    OriginalXHR.prototype.open =
      function (
        method,
        url,
        ...rest
      ) {
        this.__qtkUrl =
          url;

        return originalOpen.call(
          this,
          method,
          url,
          ...rest
        );
      };

    OriginalXHR.prototype.send =
      function (...args) {
        this.addEventListener(
          'load',
          () => {
            try {
              const url =
                this.__qtkUrl;

              if (
                url &&
                String(url).includes(
                  '/tiktok/creator/manage/item_list/v1/'
                )
              ) {
                captureItemList(
                  url,
                  JSON.parse(
                    this.responseText
                  )
                );
              }

            } catch {}
          }
        );

        return originalSend.apply(
          this,
          args
        );
      };
  }

  /* =========================
     CSRF
  ========================= */

  function getCsrfToken() {
    const entry =
      document.cookie
        .split('; ')
        .find(
          cookie =>
            cookie.startsWith(
              'tt_csrf_token='
            )
        );

    if (!entry) {
      return '';
    }

    const value =
      entry
        .split('=')
        .slice(1)
        .join('=');

    try {
      return decodeURIComponent(value);
    } catch {
      return value;
    }
  }

  /* =========================
     Insight API
  ========================= */

  async function fetchInsight(videoId) {
    const typeRequests = [
      {
        insigh_type:
          'video_info',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_traffic_source_percent_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_retention_rate_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_view_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_total_duration_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_per_duration_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_finish_rate_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_new_follower_realtime',
        aweme_id:
          videoId
      }
    ];

    const lang =
      document.documentElement.lang ||
      'zh-Hans';

    const timezone =
      Intl
        .DateTimeFormat()
        .resolvedOptions()
        .timeZone ||
      'Asia/Shanghai';

    const tzOffset =
      -new Date()
        .getTimezoneOffset() *
      60;

    const params =
      new URLSearchParams({
        locale:
          lang,

        aid:
          '1988',

        priority_region:
          'US',

        region:
          'US',

        tz_name:
          timezone,

        app_name:
          'tiktok_creator_center',

        app_language:
          lang,

        device_platform:
          'web_pc',

        channel:
          'tiktok_web',

        tz_offset:
          String(tzOffset),

        type_requests:
          JSON.stringify(
            typeRequests
          )
      });

    const headers = {
      accept:
        'application/json, text/plain, */*'
    };

    const csrf =
      getCsrfToken();

    if (csrf) {
      headers['tt-csrf-token'] =
        csrf;
    }

    const response =
      await originalFetch(
        '/aweme/v2/data/insight/?' +
        params.toString(),
        {
          credentials:
            'include',

          headers
        }
      );

    return {
      httpStatus:
        response.status,

      data:
        await response.json()
    };
  }

  /* =========================
     数据整理
  ========================= */

  function retentionAt(
    list,
    milliseconds
  ) {
    const hit =
      (list || []).find(
        item =>
          String(
            item.timestamp
          ) ===
          String(milliseconds)
      );

    return hit
      ? Number(hit.value)
      : '';
  }

  function trafficToObject(list) {
    const result = {};

    for (
      const item of list || []
    ) {
      result[item.key] =
        Number(item.value);
    }

    return result;
  }

  function extractUrl(value) {
    if (!value) {
      return '';
    }

    if (
      typeof value === 'string'
    ) {
      return value;
    }

    if (
      Array.isArray(value)
    ) {
      return (
        value.find(
          item =>
            typeof item === 'string'
        ) ||
        ''
      );
    }

    if (
      Array.isArray(
        value.url_list
      )
    ) {
      return (
        value.url_list[0] ||
        ''
      );
    }

    if (
      typeof value.url === 'string'
    ) {
      return value.url;
    }

    return '';
  }

  function pickCoverUrl(
    item,
    videoInfo
  ) {
    const candidates = [
      item?.cover_url,
      item?.cover,
      item?.video?.cover,
      item?.video?.origin_cover,
      item?.video?.dynamic_cover,

      videoInfo?.cover,
      videoInfo?.video?.cover,
      videoInfo?.video?.origin_cover,
      videoInfo?.video?.dynamic_cover
    ];

    for (
      const candidate of candidates
    ) {
      const url =
        extractUrl(candidate);

      if (url) {
        return url;
      }
    }

    return '';
  }

  function getVideoUrl(row) {
    const account =
      String(
        row?.account || ''
      ).replace(/^@/, '');

    const videoId =
      String(
        row?.video_id || ''
      );

    if (
      !account ||
      !videoId
    ) {
      return '';
    }

    return (
      `https://www.tiktok.com/@${account}/video/${videoId}`
    );
  }

  function normalizeRow(
    item,
    insight
  ) {
    const data =
      insight.data || {};

    const videoInfo =
      data.video_info || {};

    const statistics =
      videoInfo.statistics || {};

    const retention =
      data
        .video_retention_rate_realtime
        ?.value
        ?.list ||
      [];

    const traffic =
      trafficToObject(
        data
          .video_traffic_source_percent_realtime
          ?.value
          ?.value ||
        []
      );

    const durationMs =
      Number(
        videoInfo.video?.duration ||
        item.duration ||
        0
      );

    const durationSec =
      durationMs
        ? durationMs / 1000
        : '';

    const avgWatch =
      num(
        data
          .video_per_duration_realtime
          ?.value
          ?.value
      );

    const views =
      num(
        data
          .realtime_total_video_views
          ?.value
          ?.value ??
        item.play_count ??
        statistics.play_count
      );

    const likes =
      num(
        statistics.digg_count ??
        item.like_count
      );

    const comments =
      num(
        statistics.comment_count ??
        item.comment_count
      );

    const shares =
      num(
        statistics.share_count ??
        item.share_count
      );

    const favorites =
      num(
        statistics.collect_count ??
        item.favorite_count
      );

    const publishTs =
      Number(
        item.post_time ||
        videoInfo.create_time ||
        0
      );

    const account =
      videoInfo.author
        ?.unique_id ||
      item.author
        ?.unique_id ||
      item.author_unique_id ||
      '';

    return {
      account,

      nickname:
        videoInfo.author
          ?.nickname ||
        item.author
          ?.nickname ||
        '',

      video_id:
        String(
          item.item_id ||
          videoInfo.aweme_id ||
          ''
        ),

      title:
        item.desc ||
        videoInfo.desc ||
        '',

      cover_url:
        pickCoverUrl(
          item,
          videoInfo
        ),

      publish_time:
        formatDate(
          publishTs
        ),

      publish_ts:
        publishTs,

      duration_sec:
        durationSec === ''
          ? ''
          : Number(
              durationSec.toFixed(3)
            ),

      views,
      likes,
      comments,
      shares,
      favorites,

      new_followers:
        num(
          data
            .realtime_new_followers
            ?.value
            ?.value
        ),

      avg_watch_sec:
        avgWatch,

      watch_ratio:
        (
          avgWatch !== '' &&
          durationSec
        )
          ? avgWatch / durationSec
          : '',

      finish_rate:
        num(
          data
            .video_finish_rate_realtime
            ?.value
            ?.value
        ),

      total_watch_sec:
        num(
          data
            .video_total_duration_realtime
            ?.value
            ?.value
        ),

      retention_1s:
        retentionAt(
          retention,
          1000
        ),

      retention_2s:
        retentionAt(
          retention,
          2000
        ),

      retention_3s:
        retentionAt(
          retention,
          3000
        ),

      retention_5s:
        retentionAt(
          retention,
          5000
        ),

      retention_10s:
        retentionAt(
          retention,
          10000
        ),

      for_you:
        traffic[
          'For You'
        ] ?? '',

      personal_profile:
        traffic[
          'Personal Profile'
        ] ?? '',

      search:
        traffic.Search ??
        '',

      follow:
        traffic.Follow ??
        '',

      direct_message:
        traffic[
          'Direct Message'
        ] ?? '',

      sound:
        traffic.Sound ??
        '',

      others:
        traffic.Others ??
        '',

      like_rate:
        views
          ? likes / views
          : '',

      engagement_rate:
        views
          ? (
              likes +
              comments +
              shares +
              favorites
            ) / views
          : ''
    };
  }

  /* ==========================================================
     V0.3.3
     TikTok Studio 列表自动显示核心指标
  ========================================================== */

  function climbToVideoRow(element) {
    let current =
      element;

    let fallback =
      null;

    for (
      let i = 0;
      i < 10 &&
      current &&
      current !== document.body;
      i++
    ) {
      if (
        current.id ===
        'qualitell-tiktok-panel'
      ) {
        return null;
      }

      const rect =
        current.getBoundingClientRect();

      const buttons =
        current.querySelectorAll
          ? current.querySelectorAll(
              'button'
            ).length
          : 0;

      const links =
        current.querySelectorAll
          ? current.querySelectorAll(
              'a'
            ).length
          : 0;

      if (
        rect.width > 650 &&
        rect.height >= 55 &&
        rect.height <= 180
      ) {
        fallback =
          current;

        if (
          buttons >= 2 ||
          links >= 2 ||
          current.getAttribute(
            'role'
          ) === 'row' ||
          current.tagName === 'TR'
        ) {
          return current;
        }
      }

      current =
        current.parentElement;
    }

    return fallback;
  }

  function findRowByVideoId(
    videoId
  ) {
    const selectors = [
      `a[href*="${videoId}"]`,
      `[data-video-id="${videoId}"]`,
      `[data-item-id="${videoId}"]`,
      `[data-id="${videoId}"]`
    ];

    for (
      const selector of selectors
    ) {
      try {
        const element =
          document.querySelector(
            selector
          );

        if (element) {
          const row =
            climbToVideoRow(
              element
            );

          if (row) {
            return row;
          }
        }
      } catch {}
    }

    return null;
  }

  function findRowByTitle(
    item
  ) {
    const title =
      normalizeText(
        item?.desc
      );

    if (!title) {
      return null;
    }

    const key =
      title.slice(
        0,
        Math.min(
          22,
          title.length
        )
      );

    const walker =
      document.createTreeWalker(
        document.body,
        NodeFilter.SHOW_TEXT
      );

    let node;

    while (
      (
        node =
          walker.nextNode()
      )
    ) {
      const parent =
        node.parentElement;

      if (
        !parent ||
        parent.closest(
          '#qualitell-tiktok-panel'
        ) ||
        parent.closest(
          '.qtk-row-metrics'
        )
      ) {
        continue;
      }

      const text =
        normalizeText(
          node.nodeValue
        );

      if (
        text.length >= 6 &&
        (
          text.includes(key) ||
          key.includes(
            text.slice(
              0,
              15
            )
          )
        )
      ) {
        const row =
          climbToVideoRow(
            parent
          );

        if (row) {
          return row;
        }
      }
    }

    return null;
  }

  function findVideoRow(
    item
  ) {
    if (!item?.item_id) {
      return null;
    }

    return (
      findRowByVideoId(
        String(item.item_id)
      ) ||
      findRowByTitle(
        item
      )
    );
  }

  function buildMetricBox(
    videoId
  ) {
    const box =
      document.createElement(
        'div'
      );

    box.className =
      'qtk-row-metrics';

    box.dataset.videoId =
      videoId;

    box.innerHTML = `
      <div class="qtk-metric-loading">
        指标加载中…
      </div>
    `;

    return box;
  }

  function renderRowMetrics(
    box,
    row
  ) {
    if (
      !box ||
      !row
    ) {
      return;
    }

    box.classList.remove(
      'qtk-metric-error'
    );

    box.innerHTML = `
      <div
        class="qtk-mini-metric"
        title="完播率"
      >
        <span>完播</span>
        <b>${pct(row.finish_rate)}</b>
      </div>

      <div
        class="qtk-mini-metric"
        title="观看倍率 = 平均观看时长 ÷ 视频时长"
      >
        <span>倍率</span>
        <b>${pct(row.watch_ratio)}</b>
      </div>

      <div
        class="qtk-mini-metric"
        title="1秒留存"
      >
        <span>1s</span>
        <b>${pct(row.retention_1s)}</b>
      </div>

      <div
        class="qtk-mini-metric"
        title="3秒留存"
      >
        <span>3s</span>
        <b>${pct(row.retention_3s)}</b>
      </div>

      <div
        class="qtk-mini-metric"
        title="新增粉丝"
      >
        <span>+粉</span>
        <b>${row.new_followers === '' ? '-' : row.new_followers}</b>
      </div>

      <div
        class="qtk-mini-metric"
        title="互动率 = 点赞+评论+分享+收藏 ÷ 播放量"
      >
        <span>互动</span>
        <b>${pct(row.engagement_rate)}</b>
      </div>
    `;
  }

  function renderPrivateMetric(
    box,
    text
  ) {
    if (!box) {
      return;
    }

    box.classList.add(
      'qtk-metric-error'
    );

    box.innerHTML = `
      <div class="qtk-metric-message">
        ${text}
      </div>
    `;
  }

  function attachMetricBox(
    item
  ) {
    const videoId =
      String(
        item.item_id
      );

    const rowElement =
      findVideoRow(
        item
      );

    if (!rowElement) {
      return null;
    }

    rowElement.dataset.qtkItem =
      videoId;

    let box =
      rowElement.querySelector(
        `:scope > .qtk-row-metrics[data-video-id="${videoId}"]`
      );

    if (!box) {
      box =
        buildMetricBox(
          videoId
        );

      const currentPosition =
        getComputedStyle(
          rowElement
        ).position;

      if (
        currentPosition ===
        'static'
      ) {
        rowElement.style.position =
          'relative';
      }

      rowElement.appendChild(
        box
      );
    }

    const cached =
      autoMetricCache.get(
        videoId
      );

    if (
      cached?.type ===
      'public'
    ) {
      renderRowMetrics(
        box,
        cached.row
      );
    }

    if (
      cached?.type ===
      'private'
    ) {
      renderPrivateMetric(
        box,
        '非公开'
      );
    }

    if (
      cached?.type ===
      'scheduled'
    ) {
      renderPrivateMetric(
        box,
        '预约发布'
      );
    }

    if (
      cached?.type ===
      'error'
    ) {
      renderPrivateMetric(
        box,
        '暂不可用'
      );
    }

    return box;
  }

  async function loadAutoMetric(
    item
  ) {
    const videoId =
      String(
        item.item_id
      );

    if (
      autoMetricCache.has(
        videoId
      ) ||
      autoMetricLoading.has(
        videoId
      )
    ) {
      return;
    }

    const box =
      attachMetricBox(
        item
      );

    if (!box) {
      return;
    }

    if (
      isScheduled(item)
    ) {
      autoMetricCache.set(
        videoId,
        {
          type:
            'scheduled'
        }
      );

      renderPrivateMetric(
        box,
        '预约发布'
      );

      return;
    }

    autoMetricLoading.add(
      videoId
    );

    try {
      const insight =
        await fetchInsight(
          videoId
        );

      if (
        insight.httpStatus !== 200 ||
        insight.data
          ?.status_code !== 0
      ) {
        throw new Error(
          'Insight failed'
        );
      }

      const privateStatus =
        insight.data
          ?.video_info
          ?.status
          ?.private_status;

      if (
        privateStatus === 0
      ) {
        const normalized =
          normalizeRow(
            item,
            insight
          );

        autoMetricCache.set(
          videoId,
          {
            type:
              'public',

            row:
              normalized
          }
        );

        detailRows.set(
          videoId,
          normalized
        );

        renderRowMetrics(
          box,
          normalized
        );

      } else {
        autoMetricCache.set(
          videoId,
          {
            type:
              'private'
          }
        );

        privateIds.add(
          videoId
        );

        renderPrivateMetric(
          box,
          '非公开'
        );
      }

    } catch (error) {
      console.warn(
        '[TikTok Collector Auto]',
        videoId,
        error
      );

      autoMetricCache.set(
        videoId,
        {
          type:
            'error'
        }
      );

      renderPrivateMetric(
        box,
        '暂不可用'
      );

    } finally {
      autoMetricLoading.delete(
        videoId
      );

      updateUI();
    }
  }

  async function runAutoMetrics() {
    if (busy) {
      scheduleAutoMetrics(
        1000
      );

      return;
    }

    const visible = [];

    for (
      const item of items.values()
    ) {
      const row =
        findVideoRow(
          item
        );

      if (!row) {
        continue;
      }

      attachMetricBox(
        item
      );

      const rect =
        row.getBoundingClientRect();

      if (
        rect.bottom >= -150 &&
        rect.top <=
        window.innerHeight +
        400
      ) {
        visible.push(
          item
        );
      }
    }

    for (
      const item of visible
    ) {
      await loadAutoMetric(
        item
      );

      await sleep(
        350
      );
    }
  }

  function scheduleAutoMetrics(
    delay = 350
  ) {
    clearTimeout(
      autoTimer
    );

    autoTimer =
      setTimeout(
        runAutoMetrics,
        delay
      );
  }

  function startDomObserver() {
    if (
      observer ||
      !document.body
    ) {
      return;
    }

    observer =
      new MutationObserver(
        mutations => {
          let relevant =
            false;

          for (
            const mutation of mutations
          ) {
            if (
              mutation.addedNodes
                ?.length
            ) {
              relevant =
                true;

              break;
            }
          }

          if (relevant) {
            scheduleAutoMetrics(
              450
            );
          }
        }
      );

    observer.observe(
      document.body,
      {
        childList:
          true,

        subtree:
          true
      }
    );

    window.addEventListener(
      'scroll',
      () =>
        scheduleAutoMetrics(
          250
        ),
      {
        passive:
          true
      }
    );
  }

  /* =========================
     分析
  ========================= */

  function getRangeDays() {
    const value =
      $('qtk-range')
        ?.value ||
      '7';

    return (
      value === 'all'
        ? null
        : Number(value)
    );
  }

  function getAnalysisRows() {
    const rows = [
      ...detailRows.values()
    ];

    const days =
      getRangeDays();

    if (!days) {
      return rows.sort(
        (a, b) =>
          b.publish_ts -
          a.publish_ts
      );
    }

    const cutoff =
      Math.floor(
        Date.now() / 1000
      ) -
      days *
      86400;

    return rows
      .filter(
        row =>
          row.publish_ts >=
          cutoff
      )
      .sort(
        (a, b) =>
          b.publish_ts -
          a.publish_ts
      );
  }

  function summaryFor(rows) {
    return {
      count:
        rows.length,

      avgViews:
        mean(
          rows.map(
            r => r.views
          )
        ),

      medianViews:
        median(
          rows.map(
            r => r.views
          )
        ),

      avgFinish:
        mean(
          rows.map(
            r =>
              r.finish_rate
          )
        ),

      avgWatchRatio:
        mean(
          rows.map(
            r =>
              r.watch_ratio
          )
        ),

      avgR1:
        mean(
          rows.map(
            r =>
              r.retention_1s
          )
        ),

      avgR2:
        mean(
          rows.map(
            r =>
              r.retention_2s
          )
        ),

      avgR3:
        mean(
          rows.map(
            r =>
              r.retention_3s
          )
        ),

      avgR5:
        mean(
          rows.map(
            r =>
              r.retention_5s
          )
        )
    };
  }

  function topRows(
    rows,
    metric,
    limit = 5
  ) {
    return rows
      .filter(
        row =>
          Number.isFinite(
            Number(
              row?.[metric]
            )
          )
      )
      .sort(
        (a, b) =>
          Number(
            b[metric]
          ) -
          Number(
            a[metric]
          )
      )
      .slice(
        0,
        limit
      );
  }

  function metricLabel(metric) {
    const map = {
      views:
        '播放量',

      finish_rate:
        '完播率',

      watch_ratio:
        '观看倍率',

      retention_1s:
        '1秒留存',

      retention_3s:
        '3秒留存',

      new_followers:
        '新增粉丝',

      engagement_rate:
        '互动率'
    };

    return (
      map[metric] ||
      metric
    );
  }

  function metricDisplay(
    metric,
    value
  ) {
    if (
      !Number.isFinite(
        Number(value)
      )
    ) {
      return '-';
    }

    if (
      [
        'finish_rate',
        'watch_ratio',
        'retention_1s',
        'retention_3s',
        'engagement_rate'
      ].includes(metric)
    ) {
      return pct(
        Number(value)
      );
    }

    return String(value);
  }

  function buildSignals(rows) {
    if (
      rows.length < 3
    ) {
      return {
        high: [],
        hidden: [],
        weak: []
      };
    }

    const views =
      rows
        .map(
          r => r.views
        )
        .filter(
          Number.isFinite
        );

    const r3 =
      rows
        .map(
          r =>
            r.retention_3s
        )
        .filter(
          Number.isFinite
        );

    const q80Views =
      quantile(
        views,
        0.8
      );

    const medViews =
      median(
        views
      );

    const q75R3 =
      quantile(
        r3,
        0.75
      );

    const q25R3 =
      quantile(
        r3,
        0.25
      );

    return {
      high:
        rows
          .filter(
            r =>
              Number.isFinite(
                r.views
              ) &&
              r.views >=
              q80Views
          )
          .sort(
            (a, b) =>
              b.views -
              a.views
          )
          .slice(
            0,
            4
          ),

      hidden:
        rows
          .filter(
            r =>
              Number.isFinite(
                r.retention_3s
              ) &&
              Number.isFinite(
                r.views
              ) &&
              r.retention_3s >=
              q75R3 &&
              r.views <=
              medViews
          )
          .sort(
            (a, b) =>
              b.retention_3s -
              a.retention_3s
          )
          .slice(
            0,
            4
          ),

      weak:
        rows
          .filter(
            r =>
              Number.isFinite(
                r.retention_3s
              ) &&
              r.retention_3s <=
              q25R3
          )
          .sort(
            (a, b) =>
              a.retention_3s -
              b.retention_3s
          )
          .slice(
            0,
            4
          )
    };
  }

  function shortTitle(
    title,
    max = 34
  ) {
    const clean =
      String(
        title || ''
      )
        .replace(
          /\s+/g,
          ' '
        )
        .trim();

    return (
      clean.length > max
        ? (
            clean.slice(
              0,
              max
            ) +
            '…'
          )
        : clean
    );
  }

  function escapeHtml(str) {
    return String(
      str ?? ''
    )
      .replace(
        /&/g,
        '&amp;'
      )
      .replace(
        /</g,
        '&lt;'
      )
      .replace(
        />/g,
        '&gt;'
      )
      .replace(
        /"/g,
        '&quot;'
      )
      .replace(
        /'/g,
        '&#039;'
      );
  }

  function renderTopList(
    rows,
    metric
  ) {
    const list =
      topRows(
        rows,
        metric,
        5
      );

    if (!list.length) {
      return `
        <div class="qtk-empty">
          暂无数据
        </div>
      `;
    }

    return list
      .map(
        (r, i) => `
          <div class="qtk-top-item">

            <div class="qtk-rank">
              ${i + 1}
            </div>

            <div class="qtk-top-main">

              <div
                class="qtk-top-title"
                title="${escapeHtml(r.title)}"
              >
                ${escapeHtml(shortTitle(r.title))}
              </div>

              <div class="qtk-top-meta">
                播放 ${r.views || 0}
                · 完播 ${pct(r.finish_rate)}
                · 3s ${pct(r.retention_3s)}
              </div>

            </div>

            <div class="qtk-top-value">
              ${metricDisplay(metric, r[metric])}
            </div>

          </div>
        `
      )
      .join('');
  }

  function renderAnalysis() {
    if (
      !$(
        'qtk-analysis'
      )
    ) {
      return;
    }

    const rows =
      getAnalysisRows();

    const summary =
      summaryFor(
        rows
      );

    $('qtk-summary-count')
      .textContent =
      summary.count || 0;

    $('qtk-summary-avg')
      .textContent =
      summary.avgViews === ''
        ? '-'
        : Math.round(
            summary.avgViews
          );

    $('qtk-summary-median')
      .textContent =
      summary.medianViews === ''
        ? '-'
        : Math.round(
            summary.medianViews
          );

    $('qtk-summary-finish')
      .textContent =
      pct(
        summary.avgFinish
      );

    $('qtk-summary-watch')
      .textContent =
      pct(
        summary.avgWatchRatio
      );

    $('qtk-summary-r3')
      .textContent =
      pct(
        summary.avgR3
      );

    const metric =
      $('qtk-rank-metric')
        ?.value ||
      'views';

    $('qtk-top-list')
      .innerHTML =
      renderTopList(
        rows,
        metric
      );
  }

  /* =========================
     UI CSS
  ========================= */

  function injectStyle() {
    if (
      document.getElementById(
        'qualitell-tiktok-style'
      )
    ) {
      return;
    }

    const style =
      document.createElement(
        'style'
      );

    style.id =
      'qualitell-tiktok-style';

    style.textContent = `

      /* =====================
         Studio 列表核心指标
      ===================== */

      .qtk-row-metrics {
        position:absolute !important;
        right:142px !important;
        top:50% !important;
        transform:translateY(-50%) !important;

        width:132px !important;
        min-width:132px !important;

        display:grid !important;
        grid-template-columns:repeat(2,1fr) !important;
        gap:3px 5px !important;

        padding:5px 7px !important;

        background:#f8fafc !important;
        border:1px solid #e5e7eb !important;
        border-radius:7px !important;

        box-shadow:
          0 1px 3px rgba(0,0,0,.04) !important;

        z-index:10 !important;

        pointer-events:auto !important;

        font-family:
          Arial,
          "Microsoft YaHei",
          sans-serif !important;
      }

      .qtk-mini-metric {
        display:flex !important;
        align-items:center !important;
        justify-content:space-between !important;
        gap:3px !important;

        min-width:0 !important;

        line-height:15px !important;

        white-space:nowrap !important;
      }

      .qtk-mini-metric span {
        font-size:9px !important;
        color:#8b8b95 !important;
        font-weight:400 !important;
      }

      .qtk-mini-metric b {
        font-size:9.5px !important;
        color:#161823 !important;
        font-weight:600 !important;

        overflow:hidden !important;
        text-overflow:ellipsis !important;
      }

      .qtk-metric-loading,
      .qtk-metric-message {
        grid-column:1 / 3 !important;

        text-align:center !important;

        font-size:9px !important;
        color:#8b8b95 !important;

        line-height:38px !important;
      }

      .qtk-metric-error {
        background:#fafafa !important;
      }

      @media(max-width:1400px) {

        .qtk-row-metrics {
          right:125px !important;
          width:118px !important;
          min-width:118px !important;
          padding:4px 5px !important;
          gap:2px 4px !important;
        }

        .qtk-mini-metric span {
          font-size:8px !important;
        }

        .qtk-mini-metric b {
          font-size:8.5px !important;
        }

      }


      /* =====================
         浮动控制面板
      ===================== */

      #qualitell-tiktok-panel,
      #qualitell-tiktok-panel * {
        box-sizing:border-box;
      }

      #qualitell-tiktok-panel {
        position:fixed;

        right:14px;
        bottom:14px;

        z-index:2147483647;

        width:332px;
        max-height:86vh;

        overflow:hidden;

        background:
          rgba(22,22,26,.985);

        color:#fff;

        border:
          1px solid
          rgba(255,255,255,.10);

        border-radius:12px;

        box-shadow:
          0 14px 40px
          rgba(0,0,0,.35);

        font-family:
          Arial,
          "Microsoft YaHei",
          sans-serif;
      }

      .qtk-header {
        display:flex;
        align-items:center;
        justify-content:space-between;

        padding:11px 13px;

        border-bottom:
          1px solid
          rgba(255,255,255,.08);
      }

      .qtk-title {
        font-size:14px;
        font-weight:700;
      }

      .qtk-version {
        margin-top:2px;

        font-size:10px;

        color:
          rgba(255,255,255,.45);
      }

      .qtk-minimize {
        border:0;
        background:transparent;
        color:#fff;

        font-size:20px;
        cursor:pointer;
      }

      #qtk-body {
        max-height:
          calc(86vh - 52px);

        overflow:auto;

        padding:
          11px 13px 14px;
      }

      .qtk-row {
        display:flex;
        align-items:center;
        justify-content:space-between;

        gap:10px;

        padding:3px 0;

        font-size:11px;
      }

      .qtk-row span {
        color:
          rgba(255,255,255,.58);
      }

      .qtk-row b {
        font-weight:600;
      }

      .qtk-toolbar {
        display:grid;
        grid-template-columns:
          1fr 1fr;

        gap:6px;

        margin-top:9px;
      }

      .qtk-toolbar button,
      .qtk-select {
        width:100%;

        min-height:31px;

        border:0;
        border-radius:7px;

        padding:7px 7px;

        font-size:11px;
      }

      .qtk-toolbar button {
        cursor:pointer;
        font-weight:600;
      }

      #qtk-scan {
        background:#fe2c55;
        color:#fff;
      }

      #qtk-fetch {
        background:#25f4ee;
        color:#111;
      }

      #qtk-export-xlsx,
      #qtk-export-csv {
        background:#fff;
        color:#111;
      }

      #qtk-reset {
        background:#393940;
        color:#fff;
      }

      .qtk-toolbar button:disabled {
        opacity:.35;
        cursor:not-allowed;
      }

      .qtk-progress {
        height:5px;

        margin-top:9px;

        overflow:hidden;

        border-radius:999px;

        background:
          rgba(255,255,255,.10);
      }

      #qtk-progress-bar {
        height:100%;
        width:0%;

        background:#25f4ee;

        transition:
          width .2s ease;
      }

      .qtk-divider {
        height:1px;

        margin:11px 0;

        background:
          rgba(255,255,255,.08);
      }

      .qtk-section-title {
        display:flex;
        align-items:center;
        justify-content:space-between;

        margin-bottom:7px;

        font-size:12px;
        font-weight:700;
      }

      .qtk-select {
        background:#303036;
        color:#fff;
        outline:none;
      }

      .qtk-summary-grid {
        display:grid;
        grid-template-columns:
          repeat(3,1fr);

        gap:5px;
      }

      .qtk-card {
        padding:7px;

        background:#2a2a30;

        border:
          1px solid
          rgba(255,255,255,.05);

        border-radius:7px;
      }

      .qtk-card span {
        display:block;

        margin-bottom:3px;

        font-size:9px;

        color:
          rgba(255,255,255,.45);
      }

      .qtk-card b {
        font-size:13px;
      }

      .qtk-rank-controls {
        margin:
          8px 0 5px;
      }

      .qtk-top-item {
        display:flex;
        align-items:center;

        gap:6px;

        padding:6px 0;

        border-bottom:
          1px solid
          rgba(255,255,255,.05);
      }

      .qtk-rank {
        display:flex;
        align-items:center;
        justify-content:center;

        flex:none;

        width:18px;
        height:18px;

        border-radius:5px;

        background:#34343b;

        font-size:9px;
        font-weight:700;
      }

      .qtk-top-main {
        flex:1;
        min-width:0;
      }

      .qtk-top-title {
        overflow:hidden;
        text-overflow:ellipsis;
        white-space:nowrap;

        font-size:10px;
      }

      .qtk-top-meta {
        margin-top:2px;

        font-size:8px;

        color:
          rgba(255,255,255,.40);
      }

      .qtk-top-value {
        flex:none;

        font-size:10px;
        font-weight:700;

        color:#25f4ee;
      }

      .qtk-empty {
        padding:6px 0;

        font-size:10px;

        color:
          rgba(255,255,255,.42);
      }

      .qtk-tip {
        margin-top:8px;

        font-size:9px;
        line-height:1.45;

        color:
          rgba(255,255,255,.36);
      }

    `;

    (
      document.head ||
      document.documentElement
    ).appendChild(
      style
    );
  }

  /* =========================
     UI
  ========================= */

  function buildPanel() {
    if (
      !document.documentElement ||
      document.getElementById(
        'qualitell-tiktok-panel'
      )
    ) {
      return;
    }

    injectStyle();

    const panel =
      document.createElement(
        'div'
      );

    panel.id =
      'qualitell-tiktok-panel';

    panel.innerHTML = `
      <div class="qtk-header">

        <div>
          <div class="qtk-title">
            TikTok 数据采集器
          </div>

          <div class="qtk-version">
            V${VERSION} · 自动指标
          </div>
        </div>

        <button
          id="qtk-minimize"
          class="qtk-minimize"
        >
          −
        </button>

      </div>

      <div id="qtk-body">

        <div class="qtk-row">
          <span>发现作品</span>
          <b id="qtk-all">0</b>
        </div>

        <div class="qtk-row">
          <span>预约</span>
          <b id="qtk-scheduled">0</b>
        </div>

        <div class="qtk-row">
          <span>公开数据</span>
          <b id="qtk-public">0</b>
        </div>

        <div class="qtk-row">
          <span>非公开</span>
          <b id="qtk-private">0</b>
        </div>

        <div class="qtk-row">
          <span>失败/未知</span>
          <b id="qtk-failed">0</b>
        </div>

        <div class="qtk-row">
          <span>状态</span>
          <b id="qtk-status">
            自动加载中
          </b>
        </div>

        <div class="qtk-toolbar">

          <button id="qtk-scan">
            扫描全部
          </button>

          <button
            id="qtk-fetch"
            disabled
          >
            获取全部数据
          </button>

          <button
            id="qtk-export-xlsx"
            disabled
          >
            导出 XLSX
          </button>

          <button
            id="qtk-export-csv"
            disabled
          >
            导出 CSV
          </button>

          <button
            id="qtk-reset"
            style="grid-column:1/3"
          >
            清空重扫
          </button>

        </div>

        <div class="qtk-progress">
          <div id="qtk-progress-bar"></div>
        </div>

        <div class="qtk-divider"></div>

        <div id="qtk-analysis">

          <div class="qtk-section-title">

            <span>
              数据分析
            </span>

            <select
              id="qtk-range"
              class="qtk-select"
              style="width:105px"
            >
              <option value="7">
                最近7天
              </option>

              <option value="14">
                最近14天
              </option>

              <option value="30">
                最近30天
              </option>

              <option value="all">
                全部
              </option>
            </select>

          </div>

          <div class="qtk-summary-grid">

            <div class="qtk-card">
              <span>公开视频</span>
              <b id="qtk-summary-count">0</b>
            </div>

            <div class="qtk-card">
              <span>平均播放</span>
              <b id="qtk-summary-avg">-</b>
            </div>

            <div class="qtk-card">
              <span>中位播放</span>
              <b id="qtk-summary-median">-</b>
            </div>

            <div class="qtk-card">
              <span>平均完播</span>
              <b id="qtk-summary-finish">-</b>
            </div>

            <div class="qtk-card">
              <span>观看倍率</span>
              <b id="qtk-summary-watch">-</b>
            </div>

            <div class="qtk-card">
              <span>3秒留存</span>
              <b id="qtk-summary-r3">-</b>
            </div>

          </div>

          <div class="qtk-rank-controls">

            <select
              id="qtk-rank-metric"
              class="qtk-select"
            >
              <option value="views">
                播放 TOP
              </option>

              <option value="finish_rate">
                完播率 TOP
              </option>

              <option value="watch_ratio">
                观看倍率 TOP
              </option>

              <option value="retention_1s">
                1秒留存 TOP
              </option>

              <option value="retention_3s">
                3秒留存 TOP
              </option>

              <option value="new_followers">
                新增粉丝 TOP
              </option>

              <option value="engagement_rate">
                互动率 TOP
              </option>
            </select>

          </div>

          <div id="qtk-top-list">

            <div class="qtk-empty">
              正在自动采集当前页面数据…
            </div>

          </div>

        </div>

        <div class="qtk-tip">
          刷新 TikTok Studio 后自动显示：
          完播率、观看倍率、1s/3s留存、新增粉丝和互动率。
        </div>

      </div>
    `;

    document
      .documentElement
      .appendChild(
        panel
      );

    $('qtk-minimize')
      .addEventListener(
        'click',
        () => {
          minimized =
            !minimized;

          $('qtk-body')
            .style.display =
            minimized
              ? 'none'
              : '';

          $('qtk-minimize')
            .textContent =
            minimized
              ? '+'
              : '−';
        }
      );

    $('qtk-scan')
      .addEventListener(
        'click',
        scanAll
      );

    $('qtk-fetch')
      .addEventListener(
        'click',
        fetchAllInsights
      );

    $('qtk-export-xlsx')
      .addEventListener(
        'click',
        exportXlsx
      );

    $('qtk-export-csv')
      .addEventListener(
        'click',
        exportCsv
      );

    $('qtk-reset')
      .addEventListener(
        'click',
        resetAll
      );

    $('qtk-range')
      .addEventListener(
        'change',
        renderAnalysis
      );

    $('qtk-rank-metric')
      .addEventListener(
        'change',
        renderAnalysis
      );

    updateUI(
      '自动加载中'
    );
  }

  function updateUI(status) {
    if (
      !$(
        'qualitell-tiktok-panel'
      )
    ) {
      return;
    }

    $('qtk-all')
      .textContent =
      items.size;

    $('qtk-scheduled')
      .textContent =
      scheduledItems().length;

    $('qtk-public')
      .textContent =
      detailRows.size;

    $('qtk-private')
      .textContent =
      privateIds.size;

    $('qtk-failed')
      .textContent =
      failedCount +
      unknownIds.size;

    if (status) {
      $('qtk-status')
        .textContent =
        status;
    }

    $('qtk-scan')
      .disabled =
      busy;

    $('qtk-fetch')
      .disabled =
      busy ||
      candidateItems()
        .length === 0;

    $('qtk-export-xlsx')
      .disabled =
      busy ||
      detailRows.size === 0;

    $('qtk-export-csv')
      .disabled =
      busy ||
      detailRows.size === 0;

    renderAnalysis();
  }

  /* =========================
     扫描
  ========================= */

  async function scanAll() {
    if (busy) {
      return;
    }

    busy = true;

    updateUI(
      '扫描中…'
    );

    let stableRounds = 0;

    let lastCount =
      items.size;

    let lastHeight =
      document
        .documentElement
        .scrollHeight;

    for (
      let i = 0;
      i < 100;
      i++
    ) {
      page.scrollTo({
        top:
          document
            .documentElement
            .scrollHeight,

        behavior:
          'smooth'
      });

      await sleep(
        850
      );

      const newCount =
        items.size;

      const newHeight =
        document
          .documentElement
          .scrollHeight;

      if (
        newCount ===
          lastCount &&
        newHeight ===
          lastHeight
      ) {
        stableRounds++;
      } else {
        stableRounds = 0;
      }

      lastCount =
        newCount;

      lastHeight =
        newHeight;

      updateUI(
        `扫描 ${newCount} 条`
      );

      if (
        stableRounds >= 6
      ) {
        break;
      }
    }

    busy = false;

    updateUI(
      `扫描完成 ${items.size} 条`
    );

    scheduleAutoMetrics();
  }

  /* =========================
     全量获取
  ========================= */

  async function fetchAllInsights() {
    if (
      busy ||
      !items.size
    ) {
      return;
    }

    busy = true;

    failedCount = 0;

    detailRows.clear();
    privateIds.clear();
    unknownIds.clear();

    const list =
      candidateItems();

    let completed = 0;

    $('qtk-progress-bar')
      .style.width =
      '0%';

    for (
      const item of list
    ) {
      const videoId =
        String(
          item.item_id
        );

      try {
        const insight =
          await fetchInsight(
            videoId
          );

        if (
          insight.httpStatus !== 200 ||
          insight.data
            ?.status_code !== 0
        ) {
          throw new Error(
            'Insight failed'
          );
        }

        const privateStatus =
          insight.data
            ?.video_info
            ?.status
            ?.private_status;

        if (
          privateStatus === 0
        ) {
          const row =
            normalizeRow(
              item,
              insight
            );

          detailRows.set(
            videoId,
            row
          );

          autoMetricCache.set(
            videoId,
            {
              type:
                'public',
              row
            }
          );

        } else {
          privateIds.add(
            videoId
          );

          autoMetricCache.set(
            videoId,
            {
              type:
                'private'
            }
          );
        }

      } catch {
        failedCount++;
      }

      completed++;

      $('qtk-progress-bar')
        .style.width =
        `${
          Math.round(
            completed /
            list.length *
            100
          )
        }%`;

      updateUI(
        `${completed}/${list.length}`
      );

      await sleep(
        700
      );
    }

    busy = false;

    updateUI(
      `完成 ${detailRows.size} 条`
    );

    scheduleAutoMetrics();
  }

  function resetAll() {
    if (busy) {
      return;
    }

    items.clear();
    detailRows.clear();
    privateIds.clear();
    unknownIds.clear();

    autoMetricCache.clear();
    autoMetricLoading.clear();
    coverCache.clear();

    failedCount = 0;

    document
      .querySelectorAll(
        '.qtk-row-metrics'
      )
      .forEach(
        element =>
          element.remove()
      );

    updateUI(
      '已清空'
    );
  }

  /* =========================
     CSV
  ========================= */

  const exportColumns = [
    ['账号', 'account'],
    ['昵称', 'nickname'],
    ['Video ID', 'video_id'],
    ['标题', 'title'],
    ['发布时间', 'publish_time'],
    ['视频时长(s)', 'duration_sec'],
    ['播放量', 'views'],
    ['点赞', 'likes'],
    ['评论', 'comments'],
    ['分享', 'shares'],
    ['收藏', 'favorites'],
    ['新增粉丝', 'new_followers'],
    ['平均观看时间(s)', 'avg_watch_sec'],
    ['观看倍率', 'watch_ratio'],
    ['完播率', 'finish_rate'],
    ['总观看时间(s)', 'total_watch_sec'],
    ['1秒留存', 'retention_1s'],
    ['2秒留存', 'retention_2s'],
    ['3秒留存', 'retention_3s'],
    ['5秒留存', 'retention_5s'],
    ['10秒留存', 'retention_10s'],
    ['For You', 'for_you'],
    ['个人主页', 'personal_profile'],
    ['搜索', 'search'],
    ['关注', 'follow'],
    ['私信', 'direct_message'],
    ['音乐', 'sound'],
    ['其它', 'others'],
    ['点赞率', 'like_rate'],
    ['互动率', 'engagement_rate']
  ];

  const percentKeys =
    new Set([
      'watch_ratio',
      'finish_rate',
      'retention_1s',
      'retention_2s',
      'retention_3s',
      'retention_5s',
      'retention_10s',
      'for_you',
      'personal_profile',
      'search',
      'follow',
      'direct_message',
      'sound',
      'others',
      'like_rate',
      'engagement_rate'
    ]);

  function downloadBlob(
    blob,
    filename
  ) {
    const url =
      URL.createObjectURL(
        blob
      );

    const link =
      document.createElement(
        'a'
      );

    link.href =
      url;

    link.download =
      filename;

    document.body
      .appendChild(
        link
      );

    link.click();

    link.remove();

    setTimeout(
      () =>
        URL.revokeObjectURL(
          url
        ),
      3000
    );
  }

  function exportCsv() {
    const rows = [
      ...detailRows.values()
    ];

    if (!rows.length) {
      return;
    }

    const columns = [
      ...exportColumns,
      [
        '视频链接',
        '__url'
      ]
    ];

    const lines = [
      columns
        .map(
          ([label]) =>
            csvEscape(label)
        )
        .join(',')
    ];

    for (
      const row of rows
    ) {
      lines.push(
        columns
          .map(
            ([, key]) => {
              if (
                key === '__url'
              ) {
                return csvEscape(
                  getVideoUrl(row)
                );
              }

              return csvEscape(
                percentKeys.has(
                  key
                )
                  ? pct(
                      row[key]
                    )
                  : row[key]
              );
            }
          )
          .join(',')
      );
    }

    const blob =
      new Blob(
        [
          '\uFEFF' +
          lines.join('\r\n')
        ],
        {
          type:
            'text/csv;charset=utf-8'
        }
      );

    downloadBlob(
      blob,
      `TikTok数据_${new Date()
        .toISOString()
        .slice(
          0,
          10
        )}.csv`
    );
  }

  /* =========================
     Excel 封面
  ========================= */

  function gmFetchBlob(url) {
    return new Promise(
      (
        resolve,
        reject
      ) => {
        GM_xmlhttpRequest({
          method:
            'GET',

          url,

          responseType:
            'arraybuffer',

          timeout:
            15000,

          onload:
            response => {
              if (
                response.status <
                  200 ||
                response.status >=
                  300
              ) {
                reject(
                  new Error(
                    'cover failed'
                  )
                );

                return;
              }

              resolve(
                new Blob(
                  [
                    response.response
                  ],
                  {
                    type:
                      'image/jpeg'
                  }
                )
              );
            },

          onerror:
            reject,

          ontimeout:
            reject
        });
      }
    );
  }

  async function blobToDataUrl(
    blob
  ) {
    const bitmap =
      await createImageBitmap(
        blob
      );

    const canvas =
      document.createElement(
        'canvas'
      );

    const scale =
      Math.min(
        54 /
        bitmap.width,

        90 /
        bitmap.height,

        1
      );

    canvas.width =
      Math.max(
        1,
        Math.round(
          bitmap.width *
          scale
        )
      );

    canvas.height =
      Math.max(
        1,
        Math.round(
          bitmap.height *
          scale
        )
      );

    const ctx =
      canvas.getContext(
        '2d'
      );

    ctx.drawImage(
      bitmap,
      0,
      0,
      canvas.width,
      canvas.height
    );

    bitmap.close();

    return canvas.toDataURL(
      'image/jpeg',
      .72
    );
  }

  async function getCoverData(
    row
  ) {
    if (!row.cover_url) {
      return '';
    }

    if (
      coverCache.has(
        row.cover_url
      )
    ) {
      return coverCache.get(
        row.cover_url
      );
    }

    const promise =
      (
        async () => {
          try {
            return await blobToDataUrl(
              await gmFetchBlob(
                row.cover_url
              )
            );

          } catch {
            return '';
          }
        }
      )();

    coverCache.set(
      row.cover_url,
      promise
    );

    return promise;
  }

  async function prepareImages(
    rows,
    workbook
  ) {
    const map =
      new Map();

    let done = 0;

    for (
      const row of rows
    ) {
      const data =
        await getCoverData(
          row
        );

      if (data) {
        try {
          map.set(
            row.video_id,

            workbook.addImage({
              base64:
                data,

              extension:
                'jpeg'
            })
          );

        } catch {}
      }

      done++;

      updateUI(
        `封面 ${done}/${rows.length}`
      );
    }

    return map;
  }

  /* =========================
     Excel 样式
  ========================= */

  function border() {
    const side = {
      style:
        'thin',

      color: {
        argb:
          COLORS.slate200
      }
    };

    return {
      top:
        side,
      bottom:
        side,
      left:
        side,
      right:
        side
    };
  }

  function setHyperlink(
    cell,
    text,
    url
  ) {
    if (!url) {
      cell.value =
        text;

      return;
    }

    cell.value = {
      text,
      hyperlink:
        url
    };

    cell.font = {
      color: {
        argb:
          COLORS.blue
      },

      underline:
        true
    };
  }

  function buildRawSheet(
    workbook,
    rows,
    images
  ) {
    const sheet =
      workbook.addWorksheet(
        '原始数据',
        {
          views: [
            {
              state:
                'frozen',

              xSplit:
                2,

              ySplit:
                1
            }
          ]
        }
      );

    const columns = [
      ['封面', 'cover', 8],
      ['标题', 'title', 42],
      ['发布时间', 'publish_time', 20],
      ['播放', 'views', 10],
      ['点赞', 'likes', 9],
      ['评论', 'comments', 9],
      ['分享', 'shares', 9],
      ['收藏', 'favorites', 9],
      ['新增粉丝', 'new_followers', 10],
      ['平均观看(s)', 'avg_watch_sec', 13],
      ['观看倍率', 'watch_ratio', 11],
      ['完播率', 'finish_rate', 11],
      ['1秒留存', 'retention_1s', 11],
      ['2秒留存', 'retention_2s', 11],
      ['3秒留存', 'retention_3s', 11],
      ['5秒留存', 'retention_5s', 11],
      ['10秒留存', 'retention_10s', 11],
      ['For You', 'for_you', 11],
      ['搜索', 'search', 10],
      ['点赞率', 'like_rate', 10],
      ['互动率', 'engagement_rate', 10],
      ['视频', 'link', 12]
    ];

    sheet.columns =
      columns.map(
        (
          [
            header,
            key,
            width
          ]
        ) => ({
          header,
          key,
          width
        })
      );

    const header =
      sheet.getRow(1);

    header.height =
      25;

    header.eachCell(
      cell => {
        cell.font = {
          bold:
            true,

          color: {
            argb:
              COLORS.white
          }
        };

        cell.fill = {
          type:
            'pattern',

          pattern:
            'solid',

          fgColor: {
            argb:
              COLORS.navy
          }
        };

        cell.alignment = {
          vertical:
            'middle',

          horizontal:
            'center'
        };
      }
    );

    rows.forEach(
      (
        data,
        index
      ) => {
        const row =
          sheet.addRow({
            cover:
              '',
            title:
              data.title,
            publish_time:
              data.publish_time,
            views:
              data.views,
            likes:
              data.likes,
            comments:
              data.comments,
            shares:
              data.shares,
            favorites:
              data.favorites,
            new_followers:
              data.new_followers,
            avg_watch_sec:
              data.avg_watch_sec,
            watch_ratio:
              data.watch_ratio,
            finish_rate:
              data.finish_rate,
            retention_1s:
              data.retention_1s,
            retention_2s:
              data.retention_2s,
            retention_3s:
              data.retention_3s,
            retention_5s:
              data.retention_5s,
            retention_10s:
              data.retention_10s,
            for_you:
              data.for_you,
            search:
              data.search,
            like_rate:
              data.like_rate,
            engagement_rate:
              data.engagement_rate,
            link:
              '打开视频'
          });

        row.height =
          55;

        row.eachCell(
          cell => {
            cell.border =
              border();

            cell.alignment = {
              vertical:
                'middle',

              wrapText:
                true
            };

            if (
              index % 2
            ) {
              cell.fill = {
                type:
                  'pattern',

                pattern:
                  'solid',

                fgColor: {
                  argb:
                    COLORS.slate50
                }
              };
            }
          }
        );

        setHyperlink(
          row.getCell(
            'title'
          ),
          data.title ||
          '(无标题)',
          getVideoUrl(
            data
          )
        );

        setHyperlink(
          row.getCell(
            'link'
          ),
          '打开视频',
          getVideoUrl(
            data
          )
        );

        const imageId =
          images.get(
            data.video_id
          );

        if (
          imageId !==
          undefined
        ) {
          sheet.addImage(
            imageId,
            {
              tl: {
                col:
                  0.15,

                row:
                  row.number -
                  1 +
                  .08
              },

              ext: {
                width:
                  40,

                height:
                  70
              }
            }
          );
        }
      }
    );

    [
      'watch_ratio',
      'finish_rate',
      'retention_1s',
      'retention_2s',
      'retention_3s',
      'retention_5s',
      'retention_10s',
      'for_you',
      'search',
      'like_rate',
      'engagement_rate'
    ].forEach(
      key => {
        sheet.getColumn(
          key
        ).numFmt =
          '0.0%';
      }
    );

    return sheet;
  }

  function buildAnalysisSheet(
    workbook,
    rows
  ) {
    const sheet =
      workbook.addWorksheet(
        '本期分析'
      );

    const summary =
      summaryFor(
        rows
      );

    const metric =
      $('qtk-rank-metric')
        ?.value ||
      'views';

    const top =
      topRows(
        rows,
        metric,
        5
      );

    sheet.columns = [
      {
        header:
          '项目',
        width:
          52
      },

      {
        header:
          '数值',
        width:
          18
      },

      {
        header:
          '播放量',
        width:
          12
      },

      {
        header:
          '完播率',
        width:
          12
      },

      {
        header:
          '3秒留存',
        width:
          12
      },

      {
        header:
          '视频',
        width:
          14
      }
    ];

    const rowsToAdd = [
      [
        '公开视频数',
        summary.count
      ],

      [
        '平均播放',
        summary.avgViews
      ],

      [
        '中位播放',
        summary.medianViews
      ],

      [
        '平均完播率',
        summary.avgFinish
      ],

      [
        '平均观看倍率',
        summary.avgWatchRatio
      ],

      [
        '平均3秒留存',
        summary.avgR3
      ]
    ];

    rowsToAdd.forEach(
      data =>
        sheet.addRow(
          data
        )
    );

    sheet.addRow([]);

    sheet.addRow([
      `${metricLabel(metric)} TOP 5`
    ]);

    for (
      const data of top
    ) {
      const row =
        sheet.addRow([
          data.title,
          data[metric],
          data.views,
          data.finish_rate,
          data.retention_3s,
          '打开视频'
        ]);

      setHyperlink(
        row.getCell(1),
        data.title,
        getVideoUrl(
          data
        )
      );

      setHyperlink(
        row.getCell(6),
        '打开视频',
        getVideoUrl(
          data
        )
      );
    }

    sheet.getRow(1)
      .eachCell(
        cell => {
          cell.font = {
            bold:
              true,

            color: {
              argb:
                COLORS.white
            }
          };

          cell.fill = {
            type:
              'pattern',

            pattern:
              'solid',

            fgColor: {
              argb:
                COLORS.navy
            }
          };
        }
      );

    return sheet;
  }

  async function exportXlsx() {
    if (busy) {
      return;
    }

    const rows = [
      ...detailRows.values()
    ].sort(
      (a, b) =>
        b.publish_ts -
        a.publish_ts
    );

    if (!rows.length) {
      return;
    }

    busy = true;

    updateUI(
      '生成 XLSX…'
    );

    try {
      const workbook =
        new ExcelJS.Workbook();

      const images =
        await prepareImages(
          rows,
          workbook
        );

      buildRawSheet(
        workbook,
        rows,
        images
      );

      buildAnalysisSheet(
        workbook,
        getAnalysisRows()
      );

      const buffer =
        await workbook.xlsx
          .writeBuffer();

      downloadBlob(
        new Blob(
          [buffer],
          {
            type:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }
        ),

        `TikTok周报_${new Date()
          .toISOString()
          .slice(
            0,
            10
          )}.xlsx`
      );

      updateUI(
        'XLSX 已导出'
      );

    } catch (
      error
    ) {
      console.error(
        error
      );

      alert(
        'XLSX 导出失败'
      );

    } finally {
      busy = false;

      updateUI();
    }
  }

  /* =========================
     启动
  ========================= */

  function boot() {
    if (
      !document.documentElement
    ) {
      setTimeout(
        boot,
        50
      );

      return;
    }

    const start = () => {
      buildPanel();

      startDomObserver();

      setTimeout(
        () =>
          scheduleAutoMetrics(
            200
          ),
        1000
      );
    };

    if (
      document.readyState ===
      'loading'
    ) {
      document.addEventListener(
        'DOMContentLoaded',
        start,
        {
          once:
            true
        }
      );

    } else {
      start();
    }
  }

  boot();

})();
