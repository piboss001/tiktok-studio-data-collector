// ==UserScript==
// @name         TikTok Studio 数据采集器
// @namespace    qualitell.tiktok.collector
// @version      0.4.2
// @description  TikTok Studio 数据增强：独立悬浮层指标、视频标签、随机自动刷新、缓存清理、XLSX/CSV 导出。
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

  /* =========================================================
     V0.4.2
     核心：
     1. 不再向 TikTok 视频行 appendChild
     2. 不修改 TikTok 视频行 style / dataset / HTML
     3. 所有数据放在独立 fixed overlay
     4. 滚动只移动 overlay，不重新请求数据
     5. 数据刷新才更新数值
  ========================================================= */

  const page =
    typeof unsafeWindow !== 'undefined'
      ? unsafeWindow
      : window;

  const VERSION = '0.4.2';

  const STORAGE_KEY =
    'qualitell_tiktok_collector_v042_settings';

  const originalFetch =
    page.fetch.bind(page);

  const OriginalXHR =
    page.XMLHttpRequest;

  const $ = id =>
    document.getElementById(id);

  const sleep = ms =>
    new Promise(resolve =>
      setTimeout(resolve, ms)
    );

  /* =========================================================
     数据
  ========================================================= */

  const items = new Map();

  const detailRows = new Map();

  const metricCache = new Map();

  const metricLoading = new Set();

  const privateIds = new Set();

  const unknownIds = new Set();

  const coverCache = new Map();

  /*
    Video ID -> TikTok 原始 DOM 行
    这里只保存引用，不对行做任何修改。
  */
  const rowCache = new Map();

  /*
    Video ID -> Overlay DOM
  */
  const overlayCache = new Map();

  let failedCount = 0;

  let busy = false;

  let minimized = false;

  let domObserver = null;

  let metricDebounceTimer = null;

  let overlayFrame = null;

  let autoRefreshTimer = null;

  let countdownTimer = null;

  let cleanupTimer = null;

  let nextRefreshAt = 0;

  let lastRefreshAt = 0;

  let lastCleanupAt =
    Date.now();

  let headerGeometryCache =
    null;

  let headerGeometryTime =
    0;

  let rowCacheTime =
    0;

  /* =========================================================
     指标
  ========================================================= */

  const METRICS = {

    finish_rate: {
      label: '完播',
      full: '完播率',
      type: 'percent'
    },

    watch_ratio: {
      label: '倍率',
      full: '观看倍率',
      type: 'percent'
    },

    avg_watch_sec: {
      label: '平均',
      full: '平均观看时长',
      type: 'seconds'
    },

    new_followers: {
      label: '新粉',
      full: '新增粉丝',
      type: 'number'
    },

    retention_1s: {
      label: '1s',
      full: '1秒留存',
      type: 'percent'
    },

    retention_2s: {
      label: '2s',
      full: '2秒留存',
      type: 'percent'
    },

    retention_3s: {
      label: '3s',
      full: '3秒留存',
      type: 'percent'
    },

    engagement_rate: {
      label: '互动',
      full: '互动率',
      type: 'percent'
    },

    duration_sec: {
      label: '时长',
      full: '视频时长',
      type: 'seconds'
    },

    total_watch_sec: {
      label: '总观看',
      full: '总观看时长',
      type: 'duration'
    },

    retention_5s: {
      label: '5s',
      full: '5秒留存',
      type: 'percent'
    },

    retention_10s: {
      label: '10s',
      full: '10秒留存',
      type: 'percent'
    },

    shares: {
      label: '分享',
      full: '分享',
      type: 'number'
    },

    favorites: {
      label: '收藏',
      full: '收藏',
      type: 'number'
    },

    for_you: {
      label: 'For You',
      full: 'For You',
      type: 'percent'
    },

    personal_profile: {
      label: '主页',
      full: '个人主页',
      type: 'percent'
    },

    search: {
      label: '搜索',
      full: '搜索',
      type: 'percent'
    },

    like_rate: {
      label: '点赞率',
      full: '点赞率',
      type: 'percent'
    }

  };

  const METRIC_ORDER = [

    'finish_rate',
    'watch_ratio',
    'avg_watch_sec',
    'new_followers',

    'retention_1s',
    'retention_2s',
    'retention_3s',
    'engagement_rate',

    'duration_sec',
    'total_watch_sec',

    'retention_5s',
    'retention_10s',

    'shares',
    'favorites',

    'for_you',
    'personal_profile',
    'search',
    'like_rate'

  ];

  const DEFAULT_METRICS = [

    'finish_rate',
    'watch_ratio',
    'avg_watch_sec',
    'new_followers',

    'retention_1s',
    'retention_2s',
    'retention_3s',
    'engagement_rate'

  ];

  const MAX_VISIBLE_METRICS = 14;

  /* =========================================================
     设置
  ========================================================= */

  const DEFAULT_SETTINGS = {

    pageMetricsEnabled:
      true,

    videoTagEnabled:
      true,

    selectedMetrics:
      [...DEFAULT_METRICS],

    autoRefreshEnabled:
      true,

    refreshMinMinutes:
      1,

    refreshMaxMinutes:
      10,

    pauseWhenHidden:
      true,

    autoCleanupEnabled:
      true,

    cleanupMinutes:
      30

  };

  let settings =
    loadSettings();

  function loadSettings() {

    try {

      const raw =
        localStorage.getItem(
          STORAGE_KEY
        );

      if (!raw) {

        return {
          ...DEFAULT_SETTINGS,
          selectedMetrics:
            [...DEFAULT_METRICS]
        };

      }

      const parsed =
        JSON.parse(raw);

      const merged = {
        ...DEFAULT_SETTINGS,
        ...parsed
      };

      if (
        !Array.isArray(
          merged.selectedMetrics
        )
      ) {

        merged.selectedMetrics =
          [...DEFAULT_METRICS];

      }

      merged.selectedMetrics =
        merged.selectedMetrics
          .filter(
            key =>
              METRICS[key]
          )
          .slice(
            0,
            MAX_VISIBLE_METRICS
          );

      return merged;

    } catch {

      return {
        ...DEFAULT_SETTINGS,
        selectedMetrics:
          [...DEFAULT_METRICS]
      };

    }

  }

  function saveSettings() {

    try {

      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify(
          settings
        )
      );

    } catch {}

  }

  /* =========================================================
     工具
  ========================================================= */

  function num(value) {

    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {

      return '';

    }

    const n =
      Number(value);

    return Number.isNaN(n)
      ? ''
      : n;

  }

  function pct(
    value,
    digits = 1
  ) {

    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {

      return '-';

    }

    const n =
      Number(value);

    if (
      !Number.isFinite(n)
    ) {

      return '-';

    }

    return `${(
      n * 100
    ).toFixed(digits)}%`;

  }

  function compactNumber(
    value
  ) {

    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {

      return '-';

    }

    const n =
      Number(value);

    if (
      !Number.isFinite(n)
    ) {

      return '-';

    }

    if (
      n >= 1000000
    ) {

      return `${(
        n / 1000000
      ).toFixed(1)}M`;

    }

    if (
      n >= 1000
    ) {

      return `${(
        n / 1000
      ).toFixed(1)}K`;

    }

    return String(
      Math.round(n)
    );

  }

  function compactDuration(
    seconds
  ) {

    const n =
      Number(seconds);

    if (
      !Number.isFinite(n)
    ) {

      return '-';

    }

    if (
      n >= 3600
    ) {

      return `${(
        n / 3600
      ).toFixed(1)}h`;

    }

    if (
      n >= 60
    ) {

      return `${(
        n / 60
      ).toFixed(1)}m`;

    }

    return `${n.toFixed(1)}s`;

  }

  function formatMetric(
    key,
    value
  ) {

    const metric =
      METRICS[key];

    if (!metric) {

      return '-';

    }

    if (
      metric.type ===
      'percent'
    ) {

      return pct(value);

    }

    if (
      metric.type ===
      'seconds'
    ) {

      if (
        value === '' ||
        value === null ||
        value === undefined
      ) {

        return '-';

      }

      return `${Number(value).toFixed(1)}s`;

    }

    if (
      metric.type ===
      'duration'
    ) {

      return compactDuration(
        value
      );

    }

    if (
      metric.type ===
      'number'
    ) {

      return compactNumber(
        value
      );

    }

    return String(
      value ?? '-'
    );

  }

  function formatDate(
    timestamp
  ) {

    if (!timestamp) {

      return '';

    }

    return new Date(
      Number(timestamp) *
      1000
    ).toLocaleString(
      'zh-CN'
    );

  }

  function safeClone(
    value
  ) {

    try {

      return JSON.parse(
        JSON.stringify(value)
      );

    } catch {

      return null;

    }

  }

  function normalizeText(
    text
  ) {

    return String(
      text || ''
    )
      .replace(
        /\s+/g,
        ' '
      )
      .trim()
      .toLowerCase();

  }

  function normalizeCompact(
    text
  ) {

    return String(
      text || ''
    )
      .replace(
        /\s+/g,
        ''
      )
      .trim()
      .toLowerCase();

  }

  function mean(values) {

    const arr =
      values.filter(
        value =>
          typeof value ===
            'number' &&
          Number.isFinite(value)
      );

    if (!arr.length) {

      return '';

    }

    return (
      arr.reduce(
        (a, b) =>
          a + b,
        0
      ) /
      arr.length
    );

  }

  function median(values) {

    const arr =
      values
        .filter(
          value =>
            typeof value ===
              'number' &&
            Number.isFinite(value)
        )
        .sort(
          (a, b) =>
            a - b
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

  function shortTitle(
    title,
    max = 42
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

  function escapeHtml(
    text
  ) {

    return String(
      text ?? ''
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

  function csvEscape(
    value
  ) {

    if (
      value === null ||
      value === undefined
    ) {

      return '';

    }

    const text =
      String(value);

    return /[",\n\r]/.test(
      text
    )
      ? `"${text.replace(
          /"/g,
          '""'
        )}"`
      : text;

  }

  /* =========================================================
     预约判断
  ========================================================= */

  function isScheduled(
    item
  ) {

    const now =
      Math.floor(
        Date.now() / 1000
      );

    const status =
      Number(
        item?.status
      );

    const postTime =
      Number(
        item?.post_time || 0
      );

    const scheduleTime =
      Number(
        item?.schedule_time || 0
      );

    if (
      status === 140
    ) {

      return true;

    }

    if (
      scheduleTime >
      now + 60
    ) {

      return true;

    }

    if (
      postTime >
      now + 60
    ) {

      return true;

    }

    return false;

  }

  function publishedItems() {

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

  /* =========================================================
     TikTok item_list
  ========================================================= */

  function captureItemList(
    url,
    data
  ) {

    if (
      !url ||
      !String(url)
        .includes(
          '/tiktok/creator/manage/item_list/v1/'
        )
    ) {

      return;

    }

    if (
      !data ||
      !Array.isArray(
        data.item_list
      )
    ) {

      return;

    }

    for (
      const item
      of data.item_list
    ) {

      if (
        !item?.item_id
      ) {

        continue;

      }

      items.set(

        String(
          item.item_id
        ),

        safeClone(item)

      );

    }

    updateUI(
      `已发现 ${items.size} 条`
    );

    invalidateRows();

    scheduleVisibleMetrics(
      350
    );

  }

  /* =========================================================
     Hook fetch
  ========================================================= */

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
          typeof input ===
          'string'
            ? input
            : input?.url;

        if (
          url &&
          String(url)
            .includes(
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
            .catch(
              () => {}
            );

        }

      } catch {}

      return response;

    };

  /* =========================================================
     Hook XHR
  ========================================================= */

  if (
    OriginalXHR?.prototype
  ) {

    const originalOpen =
      OriginalXHR
        .prototype
        .open;

    const originalSend =
      OriginalXHR
        .prototype
        .send;

    OriginalXHR
      .prototype
      .open =
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

    OriginalXHR
      .prototype
      .send =
      function (...args) {

        this.addEventListener(
          'load',
          () => {

            try {

              const url =
                this.__qtkUrl;

              if (
                url &&
                String(url)
                  .includes(
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

  /* =========================================================
     CSRF
  ========================================================= */

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

      return decodeURIComponent(
        value
      );

    } catch {

      return value;

    }

  }

  /* =========================================================
     Insight API
  ========================================================= */

  async function fetchInsight(
    videoId
  ) {

    const types = [

      'video_info',

      'video_traffic_source_percent_realtime',

      'video_retention_rate_realtime',

      'video_view_realtime',

      'video_total_duration_realtime',

      'video_per_duration_realtime',

      'video_finish_rate_realtime',

      'video_new_follower_realtime'

    ];

    const requests =
      types.map(
        type => ({

          insigh_type:
            type,

          aweme_id:
            videoId

        })
      );

    const lang =
      document
        .documentElement
        .lang ||
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
            requests
          )

      });

    const headers = {

      accept:
        'application/json, text/plain, */*'

    };

    const csrf =
      getCsrfToken();

    if (csrf) {

      headers[
        'tt-csrf-token'
      ] =
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

  /* =========================================================
     Insight 数据
  ========================================================= */

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

  function trafficObject(
    list
  ) {

    const result = {};

    for (
      const item
      of list || []
    ) {

      result[
        item.key
      ] =
        Number(
          item.value
        );

    }

    return result;

  }

  function extractUrl(
    value
  ) {

    if (!value) {

      return '';

    }

    if (
      typeof value ===
      'string'
    ) {

      return value;

    }

    if (
      Array.isArray(value)
    ) {

      return (
        value.find(
          item =>
            typeof item ===
            'string'
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
      typeof value.url ===
      'string'
    ) {

      return value.url;

    }

    return '';

  }

  function pickCoverUrl(
    item,
    info
  ) {

    const candidates = [

      item?.cover_url,

      item?.cover,

      item?.video?.cover,

      item?.video?.origin_cover,

      item?.video?.dynamic_cover,

      info?.cover,

      info?.video?.cover,

      info?.video?.origin_cover,

      info?.video?.dynamic_cover

    ];

    for (
      const candidate
      of candidates
    ) {

      const url =
        extractUrl(
          candidate
        );

      if (url) {

        return url;

      }

    }

    return '';

  }

  function normalizeRow(
    item,
    insight
  ) {

    const data =
      insight.data || {};

    const info =
      data.video_info || {};

    const stats =
      info.statistics || {};

    const retention =
      data
        .video_retention_rate_realtime
        ?.value
        ?.list ||
      [];

    const traffic =
      trafficObject(

        data
          .video_traffic_source_percent_realtime
          ?.value
          ?.value ||
        []

      );

    const durationMs =
      Number(

        info.video?.duration ||
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

        stats.play_count

      );

    const likes =
      num(

        stats.digg_count ??
        item.like_count

      );

    const comments =
      num(

        stats.comment_count ??
        item.comment_count

      );

    const shares =
      num(

        stats.share_count ??
        item.share_count

      );

    const favorites =
      num(

        stats.collect_count ??
        item.favorite_count

      );

    const publishTs =
      Number(

        item.post_time ||
        info.create_time ||
        0

      );

    return {

      account:
        info.author
          ?.unique_id ||

        item.author
          ?.unique_id ||

        item.author_unique_id ||
        '',

      nickname:
        info.author
          ?.nickname ||

        item.author
          ?.nickname ||
        '',

      video_id:
        String(

          item.item_id ||
          info.aweme_id ||
          ''

        ),

      title:
        item.desc ||
        info.desc ||
        '',

      cover_url:
        pickCoverUrl(
          item,
          info
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
          ? avgWatch /
            durationSec
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
            ) /
            views
          : ''

    };

  }

  /* =========================================================
     视频 URL
  ========================================================= */

  function getVideoUrl(
    row
  ) {

    const account =
      String(
        row?.account || ''
      ).replace(
        /^@/,
        ''
      );

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

  /* =========================================================
     TikTok 行识别
     完全只读
  ========================================================= */

  function invalidateRows() {

    rowCacheTime = 0;

    headerGeometryCache =
      null;

    headerGeometryTime = 0;

  }

  function climbVideoRow(
    element
  ) {

    let current =
      element;

    let best =
      null;

    let bestWidth =
      0;

    for (
      let i = 0;
      i < 13 &&
      current &&
      current !== document.body;
      i++
    ) {

      if (
        current.closest?.(
          '#qualitell-tiktok-panel'
        ) ||
        current.closest?.(
          '#qtk-overlay-layer'
        )
      ) {

        return null;

      }

      const rect =
        current
          .getBoundingClientRect();

      if (
        rect.width > 650 &&
        rect.height >= 48 &&
        rect.height <= 180
      ) {

        if (
          rect.width >
          bestWidth
        ) {

          best =
            current;

          bestWidth =
            rect.width;

        }

      }

      current =
        current.parentElement;

    }

    return best;

  }

  function itemTimeTokens(
    item
  ) {

    const ts =
      Number(

        item.schedule_time ||
        item.post_time ||
        0

      );

    if (!ts) {

      return [];

    }

    const date =
      new Date(
        ts * 1000
      );

    const month =
      date.getMonth() + 1;

    const day =
      date.getDate();

    const hour =
      date.getHours();

    const minute =
      String(
        date.getMinutes()
      ).padStart(
        2,
        '0'
      );

    return [

      `${month}月${day}日${hour}:${minute}`,

      `${month}月${day}日${String(hour).padStart(2, '0')}:${minute}`,

      `${month}/${day}${hour}:${minute}`,

      `${month}-${day}${hour}:${minute}`

    ].map(
      normalizeCompact
    );

  }

  function rowStillMatchesItem(
    row,
    item
  ) {

    if (
      !row ||
      !row.isConnected
    ) {

      return false;

    }

    const id =
      String(
        item.item_id
      );

    /*
      精确 ID
    */

    try {

      if (
        row.querySelector(
          `a[href*="${id}"]`
        ) ||
        row.querySelector(
          `[data-video-id="${id}"]`
        ) ||
        row.querySelector(
          `[data-item-id="${id}"]`
        ) ||
        row.querySelector(
          `[data-aweme-id="${id}"]`
        )
      ) {

        return true;

      }

    } catch {}

    /*
      标题 + 时间
    */

    const title =
      normalizeCompact(
        item.desc || ''
      );

    if (
      !title ||
      title.length < 4
    ) {

      return false;

    }

    const titleKey =
      title.slice(
        0,
        Math.min(
          20,
          title.length
        )
      );

    const rowText =
      normalizeCompact(

        row.innerText ||
        row.textContent ||
        ''

      );

    if (
      !rowText.includes(
        titleKey
      )
    ) {

      return false;

    }

    const timeTokens =
      itemTimeTokens(
        item
      );

    if (
      !timeTokens.length
    ) {

      return true;

    }

    /*
      时间能匹配则最可靠
    */

    if (
      timeTokens.some(
        token =>
          rowText.includes(
            token
          )
      )
    ) {

      return true;

    }

    /*
      标题唯一时也允许
    */

    return true;

  }

  function collectVisibleRowCandidates() {

    const rows =
      new Set();

    /*
      视频缩略图通常是最可靠的入口
    */

    const images =
      document.querySelectorAll(
        'img'
      );

    for (
      const image
      of images
    ) {

      if (
        image.closest(
          '#qualitell-tiktok-panel'
        ) ||
        image.closest(
          '#qtk-overlay-layer'
        )
      ) {

        continue;

      }

      const rect =
        image.getBoundingClientRect();

      if (
        rect.width < 25 ||
        rect.height < 25
      ) {

        continue;

      }

      const row =
        climbVideoRow(
          image
        );

      if (!row) {

        continue;

      }

      const rowRect =
        row.getBoundingClientRect();

      if (
        rowRect.bottom < -300 ||
        rowRect.top >
          window.innerHeight + 500
      ) {

        continue;

      }

      rows.add(
        row
      );

    }

    /*
      兼容 table / role row
    */

    const explicitRows =
      document.querySelectorAll(
        '[role="row"],tr,li'
      );

    for (
      const row
      of explicitRows
    ) {

      if (
        row.closest(
          '#qualitell-tiktok-panel'
        ) ||
        row.closest(
          '#qtk-overlay-layer'
        )
      ) {

        continue;

      }

      const rect =
        row.getBoundingClientRect();

      if (
        rect.width > 650 &&
        rect.height >= 48 &&
        rect.height <= 180 &&
        rect.bottom >= -300 &&
        rect.top <=
          window.innerHeight + 500
      ) {

        rows.add(
          row
        );

      }

    }

    return [
      ...rows
    ];

  }

  function refreshRowCache(
    force = false
  ) {

    const now =
      Date.now();

    if (
      !force &&
      now - rowCacheTime < 400
    ) {

      return;

    }

    rowCacheTime =
      now;

    /*
      删除断开的 DOM
    */

    for (
      const [
        id,
        row
      ]
      of rowCache
    ) {

      if (
        !row ||
        !row.isConnected
      ) {

        rowCache.delete(id);

      }

    }

    const candidates =
      collectVisibleRowCandidates();

    if (
      !candidates.length
    ) {

      return;

    }

    const candidateData =
      candidates.map(
        row => {

          const text =
            normalizeCompact(

              row.innerText ||
              row.textContent ||
              ''

            );

          const html =
            String(
              row.innerHTML || ''
            );

          return {
            row,
            text,
            html
          };

        }
      );

    /*
      第一轮：
      优先 Video ID 精确匹配
    */

    for (
      const item
      of items.values()
    ) {

      const id =
        String(
          item.item_id
        );

      const cached =
        rowCache.get(id);

      if (
        cached &&
        rowStillMatchesItem(
          cached,
          item
        )
      ) {

        continue;

      }

      for (
        const candidate
        of candidateData
      ) {

        if (
          candidate.html.includes(id)
        ) {

          rowCache.set(
            id,
            candidate.row
          );

          break;

        }

      }

    }

    /*
      第二轮：
      标题 + 发布时间
    */

    const usedRows =
      new Set(
        rowCache.values()
      );

    for (
      const item
      of items.values()
    ) {

      const id =
        String(
          item.item_id
        );

      const cached =
        rowCache.get(id);

      if (
        cached &&
        rowStillMatchesItem(
          cached,
          item
        )
      ) {

        continue;

      }

      const title =
        normalizeCompact(
          item.desc || ''
        );

      if (
        !title ||
        title.length < 4
      ) {

        continue;

      }

      const titleKey =
        title.slice(
          0,
          Math.min(
            20,
            title.length
          )
        );

      const timeTokens =
        itemTimeTokens(
          item
        );

      let best =
        null;

      let bestScore =
        -Infinity;

      let matchingTitleCount =
        0;

      for (
        const candidate
        of candidateData
      ) {

        if (
          usedRows.has(
            candidate.row
          )
        ) {

          continue;

        }

        if (
          !candidate.text.includes(
            titleKey
          )
        ) {

          continue;

        }

        matchingTitleCount++;

        let score =
          10;

        for (
          const token
          of timeTokens
        ) {

          if (
            candidate.text.includes(
              token
            )
          ) {

            score += 100;

          }

        }

        if (
          isScheduled(item) &&
          (
            candidate.text.includes(
              '预约发布'
            ) ||
            candidate.text.includes(
              'scheduled'
            )
          )
        ) {

          score += 50;

        }

        if (
          !isScheduled(item) &&
          !candidate.text.includes(
            '预约发布'
          ) &&
          !candidate.text.includes(
            'scheduled'
          )
        ) {

          score += 5;

        }

        if (
          score >
          bestScore
        ) {

          best =
            candidate.row;

          bestScore =
            score;

        }

      }

      if (
        best &&
        (
          bestScore >= 100 ||
          matchingTitleCount === 1
        )
      ) {

        rowCache.set(
          id,
          best
        );

        usedRows.add(
          best
        );

      }

    }

  }

  function findVideoRow(
    item
  ) {

    if (
      !item?.item_id
    ) {

      return null;

    }

    const id =
      String(
        item.item_id
      );

    let row =
      rowCache.get(id);

    if (
      row &&
      rowStillMatchesItem(
        row,
        item
      )
    ) {

      return row;

    }

    rowCache.delete(id);

    refreshRowCache(
      true
    );

    row =
      rowCache.get(id);

    if (
      row &&
      rowStillMatchesItem(
        row,
        item
      )
    ) {

      return row;

    }

    return null;

  }

  /* =========================================================
     表头
  ========================================================= */

  function findHeaderNode(
    aliases
  ) {

    const nodes =
      document.querySelectorAll(
        'div,span,p'
      );

    let best =
      null;

    for (
      const node
      of nodes
    ) {

      if (
        node.closest(
          '#qualitell-tiktok-panel'
        ) ||
        node.closest(
          '#qtk-overlay-layer'
        )
      ) {

        continue;

      }

      const text =
        String(
          node.textContent || ''
        )
          .replace(
            /\s+/g,
            ''
          )
          .trim();

      if (
        !aliases.includes(
          text
        )
      ) {

        continue;

      }

      const rect =
        node.getBoundingClientRect();

      if (
        rect.width <= 0 ||
        rect.height <= 0 ||
        rect.top < -20 ||
        rect.top > 350
      ) {

        continue;

      }

      if (
        !best ||
        rect.width <
        best.rect.width
      ) {

        best = {
          node,
          rect
        };

      }

    }

    return best;

  }

  function getHeaderGeometry(
    force = false
  ) {

    const now =
      Date.now();

    if (
      !force &&
      headerGeometryCache &&
      now - headerGeometryTime <
        1500
    ) {

      return headerGeometryCache;

    }

    const privacy =
      findHeaderNode([
        '隐私'
      ]);

    const views =
      findHeaderNode([
        '观看次数'
      ]);

    const likes =
      findHeaderNode([
        '赞'
      ]);

    const comments =
      findHeaderNode([
        '评论'
      ]);

    const actions =
      findHeaderNode([
        '操作'
      ]);

    if (
      !privacy ||
      !views ||
      !likes ||
      !comments
    ) {

      return headerGeometryCache;

    }

    headerGeometryCache = {

      privacyX:
        privacy.rect.left +
        privacy.rect.width / 2,

      viewsX:
        views.rect.left +
        views.rect.width / 2,

      likesX:
        likes.rect.left +
        likes.rect.width / 2,

      commentsX:
        comments.rect.left +
        comments.rect.width / 2,

      actionsX:
        actions
          ? actions.rect.left +
            actions.rect.width / 2
          : null

    };

    headerGeometryTime =
      now;

    return headerGeometryCache;

  }

  /* =========================================================
     独立 Overlay
  ========================================================= */

  function ensureOverlayLayer() {

    let layer =
      document.getElementById(
        'qtk-overlay-layer'
      );

    if (layer) {

      return layer;

    }

    layer =
      document.createElement(
        'div'
      );

    layer.id =
      'qtk-overlay-layer';

    document.documentElement
      .appendChild(
        layer
      );

    return layer;

  }

  function visibleSelectedMetrics() {

    return METRIC_ORDER
      .filter(
        key =>
          settings
            .selectedMetrics
            .includes(key)
      )
      .slice(
        0,
        MAX_VISIBLE_METRICS
      );

  }

  function createOverlayMetric(
    key
  ) {

    const metric =
      METRICS[key];

    const item =
      document.createElement(
        'div'
      );

    item.className =
      'qtk-overlay-metric';

    item.dataset.metric =
      key;

    const label =
      document.createElement(
        'span'
      );

    label.textContent =
      metric.label;

    const value =
      document.createElement(
        'b'
      );

    value.textContent =
      '-';

    item.append(
      label,
      value
    );

    return item;

  }

  function createOverlayNode(
    videoId
  ) {

    const layer =
      ensureOverlayLayer();

    const root =
      document.createElement(
        'div'
      );

    root.className =
      'qtk-overlay-video';

    root.dataset.videoId =
      String(videoId);

    const tag =
      document.createElement(
        'div'
      );

    tag.className =
      'qtk-video-tag';

    const top =
      document.createElement(
        'div'
      );

    top.className =
      'qtk-overlay-line qtk-overlay-top';

    const bottom =
      document.createElement(
        'div'
      );

    bottom.className =
      'qtk-overlay-line qtk-overlay-bottom';

    root.append(
      tag,
      top,
      bottom
    );

    layer.appendChild(
      root
    );

    const result = {
      root,
      tag,
      top,
      bottom,
      layoutKey:
        ''
    };

    overlayCache.set(
      String(videoId),
      result
    );

    return result;

  }

  function getOverlayNode(
    videoId
  ) {

    const id =
      String(videoId);

    const cached =
      overlayCache.get(id);

    if (
      cached &&
      cached.root?.isConnected
    ) {

      return cached;

    }

    overlayCache.delete(id);

    return createOverlayNode(id);

  }

  function removeOverlay(
    videoId
  ) {

    const id =
      String(videoId);

    const overlay =
      overlayCache.get(id);

    if (overlay) {

      overlay.root?.remove();

      overlayCache.delete(id);

    }

  }

  function clearAllOverlays() {

    for (
      const overlay
      of overlayCache.values()
    ) {

      overlay.root?.remove();

    }

    overlayCache.clear();

  }

  function rebuildOverlayLayout(
    overlay,
    selected
  ) {

    const layoutKey =
      selected.join('|');

    if (
      overlay.layoutKey ===
      layoutKey
    ) {

      return;

    }

    overlay.layoutKey =
      layoutKey;

    overlay.top.textContent =
      '';

    overlay.bottom.textContent =
      '';

    const split =
      Math.ceil(
        selected.length / 2
      );

    const topKeys =
      selected.slice(
        0,
        split
      );

    const bottomKeys =
      selected.slice(
        split
      );

    for (
      const key
      of topKeys
    ) {

      overlay.top.appendChild(
        createOverlayMetric(
          key
        )
      );

    }

    for (
      const key
      of bottomKeys
    ) {

      overlay.bottom.appendChild(
        createOverlayMetric(
          key
        )
      );

    }

    overlay.root
      .classList
      .toggle(
        'qtk-overlay-dense',
        selected.length > 10
      );

  }

  function updateOverlayValues(
    overlay,
    selected,
    rowData
  ) {

    for (
      const key
      of selected
    ) {

      const valueNode =
        overlay.root
          .querySelector(
            `.qtk-overlay-metric[data-metric="${key}"] b`
          );

      if (!valueNode) {

        continue;

      }

      const next =
        formatMetric(
          key,
          rowData[key]
        );

      if (
        valueNode.textContent !==
        next
      ) {

        valueNode.textContent =
          next;

      }

    }

  }

  /* =========================================================
     标签
  ========================================================= */

  function getRecent7DayMedianViews() {

    const now =
      Math.floor(
        Date.now() / 1000
      );

    const cutoff =
      now -
      7 * 86400;

    const values = [];

    for (
      const item
      of items.values()
    ) {

      if (
        isScheduled(item)
      ) {

        continue;

      }

      const ts =
        Number(
          item.post_time || 0
        );

      if (
        ts &&
        ts < cutoff
      ) {

        continue;

      }

      const row =
        detailRows.get(
          String(
            item.item_id
          )
        );

      const views =
        Number(
          row?.views ??
          item.play_count
        );

      if (
        Number.isFinite(
          views
        )
      ) {

        values.push(
          views
        );

      }

    }

    return median(
      values
    );

  }

  function classifyVideo(
    rowData
  ) {

    const views =
      Number(
        rowData.views || 0
      );

    if (
      views < 100
    ) {

      return {
        key:
          'watching',
        text:
          '待观察'
      };

    }

    const excellent = {

      finish_rate:
        0.45,

      watch_ratio:
        1.00,

      retention_1s:
        0.85,

      retention_2s:
        0.75,

      retention_3s:
        0.65

    };

    const good = {

      finish_rate:
        0.35,

      watch_ratio:
        0.85,

      retention_1s:
        0.80,

      retention_2s:
        0.68,

      retention_3s:
        0.58

    };

    const excellentCount =
      Object.entries(
        excellent
      )
        .filter(
          ([key, threshold]) =>
            Number(
              rowData[key]
            ) >= threshold
        )
        .length;

    const isExcellent =
      Number(
        rowData.finish_rate
      ) >= 0.45 &&
      Number(
        rowData.retention_3s
      ) >= 0.65 &&
      excellentCount >= 4;

    const medianViews =
      Number(
        getRecent7DayMedianViews()
      ) || 0;

    const viralThreshold =
      Math.max(
        800,
        medianViews * 3
      );

    if (
      isExcellent &&
      views >= viralThreshold
    ) {

      return {
        key:
          'viral',
        text:
          '🔥 爆款视频'
      };

    }

    if (
      isExcellent
    ) {

      return {
        key:
          'excellent',
        text:
          '⭐ 优秀视频'
      };

    }

    const goodCount =
      Object.entries(
        good
      )
        .filter(
          ([key, threshold]) =>
            Number(
              rowData[key]
            ) >= threshold
        )
        .length;

    if (
      goodCount >= 4
    ) {

      return {
        key:
          'good',
        text:
          '👍 好视频'
      };

    }

    const badPrimary =
      Number(
        rowData.finish_rate
      ) < 0.25 ||
      Number(
        rowData.retention_3s
      ) < 0.45;

    const lowCount = [

      Number(
        rowData.watch_ratio
      ) < 0.70,

      Number(
        rowData.retention_1s
      ) < 0.70,

      Number(
        rowData.retention_2s
      ) < 0.55

    ].filter(Boolean)
      .length;

    if (
      badPrimary &&
      lowCount >= 1
    ) {

      return {
        key:
          'bad',
        text:
          '⚠ 差视频'
      };

    }

    return {
      key:
        'normal',
      text:
        '普通视频'
    };

  }

  function updateOverlayTag(
    overlay,
    rowData
  ) {

    if (
      !settings
        .videoTagEnabled
    ) {

      overlay.tag.style.display =
        'none';

      return;

    }

    overlay.tag.style.display =
      'inline-flex';

    const result =
      classifyVideo(
        rowData
      );

    const wantedClass =
      `qtk-video-tag qtk-tag-${result.key}`;

    if (
      overlay.tag.className !==
      wantedClass
    ) {

      overlay.tag.className =
        wantedClass;

    }

    if (
      overlay.tag.textContent !==
      result.text
    ) {

      overlay.tag.textContent =
        result.text;

    }

  }

  /* =========================================================
     Overlay 位置
  ========================================================= */

  function isRowVisible(
    row,
    extra = 120
  ) {

    if (
      !row ||
      !row.isConnected
    ) {

      return false;

    }

    const rect =
      row.getBoundingClientRect();

    return (
      rect.bottom >= -extra &&
      rect.top <=
        window.innerHeight + extra
    );

  }

  function positionOverlayForItem(
    item,
    rowData
  ) {

    const id =
      String(
        item.item_id
      );

    if (
      !settings
        .pageMetricsEnabled ||
      isScheduled(item)
    ) {

      removeOverlay(id);

      return;

    }

    const row =
      findVideoRow(
        item
      );

    if (
      !row ||
      !isRowVisible(
        row,
        100
      )
    ) {

      const existing =
        overlayCache.get(id);

      if (existing) {

        existing.root.style.display =
          'none';

      }

      return;

    }

    const geometry =
      getHeaderGeometry();

    if (!geometry) {

      return;

    }

    const rect =
      row.getBoundingClientRect();

    /*
      防止错误的大容器
    */

    if (
      rect.width < 650 ||
      rect.height < 45 ||
      rect.height > 180
    ) {

      return;

    }

    const overlay =
      getOverlayNode(id);

    overlay.root.style.display =
      'block';

    const selected =
      visibleSelectedMetrics();

    rebuildOverlayLayout(
      overlay,
      selected
    );

    updateOverlayValues(
      overlay,
      selected,
      rowData
    );

    updateOverlayTag(
      overlay,
      rowData
    );

    /*
      数据区：
      从 隐私 列左侧向 评论 列右侧展开。
    */

    let metricsLeft =
      Math.max(
        rect.left + 265,
        geometry.privacyX - 210
      );

    let metricsRight =
      Math.min(
        rect.right - 105,
        geometry.commentsX + 105
      );

    if (
      metricsRight -
      metricsLeft <
      340
    ) {

      metricsLeft =
        Math.max(
          rect.left + 240,
          geometry.privacyX - 185
        );

      metricsRight =
        Math.min(
          rect.right - 90,
          geometry.commentsX + 115
        );

    }

    const metricsWidth =
      Math.max(
        320,
        metricsRight -
        metricsLeft
      );

    const topY =
      Math.round(
        rect.top + 2
      );

    const bottomY =
      Math.round(
        rect.bottom - 17
      );

    overlay.top.style.left =
      `${Math.round(metricsLeft)}px`;

    overlay.top.style.top =
      `${topY}px`;

    overlay.top.style.width =
      `${Math.round(metricsWidth)}px`;

    overlay.bottom.style.left =
      `${Math.round(metricsLeft)}px`;

    overlay.bottom.style.top =
      `${bottomY}px`;

    overlay.bottom.style.width =
      `${Math.round(metricsWidth)}px`;

    /*
      标签：
      视频标题右侧 / 隐私列左侧。
    */

    const tagLeft =
      Math.max(
        rect.left + 300,
        geometry.privacyX - 190
      );

    const tagTop =
      rect.top +
      rect.height / 2;

    overlay.tag.style.left =
      `${Math.round(tagLeft)}px`;

    overlay.tag.style.top =
      `${Math.round(tagTop)}px`;

  }

  function syncOverlayPositions() {

    overlayFrame =
      null;

    if (
      !settings
        .pageMetricsEnabled
    ) {

      clearAllOverlays();

      return;

    }

    /*
      不在这里请求数据。
      只移动已有 Overlay。
    */

    for (
      const [
        videoId,
        cached
      ]
      of metricCache
    ) {

      if (
        cached?.type !==
        'public'
      ) {

        continue;

      }

      const item =
        items.get(
          String(videoId)
        );

      if (!item) {

        removeOverlay(
          videoId
        );

        continue;

      }

      if (
        isScheduled(item)
      ) {

        removeOverlay(
          videoId
        );

        continue;

      }

      positionOverlayForItem(
        item,
        cached.row
      );

    }

  }

  function scheduleOverlaySync() {

    if (overlayFrame) {

      return;

    }

    overlayFrame =
      requestAnimationFrame(
        syncOverlayPositions
      );

  }

  /* =========================================================
     当前可见作品
  ========================================================= */

  function getVisiblePublishedItems() {

    refreshRowCache();

    const result = [];

    for (
      const item
      of items.values()
    ) {

      if (
        isScheduled(item)
      ) {

        removeOverlay(
          item.item_id
        );

        continue;

      }

      const row =
        rowCache.get(
          String(
            item.item_id
          )
        );

      if (
        !row ||
        !rowStillMatchesItem(
          row,
          item
        ) ||
        !isRowVisible(
          row,
          250
        )
      ) {

        continue;

      }

      result.push(
        item
      );

    }

    return result;

  }

  /* =========================================================
     单条指标
  ========================================================= */

  async function fetchItemMetric(
    item,
    force = false
  ) {

    if (
      !item?.item_id ||
      isScheduled(item)
    ) {

      removeOverlay(
        item?.item_id
      );

      return null;

    }

    const videoId =
      String(
        item.item_id
      );

    const cached =
      metricCache.get(
        videoId
      );

    if (
      !force &&
      cached?.type ===
        'public'
    ) {

      positionOverlayForItem(
        item,
        cached.row
      );

      return cached.row;

    }

    if (
      metricLoading.has(
        videoId
      )
    ) {

      return null;

    }

    metricLoading.add(
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
          'Insight request failed'
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

        const rowData =
          normalizeRow(
            item,
            insight
          );

        metricCache.set(
          videoId,
          {

            type:
              'public',

            row:
              rowData,

            fetchedAt:
              Date.now()

          }
        );

        detailRows.set(
          videoId,
          rowData
        );

        privateIds.delete(
          videoId
        );

        unknownIds.delete(
          videoId
        );

        positionOverlayForItem(
          item,
          rowData
        );

        return rowData;

      }

      removeOverlay(
        videoId
      );

      if (
        privateStatus === null ||
        privateStatus === undefined
      ) {

        metricCache.set(
          videoId,
          {

            type:
              'unknown',

            fetchedAt:
              Date.now()

          }
        );

        unknownIds.add(
          videoId
        );

      } else {

        metricCache.set(
          videoId,
          {

            type:
              'private',

            fetchedAt:
              Date.now()

          }
        );

        privateIds.add(
          videoId
        );

      }

    } catch (error) {

      console.warn(
        '[TikTok Collector]',
        videoId,
        error
      );

      metricCache.set(
        videoId,
        {

          type:
            'error',

          fetchedAt:
            Date.now()

        }
      );

      failedCount++;

    } finally {

      metricLoading.delete(
        videoId
      );

      updateUI();

    }

    return null;

  }

  /* =========================================================
     并发
  ========================================================= */

  async function runWorkers(
    list,
    handler,
    concurrency = 2
  ) {

    if (!list.length) {

      return;

    }

    let index = 0;

    async function worker() {

      while (true) {

        const current =
          index++;

        if (
          current >=
          list.length
        ) {

          break;

        }

        await handler(
          list[current],
          current
        );

        await sleep(
          250
        );

      }

    }

    const count =
      Math.min(
        concurrency,
        list.length
      );

    await Promise.all(

      Array.from(
        {
          length:
            count
        },
        () =>
          worker()
      )

    );

  }

  /* =========================================================
     当前页自动加载
  ========================================================= */

  async function loadVisibleMetrics() {

    if (
      busy ||
      !settings
        .pageMetricsEnabled
    ) {

      return;

    }

    refreshRowCache(
      true
    );

    const list =
      getVisiblePublishedItems();

    if (!list.length) {

      scheduleOverlaySync();

      return;

    }

    /*
      已有缓存：只显示
      没缓存：才请求
    */

    const needFetch = [];

    for (
      const item
      of list
    ) {

      const videoId =
        String(
          item.item_id
        );

      const cached =
        metricCache.get(
          videoId
        );

      if (
        cached?.type ===
        'public'
      ) {

        positionOverlayForItem(
          item,
          cached.row
        );

      } else if (
        !cached &&
        !metricLoading.has(
          videoId
        )
      ) {

        needFetch.push(
          item
        );

      }

    }

    if (
      needFetch.length
    ) {

      await runWorkers(

        needFetch,

        async item => {

          await fetchItemMetric(
            item,
            false
          );

        },

        2

      );

    }

    scheduleOverlaySync();

  }

  function scheduleVisibleMetrics(
    delay = 400
  ) {

    clearTimeout(
      metricDebounceTimer
    );

    metricDebounceTimer =
      setTimeout(
        () => {

          loadVisibleMetrics()
            .catch(
              error =>
                console.warn(
                  '[TikTok Collector]',
                  error
                )
            );

        },
        delay
      );

  }

  /* =========================================================
     手动 / 自动刷新
  ========================================================= */

  async function refreshVisibleMetrics(
    source = 'manual'
  ) {

    if (busy) {

      return;

    }

    if (
      settings
        .pauseWhenHidden &&
      document.hidden
    ) {

      return;

    }

    refreshRowCache(
      true
    );

    const list =
      getVisiblePublishedItems();

    if (!list.length) {

      updateUI(
        '当前没有可刷新的公开视频'
      );

      scheduleNextAutoRefresh();

      return;

    }

    busy = true;

    updateUI(
      source === 'auto'
        ? '自动更新中…'
        : '更新中…'
    );

    let completed = 0;

    const progress =
      $('qtk-progress-bar');

    if (progress) {

      progress.style.width =
        '0%';

    }

    try {

      await runWorkers(

        list,

        async item => {

          await fetchItemMetric(
            item,
            true
          );

          completed++;

          if (progress) {

            progress.style.width =
              `${Math.round(
                completed /
                list.length *
                100
              )}%`;

          }

        },

        2

      );

      lastRefreshAt =
        Date.now();

      updateRefreshUI();

      updateUI(
        `更新完成 ${completed} 条`
      );

      scheduleOverlaySync();

    } finally {

      busy = false;

      setTimeout(
        () => {

          if (progress) {

            progress.style.width =
              '0%';

          }

        },
        800
      );

      scheduleNextAutoRefresh();

    }

  }

  /* =========================================================
     随机自动刷新
  ========================================================= */

  function randomRefreshDelay() {

    let min =
      Math.max(
        1,
        Number(
          settings
            .refreshMinMinutes
        ) || 1
      );

    let max =
      Math.max(
        min,
        Number(
          settings
            .refreshMaxMinutes
        ) || min
      );

    min =
      Math.min(
        min,
        30
      );

    max =
      Math.min(
        max,
        30
      );

    const minMs =
      min * 60000;

    const maxMs =
      max * 60000;

    return Math.floor(

      minMs +
      Math.random() *
      (
        maxMs -
        minMs +
        1
      )

    );

  }

  function clearAutoRefreshTimer() {

    if (
      autoRefreshTimer
    ) {

      clearTimeout(
        autoRefreshTimer
      );

      autoRefreshTimer =
        null;

    }

  }

  function scheduleNextAutoRefresh(
    keepExisting = false
  ) {

    clearAutoRefreshTimer();

    if (
      !settings
        .autoRefreshEnabled
    ) {

      nextRefreshAt = 0;

      updateRefreshUI();

      return;

    }

    if (
      settings
        .pauseWhenHidden &&
      document.hidden
    ) {

      updateRefreshUI();

      return;

    }

    if (
      !keepExisting ||
      !nextRefreshAt ||
      nextRefreshAt <=
        Date.now()
    ) {

      nextRefreshAt =
        Date.now() +
        randomRefreshDelay();

    }

    const delay =
      Math.max(
        500,
        nextRefreshAt -
        Date.now()
      );

    autoRefreshTimer =
      setTimeout(
        () => {

          nextRefreshAt = 0;

          refreshVisibleMetrics(
            'auto'
          );

        },
        delay
      );

    updateRefreshUI();

  }

  function formatCountdown(
    ms
  ) {

    if (
      !Number.isFinite(ms) ||
      ms <= 0
    ) {

      return '-';

    }

    const total =
      Math.ceil(
        ms / 1000
      );

    const minutes =
      Math.floor(
        total / 60
      );

    const seconds =
      total % 60;

    return (
      `${minutes}分${String(
        seconds
      ).padStart(
        2,
        '0'
      )}秒`
    );

  }

  function updateRefreshUI() {

    const last =
      $('qtk-last-refresh');

    const next =
      $('qtk-next-refresh');

    if (last) {

      const text =
        lastRefreshAt
          ? new Date(
              lastRefreshAt
            ).toLocaleTimeString(
              'zh-CN',
              {
                hour12:
                  false
              }
            )
          : '-';

      if (
        last.textContent !==
        text
      ) {

        last.textContent =
          text;

      }

    }

    if (next) {

      let text =
        '-';

      if (
        !settings
          .autoRefreshEnabled
      ) {

        text =
          '已关闭';

      } else if (
        settings
          .pauseWhenHidden &&
        document.hidden
      ) {

        text =
          '后台暂停';

      } else if (
        nextRefreshAt
      ) {

        text =
          formatCountdown(
            nextRefreshAt -
            Date.now()
          );

      } else {

        text =
          '等待安排';

      }

      if (
        next.textContent !==
        text
      ) {

        next.textContent =
          text;

      }

    }

  }

  function startCountdownTimer() {

    clearInterval(
      countdownTimer
    );

    countdownTimer =
      setInterval(
        updateRefreshUI,
        1000
      );

  }

  /* =========================================================
     缓存清理
  ========================================================= */

  function cleanupPluginCache(
    force = false
  ) {

    const minutes =
      Number(
        settings.cleanupMinutes
      ) || 30;

    const ttl =
      minutes * 60000;

    const now =
      Date.now();

    let removed = 0;

    for (
      const [
        videoId,
        cache
      ]
      of metricCache
    ) {

      const row =
        rowCache.get(
          videoId
        );

      const visible =
        row &&
        isRowVisible(
          row,
          100
        );

      if (visible) {

        continue;

      }

      const age =
        now -
        Number(
          cache?.fetchedAt || 0
        );

      if (
        force ||
        age >= ttl
      ) {

        metricCache.delete(
          videoId
        );

        removeOverlay(
          videoId
        );

        removed++;

      }

    }

    for (
      const [
        id,
        row
      ]
      of rowCache
    ) {

      if (
        !row ||
        !row.isConnected
      ) {

        rowCache.delete(id);

      }

    }

    coverCache.clear();

    lastCleanupAt =
      now;

    if (force) {

      updateUI(
        `已清理缓存 ${removed} 项`
      );

    }

  }

  function startCleanupTimer() {

    clearInterval(
      cleanupTimer
    );

    cleanupTimer =
      setInterval(
        () => {

          if (
            !settings
              .autoCleanupEnabled
          ) {

            return;

          }

          const interval =
            Number(
              settings.cleanupMinutes
            ) *
            60000;

          if (
            Date.now() -
            lastCleanupAt >=
            interval
          ) {

            cleanupPluginCache(
              false
            );

          }

        },
        60000
      );

  }

  /* =========================================================
     分析
  ========================================================= */

  function getRangeDays() {

    const value =
      $('qtk-range')
        ?.value ||
      '7';

    return value ===
      'all'
      ? null
      : Number(value);

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
      days * 86400;

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

  function summaryFor(
    rows
  ) {

    return {

      count:
        rows.length,

      avgViews:
        mean(
          rows.map(
            row =>
              row.views
          )
        ),

      medianViews:
        median(
          rows.map(
            row =>
              row.views
          )
        ),

      avgFinish:
        mean(
          rows.map(
            row =>
              row.finish_rate
          )
        ),

      avgWatch:
        mean(
          rows.map(
            row =>
              row.watch_ratio
          )
        ),

      avgR3:
        mean(
          rows.map(
            row =>
              row.retention_3s
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

  function metricRankValue(
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
      ].includes(
        metric
      )
    ) {

      return pct(value);

    }

    return compactNumber(
      value
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
        (
          row,
          index
        ) => {

          const tag =
            classifyVideo(
              row
            );

          return `
            <div class="qtk-top-item">

              <div class="qtk-rank">
                ${index + 1}
              </div>

              <div class="qtk-top-main">

                <div
                  class="qtk-top-title"
                  title="${escapeHtml(row.title)}"
                >
                  ${escapeHtml(shortTitle(row.title))}
                </div>

                <div class="qtk-top-meta">
                  ${tag.text}
                  · 播放 ${compactNumber(row.views)}
                  · 完播 ${pct(row.finish_rate)}
                </div>

              </div>

              <div class="qtk-top-value">
                ${metricRankValue(
                  metric,
                  row[metric]
                )}
              </div>

            </div>
          `;

        }
      )
      .join('');

  }

  function setTextIfChanged(
    element,
    value
  ) {

    if (!element) {

      return;

    }

    const next =
      String(value);

    if (
      element.textContent !==
      next
    ) {

      element.textContent =
        next;

    }

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

    setTextIfChanged(
      $('qtk-summary-count'),
      summary.count || 0
    );

    setTextIfChanged(
      $('qtk-summary-avg'),
      summary.avgViews === ''
        ? '-'
        : Math.round(
            summary.avgViews
          )
    );

    setTextIfChanged(
      $('qtk-summary-median'),
      summary.medianViews === ''
        ? '-'
        : Math.round(
            summary.medianViews
          )
    );

    setTextIfChanged(
      $('qtk-summary-finish'),
      pct(
        summary.avgFinish
      )
    );

    setTextIfChanged(
      $('qtk-summary-watch'),
      pct(
        summary.avgWatch
      )
    );

    setTextIfChanged(
      $('qtk-summary-r3'),
      pct(
        summary.avgR3
      )
    );

    const metric =
      $('qtk-rank-metric')
        ?.value ||
      'views';

    const html =
      renderTopList(
        rows,
        metric
      );

    const topList =
      $('qtk-top-list');

    if (
      topList &&
      topList.innerHTML !== html
    ) {

      topList.innerHTML =
        html;

    }

  }

  /* =========================================================
     CSS
  ========================================================= */

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

      /* =====================================================
         独立 Overlay
         不进入 TikTok 视频行
      ===================================================== */

      #qtk-overlay-layer {
        position:fixed !important;

        inset:0 !important;

        width:100vw !important;
        height:100vh !important;

        z-index:2147482000 !important;

        pointer-events:none !important;

        overflow:visible !important;

        contain:layout style !important;
      }

      .qtk-overlay-video {
        position:static !important;

        pointer-events:none !important;
      }

      .qtk-overlay-line {
        position:fixed !important;

        height:15px !important;

        display:flex !important;

        align-items:center !important;
        justify-content:space-around !important;

        gap:10px !important;

        padding:0 4px !important;

        pointer-events:none !important;

        white-space:nowrap !important;

        font-family:
          Arial,
          "Microsoft YaHei",
          sans-serif !important;

        z-index:2147482001 !important;

        contain:layout style !important;
      }

      .qtk-overlay-metric {
        display:inline-flex !important;

        align-items:baseline !important;
        justify-content:center !important;

        gap:3px !important;

        min-width:0 !important;

        white-space:nowrap !important;
      }

      .qtk-overlay-metric span {
        font-size:10.5px !important;

        line-height:14px !important;

        font-weight:500 !important;

        color:#8a8b91 !important;

        text-shadow:
          0 1px 0
          rgba(255,255,255,.85) !important;
      }

      .qtk-overlay-metric b {
        font-size:11.5px !important;

        line-height:14px !important;

        font-weight:700 !important;

        color:#161823 !important;

        text-shadow:
          0 1px 0
          rgba(255,255,255,.85) !important;
      }

      .qtk-overlay-dense
      .qtk-overlay-line {
        gap:5px !important;
      }

      .qtk-overlay-dense
      .qtk-overlay-metric span {
        font-size:9.5px !important;
      }

      .qtk-overlay-dense
      .qtk-overlay-metric b {
        font-size:10.5px !important;
      }


      /* =====================================================
         标签
      ===================================================== */

      .qtk-video-tag {
        position:fixed !important;

        transform:
          translateY(-50%) !important;

        z-index:2147482002 !important;

        display:inline-flex !important;

        align-items:center !important;
        justify-content:center !important;

        min-height:23px !important;

        padding:4px 9px !important;

        border-radius:999px !important;

        font-family:
          Arial,
          "Microsoft YaHei",
          sans-serif !important;

        font-size:10.5px !important;

        line-height:14px !important;

        font-weight:700 !important;

        white-space:nowrap !important;

        pointer-events:none !important;

        box-shadow:
          0 1px 3px
          rgba(0,0,0,.05) !important;
      }

      .qtk-tag-watching {
        color:#6b7280 !important;
        background:#f3f4f6 !important;
        border:1px solid #e5e7eb !important;
      }

      .qtk-tag-viral {
        color:#9a3412 !important;
        background:#ffedd5 !important;
        border:1px solid #fdba74 !important;
      }

      .qtk-tag-excellent {
        color:#92400e !important;
        background:#fef3c7 !important;
        border:1px solid #fcd34d !important;
      }

      .qtk-tag-good {
        color:#166534 !important;
        background:#dcfce7 !important;
        border:1px solid #86efac !important;
      }

      .qtk-tag-normal {
        color:#475569 !important;
        background:#f1f5f9 !important;
        border:1px solid #cbd5e1 !important;
      }

      .qtk-tag-bad {
        color:#991b1b !important;
        background:#fee2e2 !important;
        border:1px solid #fca5a5 !important;
      }


      /* =====================================================
         面板
      ===================================================== */

      #qualitell-tiktok-panel,
      #qualitell-tiktok-panel * {
        box-sizing:border-box;
      }

      #qualitell-tiktok-panel {
        position:fixed;

        right:12px;
        bottom:12px;

        z-index:2147483647;

        width:342px;

        max-height:88vh;

        overflow:hidden;

        color:#fff;

        background:
          rgba(22,22,26,.985);

        border:
          1px solid
          rgba(255,255,255,.10);

        border-radius:12px;

        box-shadow:
          0 14px 42px
          rgba(0,0,0,.36);

        font-family:
          Arial,
          "Microsoft YaHei",
          sans-serif;
      }

      .qtk-header {
        display:flex;

        align-items:center;

        justify-content:
          space-between;

        padding:
          11px 13px;

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

      .qtk-header-actions {
        display:flex;
        align-items:center;
        gap:5px;
      }

      .qtk-icon-btn {
        min-width:29px;
        height:29px;

        border:0;
        border-radius:7px;

        cursor:pointer;

        color:#fff;

        background:
          rgba(255,255,255,.08);

        font-size:13px;
      }

      .qtk-icon-btn:hover {
        background:
          rgba(255,255,255,.14);
      }

      #qtk-body {
        max-height:
          calc(88vh - 54px);

        overflow:auto;

        padding:
          11px 13px 14px;
      }

      .qtk-stat-row {
        display:flex;

        align-items:center;

        justify-content:
          space-between;

        gap:10px;

        padding:3px 0;

        font-size:11px;
      }

      .qtk-stat-row span {
        color:
          rgba(255,255,255,.55);
      }

      .qtk-stat-row b {
        text-align:right;

        font-weight:600;

        color:#fff;
      }

      .qtk-divider {
        height:1px;

        margin:10px 0;

        background:
          rgba(255,255,255,.08);
      }

      .qtk-toolbar {
        display:grid;

        grid-template-columns:
          1fr 1fr;

        gap:6px;

        margin-top:9px;
      }

      .qtk-toolbar button {
        width:100%;

        min-height:31px;

        padding:7px;

        border:0;

        border-radius:7px;

        cursor:pointer;

        font-size:11px;

        font-weight:600;
      }

      .qtk-toolbar button:disabled {
        opacity:.35;
        cursor:not-allowed;
      }

      #qtk-scan {
        color:#fff;
        background:#fe2c55;
      }

      #qtk-fetch {
        color:#111;
        background:#25f4ee;
      }

      #qtk-export-xlsx,
      #qtk-export-csv {
        color:#111;
        background:#fff;
      }

      #qtk-progress {
        height:5px;

        margin-top:9px;

        overflow:hidden;

        border-radius:999px;

        background:
          rgba(255,255,255,.10);
      }

      #qtk-progress-bar {
        width:0%;
        height:100%;

        background:#25f4ee;

        transition:
          width .2s ease;
      }

      .qtk-section-head {
        display:flex;

        align-items:center;

        justify-content:
          space-between;

        gap:8px;

        margin-bottom:7px;

        font-size:12px;

        font-weight:700;
      }

      .qtk-select {
        min-height:29px;

        padding:5px 7px;

        border:0;

        border-radius:6px;

        outline:none;

        color:#fff;

        background:#303036;

        font-size:10.5px;
      }

      .qtk-summary-grid {
        display:grid;

        grid-template-columns:
          repeat(3,1fr);

        gap:5px;
      }

      .qtk-card {
        padding:7px;

        border-radius:7px;

        background:#2a2a30;

        border:
          1px solid
          rgba(255,255,255,.05);
      }

      .qtk-card span {
        display:block;

        margin-bottom:3px;

        font-size:9px;

        color:
          rgba(255,255,255,.43);
      }

      .qtk-card b {
        font-size:13px;
      }

      .qtk-rank-select {
        margin:8px 0 5px;
      }

      .qtk-rank-select select {
        width:100%;
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
        width:18px;
        height:18px;

        flex:none;

        display:flex;

        align-items:center;
        justify-content:center;

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
        padding:7px 0;

        font-size:10px;

        color:
          rgba(255,255,255,.42);
      }


      /* =====================================================
         设置
      ===================================================== */

      #qtk-settings {
        display:none;
      }

      #qtk-settings.qtk-open {
        display:block;
      }

      .qtk-setting-block {
        padding:9px 0;

        border-bottom:
          1px solid
          rgba(255,255,255,.07);
      }

      .qtk-setting-block:last-child {
        border-bottom:0;
      }

      .qtk-setting-title {
        margin-bottom:7px;

        font-size:11px;
        font-weight:700;
      }

      .qtk-setting-line {
        display:flex;

        align-items:center;

        justify-content:
          space-between;

        gap:8px;

        min-height:28px;

        font-size:10.5px;
      }

      .qtk-setting-line label {
        color:
          rgba(255,255,255,.68);
      }

      .qtk-setting-inline {
        display:flex;

        align-items:center;

        gap:5px;
      }

      .qtk-setting-inline select {
        width:65px;
      }

      .qtk-switch {
        width:35px;
        height:19px;

        position:relative;

        display:inline-block;

        flex:none;
      }

      .qtk-switch input {
        display:none;
      }

      .qtk-switch-slider {
        position:absolute;

        inset:0;

        cursor:pointer;

        border-radius:999px;

        background:#4a4a52;

        transition:.18s ease;
      }

      .qtk-switch-slider:before {
        content:"";

        position:absolute;

        width:15px;
        height:15px;

        left:2px;
        top:2px;

        border-radius:50%;

        background:#fff;

        transition:.18s ease;
      }

      .qtk-switch input:checked +
      .qtk-switch-slider {
        background:#25f4ee;
      }

      .qtk-switch input:checked +
      .qtk-switch-slider:before {
        transform:
          translateX(16px);
      }

      .qtk-metric-options {
        display:grid;

        grid-template-columns:
          repeat(2,1fr);

        gap:4px 6px;
      }

      .qtk-check-option {
        display:flex;

        align-items:center;

        gap:5px;

        min-height:24px;

        padding:3px 5px;

        border-radius:5px;

        cursor:pointer;

        font-size:9.5px;

        color:
          rgba(255,255,255,.72);

        background:
          rgba(255,255,255,.035);
      }

      .qtk-check-option:hover {
        background:
          rgba(255,255,255,.07);
      }

      .qtk-check-option input {
        margin:0;
      }

      .qtk-setting-note {
        margin-top:6px;

        font-size:9px;

        line-height:1.5;

        color:
          rgba(255,255,255,.38);
      }

      .qtk-small-button {
        min-height:28px;

        padding:5px 9px;

        border:0;

        border-radius:6px;

        cursor:pointer;

        font-size:10px;

        color:#111;

        background:#fff;
      }

      .qtk-small-button-dark {
        color:#fff;

        background:#3a3a42;
      }

      .qtk-status-message {
        max-width:190px;

        overflow:hidden;

        text-overflow:ellipsis;

        white-space:nowrap;
      }

    `;

    (
      document.head ||
      document.documentElement
    ).appendChild(
      style
    );

  }

  /* =========================================================
     面板
  ========================================================= */

  function minuteOptions(
    selected
  ) {

    let html = '';

    for (
      let i = 1;
      i <= 30;
      i++
    ) {

      html += `
        <option
          value="${i}"
          ${
            Number(selected) === i
              ? 'selected'
              : ''
          }
        >
          ${i}分
        </option>
      `;

    }

    return html;

  }

  function cleanupOptions(
    selected
  ) {

    return [
      15,
      30,
      60
    ]
      .map(
        value => `
          <option
            value="${value}"
            ${
              Number(selected) === value
                ? 'selected'
                : ''
            }
          >
            ${value}分钟
          </option>
        `
      )
      .join('');

  }

  function buildMetricCheckboxes() {

    return METRIC_ORDER
      .map(
        key => {

          const metric =
            METRICS[key];

          const checked =
            settings
              .selectedMetrics
              .includes(key);

          return `
            <label class="qtk-check-option">

              <input
                type="checkbox"
                class="qtk-metric-check"
                data-metric="${key}"
                ${checked ? 'checked' : ''}
              >

              <span>
                ${metric.full}
              </span>

            </label>
          `;

        }
      )
      .join('');

  }

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

    ensureOverlayLayer();

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
            V${VERSION}
          </div>

        </div>


        <div class="qtk-header-actions">

          <button
            id="qtk-settings-btn"
            class="qtk-icon-btn"
            title="设置"
          >
            ⚙
          </button>

          <button
            id="qtk-minimize"
            class="qtk-icon-btn"
            title="最小化"
          >
            −
          </button>

        </div>

      </div>


      <div id="qtk-body">

        <div id="qtk-main">

          <div class="qtk-stat-row">
            <span>发现作品</span>
            <b id="qtk-all">0</b>
          </div>

          <div class="qtk-stat-row">
            <span>预约发布</span>
            <b id="qtk-scheduled">0</b>
          </div>

          <div class="qtk-stat-row">
            <span>已获取公开数据</span>
            <b id="qtk-public">0</b>
          </div>

          <div class="qtk-stat-row">
            <span>非公开</span>
            <b id="qtk-private">0</b>
          </div>

          <div class="qtk-stat-row">
            <span>失败/未知</span>
            <b id="qtk-failed">0</b>
          </div>

          <div class="qtk-stat-row">

            <span>
              状态
            </span>

            <b
              id="qtk-status"
              class="qtk-status-message"
            >
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

          </div>


          <div id="qtk-progress">

            <div id="qtk-progress-bar"></div>

          </div>


          <div class="qtk-divider"></div>


          <div id="qtk-analysis">

            <div class="qtk-section-head">

              <span>
                数据分析
              </span>

              <select
                id="qtk-range"
                class="qtk-select"
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


            <div class="qtk-rank-select">

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
                正在自动采集…
              </div>

            </div>

          </div>

        </div>


        <div id="qtk-settings">

          <div class="qtk-setting-block">

            <div class="qtk-setting-title">
              页面显示
            </div>


            <div class="qtk-setting-line">

              <label>
                页面指标
              </label>

              <label class="qtk-switch">

                <input
                  type="checkbox"
                  id="qtk-page-enabled"
                  ${
                    settings.pageMetricsEnabled
                      ? 'checked'
                      : ''
                  }
                >

                <span class="qtk-switch-slider"></span>

              </label>

            </div>


            <div class="qtk-setting-line">

              <label>
                视频标签
              </label>

              <label class="qtk-switch">

                <input
                  type="checkbox"
                  id="qtk-tag-enabled"
                  ${
                    settings.videoTagEnabled
                      ? 'checked'
                      : ''
                  }
                >

                <span class="qtk-switch-slider"></span>

              </label>

            </div>

          </div>


          <div class="qtk-setting-block">

            <div class="qtk-setting-title">
              显示指标
            </div>


            <div
              id="qtk-metric-options"
              class="qtk-metric-options"
            >
              ${buildMetricCheckboxes()}
            </div>


            <div class="qtk-setting-note">
              默认8项，最多同时显示14项。
              视频标签不计入14项。
            </div>

          </div>


          <div class="qtk-setting-block">

            <div class="qtk-setting-title">
              自动刷新
            </div>


            <div class="qtk-setting-line">

              <label>
                自动刷新
              </label>

              <label class="qtk-switch">

                <input
                  type="checkbox"
                  id="qtk-auto-enabled"
                  ${
                    settings.autoRefreshEnabled
                      ? 'checked'
                      : ''
                  }
                >

                <span class="qtk-switch-slider"></span>

              </label>

            </div>


            <div class="qtk-setting-line">

              <label>
                随机间隔
              </label>

              <div class="qtk-setting-inline">

                <select
                  id="qtk-refresh-min"
                  class="qtk-select"
                >
                  ${minuteOptions(
                    settings
                      .refreshMinMinutes
                  )}
                </select>

                <span>～</span>

                <select
                  id="qtk-refresh-max"
                  class="qtk-select"
                >
                  ${minuteOptions(
                    settings
                      .refreshMaxMinutes
                  )}
                </select>

              </div>

            </div>


            <div class="qtk-setting-line">

              <label>
                后台暂停
              </label>

              <label class="qtk-switch">

                <input
                  type="checkbox"
                  id="qtk-pause-hidden"
                  ${
                    settings.pauseWhenHidden
                      ? 'checked'
                      : ''
                  }
                >

                <span class="qtk-switch-slider"></span>

              </label>

            </div>


            <div class="qtk-setting-line">

              <label>
                最后更新
              </label>

              <b id="qtk-last-refresh">
                -
              </b>

            </div>


            <div class="qtk-setting-line">

              <label>
                下次更新
              </label>

              <b id="qtk-next-refresh">
                -
              </b>

            </div>


            <div class="qtk-setting-line">

              <button
                id="qtk-refresh-now"
                class="qtk-small-button"
              >
                ↻ 立即刷新
              </button>

            </div>

          </div>


          <div class="qtk-setting-block">

            <div class="qtk-setting-title">
              性能设置
            </div>


            <div class="qtk-setting-line">

              <label>
                自动清理缓存
              </label>

              <label class="qtk-switch">

                <input
                  type="checkbox"
                  id="qtk-cleanup-enabled"
                  ${
                    settings.autoCleanupEnabled
                      ? 'checked'
                      : ''
                  }
                >

                <span class="qtk-switch-slider"></span>

              </label>

            </div>


            <div class="qtk-setting-line">

              <label>
                清理周期
              </label>

              <select
                id="qtk-cleanup-minutes"
                class="qtk-select"
              >

                ${cleanupOptions(
                  settings
                    .cleanupMinutes
                )}

              </select>

            </div>


            <div class="qtk-setting-line">

              <button
                id="qtk-clean-now"
                class="qtk-small-button qtk-small-button-dark"
              >
                立即清理缓存
              </button>

            </div>


            <div class="qtk-setting-note">
              只清理插件自身临时缓存，不清 TikTok Cookie、
              登录状态或浏览器网页缓存。
            </div>

          </div>

        </div>

      </div>

    `;

    document
      .documentElement
      .appendChild(
        panel
      );

    bindPanelEvents();

    updateUI(
      '自动加载中'
    );

    updateRefreshUI();

  }

  /* =========================================================
     面板事件
  ========================================================= */

  function bindPanelEvents() {

    $('qtk-minimize')
      ?.addEventListener(
        'click',
        () => {

          minimized =
            !minimized;

          $('qtk-body')
            .style
            .display =
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


    $('qtk-settings-btn')
      ?.addEventListener(
        'click',
        () => {

          const settingsPanel =
            $('qtk-settings');

          const main =
            $('qtk-main');

          const open =
            settingsPanel
              .classList
              .toggle(
                'qtk-open'
              );

          main.style.display =
            open
              ? 'none'
              : '';

        }
      );


    $('qtk-scan')
      ?.addEventListener(
        'click',
        scanAll
      );


    $('qtk-fetch')
      ?.addEventListener(
        'click',
        fetchAllInsights
      );


    $('qtk-export-xlsx')
      ?.addEventListener(
        'click',
        exportXlsx
      );


    $('qtk-export-csv')
      ?.addEventListener(
        'click',
        exportCsv
      );


    $('qtk-range')
      ?.addEventListener(
        'change',
        renderAnalysis
      );


    $('qtk-rank-metric')
      ?.addEventListener(
        'change',
        renderAnalysis
      );


    $('qtk-page-enabled')
      ?.addEventListener(
        'change',
        event => {

          settings
            .pageMetricsEnabled =
            event.target.checked;

          saveSettings();

          if (
            !settings
              .pageMetricsEnabled
          ) {

            clearAllOverlays();

          } else {

            scheduleVisibleMetrics(
              50
            );

          }

        }
      );


    $('qtk-tag-enabled')
      ?.addEventListener(
        'change',
        event => {

          settings
            .videoTagEnabled =
            event.target.checked;

          saveSettings();

          scheduleOverlaySync();

        }
      );


    document
      .querySelectorAll(
        '.qtk-metric-check'
      )
      .forEach(
        checkbox => {

          checkbox.addEventListener(
            'change',
            () => {

              const key =
                checkbox.dataset
                  .metric;

              const selected =
                new Set(
                  settings
                    .selectedMetrics
                );

              if (
                checkbox.checked
              ) {

                if (
                  selected.size >=
                  MAX_VISIBLE_METRICS
                ) {

                  checkbox.checked =
                    false;

                  updateUI(
                    `最多显示 ${MAX_VISIBLE_METRICS} 项`
                  );

                  return;

                }

                selected.add(
                  key
                );

              } else {

                selected.delete(
                  key
                );

              }

              settings
                .selectedMetrics =
                METRIC_ORDER.filter(
                  metric =>
                    selected.has(
                      metric
                    )
                );

              saveSettings();

              for (
                const overlay
                of overlayCache.values()
              ) {

                overlay.layoutKey =
                  '';

              }

              scheduleOverlaySync();

            }
          );

        }
      );


    $('qtk-auto-enabled')
      ?.addEventListener(
        'change',
        event => {

          settings
            .autoRefreshEnabled =
            event.target.checked;

          saveSettings();

          nextRefreshAt = 0;

          scheduleNextAutoRefresh();

        }
      );


    $('qtk-refresh-min')
      ?.addEventListener(
        'change',
        event => {

          settings
            .refreshMinMinutes =
            Number(
              event.target.value
            );

          if (
            settings
              .refreshMinMinutes >
            settings
              .refreshMaxMinutes
          ) {

            settings
              .refreshMaxMinutes =
              settings
                .refreshMinMinutes;

            $('qtk-refresh-max')
              .value =
              String(
                settings
                  .refreshMaxMinutes
              );

          }

          saveSettings();

          nextRefreshAt = 0;

          scheduleNextAutoRefresh();

        }
      );


    $('qtk-refresh-max')
      ?.addEventListener(
        'change',
        event => {

          settings
            .refreshMaxMinutes =
            Number(
              event.target.value
            );

          if (
            settings
              .refreshMaxMinutes <
            settings
              .refreshMinMinutes
          ) {

            settings
              .refreshMinMinutes =
              settings
                .refreshMaxMinutes;

            $('qtk-refresh-min')
              .value =
              String(
                settings
                  .refreshMinMinutes
              );

          }

          saveSettings();

          nextRefreshAt = 0;

          scheduleNextAutoRefresh();

        }
      );


    $('qtk-pause-hidden')
      ?.addEventListener(
        'change',
        event => {

          settings
            .pauseWhenHidden =
            event.target.checked;

          saveSettings();

          scheduleNextAutoRefresh(
            true
          );

        }
      );


    $('qtk-refresh-now')
      ?.addEventListener(
        'click',
        () =>
          refreshVisibleMetrics(
            'manual'
          )
      );


    $('qtk-cleanup-enabled')
      ?.addEventListener(
        'change',
        event => {

          settings
            .autoCleanupEnabled =
            event.target.checked;

          saveSettings();

        }
      );


    $('qtk-cleanup-minutes')
      ?.addEventListener(
        'change',
        event => {

          settings
            .cleanupMinutes =
            Number(
              event.target.value
            );

          saveSettings();

          lastCleanupAt =
            Date.now();

        }
      );


    $('qtk-clean-now')
      ?.addEventListener(
        'click',
        () =>
          cleanupPluginCache(
            true
          )
      );

  }

  /* =========================================================
     UI
  ========================================================= */

  function updateUI(
    status
  ) {

    if (
      !$(
        'qualitell-tiktok-panel'
      )
    ) {

      return;

    }

    setTextIfChanged(
      $('qtk-all'),
      items.size
    );

    setTextIfChanged(
      $('qtk-scheduled'),
      scheduledItems()
        .length
    );

    setTextIfChanged(
      $('qtk-public'),
      detailRows.size
    );

    setTextIfChanged(
      $('qtk-private'),
      privateIds.size
    );

    setTextIfChanged(
      $('qtk-failed'),
      failedCount +
      unknownIds.size
    );

    if (status) {

      setTextIfChanged(
        $('qtk-status'),
        status
      );

    }

    $('qtk-scan')
      .disabled =
      busy;

    $('qtk-fetch')
      .disabled =
      busy ||
      publishedItems()
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

    updateRefreshUI();

  }

  /* =========================================================
     扫描全部
  ========================================================= */

  async function scanAll() {

    if (busy) {

      return;

    }

    busy = true;

    updateUI(
      '扫描中…'
    );

    let stable = 0;

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

      const currentCount =
        items.size;

      const currentHeight =
        document
          .documentElement
          .scrollHeight;

      if (
        currentCount ===
          lastCount &&
        currentHeight ===
          lastHeight
      ) {

        stable++;

      } else {

        stable = 0;

      }

      lastCount =
        currentCount;

      lastHeight =
        currentHeight;

      updateUI(
        `扫描 ${currentCount} 条`
      );

      if (
        stable >= 6
      ) {

        break;

      }

    }

    busy = false;

    invalidateRows();

    updateUI(
      `扫描完成 ${items.size} 条`
    );

    scheduleVisibleMetrics(
      150
    );

  }

  /* =========================================================
     获取全部数据
  ========================================================= */

  async function fetchAllInsights() {

    if (
      busy ||
      !items.size
    ) {

      return;

    }

    busy = true;

    failedCount = 0;

    privateIds.clear();

    unknownIds.clear();

    const list =
      publishedItems();

    let completed = 0;

    const progress =
      $('qtk-progress-bar');

    if (progress) {

      progress.style.width =
        '0%';

    }

    updateUI(
      `获取全部数据 0/${list.length}`
    );

    try {

      await runWorkers(

        list,

        async item => {

          await fetchItemMetric(
            item,
            true
          );

          completed++;

          if (progress) {

            progress.style.width =
              `${Math.round(
                completed /
                list.length *
                100
              )}%`;

          }

          updateUI(
            `获取全部数据 ${completed}/${list.length}`
          );

        },

        2

      );

      lastRefreshAt =
        Date.now();

      updateUI(
        `完成 ${detailRows.size} 条`
      );

    } finally {

      busy = false;

      setTimeout(
        () => {

          if (progress) {

            progress.style.width =
              '0%';

          }

        },
        800
      );

      scheduleNextAutoRefresh();

      scheduleOverlaySync();

    }

  }

  /* =========================================================
     CSV
  ========================================================= */

  const CSV_COLUMNS = [

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

  const PERCENT_KEYS =
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
    ].sort(
      (a, b) =>
        b.publish_ts -
        a.publish_ts
    );

    if (!rows.length) {

      return;

    }

    const columns = [

      [
        '视频标签',
        '__tag'
      ],

      ...CSV_COLUMNS,

      [
        '视频链接',
        '__url'
      ]

    ];

    const lines = [

      columns
        .map(
          ([label]) =>
            csvEscape(
              label
            )
        )
        .join(',')

    ];

    for (
      const row
      of rows
    ) {

      lines.push(

        columns
          .map(
            ([, key]) => {

              if (
                key ===
                '__tag'
              ) {

                return csvEscape(

                  classifyVideo(
                    row
                  ).text

                );

              }

              if (
                key ===
                '__url'
              ) {

                return csvEscape(

                  getVideoUrl(
                    row
                  )

                );

              }

              if (
                PERCENT_KEYS.has(
                  key
                )
              ) {

                return csvEscape(
                  pct(
                    row[key]
                  )
                );

              }

              return csvEscape(
                row[key]
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
          lines.join(
            '\r\n'
          )
        ],

        {
          type:
            'text/csv;charset=utf-8'
        }

      );

    const account =
      rows[0]
        ?.account ||
      'tiktok';

    const date =
      new Date()
        .toISOString()
        .slice(
          0,
          10
        );

    downloadBlob(

      blob,

      `${account}_TikTok数据_${date}.csv`

    );

  }

  /* =========================================================
     Excel 封面
  ========================================================= */

  function gmFetchBlob(
    url
  ) {

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
                  300 ||
                !response.response
              ) {

                reject(
                  new Error(
                    `cover http ${response.status}`
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
            () =>
              reject(
                new Error(
                  'cover error'
                )
              ),

          ontimeout:
            () =>
              reject(
                new Error(
                  'cover timeout'
                )
              )

        });

      }
    );

  }

  async function blobToThumbnail(
    blob
  ) {

    const bitmap =
      await createImageBitmap(
        blob
      );

    try {

      const scale =
        Math.min(

          60 /
          bitmap.width,

          95 /
          bitmap.height,

          1

        );

      const canvas =
        document.createElement(
          'canvas'
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

      ctx.fillStyle =
        '#ffffff';

      ctx.fillRect(

        0,
        0,
        canvas.width,
        canvas.height

      );

      ctx.drawImage(

        bitmap,
        0,
        0,
        canvas.width,
        canvas.height

      );

      return canvas.toDataURL(

        'image/jpeg',
        0.72

      );

    } finally {

      bitmap.close();

    }

  }

  async function getCoverData(
    row
  ) {

    const url =
      row?.cover_url ||
      '';

    if (!url) {

      return '';

    }

    if (
      coverCache.has(
        url
      )
    ) {

      return coverCache.get(
        url
      );

    }

    const promise =
      (
        async () => {

          try {

            const blob =
              await gmFetchBlob(
                url
              );

            return await blobToThumbnail(
              blob
            );

          } catch {

            return '';

          }

        }
      )();

    coverCache.set(
      url,
      promise
    );

    return promise;

  }

  async function prepareExcelImages(
    rows,
    workbook
  ) {

    const map =
      new Map();

    let index = 0;

    let done = 0;

    const workers =
      Math.min(
        3,
        Math.max(
          1,
          rows.length
        )
      );

    async function worker() {

      while (true) {

        const current =
          index++;

        if (
          current >=
          rows.length
        ) {

          break;

        }

        const row =
          rows[current];

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
          `准备封面 ${done}/${rows.length}`
        );

      }

    }

    await Promise.all(

      Array.from(
        {
          length:
            workers
        },
        () =>
          worker()
      )

    );

    return map;

  }

  /* =========================================================
     Excel 样式
  ========================================================= */

  function excelBorder() {

    const side = {

      style:
        'thin',

      color: {
        argb:
          'FFE2E8F0'
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

  function styleExcelHeader(
    row
  ) {

    row.height =
      26;

    row.eachCell(
      {
        includeEmpty:
          true
      },
      cell => {

        cell.font = {

          bold:
            true,

          size:
            10.5,

          color: {
            argb:
              'FFFFFFFF'
          }

        };

        cell.fill = {

          type:
            'pattern',

          pattern:
            'solid',

          fgColor: {
            argb:
              'FF0F172A'
          }

        };

        cell.alignment = {

          vertical:
            'middle',

          horizontal:
            'center',

          wrapText:
            true

        };

      }
    );

  }

  function styleExcelDataRow(
    row,
    zebra
  ) {

    row.eachCell(
      {
        includeEmpty:
          true
      },
      cell => {

        cell.font = {

          size:
            10,

          color: {
            argb:
              'FF334155'
          }

        };

        cell.border =
          excelBorder();

        cell.alignment = {

          vertical:
            'middle',

          wrapText:
            true

        };

        if (zebra) {

          cell.fill = {

            type:
              'pattern',

            pattern:
              'solid',

            fgColor: {
              argb:
                'FFF8FAFC'
            }

          };

        }

      }
    );

  }

  function setExcelLink(
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
        url,

      tooltip:
        '打开 TikTok 视频'

    };

    cell.font = {

      size:
        10,

      color: {
        argb:
          'FF2563EB'
      },

      underline:
        true

    };

  }

  function excelTagStyle(
    result
  ) {

    const map = {

      viral: {
        fill:
          'FFFFEDD5',
        font:
          'FF9A3412'
      },

      excellent: {
        fill:
          'FFFEF3C7',
        font:
          'FF92400E'
      },

      good: {
        fill:
          'FFDCFCE7',
        font:
          'FF166534'
      },

      bad: {
        fill:
          'FFFEE2E2',
        font:
          'FF991B1B'
      },

      normal: {
        fill:
          'FFF1F5F9',
        font:
          'FF475569'
      },

      watching: {
        fill:
          'FFF3F4F6',
        font:
          'FF6B7280'
      }

    };

    return (
      map[result.key] ||
      map.normal
    );

  }

  /* =========================================================
     Excel 原始数据
  ========================================================= */

  function buildRawSheet(
    workbook,
    rows,
    imageIds
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
                3,

              ySplit:
                1
            }
          ]
        }
      );

    sheet.columns = [

      {
        header:
          '封面',
        key:
          'cover',
        width:
          8
      },

      {
        header:
          '标签',
        key:
          'tag',
        width:
          14
      },

      {
        header:
          '标题',
        key:
          'title',
        width:
          44
      },

      {
        header:
          '账号',
        key:
          'account',
        width:
          17
      },

      {
        header:
          '发布时间',
        key:
          'publish_time',
        width:
          20
      },

      {
        header:
          '时长(s)',
        key:
          'duration_sec',
        width:
          10
      },

      {
        header:
          '播放',
        key:
          'views',
        width:
          11
      },

      {
        header:
          '点赞',
        key:
          'likes',
        width:
          9
      },

      {
        header:
          '评论',
        key:
          'comments',
        width:
          9
      },

      {
        header:
          '分享',
        key:
          'shares',
        width:
          9
      },

      {
        header:
          '收藏',
        key:
          'favorites',
        width:
          9
      },

      {
        header:
          '新粉',
        key:
          'new_followers',
        width:
          9
      },

      {
        header:
          '平均观看(s)',
        key:
          'avg_watch_sec',
        width:
          13
      },

      {
        header:
          '观看倍率',
        key:
          'watch_ratio',
        width:
          11
      },

      {
        header:
          '完播率',
        key:
          'finish_rate',
        width:
          11
      },

      {
        header:
          '总观看(s)',
        key:
          'total_watch_sec',
        width:
          13
      },

      {
        header:
          '1秒留存',
        key:
          'retention_1s',
        width:
          11
      },

      {
        header:
          '2秒留存',
        key:
          'retention_2s',
        width:
          11
      },

      {
        header:
          '3秒留存',
        key:
          'retention_3s',
        width:
          11
      },

      {
        header:
          '5秒留存',
        key:
          'retention_5s',
        width:
          11
      },

      {
        header:
          '10秒留存',
        key:
          'retention_10s',
        width:
          11
      },

      {
        header:
          'For You',
        key:
          'for_you',
        width:
          11
      },

      {
        header:
          '个人主页',
        key:
          'personal_profile',
        width:
          11
      },

      {
        header:
          '搜索',
        key:
          'search',
        width:
          10
      },

      {
        header:
          '点赞率',
        key:
          'like_rate',
        width:
          10
      },

      {
        header:
          '互动率',
        key:
          'engagement_rate',
        width:
          10
      },

      {
        header:
          'Video ID',
        key:
          'video_id',
        width:
          22
      },

      {
        header:
          '视频',
        key:
          'link',
        width:
          12
      }

    ];

    styleExcelHeader(
      sheet.getRow(1)
    );

    sheet.autoFilter = {

      from: {
        row:
          1,
        column:
          1
      },

      to: {
        row:
          1,
        column:
          sheet.columnCount
      }

    };

    rows.forEach(
      (
        data,
        index
      ) => {

        const result =
          classifyVideo(
            data
          );

        const videoUrl =
          getVideoUrl(
            data
          );

        const row =
          sheet.addRow({

            cover:
              '',

            tag:
              result.text,

            title:
              data.title,

            account:
              data.account,

            publish_time:
              data.publish_time,

            duration_sec:
              data.duration_sec,

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

            total_watch_sec:
              data.total_watch_sec,

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

            personal_profile:
              data.personal_profile,

            search:
              data.search,

            like_rate:
              data.like_rate,

            engagement_rate:
              data.engagement_rate,

            video_id:
              data.video_id,

            link:
              '打开视频'

          });

        row.height =
          58;

        styleExcelDataRow(
          row,
          index % 2 === 1
        );

        const tagStyle =
          excelTagStyle(
            result
          );

        row
          .getCell(
            'tag'
          )
          .fill = {

            type:
              'pattern',

            pattern:
              'solid',

            fgColor: {
              argb:
                tagStyle.fill
            }

          };

        row
          .getCell(
            'tag'
          )
          .font = {

            bold:
              true,

            color: {
              argb:
                tagStyle.font
            }

          };

        setExcelLink(

          row.getCell(
            'title'
          ),

          data.title ||
          '(无标题)',

          videoUrl

        );

        setExcelLink(

          row.getCell(
            'link'
          ),

          '打开视频',

          videoUrl

        );

        const imageId =
          imageIds.get(
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
                  0.16,

                row:
                  row.number -
                  1 +
                  0.08
              },

              ext: {
                width:
                  42,

                height:
                  74
              },

              editAs:
                'oneCell'

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

      'personal_profile',

      'search',

      'like_rate',

      'engagement_rate'

    ].forEach(
      key => {

        sheet
          .getColumn(
            key
          )
          .numFmt =
          '0.0%';

      }
    );

    return sheet;

  }

  /* =========================================================
     Excel 分析页
  ========================================================= */

  function buildAnalysisSheet(
    workbook,
    rows
  ) {

    const sheet =
      workbook.addWorksheet(
        '本期分析'
      );

    sheet.columns = [

      {
        width:
          45
      },

      {
        width:
          15
      },

      {
        width:
          15
      },

      {
        width:
          15
      },

      {
        width:
          15
      },

      {
        width:
          15
      }

    ];

    sheet.mergeCells(
      'A1:F1'
    );

    sheet.getCell(
      'A1'
    ).value =
      'TikTok 内容表现分析';

    sheet.getCell(
      'A1'
    ).font = {

      bold:
        true,

      size:
        18,

      color: {
        argb:
          'FFFFFFFF'
      }

    };

    sheet.getCell(
      'A1'
    ).fill = {

      type:
        'pattern',

      pattern:
        'solid',

      fgColor: {
        argb:
          'FF0F172A'
      }

    };

    sheet.getRow(
      1
    ).height =
      32;

    const summary =
      summaryFor(
        rows
      );

    const data = [

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
        summary.avgWatch
      ],

      [
        '平均3秒留存',
        summary.avgR3
      ]

    ];

    data.forEach(
      values =>
        sheet.addRow(
          values
        )
    );

    sheet.getCell(
      'B5'
    ).numFmt =
      '0.0%';

    sheet.getCell(
      'B6'
    ).numFmt =
      '0.0%';

    sheet.getCell(
      'B7'
    ).numFmt =
      '0.0%';

    sheet.addRow([]);

    const header =
      sheet.addRow([

        '视频标题',

        '标签',

        '播放',

        '完播率',

        '观看倍率',

        '3秒留存'

      ]);

    styleExcelHeader(
      header
    );

    const sorted =
      [...rows].sort(
        (a, b) =>
          Number(
            b.views
          ) -
          Number(
            a.views
          )
      );

    for (
      const dataRow
      of sorted
    ) {

      const tag =
        classifyVideo(
          dataRow
        );

      const row =
        sheet.addRow([

          dataRow.title,

          tag.text,

          dataRow.views,

          dataRow.finish_rate,

          dataRow.watch_ratio,

          dataRow.retention_3s

        ]);

      row.getCell(
        4
      ).numFmt =
        '0.0%';

      row.getCell(
        5
      ).numFmt =
        '0.0%';

      row.getCell(
        6
      ).numFmt =
        '0.0%';

      styleExcelDataRow(
        row,
        row.number % 2 === 0
      );

      setExcelLink(

        row.getCell(
          1
        ),

        dataRow.title,

        getVideoUrl(
          dataRow
        )

      );

    }

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

    if (
      typeof ExcelJS ===
      'undefined'
    ) {

      alert(
        'ExcelJS 未加载，请刷新页面后重试。'
      );

      return;

    }

    busy = true;

    updateUI(
      '生成 XLSX…'
    );

    try {

      const workbook =
        new ExcelJS
          .Workbook();

      workbook.creator =
        'Qualitell TikTok Studio Collector';

      workbook.created =
        new Date();

      const imageIds =
        await prepareExcelImages(
          rows,
          workbook
        );

      buildRawSheet(
        workbook,
        rows,
        imageIds
      );

      buildAnalysisSheet(
        workbook,
        getAnalysisRows()
      );

      const buffer =
        await workbook
          .xlsx
          .writeBuffer();

      const blob =
        new Blob(
          [
            buffer
          ],
          {
            type:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }
        );

      const account =
        rows[0]
          ?.account ||
        'tiktok';

      const date =
        new Date()
          .toISOString()
          .slice(
            0,
            10
          );

      downloadBlob(

        blob,

        `${account}_TikTok周报_${date}.xlsx`

      );

      updateUI(
        `XLSX 已导出 ${rows.length} 条`
      );

    } catch (error) {

      console.error(
        '[TikTok Collector]',
        error
      );

      alert(
        'XLSX 导出失败'
      );

      updateUI(
        'XLSX 导出失败'
      );

    } finally {

      coverCache.clear();

      busy = false;

    }

  }

  /* =========================================================
     Mutation Observer
     只观察 TikTok 自己的 DOM 变化
  ========================================================= */

  function isPluginNode(
    node
  ) {

    if (!node) {

      return false;

    }

    if (
      node.nodeType ===
      Node.TEXT_NODE
    ) {

      return isPluginNode(
        node.parentElement
      );

    }

    if (
      node.nodeType !==
      Node.ELEMENT_NODE
    ) {

      return false;

    }

    const element =
      node;

    return Boolean(

      element.id ===
        'qualitell-tiktok-panel' ||

      element.id ===
        'qtk-overlay-layer' ||

      element.id ===
        'qualitell-tiktok-style' ||

      element.closest?.(
        '#qualitell-tiktok-panel'
      ) ||

      element.closest?.(
        '#qtk-overlay-layer'
      )

    );

  }

  function mutationIsPluginOnly(
    mutation
  ) {

    if (
      isPluginNode(
        mutation.target
      )
    ) {

      return true;

    }

    const changed = [

      ...mutation.addedNodes,

      ...mutation.removedNodes

    ];

    return (
      changed.length > 0 &&
      changed.every(
        isPluginNode
      )
    );

  }

  function startDomObserver() {

    if (
      domObserver ||
      !document.body
    ) {

      return;

    }

    domObserver =
      new MutationObserver(
        mutations => {

          let tiktokChanged =
            false;

          for (
            const mutation
            of mutations
          ) {

            if (
              mutationIsPluginOnly(
                mutation
              )
            ) {

              continue;

            }

            tiktokChanged =
              true;

            break;

          }

          if (
            !tiktokChanged
          ) {

            return;

          }

          /*
            TikTok 虚拟列表有变化：
            失效行缓存，重新匹配。
          */

          invalidateRows();

          scheduleOverlaySync();

          scheduleVisibleMetrics(
            500
          );

        }
      );

    domObserver.observe(
      document.body,
      {

        childList:
          true,

        subtree:
          true

      }
    );


    /*
      滚动：
      只移动 Overlay。
      不访问 Insight API。
    */

    window.addEventListener(
      'scroll',
      () => {

        scheduleOverlaySync();

      },
      {
        passive:
          true
      }
    );


    let resizeTimer =
      null;

    window.addEventListener(
      'resize',
      () => {

        clearTimeout(
          resizeTimer
        );

        resizeTimer =
          setTimeout(
            () => {

              invalidateRows();

              refreshRowCache(
                true
              );

              getHeaderGeometry(
                true
              );

              scheduleOverlaySync();

            },
            180
          );

      },
      {
        passive:
          true
      }
    );

  }

  /* =========================================================
     后台暂停
  ========================================================= */

  function handleVisibilityChange() {

    if (
      !settings
        .pauseWhenHidden
    ) {

      return;

    }

    if (
      document.hidden
    ) {

      clearAutoRefreshTimer();

      updateRefreshUI();

      return;

    }

    invalidateRows();

    refreshRowCache(
      true
    );

    scheduleVisibleMetrics(
      200
    );

    scheduleOverlaySync();

    if (
      settings
        .autoRefreshEnabled
    ) {

      if (
        nextRefreshAt &&
        nextRefreshAt <=
          Date.now()
      ) {

        refreshVisibleMetrics(
          'auto'
        );

      } else {

        scheduleNextAutoRefresh(
          true
        );

      }

    }

  }

  /* =========================================================
     关闭清理
  ========================================================= */

  function cleanupBeforeUnload() {

    clearTimeout(
      metricDebounceTimer
    );

    clearAutoRefreshTimer();

    if (
      countdownTimer
    ) {

      clearInterval(
        countdownTimer
      );

    }

    if (
      cleanupTimer
    ) {

      clearInterval(
        cleanupTimer
      );

    }

    if (
      overlayFrame
    ) {

      cancelAnimationFrame(
        overlayFrame
      );

    }

    metricLoading.clear();

    metricCache.clear();

    rowCache.clear();

    overlayCache.clear();

    coverCache.clear();

    domObserver
      ?.disconnect();

  }

  /* =========================================================
     启动
  ========================================================= */

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

    const start =
      () => {

        injectStyle();

        ensureOverlayLayer();

        buildPanel();

        startDomObserver();

        startCountdownTimer();

        startCleanupTimer();

        document.addEventListener(
          'visibilitychange',
          handleVisibilityChange
        );

        window.addEventListener(
          'beforeunload',
          cleanupBeforeUnload
        );

        /*
          等 TikTok Studio 首次列表加载
        */

        setTimeout(
          () => {

            invalidateRows();

            refreshRowCache(
              true
            );

            getHeaderGeometry(
              true
            );

            scheduleVisibleMetrics(
              100
            );

            scheduleOverlaySync();

          },
          1000
        );

        scheduleNextAutoRefresh();

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
